Attribute VB_Name = "modUnserialise"
' Copyright (c) 2021-2026 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme

Option Explicit
Option Private Module
Option Base 1

#If VBA7 Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
#End If

' Reinterpret a Double as two 32-bit Longs (little-endian on Windows VBA)
Private Type TDouble
    d As Double
End Type

Private Type TLongs
    Lo As Long    ' low 32 bits
    Hi As Long    ' high 32 bits
End Type

' Reinterpret a Single as one 32-bit Long (little-endian on Windows VBA)
Private Type TSingle
    s As Single
End Type

Private Type TLong
    x As Long    ' all 32 bits
End Type

Private Type TBytes4
    B(0 To 3) As Byte
End Type

Private Type TBytes8
    B(0 To 7) As Byte
End Type

'Notes re round-tripping (Copilot assited)
'=========================================
'In Julia, string(x) for Float64 uses a shortest, round-trip algorithm
'(Ryu/Grisu class) that prints the minimal decimal digits that, when parsed
'back to a binary IEEE-754 double, reconstruct exactly the same 64-bit value.
'This ensures parse(Float64, string(x)) === x, for all Float64 values.

'VBA's CStr is not a round-trip formatter for IEEE-754 Double:
'* It typically emits ~15 significant digits, while a binary64 (Double) can
'  require 17 to guarantee an exact round-trip.
'* It obeys locale (decimal separator).
'* It may choose scientific vs. fixed forms inconsistently and trim trailing
'  zeros, none of which are guaranteed to be "shortest-round-trip".

'Data format used by Unserialise
'=============================================
'Format designed to be as fast as possible to unserialise.
'- Singleton types are prefixed with a type indicator character.
'- Dates are shown in their Excel representation as a number - faster to unserialise in VBA.
'- Floating point numbers (Double, Single) are represented in hexadecimal. See functions
'  DoubleToHex, HexToDouble, SingleToHex, HexToSingle. This ensures exact round-tripping
'  and avoids having to cope with the decimal separator being a comma.
'- Arrays are written with type indicator *, then three sections separated by semi-colons:
'  First section gives the number of dimensions (rank, up to 9) and the dimensions themselves,
'  comma delimited e.g. a 3 x 4 array would have a dimensions section "2,3,4".
'  Second section gives the lengths of the encodings of each element, comma delimited with a
'  terminating comma.
'  Third section gives the encodings, concatenated with no delimiter.
'- Note that arrays are written in column-major order.
'- Nested arrays (arrays containing arrays) are supported by the format, and by VBA but
'  cannot be returned to a worksheet.
'- Dictionaries are written with a type indicator H, then three sections separated by semi-colons:
'  First section gives the number of items in the dictionary
'  Second section gives the lengths of the encodings of the dictionary keys and items. The section
'  is comma-delimited with a terminating comma. The first element is the length of the encoding of
'  the first key, then the second item is the length of the encoding of the first item.
'  Third section gives the encodings of the dictionary keys and items, interleaved
'  first key, first item, second key second item etc.

'Type indicator characters are as follows:
' # Double, payload is hex e.g. 1.5 encoded as #3FF8000000000000
' Chr(163) (pound sterling sign) String
' T Boolean True
' F Boolean False
' D Date, payload is decimal of Excel's date representation. e.g. 22-Dec-2025 is D64013
' G DateTime, payload is hex
' E Empty
' N Null
' % Integer
' & Long
' B Byte (VBA's only unsigned type; Julia's UInt8)
' S Single, payload is hex
' C Currency - reserved, not currently implemented in Julia function encode_for_xl
' ! Error
' @ Decimal - reserved, not currently implemented in Julia function encode_for_xl
' * Array
' ^ LongLong (64-bit VBA only)
' H Dictionary
' V Array of Float64 only, no per-element type indicator or length (every element is always
'   exactly 16 hex characters). Julia -> Excel only - there is no VBA-side encoder, since
'   nothing currently sends a "V"-format string as an argument to Julia. Payload is big-endian
'   hex, decoded with the same HexToDouble used for scalar "#" - Julia produces this by
'   bulk-bswap-ing every element before hex-encoding the whole array in one operation (see
'   encode_for_xl(::Vector{Float64})/(::Matrix{Float64}) in src/encode.jl), rather than
'   reversing bytes per-element in VBA. (An earlier version of this format used little-endian
'   hex, decoded via a dedicated HexToDoubleLE, to avoid the bswap on the Julia side entirely -
'   but a direct measurement (VFormatDecodeSpeedTest, modPerformance.bas) showed
'   HexToDoubleLE's byte-by-byte reconstruction is enough slower than HexToDouble's that the
'   whole "V" format decoded slower than the general "*" format it was meant to replace. Moving
'   the bswap to Julia - cheap there, as a single bulk-broadcast intrinsic - let this format go
'   back to using the fast, already-tested HexToDouble unchanged.)
' R Range (UnitRange/StepRange/StepRangeLen/LinRange etc.) - encodes only first/step/length, not
'   every element, so the wire payload is a few dozen bytes regardless of how many elements the
'   range has (e.g. 1:1,000,000 encodes to ~15 bytes, vs ~16MB for "V"). VBA reconstructs each
'   element via plain arithmetic (first + (i-1)*step) - no per-element wire data at all. Julia ->
'   Excel only, like "V" originally was - VBA arrays are always fully materialized, so there's no
'   lazy "range" concept on the VBA side to compress this way.
'   Two sub-formats, given by the second character:
'   RI (Integer range, e.g. UnitRange{Int64}): "RI,<n>,<first>,<step>;" - first/step as plain
'     decimal (exact, matching the "^" LongLong convention), reconstructed via LongLong/Double
'     arithmetic (parseInt64, as for "^").
'   RF (Float64 range, e.g. StepRangeLen{Float64,...}): "RF,<n>;<hex first><hex step>" - first/step
'     as 16-character big-endian hex (matching scalar "#"), reconstructed via Double arithmetic
'     (HexToDouble, as for "#"/"V"). Verified (informally) to exactly reproduce Julia's own range
'     materialization, including StepRangeLen's twice-precision internal representation, for a
'     1,000,000-element case - see encode_for_xl(::AbstractRange{Float64})/(::AbstractRange{<:Integer})
'     in src/encode.jl for the encoder and the reasoning behind this.

'Examples (<pound> below stands for the single character Chr(163)):
'#3FF0000000000000 unserialises to Double 1
'&1 unserailises to Long 1
'<pound>Hello unserialises to String Hello
'T unserialises to Boolean True
'F unserialises to Boolean False
'*1,7;2,2,17,1,1,6,6,;%1%2#4008000000000000TF<pound>Hello<pound>World  unserialises to Array(1,2,3.0,True,False,"Hello","World")
'H2;2,3,4,5,;<pound>a%10<pound>abc%1000 unserialises to a Dictionary with two elements, element "a" contains 10 and element "abc" contains 1000

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : UnserialiseFromString
' Purpose    : Unserialise a result string returned directly from the Julia HTTP server.
' -----------------------------------------------------------------------------------------------------------------------
Function UnserialiseFromString(Contents As String, AllowNested As Boolean, StringLengthLimit As Long, JuliaVectorToXLColumn As Boolean)
1         On Error GoTo ErrHandler
2         Assign UnserialiseFromString, Unserialise(Contents, AllowNested, 0, StringLengthLimit, JuliaVectorToXLColumn)
3         Exit Function
ErrHandler:
4         ReThrow "UnserialiseFromString", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetStringLengthLimit
' Purpose    : Different versions of Excel have different limits for the longest string that can be an element of an
'              array passed from a VBA UDF back to Excel. I know the limit is 255 for Excel 2013 and earlier, and is
'              32,767 for Excel 365 (as of Sep 2021). But don't yet know the limit for Excel 2016 and 2019.
' Tried to get info from StackOverflow, without much joy:
' https://stackoverflow.com/questions/69303804/excel-versions-and-limits-on-the-length-of-string-elements-in-arrays-returned-by
' Note that this function returns 1 more than the maximum allowed string length, i.e. the minimum not-allowed string length.
' -----------------------------------------------------------------------------------------------------------------------
Function GetStringLengthLimit() As Long
          Static Res As Long
1         If Res = 0 Then
2             Select Case Val(Application.Version)
                  Case Is <= 15 'Excel 2010
3                     Res = 256
4                 Case Else
5                     Res = 32768 'Excel 2016, 2019, 365. Hopefully these versions (which all _
                                   return 16 as Application.Version) have the same limit.
6             End Select
7         End If
8         GetStringLengthLimit = Res
9     End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Unserialise
' Purpose    : Unserialises the contents of the results file saved by JuliaExcel julia code.
' -----------------------------------------------------------------------------------------------------------------------
Function Unserialise(Chars As String, AllowNesting As Boolean, ByRef Depth As Long, StringLengthLimit As Long, _
          JuliaVectorToXLColumn As Boolean)

1         On Error GoTo ErrHandler
2         Depth = Depth + 1
3         Select Case Asc(Left$(Chars, 1))
              Case 35    '# vbDouble
4                 Unserialise = HexToDouble(Mid$(Chars, 2))
5             Case 163    'Chr(163), pound sterling sign: vbString
6                 If StringLengthLimit > 0 Then 'Calling from worksheet formula, StringLengthLimit applies to elements of an array
7                     If Len(Chars) > IIf(Depth = 1, 32768, StringLengthLimit) Then 'Remember Chars includes an initial type indicator character of Chr(163)
8                         If StringLengthLimit = 32768 Then
9                             Throw "Data contains a string of length " & Format(Len(Chars) - 1, "###,###") & _
                                  ", too long to be returned to an Excel worksheet in Excel version " + _
                                  Application.Version() + ", for which the limit is 32,767"
10                        Else
11                            Throw "Data contains a string of length " & Format(Len(Chars) - 1, "###,###") & _
                                  ", too long to be returned to an Excel worksheet in Excel version " + _
                                  Application.Version() + ", for which the limit is " & _
                                  "32,767 for a string and " & Format(StringLengthLimit - 1, "###,###") + _
                                  " for string elements of an array"
12                        End If
13                    End If
14                End If
15                Unserialise = Mid$(Chars, 2)
16            Case 84     'T Boolean True
17                Unserialise = True
18            Case 68     'D vbDate from Date in Julia
19                Unserialise = CDate(Mid$(Chars, 2))
20            Case 70     'F Boolean False
21                Unserialise = False
22            Case 71     'G vbDate, from DateTime in Julia
23                Unserialise = CDate(HexToDouble(Mid$(Chars, 2)))
24            Case 69     'E vbEmpty
25                Unserialise = Empty
26            Case 78     'N vbNull
27                Unserialise = Null
28            Case 37     '% vbInteger
29                Unserialise = CInt(Mid$(Chars, 2))
30            Case 38     '& vbLong
31                Unserialise = CLng(Mid$(Chars, 2))
32            Case 66     'B vbByte
33                Unserialise = CByte(Mid$(Chars, 2))
34            Case 94     '^ vbLongLong
35                Unserialise = parseInt64(Mid$(Chars, 2))
36            Case 83     'S vbSingle
37                Unserialise = HexToSingle(Mid$(Chars, 2))
38            Case 67     'C vbCurrency, not currently implemented in Julia function encode_for_xl
39                Unserialise = CCur(Mid$(Chars, 2))
40            Case 33     '! vbError
41                Unserialise = CVErr(Mid$(Chars, 2))
42            Case 64     '@ vbDecimal, not currently implemented in Julia function encode_for_xl
43                Unserialise = CDec(Mid$(Chars, 2))
                  
44            Case 42     '* vbArray
45                If Depth > 1 Then If Not AllowNesting Then Throw "Excel cannot display arrays containing arrays"

                  Dim Ret() As Variant
                  Dim p1 As Long    ' Position of first ';'
                  Dim p2 As Long    ' Position of second ';'
                  Dim m As Long     ' Pointer into lengths section
                  Dim m2 As Long
                  Dim k As Long     ' Pointer into payload section
                  Dim ThisLength As Long

46                p1 = InStr(Chars, ";")
47                p2 = InStr(p1 + 1, Chars, ";")
48                m = p1 + 1
49                k = p2 + 1

                  ' Rank is the single character after '*', e.g. "*2,3,4;..."
                  Dim Rank As Long
                  'but check that the number of dimensions has only 1 digit!
50                If Mid$(Chars, 3, 1) <> "," Then Throw "Cannot unserialise arrays with " & _
                      Mid$(Chars, 2, InStr(Chars, ",") - 2) & " dimensions (max supported: 9)"
51                Rank = CInt(Mid$(Chars, 2, 1))

52                Select Case Rank
                      Case 1
                          Dim i As Long
                          Dim n As Long
53                        n = CLng(Mid$(Chars, 4, p1 - 4))
54                        If n = 0 Then
55                            If Not AllowNesting Then Throw "Excel cannot display arrays with zero elements"
56                            Unserialise = VBA.Split(vbNullString)
57                        Else
58                            If JuliaVectorToXLColumn Then
59                                ReDim Ret(1 To n, 1 To 1)
60                                For i = 1 To n
61                                    m2 = InStr(m, Chars, ",") + 1
62                                    ThisLength = CLng(Mid$(Chars, m, m2 - m - 1))
63                                    Assign Ret(i, 1), Unserialise(Mid$(Chars, k, ThisLength), AllowNesting, Depth, StringLengthLimit, JuliaVectorToXLColumn)
64                                    k = k + ThisLength
65                                    m = m2
66                                Next i
67                            Else
68                                ReDim Ret(1 To n)
69                                For i = 1 To n
70                                    m2 = InStr(m, Chars, ",") + 1
71                                    ThisLength = CLng(Mid$(Chars, m, m2 - m - 1))
72                                    Assign Ret(i), Unserialise(Mid$(Chars, k, ThisLength), AllowNesting, Depth, StringLengthLimit, JuliaVectorToXLColumn)
73                                    k = k + ThisLength
74                                    m = m2
75                                Next i
76                            End If
77                            Unserialise = Ret
78                        End If

79                    Case 2
                          Dim CommaPos As Long
                          Dim j As Long
                          Dim NC As Long
                          Dim NR As Long
80                        CommaPos = InStr(4, Chars, ",")
81                        NR = CLng(Mid$(Chars, 4, CommaPos - 4))
82                        NC = CLng(Mid$(Chars, CommaPos + 1, p1 - CommaPos - 1))
83                        If NR = 0 Or NC = 0 Then Throw "Cannot create array of size zero"
84                        ReDim Ret(1 To NR, 1 To NC)
85                        For j = 1 To NC
86                            For i = 1 To NR
87                                m2 = InStr(m, Chars, ",") + 1
88                                ThisLength = CLng(Mid$(Chars, m, m2 - m - 1))
89                                Assign Ret(i, j), Unserialise(Mid$(Chars, k, ThisLength), AllowNesting, Depth, StringLengthLimit, JuliaVectorToXLColumn)
90                                k = k + ThisLength
91                                m = m2
92                            Next i
93                        Next j
94                        Unserialise = Ret

95                    Case Else
                          ' === Section to handle >=3 dimensional arrays written by Copilot 23 Dec 2025
                          Dim Dims() As Long
96                        Dims = ParseDims(Mid$(Chars, 4, p1 - 4), Rank)  ' section between "*,<rank>," and first ';'

                          ' Guard: Excel cannot display >2-D arrays; allow only when nesting is permitted i.e. when unserialising to VBA variable
97                        If Not AllowNesting Then
98                            Throw "Excel cannot display arrays with more than 2 dimensions"
99                        End If

                          ' None of the dims may be zero
                          Dim q As Long
                          Dim Total As Long
100                        Total = 1
101                        For q = 1 To Rank
102                           If Dims(q) <= 0 Then Throw "Cannot create array of size zero"
103                           Total = Total * Dims(q)
104                       Next q

                          ' Allocate Ret() to the requested rank (up to MAX_RANK supported)
105                       ReDimVariantArray Ret, Dims

                          ' Walk in column-major order (dim 1 fastest), assigning elements
                          Dim Idx() As Long
106                       ReDim Idx(1 To Rank)
107                       For q = 1 To Rank: Idx(q) = 1: Next q

                          Dim Count As Long
                          Dim Val As Variant
108                       For Count = 1 To Total
109                           m2 = InStr(m, Chars, ",") + 1
110                           ThisLength = CLng(Mid$(Chars, m, m2 - m - 1))
111                           Assign Val, Unserialise(Mid$(Chars, k, ThisLength), AllowNesting, Depth, StringLengthLimit, JuliaVectorToXLColumn)
112                           AssignByRank Ret, Idx, Val  ' Assign Ret(i1, i2, ..., irank) = val

113                           k = k + ThisLength
114                           m = m2

                              ' Increment indices: dim 1 fastest
115                           q = 1
116                           Do While q <= Rank
117                               Idx(q) = Idx(q) + 1
118                               If Idx(q) <= Dims(q) Then Exit Do
119                               Idx(q) = 1
120                               q = q + 1
121                           Loop
122                           If q > Rank Then Exit For
123                       Next Count

124                       Unserialise = Ret
125               End Select
126           Case 72 'H Dictionary
127               If Not AllowNesting Then Throw "Excel cannot display variables of type Dictionary"
128               p1 = InStr(Chars, ";")
129               p2 = InStr(p1 + 1, Chars, ";")
130               m = p1 + 1 '"pointer" to read from lengths section. Points to the first character after each comma.
131               k = p2 + 1 '"pointer" to read from contents section. Points to the first character of each "chunk".
                  Dim DictRet As New Scripting.Dictionary
                  Dim KeyLength As Long
                  Dim m3 As Long
                  Dim ThisKey As Variant
                  Dim ThisValue As Variant
                  Dim ValueLength As Long
132               n = Mid$(Chars, 2, p1 - 2) 'Num elements in dictionary
133               For i = 1 To n
134                   m2 = InStr(m, Chars, ",") + 1
135                   m3 = InStr(m2, Chars, ",") + 1
136                   KeyLength = Mid$(Chars, m, m2 - m - 1)
137                   ValueLength = Mid$(Chars, m2, m3 - m2 - 1)
138                   Assign ThisKey, Unserialise(Mid$(Chars, k, KeyLength), AllowNesting, Depth, StringLengthLimit, JuliaVectorToXLColumn)
139                   k = k + KeyLength
140                   Assign ThisValue, Unserialise(Mid$(Chars, k, ValueLength), AllowNesting, Depth, StringLengthLimit, JuliaVectorToXLColumn)
141                   k = k + ValueLength
142                   m = m3

143                   If VarType(ThisKey) = vbLongLong Then 'LongLong not allowed as key?
144                       DictRet.Add CLng(ThisKey), ThisValue
145                   Else
146                       DictRet.Add ThisKey, ThisValue
147                   End If
148               Next i
149               Set Unserialise = DictRet
150           Case 86 'V vbArray of Float64 only, no per-element type indicator or length - the
                  '"V" indicator itself is Julia's own guarantee, via multiple dispatch, that every
                  'element really is a Float64, so unlike the general "*" array case above, nothing
                  'here needs to defend against that not being true.
                  Const VBytesPerElement As Long = 16 'BE hex (as scalar "#"), no type-indicator character per element
                  Dim Buf() As Double
151               p1 = InStr(Chars, ";")
152               If Mid$(Chars, 3, 1) <> "," Then Throw "Cannot unserialise 'V'-format arrays with " & _
                      Mid$(Chars, 2, InStr(Chars, ",") - 2) & " dimensions (max supported: 9)"
153               Rank = CInt(Mid$(Chars, 2, 1))
154               k = p1 + 1

155               Select Case Rank
                      Case 1
156                       n = CLng(Mid$(Chars, 4, p1 - 4))
157                       If Len(Chars) - k + 1 <> n * VBytesPerElement Then Throw _
                              "'V'-format string has the wrong number of hex characters for a 1-D array of " & n & " element(s)"
158                       Buf = BulkDoublesFromHex(Chars, k, n)
159                       If JuliaVectorToXLColumn Then
160                           ReDim Ret(1 To n, 1 To 1)
161                           For i = 1 To n
162                               Ret(i, 1) = Buf(i)
163                           Next i
164                       Else
165                           ReDim Ret(1 To n)
166                           For i = 1 To n
167                               Ret(i) = Buf(i)
168                           Next i
169                       End If

170                   Case 2
171                       CommaPos = InStr(4, Chars, ",")
172                       NR = CLng(Mid$(Chars, 4, CommaPos - 4))
173                       NC = CLng(Mid$(Chars, CommaPos + 1, p1 - CommaPos - 1))
174                       If Len(Chars) - k + 1 <> NR * NC * VBytesPerElement Then Throw _
                              "'V'-format string has the wrong number of hex characters for a " & NR & "x" & NC & " array"
175                       Buf = BulkDoublesFromHex(Chars, k, NR * NC)
176                       ReDim Ret(1 To NR, 1 To NC)
177                       For j = 1 To NC
178                           For i = 1 To NR
179                               Ret(i, j) = Buf((j - 1) * NR + i)
180                           Next i
181                       Next j

                      Case Else
                          ' Rank 3-9, reusing the same ParseDims/ReDimVariantArray/AssignByRank
                          ' helpers, and the same column-major index-walking scheme, as the general
                          ' "*" array format's own >=2-dimensional handling above - just without any
                          ' per-element length lookup, since every element is always exactly
                          ' VBytesPerElement hex characters.
182                       Dims = ParseDims(Mid$(Chars, 4, p1 - 4), Rank)
183                       If Not AllowNesting Then Throw "Excel cannot display arrays with more than 2 dimensions"
184                       Total = 1
185                       For q = 1 To Rank
186                           If Dims(q) <= 0 Then Throw "Cannot create array of size zero"
187                           Total = Total * Dims(q)
188                       Next q
189                       If Len(Chars) - k + 1 <> Total * VBytesPerElement Then Throw _
                              "'V'-format string has the wrong number of hex characters for a " & Rank & "-D array"
190                       Buf = BulkDoublesFromHex(Chars, k, Total)
191                       ReDimVariantArray Ret, Dims
192                       ReDim Idx(1 To Rank)
193                       For q = 1 To Rank: Idx(q) = 1: Next q
194                       For Count = 1 To Total
195                           AssignByRank Ret, Idx, Buf(Count)
196                           q = 1
197                           Do While q <= Rank
198                               Idx(q) = Idx(q) + 1
199                               If Idx(q) <= Dims(q) Then Exit Do
200                               Idx(q) = 1
201                               q = q + 1
202                           Loop
203                           If q > Rank Then Exit For
204                       Next Count
205               End Select
206               Unserialise = Ret

              Case 82 'R Range (UnitRange/StepRange/StepRangeLen/LinRange etc.) - reconstructed via
                  'arithmetic (first + (i-1)*step), no per-element wire data at all; see
                  'encode_for_xl(::AbstractRange{Float64})/(::AbstractRange{<:Integer}) in
                  'src/encode.jl. Julia -> Excel only - VBA has no lazy "range" concept to send
                  'back to Julia this way.
                  Dim HeaderParts() As String
                  Dim RFirst As Variant
                  Dim RStep As Variant
207               p1 = InStr(Chars, ";")
208               If Mid$(Chars, 2, 1) = "I" Then
209                   HeaderParts = Split(Mid$(Chars, 4, p1 - 4), ",")
210                   n = CLng(HeaderParts(0))
211                   RFirst = parseInt64(HeaderParts(1))
212                   RStep = parseInt64(HeaderParts(2))
213               ElseIf Mid$(Chars, 2, 1) = "F" Then
214                   n = CLng(Mid$(Chars, 4, p1 - 4))
215                   RFirst = HexToDouble(Mid$(Chars, p1 + 1, 16))
216                   RStep = HexToDouble(Mid$(Chars, p1 + 17, 16))
217               Else
218                   Throw "Character '" & Mid$(Chars, 2, 1) & "' is not a recognised 'R'-format sub-type (expected 'I' or 'F')"
219               End If
220               If JuliaVectorToXLColumn Then
221                   ReDim Ret(1 To n, 1 To 1)
222                   For i = 1 To n
223                       Ret(i, 1) = RFirst + (i - 1) * RStep
224                   Next i
225               Else
226                   ReDim Ret(1 To n)
227                   For i = 1 To n
228                       Ret(i) = RFirst + (i - 1) * RStep
229                   Next i
230               End If
231               Unserialise = Ret

232           Case Else
233               Throw "Character '" & Left$(Chars, 1) & "' is not recognised as a type identifier"
234       End Select

235       Exit Function
ErrHandler:
236       ReThrow "Unserialise", Err
End Function

'Values of type Int64 in Julia must be handled differently on Excel 32-bit and Excel 64bit
#If Win64 Then
      Function parseInt64(x As String)
1         parseInt64 = CLngLng(x)
      End Function
#Else
      Function parseInt64(x As String)
1         parseInt64 = CDbl(x)
      End Function
#End If

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : DoubleToHex
' Purpose    : Encode x as a 16-character hex string of its IEEE-754 bit pattern (big-endian), for
'              inclusion in the wire format passed between Excel and Julia. Uses LSet to reinterpret
'              the 8 bytes of x as a Byte array, then maps each byte through a static 256-entry
'              lookup table (seeded on first call) to produce the two-character hex pair.
' Note       : Nothing to do with Excel's own worksheet function HEX(), which converts a decimal
'              INTEGER to its hex-digit representation of that same numeric VALUE (e.g. HEX(255) is
'              "FF"). DoubleToHex instead reinterprets the IEEE-754 BIT PATTERN of a Double - a
'              completely different operation that happens to share a name with something familiar.
' -----------------------------------------------------------------------------------------------------------------------
Function DoubleToHex(ByVal x As Double) As String
          Static HexByte(0 To 255) As String
          Static Initialized As Boolean
          Dim i As Long
          Dim TB As TBytes8
          Dim TD As TDouble

1         If Not Initialized Then
2             For i = 0 To 255
3                 HexByte(i) = Right$("0" & Hex$(i), 2)
4             Next i
5             Initialized = True
6         End If

7         TD.d = x
8         LSet TB = TD   ' reinterpret the 8 bytes of the Double as a Byte array (little-endian)

9         DoubleToHex = HexByte(TB.B(7)) & HexByte(TB.B(6)) & HexByte(TB.B(5)) & HexByte(TB.B(4)) & _
              HexByte(TB.B(3)) & HexByte(TB.B(2)) & HexByte(TB.B(1)) & HexByte(TB.B(0))
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : HexToDouble
' Purpose    : Parse a 16-character hex string (uppercase or lowercase) as the IEEE-754
'              bit pattern of a Double and return the corresponding Double.
' Note       : Nothing to do with Excel's own worksheet function HEX() - see DoubleToHex's note above.
' -----------------------------------------------------------------------------------------------------------------------
Function HexToDouble(ByVal Hex As String) As Double

          Dim Hi As Long
          Dim Lo As Long
          Dim TD As TDouble
          Dim Tl As TLongs

1         On Error GoTo ErrHandler
2         If Len(Hex) <> 16 Then Throw "Hex must be 16 hex characters, but got " & Len(Hex)
3         Hi = CLng("&H" & Left$(Hex, 8))
4         Lo = CLng("&H" & Right$(Hex, 8))
5         Tl.Hi = Hi
6         Tl.Lo = Lo
7         LSet TD = Tl
8         HexToDouble = TD.d

9         Exit Function
ErrHandler:
10        ReThrow "HexToDouble", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : BulkDoublesFromHex
' Purpose    : Parses n consecutive 16-hex-character elements starting at character position
'              StartPos of Chars (the "V" format's payload, no delimiters between elements) into a
'              genuinely-typed Double() array, in one bulk operation - the same bit-for-bit decoding
'              HexToDouble produces per element, but without calling it (or LSet) once per element.
'              Each element's high/low 32-bit halves are parsed the same way HexToDouble does
'              (CLng("&H" & <8 hex chars>) - already a reasonably efficient parse, 4 bytes per call)
'              directly into a Long() buffer, then one RtlMoveMemory ("CopyMemory") call
'              reinterprets that whole buffer's bytes as the result Double() array - replacing N
'              per-element LSets and N per-element HexToDouble function calls (each of which also
'              does its own Left$/Right$ substring slicing, on top of the Mid$ slicing done here)
'              with one bulk memory copy. Measured (formerly modHexBulkPrototype.bas, a throwaway
'              prototype now deleted once its findings were wired in here - see git history) at
'              roughly 64% faster than calling HexToDouble per element, for a 100,000-element
'              array - a bigger
'              saving than the encode side's equivalent (BulkHexOfDoubleArray, modSerialise.bas),
'              since HexToDouble's per-element overhead includes that extra string-slice on top of
'              the function call and the LSet.
'              A Double's low 32 bits sit first in memory (little-endian), so the Long() buffer
'              stores each element's low half then high half, in that order, to match.
'              The result is a genuinely-typed Double() array, not the Variant() array Ret is
'              declared as elsewhere in this function - assigning a whole Double() array to a
'              Variant()-declared variable in one shot is a compile error in VBA ("Can't assign to
'              array"), so callers still copy this result into Ret element-by-element; what's saved
'              is the hex-parsing work per element, not that final copy.
'              Safety: CopyMemory is a raw memory copy - unlike an ordinary VBA error, a wrong length
'              argument here could corrupt memory or crash Excel outright, not just fail this call.
'              So immediately before calling it, both buffers' actual byte sizes are re-derived
'              independently from their own LBound/UBound (not by trusting the "n" arithmetic that
'              sized them) and compared; any mismatch throws a normal, catchable error instead of
'              proceeding. This is deliberately a hard failure, not a silent fallback to the old
'              per-element HexToDouble loop: if this ever fires, something is genuinely wrong with
'              this function's own logic, and a fallback path that (if that logic is correct) never
'              executes in practice would itself be an untested, silently bit-rotting liability.
' -----------------------------------------------------------------------------------------------------------------------
Private Function BulkDoublesFromHex(ByVal Chars As String, ByVal StartPos As Long, ByVal n As Long) As Double()
          Dim base As Long
          Dim i As Long
          Dim Raw() As Long
          Dim RawBytes As Long
          Dim Result() As Double
          Dim ResultBytes As Long

1         If n <= 0 Then Throw "BulkDoublesFromHex requires n > 0"
2         ReDim Raw(1 To n * 2)
3         For i = 1 To n
4             base = StartPos + (i - 1) * 16
5             Raw(2 * i - 1) = CLng("&H" & Mid$(Chars, base + 8, 8))   ' low 32 bits (last 8 hex chars)
6             Raw(2 * i) = CLng("&H" & Mid$(Chars, base, 8))           ' high 32 bits (first 8 hex chars)
7         Next i

8         ReDim Result(1 To n)
9         ResultBytes = (UBound(Result) - LBound(Result) + 1) * 8
10        RawBytes = (UBound(Raw) - LBound(Raw) + 1) * 4
11        If ResultBytes <> RawBytes Then Throw "BulkDoublesFromHex: Result is " & ResultBytes & _
              " bytes but Raw buffer is " & RawBytes & " bytes - refusing to call CopyMemory"
12        CopyMemory Result(1), Raw(1), ResultBytes

13        BulkDoublesFromHex = Result
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SingleToHex
' Purpose    : Encode x as an 8-character hex string of its IEEE-754 bit pattern (big-endian), for
'              inclusion in the wire format passed between Excel and Julia. Uses LSet to reinterpret
'              the 4 bytes of x as a Byte array, then maps each byte through a static 256-entry
'              lookup table (seeded on first call) to produce the two-character hex pair.
' Note       : Nothing to do with Excel's own worksheet function HEX() - see DoubleToHex's note above.
' -----------------------------------------------------------------------------------------------------------------------
Function SingleToHex(ByVal x As Single) As String
          Static HexByte(0 To 255) As String
          Static Initialized As Boolean
          Dim i As Long
          Dim TB As TBytes4
          Dim TS As TSingle

1         If Not Initialized Then
2             For i = 0 To 255
3                 HexByte(i) = Right$("0" & Hex$(i), 2)
4             Next i
5             Initialized = True
6         End If

7         TS.s = x
8         LSet TB = TS   ' reinterpret the 4 bytes of the Single as a Byte array (little-endian)

9         SingleToHex = HexByte(TB.B(3)) & HexByte(TB.B(2)) & HexByte(TB.B(1)) & HexByte(TB.B(0))
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : HexToSingle
' Purpose    : Parse an 8-character hex string (uppercase or lowercase) as the IEEE-754
'              bit pattern of a Single and return the corresponding Single.
' Note       : Nothing to do with Excel's own worksheet function HEX() - see DoubleToHex's note above.
' -----------------------------------------------------------------------------------------------------------------------
Function HexToSingle(ByVal Hex As String) As Single

          Dim Tl As TLong
          Dim TS As TSingle
          Dim Wx As Long

1         On Error GoTo ErrHandler
2         If Len(Hex) <> 8 Then Throw "Hex must be 8 hex characters, but got " & Len(Hex)
3         Wx = CLng("&H" & Hex)
4         Tl.x = Wx
5         LSet TS = Tl
6         HexToSingle = TS.s

7         Exit Function
ErrHandler:
8         ReThrow "HexToSingle", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ParseDims
' Purpose    : Parse a comma-delimited list of dimension sizes (e.g., "3,4,5") into dims(1..Rank).
' -----------------------------------------------------------------------------------------------------------------------
Private Function ParseDims(ByVal s As String, ByVal Rank As Long) As Long()
          Dim Parts() As String
1         On Error GoTo ErrHandler
2         Parts = Split(s, ",")
3         If UBound(Parts) + 1 <> Rank Then
4             Throw "Malformed array header: expected " & Rank & " dimensions, found " & (UBound(Parts) + 1)
5         End If
          Dim Dims() As Long
          Dim i As Long
6         ReDim Dims(1 To Rank)
7         For i = 1 To Rank
8             Dims(i) = CLng(Parts(i - 1))
9         Next i
10        ParseDims = Dims

11        Exit Function
ErrHandler:
12        ReThrow "ParseDims", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ReDimVariantArray
' Purpose    : ReDim Ret() to the specified dims (1..rank). Increase MAX_RANK if needed.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub ReDimVariantArray(ByRef Ret() As Variant, ByRef Dims() As Long)
          Const MAX_RANK As Long = 9
          Dim r As Long
1         On Error GoTo ErrHandler
2         r = UBound(Dims)
3         If r < 1 Or r > MAX_RANK Then
4             Throw "Cannot unserialise arrays with " & r & " dimensions (max supported: " & MAX_RANK & ")"
5         End If

6         Select Case r
              Case 1: ReDim Ret(1 To Dims(1))
7             Case 2: ReDim Ret(1 To Dims(1), 1 To Dims(2))
8             Case 3: ReDim Ret(1 To Dims(1), 1 To Dims(2), 1 To Dims(3))
9             Case 4: ReDim Ret(1 To Dims(1), 1 To Dims(2), 1 To Dims(3), 1 To Dims(4))
10            Case 5: ReDim Ret(1 To Dims(1), 1 To Dims(2), 1 To Dims(3), 1 To Dims(4), 1 To Dims(5))
11            Case 6: ReDim Ret(1 To Dims(1), 1 To Dims(2), 1 To Dims(3), 1 To Dims(4), 1 To Dims(5), 1 To Dims(6))
12            Case 7: ReDim Ret(1 To Dims(1), 1 To Dims(2), 1 To Dims(3), 1 To Dims(4), 1 To Dims(5), 1 To Dims(6), 1 To Dims(7))
13            Case 8: ReDim Ret(1 To Dims(1), 1 To Dims(2), 1 To Dims(3), 1 To Dims(4), 1 To Dims(5), 1 To Dims(6), 1 To Dims(7), 1 To Dims(8))
14            Case 9: ReDim Ret(1 To Dims(1), 1 To Dims(2), 1 To Dims(3), 1 To Dims(4), 1 To Dims(5), 1 To Dims(6), 1 To Dims(7), 1 To Dims(8), 1 To Dims(9))
15        End Select

16        Exit Sub
ErrHandler:
17        ReThrow "ReDimVariantArray", Err
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : AssignByRank
' Purpose    : Assign Ret(i1, i2, ..., irank) = Val, where idx(1..r) holds indices.
' -----------------------------------------------------------------------------------------------------------------------
Private Sub AssignByRank(ByRef Ret() As Variant, ByRef Idx() As Long, ByRef Val As Variant)
1         Select Case UBound(Idx)
              Case 1: Assign Ret(Idx(1)), Val
2             Case 2: Assign Ret(Idx(1), Idx(2)), Val
3             Case 3: Assign Ret(Idx(1), Idx(2), Idx(3)), Val
4             Case 4: Assign Ret(Idx(1), Idx(2), Idx(3), Idx(4)), Val
5             Case 5: Assign Ret(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5)), Val
6             Case 6: Assign Ret(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5), Idx(6)), Val
7             Case 7: Assign Ret(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5), Idx(6), Idx(7)), Val
8             Case 8: Assign Ret(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5), Idx(6), Idx(7), Idx(8)), Val
9             Case 9: Assign Ret(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5), Idx(6), Idx(7), Idx(8), Idx(9)), Val
10            Case Else
11                Throw "Rank > 8 not supported by AssignByRank"
12        End Select
End Sub

