Attribute VB_Name = "modUnserialise"
' Copyright (c) 2021-2026 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme

Option Explicit
Option Private Module
Option Base 1
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
' S Single, payload is hex
' C Currency - reserved, not currently implemented in Julia function encode_for_xl
' ! Error
' @ Decimal - reserved, not currently implemented in Julia function encode_for_xl
' * Array
' ^ LongLong (64-bit VBA only)
' H Dictionary

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
32            Case 94     '^ vbLongLong
33                Unserialise = parseInt64(Mid$(Chars, 2))
34            Case 83     'S vbSingle
35                Unserialise = HexToSingle(Mid$(Chars, 2))
36            Case 67     'C vbCurrency, not currently implemented in Julia function encode_for_xl
37                Unserialise = CCur(Mid$(Chars, 2))
38            Case 33     '! vbError
39                Unserialise = CVErr(Mid$(Chars, 2))
40            Case 64     '@ vbDecimal, not currently implemented in Julia function encode_for_xl
41                Unserialise = CDec(Mid$(Chars, 2))
                  
42            Case 42     '* vbArray
43                If Depth > 1 Then If Not AllowNesting Then Throw "Excel cannot display arrays containing arrays"

                  Dim Ret() As Variant
                  Dim p1 As Long    ' Position of first ';'
                  Dim p2 As Long    ' Position of second ';'
                  Dim m As Long     ' Pointer into lengths section
                  Dim m2 As Long
                  Dim k As Long     ' Pointer into payload section
                  Dim ThisLength As Long

44                p1 = InStr(Chars, ";")
45                p2 = InStr(p1 + 1, Chars, ";")
46                m = p1 + 1
47                k = p2 + 1

                  ' Rank is the single character after '*', e.g. "*2,3,4;..."
                  Dim Rank As Long
                  'but check that the number of dimensions has only 1 digit!
48                If Mid$(Chars, 3, 1) <> "," Then Throw "Cannot unserialise arrays with " & _
                      Mid$(Chars, 2, InStr(Chars, ",") - 2) & " dimensions (max supported: 8)"
49                Rank = CInt(Mid$(Chars, 2, 1))

50                Select Case Rank
                      Case 1
                          Dim i As Long
                          Dim n As Long
51                        n = CLng(Mid$(Chars, 4, p1 - 4))
52                        If n = 0 Then
53                            If Not AllowNesting Then Throw "Excel cannot display arrays with zero elements"
54                            Unserialise = VBA.Split(vbNullString)
55                        Else
56                            If JuliaVectorToXLColumn Then
57                                ReDim Ret(1 To n, 1 To 1)
58                                For i = 1 To n
59                                    m2 = InStr(m, Chars, ",") + 1
60                                    ThisLength = CLng(Mid$(Chars, m, m2 - m - 1))
61                                    Assign Ret(i, 1), Unserialise(Mid$(Chars, k, ThisLength), AllowNesting, Depth, StringLengthLimit, JuliaVectorToXLColumn)
62                                    k = k + ThisLength
63                                    m = m2
64                                Next i
65                            Else
66                                ReDim Ret(1 To n)
67                                For i = 1 To n
68                                    m2 = InStr(m, Chars, ",") + 1
69                                    ThisLength = CLng(Mid$(Chars, m, m2 - m - 1))
70                                    Assign Ret(i), Unserialise(Mid$(Chars, k, ThisLength), AllowNesting, Depth, StringLengthLimit, JuliaVectorToXLColumn)
71                                    k = k + ThisLength
72                                    m = m2
73                                Next i
74                            End If
75                            Unserialise = Ret
76                        End If

77                    Case 2
                          Dim CommaPos As Long
                          Dim j As Long
                          Dim NC As Long
                          Dim NR As Long
78                        CommaPos = InStr(4, Chars, ",")
79                        NR = CLng(Mid$(Chars, 4, CommaPos - 4))
80                        NC = CLng(Mid$(Chars, CommaPos + 1, p1 - CommaPos - 1))
81                        If NR = 0 Or NC = 0 Then Throw "Cannot create array of size zero"
82                        ReDim Ret(1 To NR, 1 To NC)
83                        For j = 1 To NC
84                            For i = 1 To NR
85                                m2 = InStr(m, Chars, ",") + 1
86                                ThisLength = CLng(Mid$(Chars, m, m2 - m - 1))
87                                Assign Ret(i, j), Unserialise(Mid$(Chars, k, ThisLength), AllowNesting, Depth, StringLengthLimit, JuliaVectorToXLColumn)
88                                k = k + ThisLength
89                                m = m2
90                            Next i
91                        Next j
92                        Unserialise = Ret

93                    Case Else
                          ' === Section to handle >=3 dimensional arrays written by Copilot 23 Dec 2025
                          Dim Dims() As Long
94                        Dims = ParseDims(Mid$(Chars, 4, p1 - 4), Rank)  ' section between "*,<rank>," and first ';'

                          ' Guard: Excel cannot display >2-D arrays; allow only when nesting is permitted i.e. when unserialising to VBA variable
95                        If Not AllowNesting Then
96                            Throw "Excel cannot display arrays with more than 2 dimensions"
97                        End If

                          ' None of the dims may be zero
                          Dim q As Long
                          Dim Total As Long
98                        Total = 1
99                        For q = 1 To Rank
100                           If Dims(q) <= 0 Then Throw "Cannot create array of size zero"
101                           Total = Total * Dims(q)
102                       Next q

                          ' Allocate Ret() to the requested rank (up to MAX_RANK supported)
103                       ReDimVariantArray Ret, Dims

                          ' Walk in column-major order (dim 1 fastest), assigning elements
                          Dim Idx() As Long
104                       ReDim Idx(1 To Rank)
105                       For q = 1 To Rank: Idx(q) = 1: Next q

                          Dim Count As Long
                          Dim Val As Variant
106                       For Count = 1 To Total
107                           m2 = InStr(m, Chars, ",") + 1
108                           ThisLength = CLng(Mid$(Chars, m, m2 - m - 1))
109                           Assign Val, Unserialise(Mid$(Chars, k, ThisLength), AllowNesting, Depth, StringLengthLimit, JuliaVectorToXLColumn)
110                           AssignByRank Ret, Idx, Val  ' Assign Ret(i1, i2, ..., irank) = val

111                           k = k + ThisLength
112                           m = m2

                              ' Increment indices: dim 1 fastest
113                           q = 1
114                           Do While q <= Rank
115                               Idx(q) = Idx(q) + 1
116                               If Idx(q) <= Dims(q) Then Exit Do
117                               Idx(q) = 1
118                               q = q + 1
119                           Loop
120                           If q > Rank Then Exit For
121                       Next Count

122                       Unserialise = Ret
123               End Select
124           Case 72 'H Dictionary
125               If Not AllowNesting Then Throw "Excel cannot display variables of type Dictionary"
126               p1 = InStr(Chars, ";")
127               p2 = InStr(p1 + 1, Chars, ";")
128               m = p1 + 1 '"pointer" to read from lengths section. Points to the first character after each comma.
129               k = p2 + 1 '"pointer" to read from contents section. Points to the first character of each "chunk".
                  Dim DictRet As New Scripting.Dictionary
                  Dim KeyLength As Long
                  Dim m3 As Long
                  Dim ThisKey As Variant
                  Dim ThisValue As Variant
                  Dim ValueLength As Long
130               n = Mid$(Chars, 2, p1 - 2) 'Num elements in dictionary
131               For i = 1 To n
132                   m2 = InStr(m, Chars, ",") + 1
133                   m3 = InStr(m2, Chars, ",") + 1
134                   KeyLength = Mid$(Chars, m, m2 - m - 1)
135                   ValueLength = Mid$(Chars, m2, m3 - m2 - 1)
136                   Assign ThisKey, Unserialise(Mid$(Chars, k, KeyLength), AllowNesting, Depth, StringLengthLimit, JuliaVectorToXLColumn)
137                   k = k + KeyLength
138                   Assign ThisValue, Unserialise(Mid$(Chars, k, ValueLength), AllowNesting, Depth, StringLengthLimit, JuliaVectorToXLColumn)
139                   k = k + ValueLength
140                   m = m3

141                   If VarType(ThisKey) = vbLongLong Then 'LongLong not allowed as key?
142                       DictRet.Add CLng(ThisKey), ThisValue
143                   Else
144                       DictRet.Add ThisKey, ThisValue
145                   End If
146               Next i
147               Set Unserialise = DictRet
148           Case Else
149               Throw "Character '" & Left$(Chars, 1) & "' is not recognised as a type identifier"
150       End Select

151       Exit Function
ErrHandler:
152       ReThrow "Unserialise", Err
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
' Procedure  : SingleToHex
' Purpose    : Encode x as an 8-character hex string of its IEEE-754 bit pattern (big-endian), for
'              inclusion in the wire format passed between Excel and Julia. Uses LSet to reinterpret
'              the 4 bytes of x as a Byte array, then maps each byte through a static 256-entry
'              lookup table (seeded on first call) to produce the two-character hex pair.
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
' Procedure  : LPad
' Purpose    : Pad s on the left with p to make it n characters long. If s is already n characters long, an equal string
'              is returned.
' -----------------------------------------------------------------------------------------------------------------------
Function LPad(s As String, n As Long, p As String)
1         If Len(s) < n Then
2             LPad = String(n - Len(s), p) & s
3         Else
4             LPad = s
5         End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : HexToSingle
' Purpose    : Parse an 8-character hex string (uppercase or lowercase) as the IEEE-754
'              bit pattern of a Single and return the corresponding Single.
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

