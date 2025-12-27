Attribute VB_Name = "modSerialise"
' Copyright (c) 2021-2025 Philip Swannell
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
'- Floating point numbers (Double, Single) are represented in hexadecimal. See functions _
 DoubleToHex, HexToDouble, SingleToHex, HexToSingle. This ensures exact round-tripping _
 and avoids having to cope with the decimal separator being a comma.
'- Arrays are written with type indicator *, then three sections separated by semi-colons:
'  First section gives the number of dimensions and the dimensions themselves, comma
'  delimited e.g. a 3 x 4 array would have a dimensions section "2,3,4".
'  Second section gives the lengths of the encodings of each element, comma delimited with a
'  terminating comma.
'  Third section gives the encodings, concatenated with no delimiter.
'- Note that arrays are written in column-major order.
'- Nested arrays (arrays containing arrays) are supported by the format, and by VBA but
'  cannot be returned to a worksheet.
'- Dictionaries are written with a type indicator ^, then three sections separated by semi-colons:
'  First section gives the number of items in the dictionary
'  Second section gives the lengths of the encodings of the dictionary keys and items. The section
'  is comma-delimited with a terminating comma. The first element is the length of the encoding of
'  the first key, then the second item is the length of the encoding of the first item.
'  Third section gives the encodings of the dictionary keys and items, interleaved
'  first key, first item, second key second item etc.

'Type indicator characters are as follows:
' # Double, payload is hex e.g. 1.5 encoded as D3FF8000000000000
' £ (pound sterling) String
' T Boolean True
' F Boolean False
' D Date, payload is decimal of Excel's date representation. e.g. 22-Dec-2025 is D64013
' G DateTime, payload is hex
' E Empty
' N Null
' % Integer
' & Long
' S Single, payload is hex
' C Currency
' ! Error
' @ Decimal
' * Array
' ^ Scripting.Dictionary

'Examples:
'#3FF0000000000000 unserialises to Double 1
'&1 unserailises to Long 1
'£Hello unserialises to String Hello
'T unserialises to Boolean True
'F unserialises to Boolean False
'*1,7;2,2,17,1,1,6,6,;%1%2#4008000000000000TF£Hello£World  unserialises to Array(1,2,3.0,True,False,"Hello","World")
'^2;2,3,4,5,;£a%10£abc%1000 unserialises to a Dictionary with two elements, element "a" contains 10 and element "abc" contains 1000

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : UnserialiseFromFile
' Purpose    : Read the file saved by the Julia code and unserialise its contents.
' -----------------------------------------------------------------------------------------------------------------------
Function UnserialiseFromFile(FileName As String, AllowNested As Boolean, StringLengthLimit As Long, JuliaVectorToXLColumn As Boolean)
          Dim Contents As String
          Dim ErrMsg As String
          Dim fso As New Scripting.FileSystemObject
          Dim TS As Scripting.TextStream

1         On Error GoTo ErrHandler
2         Set TS = fso.OpenTextFile(FileName, ForReading, , TristateTrue)
3         Contents = TS.ReadAll
4         TS.Close
5         Set TS = Nothing
6         Assign UnserialiseFromFile, Unserialise(Contents, AllowNested, 0, StringLengthLimit, JuliaVectorToXLColumn)

7         Exit Function
ErrHandler:
8         ErrMsg = ReThrow("UnserialiseFromFile", Err, True)
9         If Not TS Is Nothing Then TS.Close
10        Throw ErrMsg
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
2             Select Case val(Application.Version)
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
5             Case 163    '£ (pound sterling) vbString
6                 If StringLengthLimit > 0 Then 'Calling from worksheet formula, StringLengthLimit applies to elements of an array
7                     If Len(Chars) > IIf(Depth = 1, 32768, StringLengthLimit) Then 'Remember Chars includes an initial type indicator character of "£"
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
18            Case 68 ' D vbDate from Date in Julia
19                Unserialise = CDate(Mid$(Chars, 2))
20            Case 70     'F Boolean False
21                Unserialise = False
22            Case 71 'G vbDate, from DateTime in Julia
23                Unserialise = CDate(HexToDouble(Mid$(Chars, 2)))
24            Case 69     'E vbEmpty
25                Unserialise = Empty
26            Case 78     'N vbNull
27                Unserialise = Null
28            Case 37     '% vbInteger
29                Unserialise = CInt(Mid$(Chars, 2))
30            Case 38     '& Int64 converts to LongLong on 64bit, Double on 32bit
31                Unserialise = parseInt64(Mid$(Chars, 2))
32            Case 83     'S vbSingle
33                Unserialise = HexToSingle(Mid$(Chars, 2))
34            Case 67    'C vbCurrency
35                Unserialise = CCur(Mid$(Chars, 2))
36            Case 33     '! vbError
37                Unserialise = CVErr(Mid$(Chars, 2))
38            Case 64     '@ vbDecimal
39                Unserialise = CDec(Mid$(Chars, 2))
                  
40            Case 42     '* vbArray
41                If Depth > 1 Then If Not AllowNesting Then Throw "Excel cannot display arrays containing arrays"

                  Dim Ret() As Variant
                  Dim p1 As Long    ' Position of first ';'
                  Dim p2 As Long    ' Position of second ';'
                  Dim m As Long     ' Pointer into lengths section
                  Dim m2 As Long
                  Dim k As Long     ' Pointer into payload section
                  Dim thislength As Long

42                p1 = InStr(Chars, ";")
43                p2 = InStr(p1 + 1, Chars, ";")
44                m = p1 + 1
45                k = p2 + 1

                  ' Rank is the single character after '*', e.g. "*2,3,4;..."
                  Dim rank As Long
46                rank = CInt(Mid$(Chars, 2, 1))

47                Select Case rank
                      Case 1
                          ' === existing 1-D handling (unchanged) ===
                          Dim n As Long, i As Long
48                        n = CLng(Mid$(Chars, 4, p1 - 4))
49                        If n = 0 Then
50                            If Not AllowNesting Then Throw "Excel cannot display arrays with zero elements"
51                            Unserialise = VBA.Split(vbNullString)
52                        Else
53                            If JuliaVectorToXLColumn Then
54                                ReDim Ret(1 To n, 1 To 1)
55                                For i = 1 To n
56                                    m2 = InStr(m, Chars, ",") + 1
57                                    thislength = CLng(Mid$(Chars, m, m2 - m - 1))
58                                    Assign Ret(i, 1), Unserialise(Mid$(Chars, k, thislength), AllowNesting, Depth, StringLengthLimit, JuliaVectorToXLColumn)
59                                    k = k + thislength
60                                    m = m2
61                                Next i
62                            Else
63                                ReDim Ret(1 To n)
64                                For i = 1 To n
65                                    m2 = InStr(m, Chars, ",") + 1
66                                    thislength = CLng(Mid$(Chars, m, m2 - m - 1))
67                                    Assign Ret(i), Unserialise(Mid$(Chars, k, thislength), AllowNesting, Depth, StringLengthLimit, JuliaVectorToXLColumn)
68                                    k = k + thislength
69                                    m = m2
70                                Next i
71                            End If
72                            Unserialise = Ret
73                        End If

74                    Case 2
                          ' === existing 2-D handling (unchanged) ===
                          Dim commapos As Long, NC As Long, NR As Long, j As Long
75                        commapos = InStr(4, Chars, ",")
76                        NR = CLng(Mid$(Chars, 4, commapos - 4))
77                        NC = CLng(Mid$(Chars, commapos + 1, p1 - commapos - 1))
78                        If NR = 0 Or NC = 0 Then Throw "Cannot create array of size zero"
79                        ReDim Ret(1 To NR, 1 To NC)
80                        For j = 1 To NC
81                            For i = 1 To NR
82                                m2 = InStr(m, Chars, ",") + 1
83                                thislength = CLng(Mid$(Chars, m, m2 - m - 1))
84                                Assign Ret(i, j), Unserialise(Mid$(Chars, k, thislength), AllowNesting, Depth, StringLengthLimit, JuliaVectorToXLColumn)
85                                k = k + thislength
86                                m = m2
87                            Next i
88                        Next j
89                        Unserialise = Ret

90                    Case Else
                          ' === NEW: rank >= 3 === THIS SECTION WRITTEN BY COPILOT 23 DEC 2025
                          Dim Dims() As Long
91                        Dims = ParseDims(Mid$(Chars, 4, p1 - 4), rank)  ' section between "*,<rank>," and first ';'

                          ' Guard: Excel cannot display >2-D arrays; allow only when nesting is permitted
92                        If Not AllowNesting Then
93                            Throw "Excel cannot display arrays with more than 2 dimensions"
94                        End If

                          ' None of the dims may be zero
                          Dim q As Long, total As Long
95                        total = 1
96                        For q = 1 To rank
97                            If Dims(q) <= 0 Then Throw "Cannot create array of size zero"
98                            total = total * Dims(q)
99                        Next q

                          ' Allocate Ret() to the requested rank (up to MAX_RANK supported)
100                       ReDimVariantArray Ret, Dims

                          ' Walk in column-major order (dim 1 fastest), assigning elements
                          Dim idx() As Long
101                       ReDim idx(1 To rank)
102                       For q = 1 To rank: idx(q) = 1: Next q

                          Dim count As Long, val As Variant
103                       For count = 1 To total
104                           m2 = InStr(m, Chars, ",") + 1
105                           thislength = CLng(Mid$(Chars, m, m2 - m - 1))
106                           Assign val, Unserialise(Mid$(Chars, k, thislength), AllowNesting, Depth, StringLengthLimit, JuliaVectorToXLColumn)
107                           AssignByRank Ret, idx, val  ' Assign Ret(i1, i2, ..., irank) = val

108                           k = k + thislength
109                           m = m2

                              ' Increment indices: dim 1 fastest
110                           q = 1
111                           Do While q <= rank
112                               idx(q) = idx(q) + 1
113                               If idx(q) <= Dims(q) Then Exit Do
114                               idx(q) = 1
115                               q = q + 1
116                           Loop
117                           If q > rank Then Exit For
118                       Next count

119                       Unserialise = Ret
120               End Select
121           Case 94 '^ Dictionary
122               If Not AllowNesting Then Throw "Excel cannot display variables of type Dictionary"
123               p1 = InStr(Chars, ";")
124               p2 = InStr(p1 + 1, Chars, ";")
125               m = p1 + 1 '"pointer" to read from lengths section. Points to the first character after each comma.
126               k = p2 + 1 '"pointer" to read from contents section. Points to the first character of each "chunk".
                  Dim DictRet As New Scripting.Dictionary
                  Dim keylength As Long
                  Dim m3 As Long
                  Dim ThisKey As Variant
                  Dim ThisValue As Variant
                  Dim valuelength As Long
127               n = Mid$(Chars, 2, p1 - 2) 'Num elements in dictionary
128               For i = 1 To n
129                   m2 = InStr(m, Chars, ",") + 1
130                   m3 = InStr(m2, Chars, ",") + 1
131                   keylength = Mid$(Chars, m, m2 - m - 1)
132                   valuelength = Mid$(Chars, m2, m3 - m2 - 1)
133                   Assign ThisKey, Unserialise(Mid$(Chars, k, keylength), AllowNesting, Depth, StringLengthLimit, JuliaVectorToXLColumn)
134                   k = k + keylength
135                   Assign ThisValue, Unserialise(Mid$(Chars, k, valuelength), AllowNesting, Depth, StringLengthLimit, JuliaVectorToXLColumn)
136                   k = k + valuelength
137                   m = m3
138                   DictRet.Add ThisKey, ThisValue
139               Next i
140               Set Unserialise = DictRet
141           Case Else
142               Throw "Character '" & Left$(Chars, 1) & "' is not recognised as a type identifier"
143       End Select

144       Exit Function
ErrHandler:
145       ReThrow "Unserialise", Err
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
' Author     : Philip Swannell
' Date       : 21-Dec-2025
' Purpose    : Return a 16-character uppercase hexadecimal string representing the IEEE-754 bit pattern of `x` (Double).
'              Does not special-case NaN, +0.0 or -0.0.
' -----------------------------------------------------------------------------------------------------------------------
Function DoubleToHex(ByVal x As Double) As String

          Dim H1 As String
          Dim H2 As String
          Dim Out As String
          Dim TD As TDouble
          Dim Tl As TLongs
          
1         On Error GoTo ErrHandler
2         TD.d = x
3         LSet Tl = TD  ' reinterpret the 8 bytes of the Double as two Longs

4         Out = "0000000000000000"
5         H1 = Hex$(Tl.Hi)
6         H2 = Hex$(Tl.Lo)

7         Mid$(Out, 9 - Len(H1)) = H1
8         Mid$(Out, 17 - Len(H2)) = H2
9         DoubleToHex = Out

10        Exit Function
ErrHandler:
11        Throw "DoubleToHex (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : HexToDouble
' Author     : Philip Swannell
' Date       : 21-Dec-2025
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
10        Throw "HexToDouble (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SingleToHex
' Author     : Philip Swannell
' Date       : 22-Dec-2025
' Purpose    : Return a 8-character uppercase hexadecimal string representing the IEEE-754 bit pattern of `x` (Single).
'              Does not special-case NaN, +0.0 or -0.0.
' -----------------------------------------------------------------------------------------------------------------------
Function SingleToHex(ByVal x As Single) As String

          Dim Tl As TLong
          Dim TS As TSingle
          
1         On Error GoTo ErrHandler
2         TS.s = x
3         LSet Tl = TS  ' reinterpret the 4 bytes of the Single as a Long
4         SingleToHex = LPad(Hex$(Tl.x), 8, "0")
5         Exit Function
ErrHandler:
6         Throw "SingleToHex (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : LPad
' Author     : Philip Swannell
' Date       : 22-Dec-2025
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
' Author     : Philip Swannell
' Date       : 22-Dec-2025
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
8         Throw "HexToSingle (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function


' Parse a comma-delimited list of dimension sizes (e.g., "3,4,5") into dims(1..rank).
Private Function ParseDims(ByVal s As String, ByVal rank As Long) As Long()
          Dim parts() As String
1         parts = Split(s, ",")
2         If UBound(parts) + 1 <> rank Then
3             Throw "Malformed array header: expected " & rank & " dimensions, found " & (UBound(parts) + 1)
4         End If
          Dim Dims() As Long, i As Long
5         ReDim Dims(1 To rank)
6         For i = 1 To rank
7             Dims(i) = CLng(parts(i - 1))
8         Next i
9         ParseDims = Dims
End Function



' ReDim Ret() to the specified dims (1..rank). Increase MAX_RANK if needed.
Private Sub ReDimVariantArray(ByRef Ret() As Variant, ByRef Dims() As Long)
          Const MAX_RANK As Long = 8
1         Dim r As Long: r = UBound(Dims)
2         If r < 1 Or r > MAX_RANK Then
3             Throw "Cannot unserialise arrays with " & r & " dimensions (max supported: " & MAX_RANK & ")"
4         End If

5         Select Case r
              Case 1: ReDim Ret(1 To Dims(1))
6             Case 2: ReDim Ret(1 To Dims(1), 1 To Dims(2))
7             Case 3: ReDim Ret(1 To Dims(1), 1 To Dims(2), 1 To Dims(3))
8             Case 4: ReDim Ret(1 To Dims(1), 1 To Dims(2), 1 To Dims(3), 1 To Dims(4))
9             Case 5: ReDim Ret(1 To Dims(1), 1 To Dims(2), 1 To Dims(3), 1 To Dims(4), 1 To Dims(5))
10            Case 6: ReDim Ret(1 To Dims(1), 1 To Dims(2), 1 To Dims(3), 1 To Dims(4), 1 To Dims(5), 1 To Dims(6))
11            Case 7: ReDim Ret(1 To Dims(1), 1 To Dims(2), 1 To Dims(3), 1 To Dims(4), 1 To Dims(5), 1 To Dims(6), 1 To Dims(7))
12            Case 8: ReDim Ret(1 To Dims(1), 1 To Dims(2), 1 To Dims(3), 1 To Dims(4), 1 To Dims(5), 1 To Dims(6), 1 To Dims(7), 1 To Dims(8))
13        End Select
End Sub


' Assign Ret(i1, i2, ..., irank) = val, where idx(1..r) holds indices.
Private Sub AssignByRank(ByRef Ret() As Variant, ByRef idx() As Long, ByRef val As Variant)
1         Select Case UBound(idx)
              Case 1: Assign Ret(idx(1)), val
2             Case 2: Assign Ret(idx(1), idx(2)), val
3             Case 3: Assign Ret(idx(1), idx(2), idx(3)), val
4             Case 4: Assign Ret(idx(1), idx(2), idx(3), idx(4)), val
5             Case 5: Assign Ret(idx(1), idx(2), idx(3), idx(4), idx(5)), val
6             Case 6: Assign Ret(idx(1), idx(2), idx(3), idx(4), idx(5), idx(6)), val
7             Case 7: Assign Ret(idx(1), idx(2), idx(3), idx(4), idx(5), idx(6), idx(7)), val
8             Case 8: Assign Ret(idx(1), idx(2), idx(3), idx(4), idx(5), idx(6), idx(7), idx(8)), val
9             Case Else
10                Throw "Rank > 8 not supported by AssignByRank"
11        End Select
End Sub

