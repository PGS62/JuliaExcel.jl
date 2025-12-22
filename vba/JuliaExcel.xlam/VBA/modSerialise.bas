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

'Data format used by Serialise and Unserialise
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
33                Unserialise = CSng(Mid$(Chars, 2))
34            Case 67    'C vbCurrency
35                Unserialise = CCur(Mid$(Chars, 2))
36            Case 33     '! vbError
37                Unserialise = CVErr(Mid$(Chars, 2))
38            Case 64     '@ vbDecimal
39                Unserialise = CDec(Mid$(Chars, 2))
40            Case 42     '* vbArray
41                If Depth > 1 Then If Not AllowNesting Then Throw "Excel cannot display arrays containing arrays"
                  Dim Ret() As Variant
                  Dim p1 As Long 'Position of first semi-colon
                  Dim p2 As Long 'Position of second semi-colon
                  Dim m As Long '"pointer" to read from lengths section
                  Dim m2 As Long
                  Dim k As Long '"pointer" to read from contents section
                  Dim thislength As Long
                  Dim i As Long ' Index into Ret
                  Dim j As Long 'Index into Ret
              
42                p1 = InStr(Chars, ";")
43                p2 = InStr(p1 + 1, Chars, ";")
44                m = p1 + 1
45                k = p2 + 1
              
46                Select Case Mid$(Chars, 2, 1)
                      Case 1 '1 dimensional array
                          Dim n As Long 'Num elements in array
47                        n = Mid$(Chars, 4, p1 - 4)
48                        If n = 0 Then
49                            If Not AllowNesting Then Throw "Excel cannot display arrays with zero elements"
50                            Unserialise = VBA.Split(vbNullString) 'See discussion at https://stackoverflow.com/questions/55123413/declare-a-0-length-string-array-in-vba-impossible
51                        Else
52                            If JuliaVectorToXLColumn Then
53                                ReDim Ret(1 To n, 1 To 1)
54                                For i = 1 To n
55                                    m2 = InStr(m, Chars, ",") + 1
56                                    thislength = Mid$(Chars, m, m2 - m - 1)
57                                    Assign Ret(i, 1), Unserialise(Mid$(Chars, k, thislength), AllowNesting, Depth, StringLengthLimit, JuliaVectorToXLColumn)
58                                    k = k + thislength
59                                    m = m2
60                                Next i
61                            Else
62                                ReDim Ret(1 To n)
63                                For i = 1 To n
64                                    m2 = InStr(m, Chars, ",") + 1
65                                    thislength = Mid$(Chars, m, m2 - m - 1)
66                                    Assign Ret(i), Unserialise(Mid$(Chars, k, thislength), AllowNesting, Depth, StringLengthLimit, JuliaVectorToXLColumn)
67                                    k = k + thislength
68                                    m = m2
69                                Next i
70                            End If
71                            Unserialise = Ret
72                        End If
73                    Case 2 '2 dimensional array
                          Dim commapos As Long
                          Dim NC As Long
                          Dim NR As Long
74                        commapos = InStr(4, Chars, ",")
75                        NR = Mid$(Chars, 4, commapos - 4)
76                        NC = Mid$(Chars, commapos + 1, p1 - commapos - 1)
77                        If NR = 0 Or NC = 0 Then Throw "Cannot create array of size zero"
78                        ReDim Ret(1 To NR, 1 To NC)
79                        For j = 1 To NC
80                            For i = 1 To NR
81                                m2 = InStr(m, Chars, ",") + 1
82                                thislength = Mid$(Chars, m, m2 - m - 1)
83                                Assign Ret(i, j), Unserialise(Mid$(Chars, k, thislength), AllowNesting, Depth, StringLengthLimit, JuliaVectorToXLColumn)
84                                k = k + thislength
85                                m = m2
86                            Next i
87                        Next j
88                        Unserialise = Ret
89                    Case Else
90                        Throw "Cannot unserialise arrays with more than 2 dimensions"
91                End Select
92            Case 94 '^ Dictionary
93                If Not AllowNesting Then Throw "Excel cannot display variables of type Dictionary"
94                p1 = InStr(Chars, ";")
95                p2 = InStr(p1 + 1, Chars, ";")
96                m = p1 + 1 '"pointer" to read from lengths section. Points to the first character after each comma.
97                k = p2 + 1 '"pointer" to read from contents section. Points to the first character of each "chunk".
                  Dim DictRet As New Scripting.Dictionary
                  Dim keylength As Long
                  Dim m3 As Long
                  Dim ThisKey As Variant
                  Dim ThisValue As Variant
                  Dim valuelength As Long
98                n = Mid$(Chars, 2, p1 - 2) 'Num elements in dictionary
99                For i = 1 To n
100                   m2 = InStr(m, Chars, ",") + 1
101                   m3 = InStr(m2, Chars, ",") + 1
102                   keylength = Mid$(Chars, m, m2 - m - 1)
103                   valuelength = Mid$(Chars, m2, m3 - m2 - 1)
104                   Assign ThisKey, Unserialise(Mid$(Chars, k, keylength), AllowNesting, Depth, StringLengthLimit, JuliaVectorToXLColumn)
105                   k = k + keylength
106                   Assign ThisValue, Unserialise(Mid$(Chars, k, valuelength), AllowNesting, Depth, StringLengthLimit, JuliaVectorToXLColumn)
107                   k = k + valuelength
108                   m = m3
109                   DictRet.Add ThisKey, ThisValue
110               Next i
111               Set Unserialise = DictRet
112           Case Else
113               Throw "Character '" & Left$(Chars, 1) & "' is not recognised as a type identifier"
114       End Select

115       Exit Function
ErrHandler:
116       ReThrow "Unserialise", Err
End Function

'Values of type Int64 in Julia must be handled differently on Excel 32-bit and Excel 64bit
#If Win64 Then
    Function parseInt64(x As String)
1             parseInt64 = CLngLng(x)
    End Function
#Else
    Function parseInt64(x As String)
1             parseInt64 = CDbl(x)
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

10        Out = "0000000000000000"
11        H1 = Hex$(Tl.Hi)
12        H2 = Hex$(Tl.Lo)

13        Mid$(Out, 9 - Len(H1)) = H1
14        Mid$(Out, 17 - Len(H2)) = H2
15        DoubleToHex = Out

16        Exit Function
ErrHandler:
17        Throw "DoubleToHex (line " & CStr(Erl) + "): " & Err.Description & "!"
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

