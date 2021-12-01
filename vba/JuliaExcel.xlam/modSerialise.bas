Attribute VB_Name = "modSerialise"
' Copyright (c) 2021 - Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme

Option Explicit
Option Private Module
Option Base 1

'Data format used by Serialise and Unserialise
'=============================================
'Format designed to be as fast as possible to unserialise.
'- Singleton types are prefixed with a type indicator character.
'- Dates are shown in their Excel representation as a number - faster to unserialise in VBA.
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
' # Double
' £ (pound sterling) String
' T Boolean True
' F Boolean False
' D Date
' E Empty
' N Null
' % Integer
' & Long
' S Single
' C Currency
' ! Error
' @ Decimal
' * Array
' ^ Scripting.Dictionary

'
'Examples:
'?Serialise(CDbl(1))
'#1
'?Serialise(CLng(1))
'&1
'?Serialise("Hello")
'£Hello
'?Serialise(True)
'T
'?Serialise(False)
'F
'?Serialise(Array(1,2,3.0,True,False,"Hello","World"))
'*1,7;2,2,2,1,1,6,6,;%1%2#3TF£Hello£World

'Set foo = New Scripting.Dictionary
'foo.add "a",10
'foo.add "abc",1000
'?serialise(foo)
'^2;2,3,4,5,;£a%10£abc%1000

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : UnserialiseFromFile
' Purpose    : Read the file saved by the Julia code and unserialise its contents.
' -----------------------------------------------------------------------------------------------------------------------
Function UnserialiseFromFile(FileName As String, AllowNested As Boolean, StringLengthLimit As Long, JuliaVectorToXLColumn As Boolean)
          Dim Contents As String
          Dim ErrMsg As String
          Dim FSO As New Scripting.FileSystemObject
          Dim ts As Scripting.TextStream

1         On Error GoTo ErrHandler
2         Set ts = FSO.OpenTextFile(FileName, ForReading, , TristateTrue)
3         Contents = ts.ReadAll
4         ts.Close
5         Set ts = Nothing
6         Assign UnserialiseFromFile, Unserialise(Contents, AllowNested, 0, StringLengthLimit, JuliaVectorToXLColumn)

7         Exit Function
ErrHandler:
8         ErrMsg = "#UnserialiseFromFile (line " & CStr(Erl) + "): " & Err.Description & "!"
9         If Not ts Is Nothing Then ts.Close
10        Throw ErrMsg
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetStringLengthLimit
' Purpose    : Different versions of Excel have different limits for the longest string that can be an element of an
'              array passed from a VBA UDF back to Excel. I know the limit is 255 for Excel 2010 and earlier, and is
'              32,767 for Excel 365 (as of Sep 2021). But don't yet know the limit for Excel 2013, 2016 and 2019.
' Tried to get info from StackOverflow, without much joy:
' https://stackoverflow.com/questions/69303804/excel-versions-and-limits-on-the-length-of-string-elements-in-arrays-returned-by
' Note that this function returns 1 more than the maximum allowed string length
' -----------------------------------------------------------------------------------------------------------------------
Function GetStringLengthLimit() As Long
          Static Res As Long
1         If Res = 0 Then
2             Select Case Val(Application.Version)
                  Case Is <= 14 'Excel 2010
3                     Res = 256
4                 Case 15
5                     Res = 32768 'Don't yet know if this is correct for Excel 2013
6                 Case Else
7                     Res = 32768 'Excel 2016, 2019, 365. Hopefully these versions (which all _
                                   return 16 as Application.Version) have the same limit.
8             End Select
9         End If
10        GetStringLengthLimit = Res
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Unserialise
' Purpose    : Unserialises the contents of the results file saved by JuliaExcel julia code.
' -----------------------------------------------------------------------------------------------------------------------
Function Unserialise(Chars As String, AllowNesting As Boolean, ByRef Depth As Long, StringLengthLimit As Long, JuliaVectorToXLColumn As Boolean)

1         On Error GoTo ErrHandler
2         Depth = Depth + 1
3         Select Case Asc(Left$(Chars, 1))
              Case 35    '# vbDouble
4                 Unserialise = CDbl(Mid$(Chars, 2))
5             Case 163    '£ (pound sterling) vbString
6                 If StringLengthLimit > 0 Then
7                     If Len(Chars) > StringLengthLimit Then
8                         Throw "Data contains a string of length " & Format(Len(Chars) - 1, "###,###") & _
                              ", too long to display in Excel version " + Application.Version() + " (the limit is " _
                              & Format(StringLengthLimit - 1, "###,###") + ")"
9                     End If
10                End If
11                Unserialise = Mid$(Chars, 2)
12            Case 84     'T Boolean True
13                Unserialise = True
14            Case 70     'F Boolean False
15                Unserialise = False
16            Case 68     'D vbDate
17                Unserialise = CDate(Mid$(Chars, 2))
18            Case 69     'E vbEmpty
19                Unserialise = Empty
20            Case 78     'N vbNull
21                Unserialise = Null
22            Case 37     '% vbInteger
23                Unserialise = CInt(Mid$(Chars, 2))
24            Case 38     '& Int64 converts to LongLong on 64bit, Double on 32bit
25                Unserialise = parseInt64(Mid$(Chars, 2))
26            Case 83     'S vbSingle
27                Unserialise = CSng(Mid$(Chars, 2))
28            Case 67    'C vbCurrency
29                Unserialise = CCur(Mid$(Chars, 2))
30            Case 33     '! vbError
31                Unserialise = CVErr(Mid$(Chars, 2))
32            Case 64     '@ vbDecimal
33                Unserialise = CDec(Mid$(Chars, 2))
34            Case 42     '* vbArray
35                If Depth > 1 Then If Not AllowNesting Then Throw "Excel cannot display arrays containing arrays"
                  Dim Ret() As Variant
                  Dim p1 As Long 'Position of first semi-colon
                  Dim p2 As Long 'Position of second semi-colon
                  Dim m As Long '"pointer" to read from lengths section
                  Dim m2 As Long
                  Dim k As Long '"pointer" to read from contents section
                  Dim thislength As Long
                  Dim i As Long ' Index into Ret
                  Dim j As Long 'Index into Ret
              
36                p1 = InStr(Chars, ";")
37                p2 = InStr(p1 + 1, Chars, ";")
38                m = p1 + 1
39                k = p2 + 1
              
40                Select Case Mid$(Chars, 2, 1)
                      Case 1 '1 dimensional array
                          Dim n As Long 'Num elements in array
41                        n = Mid$(Chars, 4, p1 - 4)
42                        If n = 0 Then
43                            If Not AllowNesting Then Throw "Excel cannot display arrays with zero elements"
44                            Unserialise = VBA.Split(vbNullString) 'See discussion at https://stackoverflow.com/questions/55123413/declare-a-0-length-string-array-in-vba-impossible
45                        Else
46                            If JuliaVectorToXLColumn Then
47                                ReDim Ret(1 To n, 1 To 1)
48                                For i = 1 To n
49                                    m2 = InStr(m, Chars, ",") + 1
50                                    thislength = Mid$(Chars, m, m2 - m - 1)
51                                    Assign Ret(i, 1), Unserialise(Mid$(Chars, k, thislength), AllowNesting, Depth, StringLengthLimit, JuliaVectorToXLColumn)
52                                    k = k + thislength
53                                    m = m2
54                                Next i
55                            Else
56                                ReDim Ret(1 To n)
57                                For i = 1 To n
58                                    m2 = InStr(m, Chars, ",") + 1
59                                    thislength = Mid$(Chars, m, m2 - m - 1)
60                                    Assign Ret(i), Unserialise(Mid$(Chars, k, thislength), AllowNesting, Depth, StringLengthLimit, JuliaVectorToXLColumn)
61                                    k = k + thislength
62                                    m = m2
63                                Next i
64                            End If
65                            Unserialise = Ret
66                        End If
67                    Case 2 '2 dimensional array
                          Dim commapos As Long
                          Dim NC As Long
                          Dim NR As Long
68                        commapos = InStr(4, Chars, ",")
69                        NR = Mid$(Chars, 4, commapos - 4)
70                        NC = Mid$(Chars, commapos + 1, p1 - commapos - 1)
71                        If NR = 0 Or NC = 0 Then Throw "Cannot create array of size zero"
72                        ReDim Ret(1 To NR, 1 To NC)
73                        For j = 1 To NC
74                            For i = 1 To NR
75                                m2 = InStr(m, Chars, ",") + 1
76                                thislength = Mid$(Chars, m, m2 - m - 1)
77                                Assign Ret(i, j), Unserialise(Mid$(Chars, k, thislength), AllowNesting, Depth, StringLengthLimit, JuliaVectorToXLColumn)
                                '  Ret(i, j) = Unserialise(Mid$(Chars, k, thislength), AllowNesting, Depth, StringLengthLimit, JuliaVectorToXLColumn)
78                                k = k + thislength
79                                m = m2
80                            Next i
81                        Next j
82                        Unserialise = Ret
83                    Case Else
84                        Throw "Cannot unserialise arrays with more than 2 dimensions"
85                End Select
86            Case 94 '^ Dictionary
87                If Not AllowNesting Then Throw "Excel cannot display variables of type Dictionary"
88                p1 = InStr(Chars, ";")
89                p2 = InStr(p1 + 1, Chars, ";")
90                m = p1 + 1 '"pointer" to read from lengths section. Points to the first character after each comma.
91                k = p2 + 1 '"pointer" to read from contents section. Points to the first character of each "chunk".
                  Dim DictRet As New Scripting.Dictionary
                  Dim keylength As Long
                  Dim m3 As Long
                  Dim ThisKey As Variant
                  Dim ThisValue As Variant
                  Dim valuelength As Long
92                n = Mid$(Chars, 2, p1 - 2) 'Num elements in dictionary
93                For i = 1 To n
94                    m2 = InStr(m, Chars, ",") + 1
95                    m3 = InStr(m2, Chars, ",") + 1
96                    keylength = Mid$(Chars, m, m2 - m - 1)
97                    valuelength = Mid$(Chars, m2, m3 - m2 - 1)
98                    Assign ThisKey, Unserialise(Mid$(Chars, k, keylength), AllowNesting, Depth, StringLengthLimit, JuliaVectorToXLColumn)
99                    k = k + keylength
100                   Assign ThisValue, Unserialise(Mid$(Chars, k, valuelength), AllowNesting, Depth, StringLengthLimit, JuliaVectorToXLColumn)
101                   k = k + valuelength
102                   m = m3
103                   DictRet.Add ThisKey, ThisValue
104               Next i
105               Set Unserialise = DictRet
106           Case Else
107               Throw "Character '" & Left$(Chars, 1) & "' is not recognised as a type identifier"
108       End Select

109       Exit Function
ErrHandler:
110       Throw "#Unserialise (line " & CStr(Erl) + "): " & Err.Description & "!"
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
' Procedure  : SerialiseToFile
' Purpose    : Serialise Data and write to file, the inverse of UnserialiseFromFile. Currently this procedure is not used
'              but might be useful for writing tests of UnserialiseFromFile.
' -----------------------------------------------------------------------------------------------------------------------
Function SerialiseToFile(Data, FileName As String)

          Dim ErrMsg As String
          Dim FSO As New Scripting.FileSystemObject
          Dim ts As Scripting.TextStream

1         On Error GoTo ErrHandler
2         If TypeName(Data) = "Range" Then Data = Data.Value2
3         Set ts = FSO.OpenTextFile(FileName, ForWriting, True, TristateTrue)
4         ts.Write Serialise(Data)
5         ts.Close
6         Set ts = Nothing
7         SerialiseToFile = FileName

8         Exit Function
ErrHandler:
9         ErrMsg = "#SerialiseToFile (line " & CStr(Erl) + ") error writing'" & FileName & "' " & Err.Description & "!"
10        If Not ts Is Nothing Then ts.Close
11        Throw ErrMsg
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Serialise
' Date       : 04-Nov-2021
' Purpose    : Equivalent to the julia function in JuliaExcel.encode_for_xl and serialises to the same format, though this
'              VBA version is not currently used.
' -----------------------------------------------------------------------------------------------------------------------
Function Serialise(x As Variant) As String

          Dim ContentsArray() As String
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim KeysArray() As String
          Dim LengthsArray() As String
          Dim NC As Long
          Dim NR As Long

1         On Error GoTo ErrHandler
2         Select Case VarType(x)
              Case vbEmpty
3                 Serialise = "E"
4             Case vbNull
5                 Serialise = "N"
6             Case vbInteger
7                 Serialise = "%" & CStr(x)
8             Case vbLong
9                 Serialise = "&" & CStr(x)
10            Case vbSingle
11                Serialise = "S" & CStr(x)
12            Case vbDouble
13                Serialise = "#" & CStr(x)
14            Case vbCurrency
15                Serialise = "C" & CStr(x)
16            Case vbDate
17                Serialise = "D" & CStr(CDbl(x))
18            Case vbString
19                Serialise = "£" & x
20            Case vbError
21                Serialise = "!" & CStr(CLng(x))
22            Case vbBoolean
23                Serialise = IIf(x, "T", "F")
24            Case vbDecimal
25                Serialise = "@" & CStr(x)
26            Case vbObject
27                If TypeName(x) <> "Dictionary" Then Throw "Cannot serialise object of type " + TypeName(x)
28                ReDim LengthsArray(0 To x.Count - 1)
29                ReDim ContentsArray(0 To x.Count - 1)
                  Dim key
                  Dim ThisItem As String
                  Dim ThisKey As String
30                i = 0
31                For Each key In x.Keys
32                    ThisKey = Serialise(key)
33                    ThisItem = Serialise(x(key))
34                    ContentsArray(i) = ThisKey & ThisItem
35                    LengthsArray(i) = CStr(Len(ThisKey)) & "," & CStr(Len(ThisItem))
36                    i = i + 1
37                Next key
38                Serialise = "^" & CStr(x.Count) & ";" & VBA.Join(LengthsArray, ",") & ",;" & VBA.Join(ContentsArray, "")
39            Case Is >= vbArray
40                Select Case NumDimensions(x)
                      Case 1
41                        ReDim LengthsArray(LBound(x) To UBound(x))
42                        ReDim ContentsArray(LBound(x) To UBound(x))
43                        For i = LBound(x) To UBound(x)
44                            ContentsArray(i) = Serialise(x(i))
45                            LengthsArray(i) = CStr(Len(ContentsArray(i)))
46                        Next i
47                        Serialise = "*1," & CStr(UBound(x) - LBound(x) + 1) & ";" & VBA.Join(LengthsArray, ",") & ",;" & VBA.Join(ContentsArray, "")
48                    Case 2
49                        NR = UBound(x, 1) - LBound(x, 1) + 1
50                        NC = UBound(x, 2) - LBound(x, 2) + 1
51                        k = 0
52                        ReDim LengthsArray(NR * NC)
53                        ReDim ContentsArray(NR * NC)
54                        For j = LBound(x, 2) To UBound(x, 2)
55                            For i = LBound(x, 1) To UBound(x, 1)
56                                k = k + 1
57                                ContentsArray(k) = Serialise(x(i, j))
58                                LengthsArray(k) = CStr(Len(ContentsArray(k)))
59                            Next i
60                        Next j
61                        Serialise = "*2," & CStr(UBound(x, 1) - LBound(x, 1) + 1) & "," & CStr(UBound(x, 2) - LBound(x, 2) + 1) & ";" & VBA.Join(LengthsArray, ",") & ",;" & VBA.Join(ContentsArray, "")
62                    Case Else
63                        Throw "Cannot serialise array with " + CStr(NumDimensions(x)) + " dimensions"
64                End Select
65            Case Else
66                Throw "Cannot serialise variable of type " & TypeName(x)
67        End Select

68        Exit Function
ErrHandler:
69        Throw "#Serialise (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

