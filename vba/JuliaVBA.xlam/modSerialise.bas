Attribute VB_Name = "modSerialise"
' Copyright (c) 2021 - Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaVBA.jl#readme

Option Explicit
Option Private Module
Option Base 1

'Data format used by Serialise and Unserialise
'=============================================
'Format designed to be as fast as possible to unserialise.
'- Singleton types are prefixed with a type indicator character.
'- Dates are converted to their Excel representation - faster to unserialise in VBA.
'- Arrays are written with type indicator *, then three sections separated by semi-colons:
'  First section gives the number of dimensions and the dimensions themselves, comma
'  delimited e.g. a 3 x 4 array would have a dimensions section "2,3,4".
'  Second section gives the lengths of the encodings of each element, comma delimited with a
'  terminating comma.
'  Third section gives the encodings, concatenated with no delimiter.
'- Note that arrays are written in column-major order.
'- Nested arrays (arrays containing arrays) are supported by the format, and by VBA but
'  cannot be returned to a worksheet.

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

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : UnserialiseFromFile
' Purpose    : Read the file saved by the Julia code and unserialise its contents.
' -----------------------------------------------------------------------------------------------------------------------
Function UnserialiseFromFile(FileName As String)
          Dim AllowNesting As Boolean
          Dim Contents As String
          Dim ErrMsg As String
          Dim FSO As New Scripting.FileSystemObject
          Dim ts As Scripting.TextStream

1         On Error GoTo ErrHandler
2         Set ts = FSO.OpenTextFile(FileName, ForReading, , TristateTrue)
3         Contents = ts.ReadAll
4         ts.Close
5         Set ts = Nothing
6         AllowNesting = TypeName(Application.Caller) <> "Range"
7         UnserialiseFromFile = Unserialise(Contents, AllowNesting)
8         Exit Function
ErrHandler:
9         ErrMsg = "#UnserialiseFromFile (line " & CStr(Erl) + "): " & Err.Description & "!"
10        If Not ts Is Nothing Then ts.Close
11        Throw ErrMsg
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Unserialise
' Purpose    : Unserialises the contents of the results file saved by JuliaVBA julia code.
' -----------------------------------------------------------------------------------------------------------------------
Function Unserialise(Chars As String, AllowNesting As Boolean, Optional ByRef Depth As Long)

1         On Error GoTo ErrHandler
2         Depth = Depth + 1
3         Select Case Asc(Left$(Chars, 1))
              Case 35    '# vbDouble
4                 Unserialise = CDbl(Mid$(Chars, 2))
5             Case 163    '£ (pound sterling) vbString
6                 Unserialise = Mid$(Chars, 2)
7             Case 84     'T Boolean True
8                 Unserialise = True
9             Case 70     'F Boolean False
10                Unserialise = False
11            Case 68     'D vbDate
12                Unserialise = CDate(Mid$(Chars, 2))
13            Case 69     'E vbEmpty
14                Unserialise = Empty
15            Case 78     'N vbNull
16                Unserialise = Null
17            Case 37     '% vbInteger
18                Unserialise = CInt(Mid$(Chars, 2))
19            Case 38     '& Int64 converts to LongLong on 64bit, Double on 32bit
20                Unserialise = parseInt64(Mid$(Chars, 2))
21            Case 83     'S vbSingle
22                Unserialise = CSng(Mid$(Chars, 2))
23            Case 67    'C vbCurrency
24                Unserialise = CCur(Mid$(Chars, 2))
25            Case 33     '! vbError
26                Unserialise = CVErr(Mid$(Chars, 2))
27            Case 64     '@ vbDecimal
28                Unserialise = CDec(Mid$(Chars, 2))
29            Case 42     '* vbArray
30                If Depth > 1 Then If Not AllowNesting Then Throw "Excel cannot display arrays containing arrays"
                  Dim Ret() As Variant
                  Dim p1 As Long 'Position of first semi-colon
                  Dim p2 As Long 'Position of second semi-colon
                  Dim m As Long '"pointer" to read from lengths section
                  Dim m2 As Long
                  Dim k As Long '"pointer" to read from contents section
                  Dim thislength As Long
                  Dim i As Long ' Index into Ret
                  Dim j As Long 'Index into Ret
              
31                p1 = InStr(Chars, ";")
32                p2 = InStr(p1 + 1, Chars, ";")
33                m = p1 + 1
34                k = p2 + 1
              
35                Select Case Mid$(Chars, 2, 1)
                      Case 1 '1 dimensional array
                          Dim N As Long 'Num elements in array
36                        N = Mid$(Chars, 4, p1 - 4)
37                        If N = 0 Then Throw "Cannot create array of size zero"
38                        ReDim Ret(1 To N)
39                        For i = 1 To N
40                            m2 = InStr(m + 1, Chars, ",")
41                            thislength = Mid$(Chars, m, m2 - m)
42                            Ret(i) = Unserialise(Mid$(Chars, k, thislength), AllowNesting, Depth)
43                            k = k + thislength
44                            m = m2 + 1
45                        Next i
46                        Unserialise = Ret
47                    Case 2 '2 dimensional array
                          Dim commapos As Long
                          Dim NC As Long
                          Dim NR As Long
48                        commapos = InStr(4, Chars, ",")
49                        NR = Mid$(Chars, 4, commapos - 4)
50                        NC = Mid$(Chars, commapos + 1, p1 - commapos - 1)
51                        If NR = 0 Or NC = 0 Then Throw "Cannot create array of size zero"
52                        ReDim Ret(1 To NR, 1 To NC)
53                        For j = 1 To NC
54                            For i = 1 To NR
55                                m2 = InStr(m + 1, Chars, ",")
56                                thislength = Mid$(Chars, m, m2 - m)
57                                Ret(i, j) = Unserialise(Mid$(Chars, k, thislength), AllowNesting, Depth)
58                                k = k + thislength
59                                m = m2 + 1
60                            Next i
61                        Next j
62                        Unserialise = Ret
63                    Case Else
64                        Throw "Cannot unserialise arrays with more than 2 dimensions"
65                End Select
66            Case Else
67                Throw "Character '" & Left$(Chars, 1) & "' is not recognised as a type identifier"
68        End Select

69        Exit Function
ErrHandler:
70        Throw "#Unserialise (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'Values of type Int64 in Julia must be handled differently on Excel 32-bit and Excel 64bit
#If Win64 Then
    Function parseInt64(x As String)
        parseInt64 = CLngLng(x)
    End Function
#Else
    Function parseInt64(x As String)
        parseInt64 = CDbl(x)
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
' Purpose    : Equivalent to the julia function in JuliaVBA.encode_for_xl and serialises to the same format, though this
'              VBA version is not currently used.
' -----------------------------------------------------------------------------------------------------------------------
Function Serialise(x As Variant) As String

          Dim contentsArray() As String
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim lengthsArray() As String
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
26            Case Is >= vbArray
27                Select Case NumDimensions(x)
                      Case 1
28                        ReDim lengthsArray(LBound(x) To UBound(x))
29                        ReDim contentsArray(LBound(x) To UBound(x))
30                        For i = LBound(x) To UBound(x)
31                            contentsArray(i) = Serialise(x(i))
32                            lengthsArray(i) = CStr(Len(contentsArray(i)))
33                        Next i
34                        Serialise = "*1," & CStr(UBound(x) - LBound(x) + 1) & ";" & VBA.Join(lengthsArray, ",") & ",;" & VBA.Join(contentsArray, "")
35                    Case 2
36                        NR = UBound(x, 1) - LBound(x, 1) + 1
37                        NC = UBound(x, 2) - LBound(x, 2) + 1
38                        k = 0
39                        ReDim lengthsArray(NR * NC)
40                        ReDim contentsArray(NR * NC)
41                        For j = LBound(x, 2) To UBound(x, 2)
42                            For i = LBound(x, 1) To UBound(x, 1)
43                                k = k + 1
44                                contentsArray(k) = Serialise(x(i, j))
45                                lengthsArray(k) = CStr(Len(contentsArray(k)))
46                            Next i
47                        Next j
48                        Serialise = "*2," & CStr(UBound(x, 1) - LBound(x, 1) + 1) & "," & CStr(UBound(x, 2) - LBound(x, 2) + 1) & ";" & VBA.Join(lengthsArray, ",") & ",;" & VBA.Join(contentsArray, "")
49                    Case Else
50                        Throw "Cannot serialise array with " + CStr(NumDimensions(x)) + " dimensions"
51                End Select
52            Case Else
53                Throw "Cannot serialise variable of type " & TypeName(x)
54        End Select

55        Exit Function
ErrHandler:
56        Throw "#Serialise (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
