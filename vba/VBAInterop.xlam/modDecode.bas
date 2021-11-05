Attribute VB_Name = "modDecode"
Option Explicit
Option Private Module
Option Base 1

#If Win64 Then
    Function parseInt64(x As String)
        parseInt64 = CLngLng(x)
    End Function
#Else
    Function parseInt64(x As String)
        parseInt64 = CDbl(x)
    End Function
#End If



'Decode implements a data un-serialisation for a format that's easier and faster to
'unserialise than csv.
'- Singleton types are prefixed with a type indicator character.
'- Dates are converted to their Excel representation - faster to unserialise in VBA.
'- Arrays are written with type indicator *, then three sections separated by semi-colons:
'  First section gives the number of dimensions and the dimensions themselves, comma
'  delimited e.g. a 3 x 4 array would have a dimensions section "2,3,4".
'  Second section gives the lengths of the encodings of each element, comma delimited with a
'  terminating comma.
'  Third section gives the encodings, concatenated with no delimiter.
'  - Note that arrays are written in column-major order.
'Type indicator characters are as follows:
'# vbDouble
'£ (pound sterling) vbString
'T Boolean True
'F Boolean False
'D vbDate
'E vbEmpty
'N vbNull
'% vbInteger
'& vbLong
'S vbSingle
'C vbCurrency
'! vbError
'@ vbDecimal
'* vbArray

'
'Examples:
'?encode(CDbl(1))
'#1
'?encode(CLng(1))
'&1
'?encode("Hello")
'£Hello
'?encode(True)
'T
'?encode(False)
'F
'?encode(Array(1,2,3.0,True,False,"Hello","World"))
'*1,7;2,2,2,1,1,6,6,;%1%2#3TF£Hello£World

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Decode
' Author     : Philip Swannell
' Date       : 04-Nov-2021
' Purpose    : Decodes (unserializes) the contents of the results file saved by JuliaInterop julia code.
' -----------------------------------------------------------------------------------------------------------------------
Function Decode(Chars As String, Optional ByRef Depth As Long)

1         On Error GoTo ErrHandler
2         Depth = Depth + 1
3         Select Case Asc(Left$(Chars, 1))
              Case 35    '# vbDouble
4                 Decode = CDbl(Mid$(Chars, 2))
5             Case 163    '£ (pound sterling) vbString
6                 Decode = Mid$(Chars, 2)
7             Case 84     'T Boolean True
8                 Decode = True
9             Case 70     'F Boolean False
10                Decode = False
11            Case 68     'D vbDate
12                Decode = CDate(Mid$(Chars, 2))
13            Case 69     'E vbEmpty
14                Decode = Empty
15            Case 78     'N vbNull
16                Decode = Null
17            Case 37     '% vbInteger
18                Decode = CInt(Mid$(Chars, 2))
19            Case 38     '& Int64 converts to LongLong on 64bit, Double on 32bit
20                Decode = parseInt64(Mid$(Chars, 2))
21            Case 83     'S vbSingle
22                Decode = CSng(Mid$(Chars, 2))
23            Case 67    'C vbCurrency
24                Decode = CCur(Mid$(Chars, 2))
25            Case 33     '! vbError
26                Decode = CVErr(Mid$(Chars, 2))
27            Case 64     '@ vbDecimal
28                Decode = CDec(Mid$(Chars, 2))
29            Case 42     '* vbArray
30                If Depth > 1 Then Throw "Excel cannot display arrays containing arrays"
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
42                            Ret(i) = Decode(Mid$(Chars, k, thislength), Depth)
43                            k = k + thislength
44                            m = m2 + 1
45                        Next i
46                        Decode = Ret
47                    Case 2 '2 dimensional array
                          Dim NR As Long, NC As Long, commapos As Long
48                        commapos = InStr(4, Chars, ",")
49                        NR = Mid$(Chars, 4, commapos - 4)
50                        NC = Mid$(Chars, commapos + 1, p1 - commapos - 1)
51                        If NR = 0 Or NC = 0 Then Throw "Cannot create array of size zero"
52                        ReDim Ret(1 To NR, 1 To NC)
53                        For j = 1 To NC
54                            For i = 1 To NR
55                                m2 = InStr(m + 1, Chars, ",")
56                                thislength = Mid$(Chars, m, m2 - m)
57                                Ret(i, j) = Decode(Mid$(Chars, k, thislength), Depth)
58                                k = k + thislength
59                                m = m2 + 1
60                            Next i
61                        Next j
62                        Decode = Ret
63                    Case Else
64                        Throw "Cannot decode arrays with more than 2 dimensions"
65                End Select
66            Case Else
67                Throw "Character '" & Left$(Chars, 1) & "' is not recognised as a type identifier"
68        End Select

69        Exit Function
ErrHandler:
70        Throw "#Decode (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Encode
' Author     : Philip Swannell
' Date       : 04-Nov-2021
' Purpose    : Equivalent to the julia function in VBAInterop.encode_for_xl and encodes to the same format, though this
'              VBA version is not currently used.
' Parameters :
'  x:
' -----------------------------------------------------------------------------------------------------------------------
Function Encode(x) As String

          Dim lengthsArray() As String
          Dim contentsArray() As String
          Dim i As Long, j As Long, k As Long, NR As Long, NC As Long

1         On Error GoTo ErrHandler
2         Select Case VarType(x)
              Case vbEmpty
3                 Encode = "E"
4             Case vbNull
5                 Encode = "N"
6             Case vbInteger
7                 Encode = "%" & CStr(x)
8             Case vbLong
9                 Encode = "&" & CStr(x)
10            Case vbSingle
11                Encode = "S" & CStr(x)
12            Case vbDouble
13                Encode = "#" & CStr(x)
14            Case vbCurrency
15                Encode = "C" & CStr(x)
16            Case vbDate
17                Encode = "D" & CStr(CDbl(x))
18            Case vbString
19                Encode = "£" & x
20            Case vbError
21                Encode = "!" & CStr(CLng(x))
22            Case vbBoolean
23                Encode = IIf(x, "T", "F")
24            Case vbDecimal
25                Encode = "@" & CStr(x)
26            Case Is >= vbArray
27                Select Case NumDimensions(x)
                      Case 1
28                        ReDim lengthsArray(LBound(x) To UBound(x))
29                        ReDim contentsArray(LBound(x) To UBound(x))
30                        For i = LBound(x) To UBound(x)
31                            contentsArray(i) = Encode(x(i))
32                            lengthsArray(i) = CStr(Len(contentsArray(i)))
33                        Next i
34                        Encode = "*1," & CStr(UBound(x) - LBound(x) + 1) & ";" & VBA.Join(lengthsArray, ",") & ",;" & VBA.Join(contentsArray, "")
35                    Case 2
36                        NR = UBound(x, 1) - LBound(x, 1) + 1
37                        NC = UBound(x, 2) - LBound(x, 2) + 1
38                        k = 0
39                        ReDim lengthsArray(NR * NC)
40                        ReDim contentsArray(NR * NC)
41                        For j = LBound(x, 2) To UBound(x, 2)
42                            For i = LBound(x, 1) To UBound(x, 1)
43                                k = k + 1
44                                contentsArray(k) = Encode(x(i, j))
45                                lengthsArray(k) = CStr(Len(contentsArray(k)))
46                            Next i
47                        Next j
48                        Encode = "*2," & CStr(UBound(x, 1) - LBound(x, 1) + 1) & "," & CStr(UBound(x, 2) - LBound(x, 2) + 1) & ";" & VBA.Join(lengthsArray, ",") & ",;" & VBA.Join(contentsArray, "")
49                    Case Else
50                        Throw "Cannot encode array with " + CStr(NumDimensions(x)) + " dimensions"
51                End Select
52            Case Else
53                Throw "Cannot encode variable of type " & TypeName(x)
54        End Select

55        Exit Function
ErrHandler:
56        Throw "#Encode (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function


Function EncodeAndWriteFile(Data, FileName As String)

          Dim FSO As New Scripting.FileSystemObject
          Dim ts As Scripting.TextStream
          Dim ErrMsg As String

1         On Error GoTo ErrHandler
2         If TypeName(Data) = "Range" Then Data = Data.Value2
3         Set ts = FSO.OpenTextFile(FileName, ForWriting, True, TristateTrue)
4         ts.Write Encode(Data)
5         ts.Close
6         Set ts = Nothing
7         EncodeAndWriteFile = FileName

8         Exit Function
ErrHandler:
9         ErrMsg = "#EncodeAndWriteFile (line " & CStr(Erl) + ") error writing'" & FileName & "' " & Err.Description & "!"
10        If Not ts Is Nothing Then ts.Close
11        Throw ErrMsg
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ReadFileAndDecode
' Author     : Philip Swannell
' Date       : 04-Nov-2021
' Purpose    : Read the file saved by the Julia code and unserialize its contents.
' Parameters :
'  FileName:
' -----------------------------------------------------------------------------------------------------------------------
Function ReadFileAndDecode(FileName As String)
          Dim FSO As New Scripting.FileSystemObject
          Dim ts As Scripting.TextStream
          Dim ErrMsg As String
          Dim Contents As String

1         On Error GoTo ErrHandler
2         Set ts = FSO.OpenTextFile(FileName, ForReading, , TristateTrue)
3         Contents = ts.ReadAll
4         ts.Close
5         Set ts = Nothing
6         ReadFileAndDecode = Decode(Contents)

7         Exit Function
ErrHandler:
8         ErrMsg = "#ReadFileAndDecode (line " & CStr(Erl) + ") error reading'" & FileName & "' " & Err.Description & "!"
9         If Not ts Is Nothing Then ts.Close
10        Throw ErrMsg
End Function

