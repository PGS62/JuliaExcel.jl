Attribute VB_Name = "modEncode"
Option Explicit
Option Private Module
Option Base 1

'Experimenting with a file format other than CSV for transporting data back from Julia to Excel.
'looks like MyFileFormat is about 2 to 3 times faster to read than CSV, no type conversion

'Examples Array(1#, "2.3")
' *1,2;2,4,;#1$2.3

' * = it's an array
' 1 = it has one dimension
' , = delimiter in dimensions section
' 2 = it's first dimension is of length 2
' ; = delimiter between sections
' 2 = the encoding of the first element of the contents section is of length 2
' , = delimiter in lengths section
' 4 = the encoding of the second element of the contents section is of length 4
' , = delimiter in the lengths section, note this terminatng delimiter
' ; = delimiter between the lengths section and the contents section
' #1 = encoding of the first element, # indicates Double
' $2.3 = encoding of the second element, $ indicates string


'FileFormat = TypeIndicatorNumDims;NR;NC,lengths;contents


'assumes lower bounds of 1. Write code would be in Julia but here write in VBA to test read speed


Sub speedtest()
          Dim Data
          Const FileName1 = "c:\Temp\test.csv"
          Const FileName2 = "c:\Temp\test.mff"
          Dim Res1 As Variant
          Dim Res2 As Variant
          Dim t1 As Double, t2 As Double, t3 As Double

          'Data = sreshape(sarraystack("xxx", 2, True, False), 1000, 1000) '2.13 times faster
          'Data = sReshape(CVErr(xlErrNA), 1000, 1000) ' 1.4 times faster
          'Data = sreshape(True, 1000, 1000) '2.19 times faster
          'Data = sreshape(False, 1000, 1000) '2 times faster
          'Data = sreshape(Date, 1000, 1000) '4 times faster
          'Data = sreshape("xxx", 1000, 1000) ' 2 times faster
          'Data = sreshape(CDate("2021-11-03 18:00:00"), 1000, 1000) '3.5 times faster
          'Data = Application.WorksheetFunction.Sequence(1000, 1000) '1.8 times faster
1         'Data = Application.WorksheetFunction.RandArray(1000, 100) '1.4 times faster

2         Application.Run "sCSVWrite", Data, FileName1, True
3         EncodeAndWriteFile Data, FileName2

4         t1 = ElapsedTime()
5         Res1 = Application.Run("sCSVRead", FileName1, "NDBE", ",", , "ISO")
6         t2 = ElapsedTime()
7         Res2 = ReadFileAndDecode(FileName2)
8         t3 = ElapsedTime()

9         Debug.Print Application.Run("sArraysNearlyIdentical", Res1, Res2)
10        Debug.Print t2 - t1, t3 - t2, (t2 - t1) / (t3 - t2)

End Sub

Function Decode(Chars As String)

1         On Error GoTo ErrHandler
2         Select Case Left$(Chars, 1)
              Case "#"    'vbDouble
3                 Decode = CDbl(Mid$(Chars, 2))
4             Case "£"    'vbString
5                 Decode = Mid$(Chars, 2)
6             Case "T" 'Boolean True
7                 Decode = True
8             Case "F" 'Boolean False
9                 Decode = False
10            Case "D"    'vbDate
11                Decode = CDate(Mid$(Chars, 2))
12            Case "E" 'vbEmpty
13                Decode = Empty
14            Case "N" 'vbNull
15                Decode = Null
16            Case "%" 'vbInteger
17                Decode = CInt(Mid$(Chars, 2))
18            Case "&" 'vbLong
19                Decode = CLng(Mid$(Chars, 2))
20            Case "S"    'vbSingle
21                Decode = CSng(Mid$(Chars, 2))
22            Case "C"    'vbCurrency
23                Decode = CCur(Mid$(Chars, 2))
24            Case "!"    'vbError
25                Decode = CVErr(Mid$(Chars, 2))
26            Case "@"    'vbDecimal
27                Decode = CDec(Mid$(Chars, 2))
28            Case "L"    'vbLongLong
29                Decode = CLngLng(Mid$(Chars, 2))
30            Case "*" ' vbArray
              
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
42                            Ret(i) = Decode(Mid$(Chars, k, thislength))
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
57                                Ret(i, j) = Decode(Mid$(Chars, k, thislength))
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
26            Case vbLongLong
27                Encode = "L" & CStr(x)
28            Case Is >= vbArray
29                Select Case NumDimensions(x)
                      Case 1
30                        ReDim lengthsArray(LBound(x) To UBound(x))
31                        ReDim contentsArray(LBound(x) To UBound(x))
32                        For i = LBound(x) To UBound(x)
33                            contentsArray(i) = Encode(x(i))
34                            lengthsArray(i) = CStr(Len(contentsArray(i)))
35                        Next i
36                        Encode = "*1," & CStr(UBound(x) - LBound(x) + 1) & ";" & VBA.Join(lengthsArray, ",") & ",;" & VBA.Join(contentsArray, "")
37                    Case 2
38                        NR = UBound(x, 1) - LBound(x, 1) + 1
39                        NC = UBound(x, 2) - LBound(x, 2) + 1
40                        k = 0
41                        ReDim lengthsArray(NR * NC)
42                        ReDim contentsArray(NR * NC)
43                        For j = LBound(x, 2) To UBound(x, 2)
44                            For i = LBound(x, 1) To UBound(x, 1)
45                                k = k + 1
46                                contentsArray(k) = Encode(x(i, j))
47                                lengthsArray(k) = CStr(Len(contentsArray(k)))
48                            Next i
49                        Next j
50                        Encode = "*2," & CStr(UBound(x, 1) - LBound(x, 1) + 1) & "," & CStr(UBound(x, 2) - LBound(x, 2) + 1) & ";" & VBA.Join(lengthsArray, ",") & ",;" & VBA.Join(contentsArray, "")
51                    Case Else
52                        Throw "Cannot encode array with " + CStr(NumDimensions(x)) + " dimensions"
53                End Select
54            Case Else
55                Throw "Cannot encode variable of type " & TypeName(x)
56        End Select

57        Exit Function
ErrHandler:
58        Throw "#Encode (line " & CStr(Erl) + "): " & Err.Description & "!"
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

