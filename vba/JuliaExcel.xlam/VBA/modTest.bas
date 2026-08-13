Attribute VB_Name = "modTest"
' Copyright (c) 2021-2026 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme

Option Explicit

Function TestExitAndRelaunch()

1         On Error GoTo ErrHandler
2         JuliaEval "exit()"
3         ThrowIfError JuliaLaunch()

4         Exit Function
ErrHandler:
5         TestExitAndRelaunch = ReThrow("TestExitAndRelaunch", Err, True)
6     Debug.Print TestExitAndRelaunch
End Function


' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RunTests
' Purpose    : Test JuliaCall for a variety of data types. For each data type we check that x is identical to
'              JuliaCall("identity", x). Prints results to Immediate window and to a MsgBox. Assigned to button
'              "Run Tests!" on worksheet Audit.
' -----------------------------------------------------------------------------------------------------------------------
Function RunTests(Optional SilentMode = False)

          Const Title = "JuliaExcel RunTests"
          Dim NumFailed As Long
          Dim NumPassed As Long
          Dim Prompt As String
          
1         On Error GoTo ErrHandler

2         JuliaEval "exit()"
3         PreciseSleep 1000 'Give time to shut down properly, otherwise launch can fail thinking Julia still running but unresponsive
4         ThrowIfError JuliaLaunch()

5         PrintTwice vbLf & String(80, "=")
6         PrintTwice "JuliaExcel RunTests"
7         PrintTwice Format(Now, "yyyy-mm-dd hh:mm:ss")
8         PrintTwice "ComputerName = " & Environ("ComputerName")

9         AccResult "TestEmpty", TestEmpty, NumPassed, NumFailed
10        AccResult "TestBoolean", TestBoolean, NumPassed, NumFailed
11        AccResult "TestDouble", TestDouble, NumPassed, NumFailed
12        AccResult "TestString", TestString, NumPassed, NumFailed
13        AccResult "TestWideString", TestWideString, NumPassed, NumFailed
14        AccResult "TestLong", TestLong, NumPassed, NumFailed
15        AccResult "TestLongLong", TestLongLong, NumPassed, NumFailed
16        AccResult "TestSingle", TestSingle, NumPassed, NumFailed
17        AccResult "TestDate", TestDate, NumPassed, NumFailed
18        AccResult "TestDateTime", TestDateTime, NumPassed, NumFailed
19        AccResult "Test1DArrayOfDoubles", Test1DArrayOfDoubles, NumPassed, NumFailed
20        AccResult "Test2DArrayOfMixedType", Test2DArrayOfMixedType, NumPassed, NumFailed
21        AccResult "Test3DArray", Test3DArray, NumPassed, NumFailed
22        AccResult "Test4DArray", Test4DArray, NumPassed, NumFailed
23        AccResult "TestDictionary", TestDictionary, NumPassed, NumFailed
24        AccResult "TestExactRoundTripping", TestExactRoundTripping, NumPassed, NumFailed
25        AccResult "TestArrayOfDictionaries", TestArrayOfDictionaries, NumPassed, NumFailed
26        AccResult "TestDictionaryOfArrays", TestDictionaryOfArrays, NumPassed, NumFailed
27        AccResult "TestDictionaryOfTypes", TestDictionaryOfTypes, NumPassed, NumFailed
28        AccResult "TestOneDArraysDisplayAsOneColumnOnSheet", TestOneDArraysDisplayAsOneColumnOnSheet, NumPassed, NumFailed
29        AccResult "TestElType", TestElType, NumPassed, NumFailed
30        AccResult "TestBroadcasting", TestBroadcasting, NumPassed, NumFailed

31        Prompt = NumPassed & " test(s) passed" & vbLf & _
              NumFailed & " test(s) failed"

32        If NumFailed > 0 Then
33            Prompt = Prompt & vbLf & vbLf & _
                  "See VBA Immediate window for details"
34        End If

35        PrintTwice NumPassed & " test(s) passed"
36        PrintTwice NumFailed & " test(s) failed"
37        PrintTwice String(80, "=")

38        If Not SilentMode Then
39            MsgBox Prompt, IIf(NumFailed = 0, vbInformation, vbCritical), Title
40        End If

41        RunTests = NumFailed = 0

42        Exit Function
ErrHandler:
43        If Not SilentMode Then
44            MsgBox ReThrow("RunTests", Err, True), vbCritical, Title
45        End If
46        RunTests = False
End Function

Sub PrintTwice(Text As String)

1         On Error GoTo ErrHandler
2         ThrowIfError JuliaEval("println(" & MakeJuliaLiteral(Text) & ")")
3         Debug.Print Text

4         Exit Sub
ErrHandler:
5         ReThrow "PrintTwice", Err
End Sub

Sub AccResult(TestName As String, Result As Boolean, ByRef NumPassed, ByRef NumFailed)
1         On Error GoTo ErrHandler
2         PrintTwice "Test " & TestName & " completed"
3         If Result Then
4             NumPassed = NumPassed + 1
5         Else
6             PrintTwice "Test " & TestName & " Failed!"
7             NumFailed = NumFailed + 1
8         End If
9         Exit Sub
ErrHandler:
10        ReThrow "AccResult", Err
End Sub

Function TestEmpty()
1         On Error GoTo ErrHandler
3         TestEmpty = IsEmpty(JuliaCall("identity", Empty))
4         Exit Function
ErrHandler:
5         PrintTwice ReThrow("TestEmpty", Err, True)
6         TestEmpty = False
End Function

Function TestBoolean()
1         On Error GoTo ErrHandler
2         TestBoolean = (JuliaCall("identity", True) = True) And (JuliaCall("identity", False) = False)
3         Exit Function
ErrHandler:
4         PrintTwice ReThrow("TestBoolean", Err, True)
5         TestBoolean = False
End Function

Function TestDouble()
          Dim x As Double
          Dim y As Variant
1         On Error GoTo ErrHandler
2         x = Application.WorksheetFunction.Pi()
3         y = ThrowIfError(JuliaCall("identity", x))
4         TestDouble = (x = y) And (VarType(y) = vbDouble)
5         Exit Function
ErrHandler:
6         PrintTwice ReThrow("TestDouble", Err, True)
7         TestDouble = False
End Function

Function TestString()
          Dim x As String
          Dim y As Variant
1         On Error GoTo ErrHandler
2         x = "FooBar"
3         y = ThrowIfError(JuliaCall("identity", x))
4         TestString = x = y

5         Exit Function
ErrHandler:
6         PrintTwice ReThrow("TestString", Err, True)
7         TestString = False
End Function

Function TestWideString()
          Dim i As Long
          Dim x As String
          Dim y As Variant
1         On Error GoTo ErrHandler
2         x = String(10000, " ")
3         For i = 1 To 1000
4             Mid$(x, i, 1) = ChrW(i)
5         Next i
6         y = ThrowIfError(JuliaCall("identity", x))
7         TestWideString = (x = y) And VarType(y) = vbString

8         Exit Function
ErrHandler:
9         PrintTwice ReThrow("TestWideString", Err, True)
10        TestWideString = False
End Function

Function TestLong()
          Dim x As Long
          Dim y As Variant
1         On Error GoTo ErrHandler
2         x = 123456789
3         y = ThrowIfError(JuliaCall("identity", x))
4         TestLong = (x = y) And VarType(y) = vbLong

5         Exit Function
ErrHandler:
6         PrintTwice ReThrow("TestLong", Err, True)
7         TestLong = False
End Function

Function TestLongLong()
          Dim x As LongLong
          Dim y As Variant
1         On Error GoTo ErrHandler
2         x = 123456789^
3         y = ThrowIfError(JuliaCall("identity", x))
4         TestLongLong = (x = y) And VarType(y) = vbLongLong

5         Exit Function
ErrHandler:
6         PrintTwice ReThrow("TestLongLong", Err, True)
7         TestLongLong = False
End Function

Function TestSingle()
          Dim x As Single
          Dim y As Variant
1         On Error GoTo ErrHandler
2         x = CSng(1 / 3)
3         y = ThrowIfError(JuliaCall("identity", x))
4         TestSingle = (x = y) And (VarType(y) = vbSingle)

5         Exit Function
ErrHandler:
6         PrintTwice ReThrow("TestSingle", Err, True)
7         TestSingle = False
End Function

Function TestDate()
          Dim x As Date
          Dim y As Variant
1         On Error GoTo ErrHandler
2         x = DateSerial(2025, 12, 22)
3         y = ThrowIfError(JuliaCall("identity", x))
4         TestDate = (x = y) And VarType(y) = vbDate

5         Exit Function
ErrHandler:
6         PrintTwice ReThrow("TestDate", Err, True)
7         TestDate = False
End Function

Function TestDateTime()
          Dim x As Date
          Dim y As Variant
1         On Error GoTo ErrHandler
2         x = DateSerial(2025, 12, 22) + TimeValue("03:40:33")
3         y = ThrowIfError(JuliaCall("identity", x))
4         TestDateTime = x = y

5         Exit Function
ErrHandler:
6         PrintTwice ReThrow("TestDateTime", Err, True)
7         TestDateTime = False
End Function

Function Test1DArrayOfDoubles()
          Dim x() As Double
          Dim y As Variant

1         On Error GoTo ErrHandler
2         ReDim x(1 To 3)
3         x(1) = 1 / 3
4         x(2) = 1E+100
5         x(3) = 0
6         y = JuliaCallVBA("identity", x)
7         Test1DArrayOfDoubles = ArraysIdentical(x, y)

8         Exit Function
ErrHandler:
9         PrintTwice ReThrow("Test1DArrayOfDoubles", Err, True)
10        Test1DArrayOfDoubles = False
End Function

Function Test2DArrayOfMixedType()
          Dim x() As Variant
          Dim y As Variant

1         On Error GoTo ErrHandler
2         ReDim x(1 To 3, 1 To 3)
3         x(1, 1) = 1 / 3:  x(1, 2) = 1E+100:    x(1, 3) = CLng(100)
4         x(2, 1) = "Foo":  x(2, 2) = CSng(3):   x(2, 3) = CLngLng(100)
5         x(3, 1) = "Foo":  x(3, 2) = CSng(3):   x(3, 3) = CInt(100)

6         y = ThrowIfError(JuliaCall("identity", x))
7         Test2DArrayOfMixedType = ArraysIdentical(x, y)

8         Exit Function
ErrHandler:
9         PrintTwice ReThrow("Test2DArrayOfMixedType", Err, True)
10        Test2DArrayOfMixedType = False
End Function

Function Test3DArray()
          Dim i As Long
          Dim x() As Variant
          Dim y As Variant

1         On Error GoTo ErrHandler
2         ReDim x(1 To 2, 1 To 3, 1 To 4)
3         For i = 1 To 24
4             SetAtLinear x, i, ChrW(i)
5         Next

6         y = JuliaCallVBA("identity", x)
7         Test3DArray = ArraysIdentical(x, y)

8         Exit Function
ErrHandler:
9         PrintTwice ReThrow("Test3DArray", Err, True)
10        Test3DArray = False
End Function

Function Test4DArray()
          Dim i As Long
          Dim x() As Variant
          Dim y As Variant

1         On Error GoTo ErrHandler
2         ReDim x(1 To 2, 1 To 3, 1 To 4, 1 To 5)

3         For i = 1 To 120
4             SetAtLinear x, i, i
5         Next

6         y = JuliaCallVBA("identity", x)
7         Test4DArray = ArraysIdentical(x, y)

8         Exit Function
ErrHandler:
9         PrintTwice ReThrow("Test4DArray", Err, True)
10        Test4DArray = False
End Function

Function TestDictionary()

          Dim x As New Scripting.Dictionary
          Dim y As Scripting.Dictionary
          Dim z As New Scripting.Dictionary
            
1         On Error GoTo ErrHandler
2         z.Add "alpha", 100
3         z.Add "beta", 200

4         x.Add "a", 1
5         x.Add "b", 2
6         x.Add "c", "d"
7         x.Add "d", Array(1, 2, 3)
8         x.Add "e", z

9         Set y = JuliaCallVBA("identity", x)

10        ThrowIfError JuliaSetVar("first_dictionary", x)
11        ThrowIfError JuliaSetVar("second_dictionary", y)

12        TestDictionary = JuliaEval("first_dictionary == second_dictionary")

13        Exit Function
ErrHandler:
14        PrintTwice ReThrow("TestDictionary", Err, True)
15        TestDictionary = False
End Function

Function TestExactRoundTripping()
          Dim i As Long
          Dim x() As Double
          Dim y As Variant

1         On Error GoTo ErrHandler
2         ReDim x(1 To 1000)
3         For i = 1 To 1000
4             x(i) = Sqr(i)
5         Next i
6         y = JuliaCallVBA("identity", x)
7         TestExactRoundTripping = ArraysIdentical(x, y)

8         Exit Function
ErrHandler:
9         PrintTwice ReThrow("TestExactRoundTripping", Err, True)
10        TestExactRoundTripping = False
End Function

Function TestArrayOfDictionaries()
          Dim x() As Variant
          Dim y As Variant
          Dim z As New Scripting.Dictionary

1         On Error GoTo ErrHandler
2         z("a") = 1
3         z("b") = 2

4         ReDim x(1 To 2, 1 To 2)

5         Set x(1, 1) = z: Set x(1, 2) = z
6         Set x(2, 1) = z: Set x(2, 2) = z
7         y = JuliaCallVBA("identity", x)
8         ThrowIfError JuliaSetVar("first_array_of_dictionaries", x)
9         ThrowIfError JuliaSetVar("second_array_of_dictionaries", y)

10        TestArrayOfDictionaries = JuliaEval("first_array_of_dictionaries == second_array_of_dictionaries")

11        Exit Function
ErrHandler:
12        PrintTwice ReThrow("TestArrayOfDictionaries", Err, True)
13        TestArrayOfDictionaries = False
End Function

Function TestDictionaryOfArrays()

          Dim x As New Scripting.Dictionary
          Dim y As Scripting.Dictionary

1         On Error GoTo ErrHandler

2         x("doubles") = Array(1#, 2#, 3#)
3         x("strings") = Array("foo", "bar")

4         Set y = JuliaCallVBA("identity", x)

5         ThrowIfError JuliaSetVar("first_dict_of_arrays", x)
6         ThrowIfError JuliaSetVar("second_dict_of_arrays", y)

7         TestDictionaryOfArrays = JuliaEval("first_dict_of_arrays == second_dict_of_arrays")

8         Exit Function
ErrHandler:
9         PrintTwice ReThrow("TestDictionaryOfArrays", Err, True)
10        TestDictionaryOfArrays = False
End Function

Function TestDictionaryOfTypes()

          Dim x As New Scripting.Dictionary
          Dim y As Scripting.Dictionary
            
1         On Error GoTo ErrHandler

2         x.Add "Integer", 1
3         x.Add "Long", CLng(1)
4         x.Add "LongLong", CLngLng(1)
5         x.Add "Single", CSng(1)
6         x.Add "Double", CDbl(1)

7         Set y = JuliaCallVBA("identity", x)

8         TestDictionaryOfTypes = True
          Dim k As Variant
9         For Each k In y.Keys
10            If TypeName(y(k)) <> k Then
11                TestDictionaryOfTypes = False
12            End If
13        Next

14        Exit Function
ErrHandler:
15        PrintTwice ReThrow("TestDictionary", Err, True)
16        TestDictionaryOfTypes = False
End Function

Function TestOneDArraysDisplayAsOneColumnOnSheet()

          Dim OneDArray() As Variant

1         On Error GoTo ErrHandler
2         ReDim OneDArray(1 To 3)
3         OneDArray(1) = 1
4         OneDArray(2) = 2
5         OneDArray(3) = 3

6         TestOneDArraysDisplayAsOneColumnOnSheet = _
              NumDimensions(JuliaCall("identity", OneDArray)) = 2 And _
              NumDimensions(JuliaCallVBA("identity", OneDArray)) = 1

7         Exit Function
ErrHandler:
8         PrintTwice ReThrow("TestOneDArraysDisplayAsOneColumnOnSheet", Err, True)
9         TestOneDArraysDisplayAsOneColumnOnSheet = False
End Function

Function TestElType()

1         On Error GoTo ErrHandler

2         TestElType = _
              JuliaCall("eltype", Array(1, 2, 3)) = "Int16" And _
              JuliaCall("eltype", Array(1!, 2!, 3!)) = "Float32" And _
              JuliaCall("eltype", Array(1#, 2#, 3#)) = "Float64" And _
              JuliaCall("eltype", Array(1, 2#, 3#)) = "Any" And _
              JuliaCall("eltype", Array("a", "b", "c")) = "String" And _
              JuliaCall("eltype", Array("a", 1, True)) = "Any" And _
              JuliaCall("eltype", Array(True, False, True)) = "Bool" And _
              JuliaCall("eltype", Array(Empty, Empty, Empty)) = "Missing"

3         Exit Function
ErrHandler:
4         PrintTwice ReThrow("TestElType", Err, True)
5         TestElType = False
End Function

Function TestBroadcasting()

          Dim ExpRes(1 To 2, 1 To 2) As Variant
          Dim ObsRes As Variant
          Dim Xs(1 To 1, 1 To 2) As Variant
          Dim Ys(1 To 2, 1 To 1) As Variant

1         On Error GoTo ErrHandler

2         Xs(1, 1) = 7#: Xs(1, 2) = 5#
3         Ys(1, 1) = 2#: Ys(2, 1) = 3#

4         ExpRes(1, 1) = 20#: ExpRes(1, 2) = 16#
5         ExpRes(2, 1) = 23#: ExpRes(2, 2) = 19#

6         ThrowIfError JuliaEval("fn(x,y) = 2x + 3y")
          'Note the dot character below, makes this a broadcast use of fn
7         ObsRes = JuliaCall("fn.", Xs, Ys)

8         TestBroadcasting = ArraysIdentical(ExpRes, ObsRes)
          
9         Exit Function
ErrHandler:
10        PrintTwice ReThrow("TestBroadcasting", Err, True)
11        TestBroadcasting = False
End Function

