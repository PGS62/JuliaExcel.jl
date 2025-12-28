Attribute VB_Name = "modTest"
' Copyright (c) 2021-2025 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme

Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RunTests
' Purpose    : Test JuliaCall for a variety of data types. For each data type we check that x is identical to
'              JuliaCall("identity", x). Prints results to Immediate window and to a MsgBox. Assigned to button
'              "Run Tests!" on worksheet Audit.
' -----------------------------------------------------------------------------------------------------------------------
Sub RunTests()

          Const Title = "JuliaExcel RunTests"
          Dim NumFailed As Long
          Dim NumPassed As Long
          Dim Prompt As String
          
1         On Error GoTo ErrHandler

2         Debug.Print String(80, "=")
3         Debug.Print "JuliaExcel RunTests"
4         Debug.Print Now()
5         Debug.Print "ComputerName = " & Environ("ComputerName")
          
6         ThrowIfError JuliaLaunch()
7         AccResult "TestEmpty", TestEmpty, NumPassed, NumFailed
8         AccResult "TestBoolean", TestBoolean, NumPassed, NumFailed
9         AccResult "TestDouble", TestDouble, NumPassed, NumFailed
10        AccResult "TestString", TestString, NumPassed, NumFailed
11        AccResult "TestWideString", TestWideString, NumPassed, NumFailed
12        AccResult "TestLong", TestLong, NumPassed, NumFailed
13        AccResult "TestLongLong", TestLongLong, NumPassed, NumFailed
14        AccResult "TestSingle", TestSingle, NumPassed, NumFailed
15        AccResult "TestDate", TestDate, NumPassed, NumFailed
16        AccResult "TestDateTime", TestDateTime, NumPassed, NumFailed
17        AccResult "Test1DArrayOfDoubles", Test1DArrayOfDoubles, NumPassed, NumFailed
18        AccResult "Test2DArrayOfMixedType", Test2DArrayOfMixedType, NumPassed, NumFailed
19        AccResult "Test3DArray", Test3DArray, NumPassed, NumFailed
20        AccResult "Test4DArray", Test4DArray, NumPassed, NumFailed
21        AccResult "TestDictionary", TestDictionary, NumPassed, NumFailed
22        AccResult "TestExactRoundTripping", TestExactRoundTripping, NumPassed, NumFailed
23        AccResult "TestArrayOfDictionaries", TestArrayOfDictionaries, NumPassed, NumFailed
24        AccResult "TestDictionaryOfTypes", TestDictionaryOfTypes, NumPassed, NumFailed

25        Prompt = NumPassed & " test(s) passed" & vbLf & _
              NumFailed & " test(s) failed"

26        If NumFailed > 0 Then
27            Prompt = Prompt & vbLf & vbLf & _
                  "See VBA Immediate window for details"
28        End If

29        Debug.Print NumPassed & " test(s) passed"
30        Debug.Print NumFailed & " test(s) failed"
31        Debug.Print String(80, "=")

32        MsgBox Prompt, IIf(NumFailed = 0, vbInformation, vbCritical), Title

33        Exit Sub
ErrHandler:
34        MsgBox ReThrow("RunTests", Err, True), vbCritical, Title
End Sub

Sub AccResult(TestName As String, Result As Boolean, ByRef NumPassed, ByRef NumFailed)
1         On Error GoTo ErrHandler
2         If Result Then
3             NumPassed = NumPassed + 1
4         Else
5             Debug.Print "Test " & TestName & " Failed!"
6             NumFailed = NumFailed + 1
7         End If
8         Exit Sub
ErrHandler:
9         ReThrow "AccResult", Err
End Sub

Function TestEmpty()
1         On Error GoTo ErrHandler
2         TestEmpty = IsEmpty(JuliaCall("identity", Empty))
3         Exit Function
ErrHandler:
4         Debug.Print ReThrow("TestEmpty", Err, True)
5         TestEmpty = False
End Function

Function TestBoolean()
1         On Error GoTo ErrHandler
2         TestBoolean = (JuliaCall("identity", True) = True) And (JuliaCall("identity", False) = False)
3         Exit Function
ErrHandler:
4         Debug.Print ReThrow("TestBoolean", Err, True)
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
6         Debug.Print ReThrow("TestDouble", Err, True)
TestDouble = False
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
6         Debug.Print ReThrow("TestString", Err, True)
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
9         Debug.Print ReThrow("TestWideString", Err, True)
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
6         Debug.Print ReThrow("TestLong", Err, True)
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
6         Debug.Print ReThrow("TestLongLong", Err, True)
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
6         Debug.Print ReThrow("TestSingle", Err, True)
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
6         Debug.Print ReThrow("TestDate", Err, True)
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
6         Debug.Print ReThrow("TestDateTime", Err, True)
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
9         Debug.Print ReThrow("Test1DArrayOfDoubles", Err, True)
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
9         Debug.Print ReThrow("Test2DArrayOfMixedType", Err, True)
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
9         Debug.Print ReThrow("Test3DArray", Err, True)
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
9         Debug.Print ReThrow("Test4DArray", Err, True)
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
14        Debug.Print ReThrow("TestDictionary", Err, True)
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
9         Debug.Print ReThrow("TestExactRoundTripping", Err, True)
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
12        Debug.Print ReThrow("TestArrayOfDictionaries", Err, True)
13        TestArrayOfDictionaries = False
End Function

Function TestDictionaryOfTypes()

          Dim x As New Scripting.Dictionary
          Dim y As Scripting.Dictionary
          Dim z As New Scripting.Dictionary
            
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
15        Debug.Print ReThrow("TestDictionary", Err, True)
16        TestDictionaryOfTypes = False
End Function

