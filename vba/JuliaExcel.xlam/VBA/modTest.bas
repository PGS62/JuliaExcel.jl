Attribute VB_Name = "modTest"
Option Explicit

Sub RunTests()
          Dim NumPassed As Long
          Dim NumFailed As Long
          
1         Debug.Print String(80, "=")
2         Debug.Print "JuliaExcel RunTests"
3         Debug.Print Now()
4         Debug.Print "ComputerName = " & Environ("ComputerName")
          
5         AccResult "TestEmpty", TestEmpty, NumPassed, NumFailed
6         AccResult "TestBoolean", TestBoolean, NumPassed, NumFailed
7         AccResult "TestDouble", TestDouble, NumPassed, NumFailed
8         AccResult "TestString", TestString, NumPassed, NumFailed
9         AccResult "TestWideString", TestWideString, NumPassed, NumFailed
10        AccResult "TestLong", TestLong, NumPassed, NumFailed
11        AccResult "TestSingle", TestSingle, NumPassed, NumFailed
12        AccResult "TestDate", TestDate, NumPassed, NumFailed
13        AccResult "TestDateTime", TestDateTime, NumPassed, NumFailed
14        AccResult "Test1DArrayOfDoubles", Test1DArrayOfDoubles, NumPassed, NumFailed
15        AccResult "Test2DArrayOfMixedType", Test2DArrayOfMixedType, NumPassed, NumFailed
          
          
16        Debug.Print NumPassed & " test(s) passed"
17        Debug.Print NumFailed & " test(s) failed"
18        Debug.Print String(80, "=")
          
End Sub

Sub AccResult(TestName As String, Result As Boolean, ByRef NumPassed, ByRef NumFailed)
1         If Result Then
2             NumPassed = NumPassed + 1
3         Else
4             Debug.Print "Test " & TestName & " Failed!"
5             NumFailed = NumFailed + 1
6         End If
End Sub

Function TestEmpty()
1         TestEmpty = IsEmpty(JuliaCall("identity", Empty))
End Function

Function TestBoolean()
1         TestBoolean = (JuliaCall("identity", True) = True) And (JuliaCall("identity", False) = False)
End Function

Function TestDouble()
          Dim x As Double
          Dim y As Variant
1         x = Application.WorksheetFunction.Pi()
2         y = JuliaCall("identity", x)
3         TestDouble = x = y
End Function

Function TestString()
          Dim x As String
          Dim y As Variant
1         x = "FooBar"
2         y = JuliaCall("identity", x)
3         TestString = x = y
End Function

Function TestWideString()
          Dim x As String
          Dim i As Long
          Dim y As Variant
1         x = String(10000, " ")
2         For i = 1 To 1000
3             Mid$(x, i, 1) = ChrW(i)
4         Next i
5         y = JuliaCall("identity", x)
6         TestWideString = x = y
End Function

Function TestLong()
          Dim x As Long
          Dim y As Variant
1         x = 123456789
2         y = JuliaCall("identity", x)
3         TestLong = x = y
End Function

Function TestSingle()
          Dim x As Single
          Dim y As Variant
1         x = CSng(1 / 3)
2         y = JuliaCall("identity", x)
3         TestSingle = x = y
End Function

Function TestDate()
          Dim x As Date
          Dim y As Variant
1         x = DateSerial(2025, 12, 22)
2         y = JuliaCall("identity", x)
3         TestDate = x = y
End Function
Function TestDateTime()
          Dim x As Date
          Dim y As Variant
1         x = DateSerial(2025, 12, 22) + TimeValue("03:40:33")
2         y = JuliaCall("identity", x)
3         TestDateTime = x = y
End Function

Function Test1DArrayOfDoubles()
          Dim x() As Double
          Dim y As Variant

1         ReDim x(1 To 3)
2         x(1) = 1 / 3
3         x(2) = 1E+100
4         x(3) = 0
5         y = JuliaCallVBA("identity", x)
6         Test1DArrayOfDoubles = ArraysIdentical(x, y)

End Function

Function Test2DArrayOfMixedType()
          Dim x() As Variant
          Dim y As Variant

1         ReDim x(1 To 3, 1 To 3)
2         x(1, 1) = 1 / 3:  x(1, 2) = 1E+100:    x(1, 3) = 0
3         x(2, 1) = "Foo":  x(2, 2) = CSng(3):   x(2, 3) = 0
4         x(3, 1) = "Foo":  x(3, 2) = CSng(3):   x(3, 3) = 0

5         y = JuliaCall("identity", x)
6         Test2DArrayOfMixedType = ArraysIdentical(x, y)

End Function


Function Test3DArray()
          Dim x() As Variant
          Dim y As Variant
          Dim i As Long

1         ReDim x(1 To 2, 1 To 3, 1 To 4)
2         For i = 1 To 24
3             ArraySetLinear x, i, String(i, "x")
4         Next

10        y = JuliaCallVBA("identity", x)
11        Test3DArray = ArraysIdentical(x, y)

End Function


Function Test4DArray()
          Dim x() As Variant
          Dim y As Variant
          Dim i As Long

1         ReDim x(1 To 2, 1 To 3, 1 To 4, 1 To 5)

2         For i = 1 To 120
3             ArraySetLinear x, i, i
4         Next

5         y = JuliaCallVBA("identity", x)
6         Test4DArray = ArraysIdentical(x, y)


End Function







