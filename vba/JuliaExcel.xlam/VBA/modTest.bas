Attribute VB_Name = "modTest"
' Copyright (c) 2021-2026 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme

Option Explicit
Option Private Module

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
4         ThrowIfError JuliaLaunch(, , gTestCommandOptions)

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
31        AccResult "TestNaNFallback", TestNaNFallback, NumPassed, NumFailed
32        AccResult "TestInfFallback", TestInfFallback, NumPassed, NumFailed
33        AccResult "TestVFormatVectorColumn", TestVFormatVectorColumn, NumPassed, NumFailed
34        AccResult "TestVFormatMatrixNonSquare", TestVFormatMatrixNonSquare, NumPassed, NumFailed
35        AccResult "TestMatrixNaNFallback", TestMatrixNaNFallback, NumPassed, NumFailed
36        AccResult "TestVFormatEmptyArrayFallback", TestVFormatEmptyArrayFallback, NumPassed, NumFailed
37        AccResult "TestSerialiseArrayAsV", TestSerialiseArrayAsV, NumPassed, NumFailed
38        AccResult "TestVEncodeFromVariantArray", TestVEncodeFromVariantArray, NumPassed, NumFailed
39        AccResult "TestVFormat3DArray", TestVFormat3DArray, NumPassed, NumFailed
40        AccResult "TestRangeIntegerVsCollect", TestRangeIntegerVsCollect, NumPassed, NumFailed
41        AccResult "TestRangeStepVsCollect", TestRangeStepVsCollect, NumPassed, NumFailed
42        AccResult "TestRangeFloatVsCollect", TestRangeFloatVsCollect, NumPassed, NumFailed
43        AccResult "TestRangeJuliaEvalVBAIsOneD", TestRangeJuliaEvalVBAIsOneD, NumPassed, NumFailed
44        AccResult "TestRangeJuliaEvalIsColumn", TestRangeJuliaEvalIsColumn, NumPassed, NumFailed
45        AccResult "TestRangeWireFormat", TestRangeWireFormat, NumPassed, NumFailed
46        AccResult "TestNaNInfRoundTripViaV", TestNaNInfRoundTripViaV, NumPassed, NumFailed
47        AccResult "TestExcelErrorRoundTrip", TestExcelErrorRoundTrip, NumPassed, NumFailed
48        AccResult "TestByte", TestByte, NumPassed, NumFailed
49        AccResult "TestVFormatBoundarySizes", TestVFormatBoundarySizes, NumPassed, NumFailed
50        AccResult "TestMakeJuliaLiteral", TestMakeJuliaLiteral, NumPassed, NumFailed
51        AccResult "TestUnserialiseStringTooLong", TestUnserialiseStringTooLong, NumPassed, NumFailed
52        AccResult "TestUnserialiseMalformedRank", TestUnserialiseMalformedRank, NumPassed, NumFailed
53        AccResult "TestConcatenateExpressions", TestConcatenateExpressions, NumPassed, NumFailed
54        AccResult "TestTrim255", TestTrim255, NumPassed, NumFailed

55        Prompt = NumPassed & " test(s) passed" & vbLf & _
              NumFailed & " test(s) failed"

56        If NumFailed > 0 Then
57            Prompt = Prompt & vbLf & vbLf & _
                  "See VBA Immediate window for details"
58        End If

59        PrintTwice NumPassed & " test(s) passed"
60        PrintTwice NumFailed & " test(s) failed"
61        PrintTwice String(80, "=")

62        If Not SilentMode Then
63            AppActivate Application.Caption
64            MsgBox Prompt, IIf(NumFailed = 0, vbInformation, vbCritical), Title
65        End If

66        RunTests = NumFailed = 0

67        Exit Function
ErrHandler:
68        If Not SilentMode Then
69            MsgBox ReThrow("RunTests", Err, True), vbCritical, Title
70        End If
71        RunTests = False
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

Function TestByte()
          Dim x As Byte
          Dim y As Variant
1         On Error GoTo ErrHandler
2         x = 200
3         y = ThrowIfError(JuliaCall("identity", x))
4         TestByte = (x = y) And VarType(y) = vbByte

5         Exit Function
ErrHandler:
6         PrintTwice ReThrow("TestByte", Err, True)
7         TestByte = False
End Function

' Confirms TrySerialiseArrayAsV/Unserialise's "V" format round-trips correctly across a range of
' array sizes, particularly at small boundaries (1, 2, 3) and well beyond the usual 100,000-element
' benchmark size - added 2026-08-19 alongside the bulk CopyMemory-based rewrite of both the encode
' (BulkHexOfDoubleArray, modSerialise.bas) and decode (BulkDoublesFromHex, modUnserialise.bas)
' helpers, to directly exercise the byte-count arithmetic at sizes the standard performance tests
' never varied. TrySerialiseArrayAsV/Unserialise are called directly (the real production functions,
' not the now-deleted throwaway prototype - see git history) so this tests the actual code path.
Function TestVFormatBoundarySizes() As Boolean
          Dim EncodedV As String
          Dim i As Long
          Dim Idx As Long
          Dim OK As Boolean
          Dim Sizes As Variant
          Dim Sz As Long
          Dim x() As Double
          Dim y As Variant

1         On Error GoTo ErrHandler
2         Sizes = Array(1, 2, 3, 7, 8, 9, 17, 100, 1000, 1000000)

3         For Idx = LBound(Sizes) To UBound(Sizes)
4             Sz = Sizes(Idx)
5             ReDim x(1 To Sz)
6             For i = 1 To Sz
7                 x(i) = (i - Sz / 2) * 1.5 + 0.25   ' varied: negative, positive, fractional values
8             Next i
9             OK = TrySerialiseArrayAsV(x, EncodedV)
10            If Not OK Then Throw "TrySerialiseArrayAsV declined for size " & Sz
11            y = UnserialiseFromString(EncodedV, False, GetStringLengthLimit(), False)
12            If Not ArraysIdentical(x, y) Then Throw "Round trip mismatch for size " & Sz
13        Next Idx

14        TestVFormatBoundarySizes = True
15        Exit Function
ErrHandler:
16        PrintTwice ReThrow("TestVFormatBoundarySizes", Err, True)
17        TestVFormatBoundarySizes = False
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

' Checks that a Vector{Float64} containing a NaN falls back from the compact "V" wire format to
' the general "*" format (see encode_for_xl(::Vector{Float64}) in src/encode.jl), so that the NaN
' element arrives as the Excel error #N/A rather than corrupting the fast binary decode.
Function TestNaNFallback()
          Dim Expected(1 To 3, 1 To 1) As Variant
          Dim y As Variant

1         On Error GoTo ErrHandler
2         Expected(1, 1) = 1#
3         Expected(2, 1) = 2#
4         Expected(3, 1) = CVErr(2042) '#N/A
5         y = ThrowIfError(JuliaEval("[1.0,2.0,NaN]"))
6         TestNaNFallback = ArraysIdentical(Expected, y)

7         Exit Function
ErrHandler:
8         PrintTwice ReThrow("TestNaNFallback", Err, True)
9         TestNaNFallback = False
End Function

' As TestNaNFallback, but for Inf, which falls back to the general format because it also can't be
' round-tripped through the compact "V" format's plain hex-encoded Double representation.
Function TestInfFallback()
          Dim Expected(1 To 3, 1 To 1) As Variant
          Dim y As Variant

1         On Error GoTo ErrHandler
2         Expected(1, 1) = 1#
3         Expected(2, 1) = 2#
4         Expected(3, 1) = CVErr(2036) '#NUM!
5         y = ThrowIfError(JuliaEval("[1.0,2.0,Inf]"))
6         TestInfFallback = ArraysIdentical(Expected, y)

7         Exit Function
ErrHandler:
8         PrintTwice ReThrow("TestInfFallback", Err, True)
9         TestInfFallback = False
End Function

' Exercises the Case 86 'V' branch's "JuliaVectorToXLColumn = True" path (modUnserialise.bas), used
' by JuliaCall/JuliaEval but not by JuliaCallVBA - Test1DArrayOfDoubles/TestExactRoundTripping only
' cover the False (native VBA 1-D array) path, via JuliaCallVBA.
Function TestVFormatVectorColumn()
          Dim Expected(1 To 4, 1 To 1) As Variant
          Dim x() As Double
          Dim y As Variant

1         On Error GoTo ErrHandler
2         ReDim x(1 To 4)
3         x(1) = 1.5: x(2) = -2.25: x(3) = 0#: x(4) = 100000.125
4         Expected(1, 1) = x(1): Expected(2, 1) = x(2)
5         Expected(3, 1) = x(3): Expected(4, 1) = x(4)
6         y = ThrowIfError(JuliaCall("identity", x))
7         TestVFormatVectorColumn = ArraysIdentical(Expected, y)

8         Exit Function
ErrHandler:
9         PrintTwice ReThrow("TestVFormatVectorColumn", Err, True)
10        TestVFormatVectorColumn = False
End Function

' Exercises the Case 86 'V' branch's rank-2 (matrix) path with a non-square shape and distinct
' values in every cell, to catch a row/column transposition bug that a square test could mask.
Function TestVFormatMatrixNonSquare()
          Dim Expected(1 To 2, 1 To 3) As Variant
          Dim x(1 To 2, 1 To 3) As Double
          Dim y As Variant

1         On Error GoTo ErrHandler
2         x(1, 1) = 1.1: x(1, 2) = 2.2: x(1, 3) = 3.3
3         x(2, 1) = 4.4: x(2, 2) = 5.5: x(2, 3) = 6.6
4         Expected(1, 1) = x(1, 1): Expected(1, 2) = x(1, 2): Expected(1, 3) = x(1, 3)
5         Expected(2, 1) = x(2, 1): Expected(2, 2) = x(2, 2): Expected(2, 3) = x(2, 3)
6         y = JuliaCallVBA("identity", x)
7         TestVFormatMatrixNonSquare = ArraysIdentical(Expected, y)

8         Exit Function
ErrHandler:
9         PrintTwice ReThrow("TestVFormatMatrixNonSquare", Err, True)
10        TestVFormatMatrixNonSquare = False
End Function

' As TestNaNFallback, but for encode_for_xl(::Matrix{Float64}) - a separate method in src/encode.jl
' with its own NaN/Inf fallback guard, not covered by the vector-only NaN/Inf tests above.
Function TestMatrixNaNFallback()
          Dim Expected(1 To 2, 1 To 3) As Variant
          Dim y As Variant

1         On Error GoTo ErrHandler
2         Expected(1, 1) = 1#: Expected(1, 2) = 2#: Expected(1, 3) = 3#
3         Expected(2, 1) = 4#
4         Expected(2, 2) = CVErr(2042) '#N/A
5         Expected(2, 3) = 6#
6         y = ThrowIfError(JuliaEval("[1.0 2.0 3.0; 4.0 NaN 6.0]"))
7         TestMatrixNaNFallback = ArraysIdentical(Expected, y)

8         Exit Function
ErrHandler:
9         PrintTwice ReThrow("TestMatrixNaNFallback", Err, True)
10        TestMatrixNaNFallback = False
End Function

' Checks the n = 0 guard in encode_for_xl(::Vector{Float64}) (src/encode.jl) correctly routes an
' empty Float64 vector to the general "*" array format rather than attempting a zero-length "V"
' payload. The general format represents a zero-element 1-D array (when AllowNesting is True, as it
' is for JuliaEvalVBA) as a zero-length array, via VBA.Split(vbNullString) in Unserialise
' (modUnserialise.bas) - not something introduced by "V".
Function TestVFormatEmptyArrayFallback()
          Dim y As Variant

1         On Error GoTo ErrHandler
2         y = ThrowIfError(JuliaEvalVBA("Float64[]"))
3         TestVFormatEmptyArrayFallback = IsArray(y) And (UBound(y) < LBound(y))

4         Exit Function
ErrHandler:
5         PrintTwice ReThrow("TestVFormatEmptyArrayFallback", Err, True)
6         TestVFormatEmptyArrayFallback = False
End Function

' Direct unit test of TrySerialiseArrayAsV (modSerialise.bas) - the Excel -> Julia direction "V"
' encoder wired into SerialiseElement. Checks the exact wire string for a vector and a (non-square,
' column-major) matrix, that it correctly declines (returns False) for a mixed-type array, and
' (since 2026-08-18 - see TrySerialiseArrayAsV's own docstring for why no NaN/Inf check is needed on
' this side) that it now SUCCEEDS for an array containing NaN/Infinity, producing the same raw
' bit-pattern encoding as any other Double. NaN/Infinity Doubles can't be produced by ordinary VBA
' arithmetic (which raises a runtime error on overflow rather than returning a special value), so
' they're constructed here via HexToDouble on the standard IEEE-754 bit patterns, the same trick
' DoubleToHex/HexToDouble themselves rely on.
Function TestSerialiseArrayAsV()
          Dim EncodedV As String
          Dim OK As Boolean
          Dim x() As Double
          Dim x2D(1 To 2, 1 To 3) As Double
          Dim xInf(1 To 2) As Double
          Dim xMixed(1 To 3) As Variant
          Dim xNaN(1 To 2) As Double

1         On Error GoTo ErrHandler

2         ReDim x(1 To 3)
3         x(1) = 1#: x(2) = -2.5: x(3) = 3.14159265358979
4         OK = TrySerialiseArrayAsV(x, EncodedV)
5         If Not OK Then Throw "Expected TrySerialiseArrayAsV to succeed for an all-Double vector"
6         If EncodedV <> "V1,3;" & DoubleToHex(x(1)) & DoubleToHex(x(2)) & DoubleToHex(x(3)) Then _
              Throw "Unexpected 'V' encoding for a Double vector"

7         x2D(1, 1) = 1: x2D(1, 2) = 2: x2D(1, 3) = 3
8         x2D(2, 1) = 4: x2D(2, 2) = 5: x2D(2, 3) = 6
9         OK = TrySerialiseArrayAsV(x2D, EncodedV)
10        If Not OK Then Throw "Expected TrySerialiseArrayAsV to succeed for an all-Double matrix"
11        If EncodedV <> "V2,2,3;" & DoubleToHex(1#) & DoubleToHex(4#) & DoubleToHex(2#) & _
              DoubleToHex(5#) & DoubleToHex(3#) & DoubleToHex(6#) Then _
              Throw "Unexpected 'V' encoding for a Double matrix (expected column-major)"

12        xMixed(1) = 1#: xMixed(2) = "foo": xMixed(3) = 3#
13        OK = TrySerialiseArrayAsV(xMixed, EncodedV)
14        If OK Then Throw "Expected TrySerialiseArrayAsV to decline a mixed-type array"

15        xNaN(1) = 1#
16        xNaN(2) = HexToDouble("7FF8000000000000") 'quiet NaN
17        OK = TrySerialiseArrayAsV(xNaN, EncodedV)
18        If Not OK Then Throw "Expected TrySerialiseArrayAsV to succeed for an array containing NaN"
19        If EncodedV <> "V1,2;" & DoubleToHex(xNaN(1)) & DoubleToHex(xNaN(2)) Then _
              Throw "Unexpected 'V' encoding for an array containing NaN"

20        xInf(1) = 1#
21        xInf(2) = HexToDouble("7FF0000000000000") 'positive infinity
22        OK = TrySerialiseArrayAsV(xInf, EncodedV)
23        If Not OK Then Throw "Expected TrySerialiseArrayAsV to succeed for an array containing Infinity"
24        If EncodedV <> "V1,2;" & DoubleToHex(xInf(1)) & DoubleToHex(xInf(2)) Then _
              Throw "Unexpected 'V' encoding for an array containing Infinity"

25        TestSerialiseArrayAsV = True

26        Exit Function
ErrHandler:
27        PrintTwice ReThrow("TestSerialiseArrayAsV", Err, True)
28        TestSerialiseArrayAsV = False
End Function

' Confirms that removing TrySerialiseArrayAsV's NaN/Inf check (2026-08-18) didn't change end-to-end
' behaviour: an array containing NaN/Infinity now travels Excel -> Julia via the fast "V" path (raw
' bit pattern, no translation) rather than falling back to the general format, but Julia's own
' outbound encode_for_xl(::Float64) still correctly maps NaN -> #N/A and Infinity -> #NUM! on the
' way back - exactly as it does for a literal Julia-side NaN/Inf (TestNaNFallback/TestInfFallback).
Function TestNaNInfRoundTripViaV()
          Dim Expected(1 To 4, 1 To 1) As Variant
          Dim x(1 To 4) As Double
          Dim y As Variant

1         On Error GoTo ErrHandler
2         x(1) = 1#
3         x(2) = HexToDouble("7FF8000000000000") 'quiet NaN
4         x(3) = 2#
5         x(4) = HexToDouble("7FF0000000000000") 'positive infinity

6         Expected(1, 1) = 1#
7         Expected(2, 1) = CVErr(2042) '#N/A
8         Expected(3, 1) = 2#
9         Expected(4, 1) = CVErr(2036) '#NUM!

10        y = ThrowIfError(JuliaCall("identity", x))
11        TestNaNInfRoundTripViaV = ArraysIdentical(Expected, y)

12        Exit Function
ErrHandler:
13        PrintTwice ReThrow("TestNaNInfRoundTripViaV", Err, True)
14        TestNaNInfRoundTripViaV = False
End Function

' Confirms Excel errors now round-trip correctly through Julia via JuliaCall("identity", ...):
' Julia decodes an incoming VBA error to the ExcelError type (JuliaExcel.jl, added 2026-08-18)
' rather than a plain String, so a function like "identity" that doesn't know about errors
' specifically passes the value straight through unchanged, and Julia's own
' encode_for_xl(::ExcelError) re-emits the same wire "!<code>" - unlike a String, which would
' encode as an ordinary text value. Exercises all 14 Excel error codes, not just the two (2036,
' 2042) Julia can generate automatically from Inf/NaN. Built with ReDim (1-based), not VBA.Array()
' (0-based), to match Unserialise's own 1-based convention for the general "*" array format - see
' TestVEncodeFromVariantArray's comment for the same trap.
Function TestExcelErrorRoundTrip()
          Dim Codes(1 To 14) As Long
          Dim Expected(1 To 14, 1 To 1) As Variant
          Dim i As Long
          Dim y As Variant

1         On Error GoTo ErrHandler
2         Codes(1) = 2000: Codes(2) = 2007: Codes(3) = 2015: Codes(4) = 2023: Codes(5) = 2029
3         Codes(6) = 2036: Codes(7) = 2042: Codes(8) = 2043: Codes(9) = 2045: Codes(10) = 2046
4         Codes(11) = 2047: Codes(12) = 2048: Codes(13) = 2049: Codes(14) = 2050

5         For i = 1 To 14
6             Expected(i, 1) = CVErr(Codes(i))
7         Next i

8         y = ThrowIfError(JuliaCall("identity", Expected))
9         TestExcelErrorRoundTrip = ArraysIdentical(Expected, y)

10        Exit Function
ErrHandler:
11        PrintTwice ReThrow("TestExcelErrorRoundTrip", Err, True)
12        TestExcelErrorRoundTrip = False
End Function

' Live round trip through a Variant() array (not a genuinely-typed Double() array) - the realistic
' "worst case" TrySerialiseArrayAsV has to handle in practice, since Range.Value2 always arrives
' this way, even when every element holds a number (measured historically via the now-removed
' prototype benchmark VEncodeSpeedTest, modPerformance.bas). Test1DArrayOfDoubles/TestExactRoundTripping
' already cover the genuinely-typed Double() case via JuliaCallVBA.
' Built with ReDim (1-based), not the VBA.Array() function - Array() returns a 0-based array, which
' would make ArraysIdentical report a spurious mismatch against the decoder's always-1-based result
' (Unserialise's Case 86 and Case 42 both ReDim their result 1 To n) even when every value round-
' tripped correctly - a real trap hit while writing this test, not a decoder bug.
Function TestVEncodeFromVariantArray()
          Dim x() As Variant
          Dim y As Variant

1         On Error GoTo ErrHandler
2         ReDim x(1 To 5)
3         x(1) = 1#: x(2) = -2.5: x(3) = 3.14159265358979: x(4) = 0#: x(5) = 1000000#
4         y = JuliaCallVBA("identity", x)
5         TestVEncodeFromVariantArray = ArraysIdentical(x, y)

6         Exit Function
ErrHandler:
7         PrintTwice ReThrow("TestVEncodeFromVariantArray", Err, True)
8         TestVEncodeFromVariantArray = False
End Function

' Exercises the Case 86 'V' branch's rank 3-9 handling (Unserialise, modUnserialise.bas), added
' alongside the Julia-side encode_for_xl(x::Array{Float64,N}) generalisation - reuses
' ParseDims/ReDimVariantArray/AssignByRank, the same helpers the general "*" format's own
' >=3-dimensional handling already used. A genuinely-typed Double() array, so both directions of
' this round trip go via "V": VBA -> Julia through TrySerialiseArrayAsV, Julia -> Excel through the
' new Array{Float64,N} method. Distinct values at every position (i + 10*j + 100*k) so a bug in the
' column-major index walk (either side) would show up as a value in the wrong place, not just a
' wrong-shaped result.
Function TestVFormat3DArray()
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim x(1 To 2, 1 To 3, 1 To 4) As Double
          Dim y As Variant

1         On Error GoTo ErrHandler
2         For k = 1 To 4
3             For j = 1 To 3
4                 For i = 1 To 2
5                     x(i, j, k) = i + 10 * j + 100 * k
6                 Next i
7             Next j
8         Next k

9         y = JuliaCallVBA("identity", x)
10        TestVFormat3DArray = ArraysIdentical(x, y)

11        Exit Function
ErrHandler:
12        PrintTwice ReThrow("TestVFormat3DArray", Err, True)
13        TestVFormat3DArray = False
End Function

' Exercises the 'R' branch's integer sub-format (RI, Unserialise, modUnserialise.bas), added
' alongside encode_for_xl(::AbstractRange{<:Integer}) in src/encode.jl. A UnitRange is encoded as
' just first/step/length and reconstructed via arithmetic in VBA, rather than transmitting every
' element - this confirms that fast path gives an identical result to fully materializing the
' range in Julia first.
Function TestRangeIntegerVsCollect()
          Dim y1 As Variant
          Dim y2 As Variant

1         On Error GoTo ErrHandler
2         y1 = ThrowIfError(JuliaEval("1:1000"))
3         y2 = ThrowIfError(JuliaEval("collect(1:1000)"))
4         TestRangeIntegerVsCollect = ArraysIdentical(y1, y2)

5         Exit Function
ErrHandler:
6         PrintTwice ReThrow("TestRangeIntegerVsCollect", Err, True)
7         TestRangeIntegerVsCollect = False
End Function

' As TestRangeIntegerVsCollect, but for a StepRange with a non-1 step (5:3:47) - confirms the
' reconstruction arithmetic handles a real step, not just the step=1 case UnitRange always has.
Function TestRangeStepVsCollect()
          Dim y1 As Variant
          Dim y2 As Variant

1         On Error GoTo ErrHandler
2         y1 = ThrowIfError(JuliaEval("5:3:47"))
3         y2 = ThrowIfError(JuliaEval("collect(5:3:47)"))
4         TestRangeStepVsCollect = ArraysIdentical(y1, y2)

5         Exit Function
ErrHandler:
6         PrintTwice ReThrow("TestRangeStepVsCollect", Err, True)
7         TestRangeStepVsCollect = False
End Function

' As TestRangeIntegerVsCollect, but for the float sub-format (RF) - a StepRangeLen{Float64,...}
' from broadcasting a scalar multiply over a range. This is the case that matters most for exact
' round-tripping: confirms VBA's naive "first + (i-1)*step" arithmetic exactly reproduces Julia's
' own (twice-precision) range materialization, not just approximately.
Function TestRangeFloatVsCollect()
          Dim y1 As Variant
          Dim y2 As Variant

1         On Error GoTo ErrHandler
2         y1 = ThrowIfError(JuliaEval("(1:1000).*pi"))
3         y2 = ThrowIfError(JuliaEval("collect((1:1000).*pi)"))
4         TestRangeFloatVsCollect = ArraysIdentical(y1, y2)

5         Exit Function
ErrHandler:
6         PrintTwice ReThrow("TestRangeFloatVsCollect", Err, True)
7         TestRangeFloatVsCollect = False
End Function

' Confirms JuliaEvalVBA (JuliaVectorToXLColumn=False) gives a genuine 1-D array for a range result,
' matching the same distinction Case 86 ('V') already makes for rank-1 arrays - the 'R' branch
' needs to honour JuliaVectorToXLColumn just as much as any other array-producing format.
Function TestRangeJuliaEvalVBAIsOneD()
          Dim y As Variant

1         On Error GoTo ErrHandler
2         y = ThrowIfError(JuliaEvalVBA("1:1000"))
3         TestRangeJuliaEvalVBAIsOneD = (NumDimensions(y) = 1) And (UBound(y) - LBound(y) + 1 = 1000) And _
              (y(LBound(y)) = 1) And (y(UBound(y)) = 1000)

4         Exit Function
ErrHandler:
5         PrintTwice ReThrow("TestRangeJuliaEvalVBAIsOneD", Err, True)
6         TestRangeJuliaEvalVBAIsOneD = False
End Function

' As TestRangeJuliaEvalVBAIsOneD, but via JuliaEval (JuliaVectorToXLColumn=True) - must give a 2-D,
' single-column array, matching how a worksheet formula needs to display a vector result.
Function TestRangeJuliaEvalIsColumn()
          Dim y As Variant

1         On Error GoTo ErrHandler
2         y = ThrowIfError(JuliaEval("1:1000"))
3         TestRangeJuliaEvalIsColumn = (NumDimensions(y) = 2) And (UBound(y, 1) = 1000) And (UBound(y, 2) = 1) And _
              (y(1, 1) = 1) And (y(1000, 1) = 1000)

4         Exit Function
ErrHandler:
5         PrintTwice ReThrow("TestRangeJuliaEvalIsColumn", Err, True)
6         TestRangeJuliaEvalIsColumn = False
End Function

' Directly confirms the compact "R" wire format is actually used (not just that results happen to
' be correct via some fallback) - checks the literal prefix of the raw string
' encode_for_xl produces, for both the integer and float sub-formats.
Function TestRangeWireFormat()
          Dim s As Variant

1         On Error GoTo ErrHandler
2         s = ThrowIfError(JuliaEvalVBA("JuliaExcel.encode_for_xl(1:1000)"))
3         If Left$(s, 3) <> "RI," Then Throw "Expected integer range to use 'RI' wire format, got: " & Left$(s, 10)
4         s = ThrowIfError(JuliaEvalVBA("JuliaExcel.encode_for_xl((1:1000).*pi)"))
5         If Left$(s, 3) <> "RF," Then Throw "Expected float range to use 'RF' wire format, got: " & Left$(s, 10)
6         TestRangeWireFormat = True

7         Exit Function
ErrHandler:
8         PrintTwice ReThrow("TestRangeWireFormat", Err, True)
9         TestRangeWireFormat = False
End Function

' Confirms MakeJuliaLiteral's escaping order and coverage: backslash must be escaped first (so later
' substitutions' own inserted backslashes aren't re-escaped), Trojan-Source bidi control characters
' (from both guarded ranges, 8234-8238 and 8294-8297) become \uXXXX, and CR/LF/$/an embedded double
' quote are each escaped. Previously only ever exercised via plain ASCII diagnostic strings passed to
' PrintTwice, so none of this was actually checked.
Function TestMakeJuliaLiteral() As Boolean
          Dim OK As Boolean
          Dim Res As String
          Dim x As String

1         On Error GoTo ErrHandler
2         OK = True

          'One character from each guarded bidi range, plus backslash, CR, LF, $ and an embedded
          'double quote, all in one string.
3         x = "a\b" & vbCr & vbLf & "$" & Chr(34) & ChrW(8234) & ChrW(8296)
4         Res = MakeJuliaLiteral(x)

5         OK = OK And Left$(Res, 1) = Chr(34)                  'outer quoting
6         OK = OK And Right$(Res, 1) = Chr(34)
7         OK = OK And InStr(Res, "\\") > 0                     'backslash doubled
8         OK = OK And InStr(Res, "\r") > 0                     'CR
9         OK = OK And InStr(Res, "\n") > 0                     'LF
10        OK = OK And InStr(Res, "\$") > 0                     '$
11        OK = OK And InStr(Res, "\" & Chr(34)) > 0            'embedded quote
12        OK = OK And InStr(Res, "\u202a") > 0                 'ChrW(8234), first guarded range
13        OK = OK And InStr(Res, "\u2068") > 0                 'ChrW(8296), second guarded range

          'A string with none of the above passes through unchanged, just quoted.
14        OK = OK And MakeJuliaLiteral("hello") = Chr(34) & "hello" & Chr(34)

15        TestMakeJuliaLiteral = OK
16        Exit Function
ErrHandler:
17        PrintTwice ReThrow("TestMakeJuliaLiteral", Err, True)
18        TestMakeJuliaLiteral = False
End Function

' Confirms the Case 163 (string) length guard in Unserialise: the 32,767-character Excel-worksheet-
' string limit always applies at the top level (Depth=1) regardless of StringLengthLimit; a shorter
' StringLengthLimit only applies to array elements (Depth>1); the message wording differs depending
' on whether StringLengthLimit is exactly 32768; and the whole check is skipped when
' StringLengthLimit=0 (called from VBA, not a worksheet formula). None of this was previously
' exercised - real calls never generate strings this large.
Function TestUnserialiseStringTooLong() As Boolean
          Dim d As Long
          Dim OK As Boolean
          Dim Threw As Boolean

1         On Error GoTo ErrHandler
2         OK = True

          'Top level (Depth=1), string just over 32,767 chars, StringLengthLimit=32768 exactly ->
          'the shorter message variant, mentioning only the worksheet-cell limit.
3         Threw = False
4         On Error Resume Next
5         UnserialiseFromString Chr(163) & String(35000, "x"), False, 32768, False
6         Threw = InStr(Err.Description, "limit is 32,767") > 0 And _
                  InStr(Err.Description, "string elements of an array") = 0
7         On Error GoTo ErrHandler
8         OK = OK And Threw

          'Top level (Depth=1), same oversized string, but StringLengthLimit=500 (<> 32768) -> the
          'longer message variant, still triggered by the 32,767 top-level limit, not 500.
9         Threw = False
10        On Error Resume Next
11        UnserialiseFromString Chr(163) & String(35000, "x"), False, 500, False
12        Threw = InStr(Err.Description, "499 for string elements of an array") > 0
13        On Error GoTo ErrHandler
14        OK = OK And Threw

          'Nested (Depth=1 passed in, becomes 2 inside Unserialise), a 601-char string with
          'StringLengthLimit=500 -> throws, because a non-top-level element is checked against
          'StringLengthLimit, not the 32,767 top-level limit (601 alone would not throw at Depth=1).
15        Threw = False
16        d = 1
17        On Error Resume Next
18        Unserialise Chr(163) & String(600, "x"), False, d, 500, False
19        Threw = InStr(Err.Description, "499 for string elements of an array") > 0
20        On Error GoTo ErrHandler
21        OK = OK And Threw

          'Same 601-char string at top level (Depth=1) does not throw - it's under the 32,767 limit.
22        OK = OK And UnserialiseFromString(Chr(163) & String(600, "x"), False, 500, False) = String(600, "x")

          'StringLengthLimit=0 means "not called from a worksheet formula" - the check is skipped
          'entirely, however long the string.
23        OK = OK And Len(UnserialiseFromString(Chr(163) & String(40000, "x"), False, 0, False)) = 40000

24        TestUnserialiseStringTooLong = OK
25        Exit Function
ErrHandler:
26        PrintTwice ReThrow("TestUnserialiseStringTooLong", Err, True)
27        TestUnserialiseStringTooLong = False
End Function

' Confirms Unserialise rejects a wire-format array header with a multi-digit rank (e.g. "*10,...") -
' the rank digit is assumed to be a single character elsewhere in the parsing, so a malformed or
' out-of-range rank should be caught here rather than silently misparsed. Never previously exercised
' - a genuine 10+ dimensional array is never actually produced by encode_for_xl (Julia's own side
' caps out at 9 dimensions), so this simulates a corrupted/malformed header directly.
Function TestUnserialiseMalformedRank() As Boolean
          Dim d As Long
          Dim OK As Boolean
          Dim Threw As Boolean

1         On Error GoTo ErrHandler
2         d = 0
3         Threw = False
4         On Error Resume Next
5         Unserialise "*10,1,1;1;x", False, d, 0, False
6         Threw = InStr(Err.Description, "10 dimensions (max supported: 9)") > 0
7         On Error GoTo ErrHandler
8         OK = Threw

9         TestUnserialiseMalformedRank = OK
10        Exit Function
ErrHandler:
11        PrintTwice ReThrow("TestUnserialiseMalformedRank", Err, True)
12        TestUnserialiseMalformedRank = False
End Function

' Confirms ConcatenateExpressions' handling of each accepted input shape (scalar, 1-D array, single-
' column 2-D array) and its two rejection paths (multi-column 2-D array, rank 3+) - previously
' unexercised, since every existing call to JuliaEval/JuliaCall in the test suite passes a plain
' string.
Function TestConcatenateExpressions() As Boolean
          Dim Arr1Col() As Variant
          Dim Arr2Col() As Variant
          Dim Arr3D() As Variant
          Dim OK As Boolean
          Dim Threw As Boolean

1         On Error GoTo ErrHandler
2         OK = True

3         OK = OK And ConcatenateExpressions("a=1") = "a=1"
4         OK = OK And ConcatenateExpressions(Array("a=1", "b=2")) = "a=1;b=2"

5         ReDim Arr1Col(1 To 2, 1 To 1)
6         Arr1Col(1, 1) = "a=1"
7         Arr1Col(2, 1) = "b=2"
8         OK = OK And ConcatenateExpressions(Arr1Col) = "a=1;b=2"

9         ReDim Arr2Col(1 To 1, 1 To 2)
10        Arr2Col(1, 1) = "a=1"
11        Arr2Col(1, 2) = "b=2"
12        Threw = False
13        On Error Resume Next
14        ConcatenateExpressions Arr2Col
15        Threw = InStr(Err.Description, "2 columns") > 0
16        On Error GoTo ErrHandler
17        OK = OK And Threw

18        ReDim Arr3D(1 To 1, 1 To 1, 1 To 1)
19        Threw = False
20        On Error Resume Next
21        ConcatenateExpressions Arr3D
22        Threw = InStr(Err.Description, "Too many dimensions") > 0
23        On Error GoTo ErrHandler
24        OK = OK And Threw

25        TestConcatenateExpressions = OK
26        Exit Function
ErrHandler:
27        PrintTwice ReThrow("TestConcatenateExpressions", Err, True)
28        TestConcatenateExpressions = False
End Function

