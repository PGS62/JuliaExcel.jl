Attribute VB_Name = "modPerformance"
Option Explicit

'========================================================================================================================
'Running method PerformanceTest
'Time now = 2026-08-03 18:01:11
'JuliaExcel Version = 107
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 7.96316399984062 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of length 100,000) = 1.48447485999204 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of length 100,000) = 1.27661278001033 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 0.294781249994412 seconds (averaged over 10 calls)

'========================================================================================================================
'Running method PerformanceTest
'Time now = 2026-08-04 15:53:30
'JuliaExcel Version = 109 <- Version 109 had experimental code to use UTF-8 for Excel-> Julia and Julia->Excel _
communication.Julia To Excel gets faster Excel To Julia gets slower.
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 8.14005120001093 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of length 100,000) = 1.82576665000015 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of length 100,000) = 1.57798423999993 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 0.187884929999564 seconds (averaged over 10 calls)

'Running method PerformanceTest
'Time now = 2026-08-04 16:25:47
'JuliaExcel Version = 110 < using UTF-8 for Julia -> Excel, UTF-16 for Excel -> Julia
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 14.7025480000011 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of length 100,000) = 2.14857505000036 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of length 100,000) = 2.03623397000047 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 0.222088039999653 seconds (averaged over 10 calls)

'========================================================================================================================
'Running method PerformanceTest
'Time now = 2026-08-04 17:07:48
'JuliaExcel Version = 111
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 10.9454097999987 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of length 100,000) = 2.14352065000057 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of length 100,000) = 1.94212056999968 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 0.31337268999996 seconds (averaged over 10 calls)

'========================================================================================================================
'Running method PerformanceTest
'Time now = 2026-08-06 17:40:36
'JuliaExcel Version = 119
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 11.8458608000074 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of length 100,000) = 0.555459060001885 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of length 100,000) = 0.154651710001053 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 0.479368480000994 seconds (averaged over 10 calls)
'========================================================================================================================
'Running method PerformanceTest
'Time now = 2026-08-07 13:53:55
'JuliaExcel Version = 122
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 2.09547800000291 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of length 100,000) = 0.278066150000086 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of length 100,000) = 0.104752339998959 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 0.154756969999289 seconds (averaged over 10 calls)




Sub PerformanceTest()
          Const NumCallsOnePlusOne As Long = 500
          Const NumCallsVectors = 10
          Const VectorLength As Long = 100000
          Dim i As Long
          Dim InputData As Variant
          Dim j As Long
          Dim JuliaFunction As String
          Dim Res As Variant
          Dim t1 As Double
          Dim t2 As Double
          Dim WhatWasExecuted As String
          
1         On Error GoTo ErrHandler
2         Debug.Print "'" & String(120, "=")
3         Debug.Print "'Running method PerformanceTest"
4         Debug.Print "'Time now = " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
          
          'Warm up
5         JuliaEval "exit()" 'shuts down Julia if it's running
6         JuliaLaunch
7         ThrowIfError JuliaEval("1+1")
8         InputData = Application.Evaluate("=RANDARRAY(100)")
9         ThrowIfError JuliaCall("identity", InputData)
10        ThrowIfError JuliaCall("sum", InputData)
11        ThrowIfError JuliaEval("collect((1:" & VectorLength & ").*pi)")
          
          'Latency test
12        t1 = ElapsedTime
13        For i = 1 To NumCallsOnePlusOne
14            Res = JuliaEval("1+1")
15        Next i
16        t2 = ElapsedTime
          
17        Debug.Print "'JuliaExcel Version = " & CStr(shAudit.Range("Headers").Cells(2, 1).Value)
18        Debug.Print "'Computer = " & Environ$("ComputerName")
19        Debug.Print "'Latency test"
20        Debug.Print "'Average time for JuliaEval(""1+1"") = " & CStr(1000 * (t2 - t1) / NumCallsOnePlusOne) & _
              " miliseconds (averaged over " & CStr(NumCallsOnePlusOne) & " calls)"
          
21        InputData = Application.Evaluate("=RANDARRAY(" & VectorLength & ")")
          
          'Data transport tests
22        For j = 1 To 3
23            JuliaFunction = Choose(j, "identity", "sum", "collect")
24            If JuliaFunction = "collect" Then
25                t1 = ElapsedTime()
26                For i = 1 To NumCallsVectors
27                    Res = JuliaEval("collect((1:" & VectorLength & ").*pi)")
28                Next i
29                t2 = ElapsedTime()
30            Else
31                t1 = ElapsedTime()
32                For i = 1 To NumCallsVectors
33                    Res = JuliaCall(JuliaFunction, InputData)
34                Next i
35                t2 = ElapsedTime()
36            End If
        
37            If JuliaFunction = "identity" Then
38                If Not ArraysIdentical(Res, InputData) Then
39                    Throw "Ohoh, return from Julia function identity is not equal to its input"
40                End If
41            End If
        
42            If JuliaFunction = "identity" Then
43                WhatWasExecuted = "JuliaCall(""identity"", vector of length " & Format(VectorLength, "###,###") & ")"
44                Debug.Print "'Two-way data transport test"
45            ElseIf JuliaFunction = "sum" Then
46                WhatWasExecuted = "JuliaCall(""sum"", vector of length " & Format(VectorLength, "###,###") & ")"
47                Debug.Print "'One-way data transport test, Excel to Julia"
48            ElseIf JuliaFunction = "collect" Then
49                WhatWasExecuted = "JuliaEval(""collect((1:" & CStr(VectorLength) & ").*pi)"")"
50                Debug.Print "'One-way data transport test, Julia to Excel"
51            End If
        
52            Debug.Print "'Average time for " & WhatWasExecuted & " = " & _
                  CStr((t2 - t1) / NumCallsVectors) & " seconds (averaged over " & CStr(NumCallsVectors) & " calls)"
53        Next j
          
54        Exit Sub
ErrHandler:
55        ReThrow "PerformanceTest", Err
End Sub

'========================================================================================================================
'Running method BenchmarkDoubleToHex
'JuliaExcel Version = 117
'Time now = 2026-08-05 20:08:07
'Correctness check: PASSED (10 values agree)
'DoubleToHexOld:    1.170 microseconds/call (1,000,000 calls)
'DoubleToHex: 0.294 microseconds/call (1,000,000 calls)
'Speedup:        4.0x

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : BenchmarkDoubleToHex
' Purpose    : Correctness check and speed comparison of DoubleToHexOld (original, Hex$/Mid$ approach)
'              vs DoubleToHex (lookup table approach). Run from the Immediate window or press F5.
' -----------------------------------------------------------------------------------------------------------------------
Sub BenchmarkDoubleToHex()
          Const NumOuterLoops As Long = 1000
          Const NumInnerLoops As Long = 1000
          Dim CheckValues(1 To 10) As Double
          Dim i As Long
          Dim j As Long
          Dim Mismatches As Long
          Dim NumCalls As Long
          Dim t1 As Double
          Dim t2 As Double
          Dim TestData() As Double
          Dim Tmp As String
          Dim tNew As Double
          Dim tOld As Double

1         On Error GoTo ErrHandler
2         NumCalls = NumOuterLoops * NumInnerLoops

3         ReDim TestData(1 To NumInnerLoops)

4         Debug.Print "'" & String(120, "=")
5         Debug.Print "'Running method BenchmarkDoubleToHex"
6         Debug.Print "'JuliaExcel Version = " & CStr(shAudit.Range("Headers").Cells(2, 1).Value)
7         Debug.Print "'Time now = " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")

          ' ---- Correctness check: both functions must agree on 10 varied inputs ----
8         CheckValues(1) = 0#
9         CheckValues(2) = 1#
10        CheckValues(3) = -1#
11        CheckValues(4) = 4# * Atn(1#)              ' pi
12        CheckValues(5) = -4# * Atn(1#)             ' -pi
13        CheckValues(6) = 0.1
14        CheckValues(7) = 123456789.123456
15        CheckValues(8) = 1.23456789012345E+300
16        CheckValues(9) = 1.23456789012345E-300
17        CheckValues(10) = -9.87654321098765E+123

18        Mismatches = 0
19        For i = 1 To 10
20            If DoubleToHexOld(CheckValues(i)) <> DoubleToHex(CheckValues(i)) Then
21                Debug.Print "MISMATCH for input " & i & " (" & CheckValues(i) & "): " & _
                      "DoubleToHex=" & DoubleToHexOld(CheckValues(i)) & _
                      " DoubleToHexNew=" & DoubleToHex(CheckValues(i))
22                Mismatches = Mismatches + 1
23            End If
24        Next i
25        If Mismatches = 0 Then
26            Debug.Print "'Correctness check: PASSED (10 values agree)"
27        Else
28            Debug.Print "'Correctness check: FAILED (" & Mismatches & " mismatches)"
29        End If

          ' ---- Speed benchmark ----
          ' Seed the static lookup table in DoubleToHex before timing begins
30        Tmp = DoubleToHex(1#)

31        For i = 1 To NumInnerLoops
32            TestData(i) = CDbl(i) * 3.14159265358979
33        Next i

34        t1 = ElapsedTime()
35        For j = 1 To NumOuterLoops
36            For i = 1 To NumInnerLoops
37                Tmp = DoubleToHexOld(TestData(i))
38            Next i
39        Next j
40        t2 = ElapsedTime()
41        tOld = t2 - t1

42        t1 = ElapsedTime()
43        For j = 1 To NumOuterLoops
44            For i = 1 To NumInnerLoops
45                Tmp = DoubleToHex(TestData(i))
46            Next i
47        Next j
48        t2 = ElapsedTime()
49        tNew = t2 - t1

50        Debug.Print "'DoubleToHexOld:    " & Format$(tOld / NumCalls * 1000000#, "0.000") & _
              " microseconds/call (" & Format(NumCalls, "###,###") & " calls)"
51        Debug.Print "'DoubleToHex: " & Format$(tNew / NumCalls * 1000000#, "0.000") & _
              " microseconds/call (" & Format(NumCalls, "###,###") & " calls)"
52        Debug.Print "'Speedup:        " & Format$(tOld / tNew, "0.0") & "x"

53        Exit Sub
ErrHandler:
54        ReThrow "BenchmarkDoubleToHex", Err
End Sub

'========================================================================================================================
'Running method BenchmarkSingleToHex
'JuliaExcel Version = 118
'Time now = 2026-08-06 07:31:07
'Correctness check: PASSED (10 values agree)
'SingleToHexOld: 0.215 microseconds/call (10,000,000 calls)
'SingleToHex:    0.182 microseconds/call (10,000,000 calls)
'Speedup:        1.2x
' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : BenchmarkSingleToHex
' Purpose    : Correctness check and speed comparison of SingleToHexOld (original, Hex$/LPad approach)
'              vs SingleToHex (lookup table approach). Run from the Immediate window or press F5.
' -----------------------------------------------------------------------------------------------------------------------
Sub BenchmarkSingleToHex()
          Const NumOuterLoops As Long = 10000
          Const NumInnerLoops As Long = 1000
          Dim CheckValues(1 To 10) As Single
          Dim i As Long
          Dim j As Long
          Dim Mismatches As Long
          Dim NumCalls As Long
          Dim t1 As Double
          Dim t2 As Double
          Dim TestData(1 To 1000) As Single
          Dim Tmp As String
          Dim tNew As Double
          Dim tOld As Double

1         On Error GoTo ErrHandler
2         NumCalls = NumOuterLoops * NumInnerLoops

3         Debug.Print "'" & String(120, "=")
4         Debug.Print "'Running method BenchmarkSingleToHex"
5         Debug.Print "'JuliaExcel Version = " & CStr(shAudit.Range("Headers").Cells(2, 1).Value)
6         Debug.Print "'Time now = " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")

          ' ---- Correctness check: both functions must agree on 10 varied inputs ----
7         CheckValues(1) = 0!
8         CheckValues(2) = 1!
9         CheckValues(3) = -1!
10        CheckValues(4) = 4! * Atn(1!)          ' pi as Single
11        CheckValues(5) = -4! * Atn(1!)          ' -pi as Single
12        CheckValues(6) = 0.1!
13        CheckValues(7) = 12345.68!
14        CheckValues(8) = 1.234568E+37!
15        CheckValues(9) = 1.234568E-37!
16        CheckValues(10) = -9.876544E+12!

17        Mismatches = 0
18        For i = 1 To 10
19            If SingleToHexOld(CheckValues(i)) <> SingleToHex(CheckValues(i)) Then
20                Debug.Print "MISMATCH for input " & i & " (" & CheckValues(i) & "): " & _
                      "SingleToHexOld=" & SingleToHexOld(CheckValues(i)) & _
                      " SingleToHex=" & SingleToHex(CheckValues(i))
21                Mismatches = Mismatches + 1
22            End If
23        Next i
24        If Mismatches = 0 Then
25            Debug.Print "'Correctness check: PASSED (10 values agree)"
26        Else
27            Debug.Print "'Correctness check: FAILED (" & Mismatches & " mismatches)"
28        End If

          ' ---- Speed benchmark ----
          ' Seed the static lookup table in SingleToHex before timing begins
29        Tmp = SingleToHex(1!)

30        For i = 1 To NumInnerLoops
31            TestData(i) = CSng(i) * 3.14159!
32        Next i

33        t1 = ElapsedTime()
34        For j = 1 To NumOuterLoops
35            For i = 1 To NumInnerLoops
36                Tmp = SingleToHexOld(TestData(i))
37            Next i
38        Next j
39        t2 = ElapsedTime()
40        tOld = t2 - t1

41        t1 = ElapsedTime()
42        For j = 1 To NumOuterLoops
43            For i = 1 To NumInnerLoops
44                Tmp = SingleToHex(TestData(i))
45            Next i
46        Next j
47        t2 = ElapsedTime()
48        tNew = t2 - t1

49        Debug.Print "'SingleToHexOld: " & Format$(tOld / NumCalls * 1000000#, "0.000") & _
              " microseconds/call (" & Format(NumCalls, "###,###") & " calls)"
50        Debug.Print "'SingleToHex:    " & Format$(tNew / NumCalls * 1000000#, "0.000") & _
              " microseconds/call (" & Format(NumCalls, "###,###") & " calls)"
51        Debug.Print "'Speedup:        " & Format$(tOld / tNew, "0.0") & "x"

52        Exit Sub
ErrHandler:
53        ReThrow "BenchmarkSingleToHex", Err
End Sub

