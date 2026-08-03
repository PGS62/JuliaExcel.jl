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
5         JuliaLaunch
6         ThrowIfError JuliaEval("1+1")
7         InputData = Application.Evaluate("=RANDARRAY(100)")
8         ThrowIfError JuliaCall("identity", InputData)
9         ThrowIfError JuliaCall("sum", InputData)

          'Latency test
10        t1 = ElapsedTime
11        For i = 1 To NumCallsOnePlusOne
12            Res = JuliaEval("1+1")
13        Next i
14        t2 = ElapsedTime

15        Debug.Print "'JuliaExcel Version = " & CStr(shAudit.Range("Headers").Cells(2, 1).Value)
16        Debug.Print "'Computer = " & Environ$("ComputerName")
17        Debug.Print "'Latency test"
18        Debug.Print "'Average time for JuliaEval(""1+1"") = " & CStr(1000 * (t2 - t1) / NumCallsOnePlusOne) & _
              " miliseconds (averaged over " & CStr(NumCallsOnePlusOne) & " calls)"

19        InputData = Application.Evaluate("=RANDARRAY(" & VectorLength & ")")

          'Data transport tests
20        For j = 1 To 3
21            JuliaFunction = Choose(j, "identity", "sum", "collect")
22            If JuliaFunction = "collect" Then
23                t1 = ElapsedTime()
24                For i = 1 To NumCallsVectors
25                    Res = JuliaEval("collect((1:" & VectorLength & ").*pi)")
26                Next i
27                t2 = ElapsedTime()
28            Else
29                t1 = ElapsedTime()
30                For i = 1 To NumCallsVectors
31                    Res = JuliaCall(JuliaFunction, InputData)
32                Next i
33                t2 = ElapsedTime()
34            End If

35            If JuliaFunction = "identity" Then
36                If Not ArraysIdentical(Res, InputData) Then
37                    Throw "Ohoh, return from Julia function identity is not equal to its input"
38                End If
39            End If

40            If JuliaFunction = "identity" Then
41                WhatWasExecuted = "JuliaCall(""identity"", vector of length " & Format(VectorLength, "###,###") & ")"
42                Debug.Print "'Two-way data transport test"
43            ElseIf JuliaFunction = "sum" Then
44                WhatWasExecuted = "JuliaCall(""sum"", vector of length " & Format(VectorLength, "###,###") & ")"
45                Debug.Print "'One-way data transport test, Excel to Julia"
46            ElseIf JuliaFunction = "collect" Then
47                WhatWasExecuted = "JuliaEval(""collect((1:" & CStr(VectorLength) & ").*pi)"")"
48                Debug.Print "'One-way data transport test, Julia to Excel"
49            End If

50            Debug.Print "'Average time for " & WhatWasExecuted & " = " & _
                  CStr((t2 - t1) / NumCallsVectors) & " seconds (averaged over " & CStr(NumCallsVectors) & " calls)"
51        Next j

52        Exit Sub
ErrHandler:
53        ReThrow "PerformanceTest", Err
End Sub










