Attribute VB_Name = "modPerformance"
' Copyright (c) 2021-2026 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme

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

'========================================================================================================================
'Running method PerformanceTest
'Time now = 2026-08-10 13:17:53
'JuliaExcel Version = 128
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 1.16016960004345 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of length 100,000) = 0.557932109991089 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of length 100,000) = 0.276260319992434 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 0.370191970001906 seconds (averaged over 10 calls)

'========================================================================================================================
'Running method PerformanceTest
'Time now = 2026-08-10 14:39:57
'JuliaExcel Version = 128
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 1.09734679991379 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of length 100,000) = 0.274450039991643 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of length 100,000) = 0.112631449999753 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 0.167135900002904 seconds (averaged over 10 calls)

'========================================================================================================================
'Running method PerformanceTest
'Time now = 2026-08-13 07:35:29
'JuliaExcel Version = 131
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 1.31439220000175 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of length 100,000) = 0.257647519999591 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of length 100,000) = 0.102828399999999 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 0.148632560000988 seconds (averaged over 10 calls)

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
55        MsgBox ReThrow("PerformanceTest", Err, True)
End Sub

