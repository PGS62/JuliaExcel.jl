Attribute VB_Name = "modPerformance"
' Copyright (c) 2021-2026 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme

Option Explicit
Option Private Module

'--------------------------------------------------
'05-Nov-2021 16:18:37        DESKTOP-0VD2AF0
'Expression = fill("xxx", 1000, 1000)
'Average time in JuliaEval    1.47189380999916
'--------------------------------------------------
'06-Nov-2021 12:28:58        PHILIP-LAPTOP
'Expression = fill("xxx", 1000, 1000)
'Average time in JuliaEval    1.9295860900078
'--------------------------------------------------
'30-Nov-2021 10:13:30        PHILIP-LAPTOP
'Expression = fill("xxx", 1000, 1000)
'Average time in JuliaEval    2.82354638000252  <--- Mmm, why the slowdown since 6-Nov version? Use of Assign?
'--------------------------------------------------
'01-Dec-2021 10:30:10       DESKTOP-0VD2AF0
'Expression = fill("xxx",1000,1000)
'Average time in JuliaEval   2.25666286000051   <-- also seeing slowdown on PC in the office
'--------------------------------------------------
'20-Sep-2023 16:34:52       DESKTOP-HSGAM5S
'Expression = fill("xxx",1000,1000)
'Average time in JuliaEval   1.42395350000006  <-- higher spec PC
'--------------------------------------------------
'29-Oct-2025 18:40:16       PHILIP-HPZ1
'Expression = fill("xxx",1000,1000)
'Average time in JuliaEval   2.66512269999985           Averaged over 10 calls
'--------------------------------------------------
'22-Dec-2025 15:57:14       MSI
'Expression = fill("xxx",1000,1000)
'Average time in JuliaEval   1.7418744300001            Averaged over 20 calls
'--------------------------------------------------
'--------------------------------------------------
'18-Aug-2026 20:06:31       MSI
'Expression = fill("xxx",1000,1000)
'Average time in JuliaEval   0.920728114999656          Averaged over 20 calls
'--------------------------------------------------
Private Sub SimpleSpeedTest()

          'Const Expression As String = "fill(""xxx"",1000,1000)"
          'Const NumCalls = 20
          Const Expression As String = "1+1"
          Const NumCalls = 1000
          Const UseLinux As Boolean = False
          Dim i As Long
          Dim Res As Variant
          Dim t1 As Double
          Dim t2 As Double

1         JuliaLaunch UseLinux
2         t1 = ElapsedTime
3         For i = 1 To NumCalls
4             Res = JuliaEval(Expression)
5         Next i
6         t2 = ElapsedTime

7         Debug.Print "'" & Format(Now(), "dd-mmm-yyyy hh:mm:ss"), Environ("ComputerName")
8         Debug.Print "'Expression = " & Expression
9         Debug.Print "'Average time in JuliaEval", (t2 - t1) / NumCalls, "Averaged over " & CStr(NumCalls) & " calls"
10        Debug.Print "'--------------------------------------------------"
11    End Sub

      '--------------------------------------------------
      '29-Oct-2025 18:37:22       PHILIP-HPZ1
      'Expression = 1+1
      'Average time in JuliaEval   6.16188229999898E-03       Averaged over 1000 calls
      '--------------------------------------------------
      '22-Dec-2025 15:58:22       MSI
      'Expression = 1+1
      'Average time in JuliaEval   0.015039338300001          Averaged over 1000 calls
      '--------------------------------------------------
      '18-Aug-2026 20:09:02       MSI
      'Expression = 1+1
      'Average time in JuliaEval   1.32635509999818E-03       Averaged over 1000 calls
      '--------------------------------------------------

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
'========================================================================================================================
'Running method PerformanceTest
'Time now = 2026-08-16 15:04:10
'JuliaExcel Version = 138
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 1.20061220001662 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of length 100,000) = 0.255838729999959 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of length 100,000) = 0.100987849998637 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 0.146759319998091 seconds (averaged over 10 calls)
'========================================================================================================================
'Running method PerformanceTest
'Time now = 2026-08-17 11:56:58
'JuliaExcel Version = 139
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 1.64528539997991 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of 100,000 doubles) = 0.199695459997747 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of 100,000 doubles) = 0.112783389998367 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 9.42291199986357E-02 seconds (averaged over 10 calls)
'========================================================================================================================
'Running method PerformanceTest - NOW WITH THE VBA -> JULIA "V" ENCODER ALSO WIRED IN
'(TrySerialiseArrayAsV in modSerialise.bas, decode_xl_array_v in src/decode.jl) - previously only
'Julia -> Excel had the fast "V" path. Compare "sum" (Excel -> Julia) and "identity" (two-way)
'against the runs above and against the pre-V baseline (v138, 2026-08-16 15:04:10: identity =
'0.2558s, sum = 0.1010s, collect = 0.1468s).
'Time now = 2026-08-17 14:25:31
'JuliaExcel Version = 139
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 1.81287340004928 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of 100,000 doubles) = 0.150152799999341 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of 100,000 doubles) = 6.95522399968468E-02 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 8.78807900007814E-02 seconds (averaged over 10 calls)
'========================================================================================================================
'Running method PerformanceTest
'Time now = 2026-08-18 11:31:05. Slowdown since 2026-08-17 14:25:31 - ROOT CAUSE FOUND AND FIXED (see
'run below): TrySerialiseArrayAsV's finiteness check (modSerialise.bas) called a separate Function
'(IsFiniteHex, ByVal String parameter) once per element - measured at ~80ms of pure VBA function-
'call/BSTR-copy overhead alone for a 100,000-element array, on top of the encoding itself. Fixed by
'inlining the check at each of the three call sites; IsFiniteHex deleted. Not thermal throttling -
'the PC had been idle overnight and was cool when this run was taken.
'JuliaExcel Version = 142
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 1.56602779999957 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of 100,000 doubles) = 0.254055260000314 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of 100,000 doubles) = 0.157926959999895 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 9.43088400003035E-02 seconds (averaged over 10 calls)
'One-way data transport (AbstractRange), Julia to Excel
'Average time for JuliaEval("(1:100000).*pi") = 1.29171199994744E-02 seconds (averaged over 10 calls)
'========================================================================================================================
'Running method PerformanceTest - AFTER inlining TrySerialiseArrayAsV's finiteness check (removing
'the per-element IsFiniteHex Function call). Confirms recovery to, and slightly better than, the
'2026-08-17 14:25:31 numbers (identity = 0.1502s, sum = 0.0696s).
'Time now = 2026-08-18 11:52:28
'JuliaExcel Version = 142
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 1.20299299999897 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of 100,000 doubles) = 0.158094169999822 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of 100,000 doubles) = 0.071729609999602 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 8.29558800003724E-02 seconds (averaged over 10 calls)
'One-way data transport (AbstractRange), Julia to Excel
'Average time for JuliaEval("(1:100000).*pi") = 1.24318100002711E-02 seconds (averaged over 10 calls)
'========================================================================================================================
'Running method PerformanceTest - AFTER removing TrySerialiseArrayAsV's NaN/Inf finiteness check
'entirely (not just inlining it): shown unnecessary for the Excel -> Julia direction, since the
'general per-scalar Double encode does no NaN/Inf translation either, and Julia's own decode
'reconstructs the identical Float64 bit pattern via either path - unlike the Julia -> Excel
'direction, where NaN -> #N/A / Inf -> #NUM! translation is genuine and stays in place (see
'TrySerialiseArrayAsV's docstring, modSerialise.bas). Best numbers yet.
'Time now = 2026-08-18 12:12:01
'JuliaExcel Version = 143
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 1.2095915999962 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of 100,000 doubles) = 0.14551833000005 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of 100,000 doubles) = 6.27044100001513E-02 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 8.13102000000072E-02 seconds (averaged over 10 calls)
'One-way data transport (AbstractRange), Julia to Excel
'Average time for JuliaEval("(1:100000).*pi") = 1.25317000005452E-02 seconds (averaged over 10 calls)
'========================================================================================================================
'Running method PerformanceTest - AFTER wiring the bulk-CopyMemory hex trick into production V-format
'encode (TrySerialiseArrayAsV, modSerialise.bas) and decode (Unserialise Case 86, modUnserialise.bas)
'- one bulk memory copy instead of N per-element LSet + function-call operations on each side. See
'BulkHexOfDoubleArray/BulkDoublesFromHex's own docstrings, and the now-deleted
'modHexBulkPrototype.bas (see git history) for the prototype that measured this first, in
'isolation. "sum" and "identity" (encode) and "collect" (decode) all improve substantially;
'"collect" roughly halves.
'Time now = 2026-08-19 09:02:52
'JuliaExcel Version = 145
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 1.24485660000937 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of 100,000 doubles) = 0.10265148000035 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of 100,000 doubles) = 5.39971099991817E-02 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 4.13354000018444E-02 seconds (averaged over 10 calls)
'One-way data transport (AbstractRange), Julia to Excel
'Average time for JuliaEval("(1:100000).*pi") = 1.41443200001959E-02 seconds (averaged over 10 calls)
'========================================================================================================================
'Running method PerformanceTest
'Time now = 2026-08-23 18:51:59
'JuliaExcel Version = 148
'Computer = MSI
'Latency test
'Average time for JuliaEval("1+1") = 1.10643340006936 miliseconds (averaged over 500 calls)
'Two-way data transport test
'Average time for JuliaCall("identity", vector of 100,000 doubles) = 9.43825400026981E-02 seconds (averaged over 10 calls)
'One-way data transport test, Excel to Julia
'Average time for JuliaCall("sum", vector of 100,000 doubles) = 5.68173000006936E-02 seconds (averaged over 10 calls)
'One-way data transport test, Julia to Excel
'Average time for JuliaEval("collect((1:100000).*pi)") = 4.60276499972679E-02 seconds (averaged over 10 calls)
'One-way data transport (AbstractRange), Julia to Excel
'Average time for JuliaEval("(1:100000).*pi") = 1.36371799977496E-02 seconds (averaged over 10 calls)

Function PerformanceTest() As String
          Const NumCallsOnePlusOne As Long = 500
          Const NumCallsVectors = 10
          Const VectorLength As Long = 100000
          Dim i As Long
          Dim InputData As Variant
          Dim j As Long
          Dim JuliaFunction As String
          Dim Report As String
          Dim Res As Variant
          Dim t1 As Double
          Dim t2 As Double
          Dim WhatWasExecuted As String

1         On Error GoTo ErrHandler
2         Report = Report & String(120, "=") & vbLf
3         Report = Report & "Running method PerformanceTest" & vbLf
4         Report = Report & "Time now = " & Format$(Now(), "yyyy-mm-dd hh:mm:ss") & vbLf

          'Warm up
5         JuliaEval "exit()" 'shuts down Julia if it's running
6         PreciseSleep 1000
7         JuliaLaunch , , gTestCommandOptions
8         ThrowIfError JuliaEval("1+1")
9         InputData = Application.Evaluate("=RANDARRAY(100)")
10        ThrowIfError JuliaCall("identity", InputData)
11        ThrowIfError JuliaCall("sum", InputData)
12        ThrowIfError JuliaEval("collect((1:" & VectorLength & ").*pi)")

          'Latency test
13        t1 = ElapsedTime
14        For i = 1 To NumCallsOnePlusOne
15            Res = JuliaEval("1+1")
16        Next i
17        t2 = ElapsedTime

18        Report = Report & "JuliaExcel Version = " & CStr(shAudit.Range("Headers").Cells(2, 1).Value) & vbLf
19        Report = Report & "Computer = " & Environ$("ComputerName") & vbLf
20        Report = Report & "Latency test" & vbLf
21        Report = Report & "Average time for JuliaEval(""1+1"") = " & CStr(1000 * (t2 - t1) / NumCallsOnePlusOne) & _
              " miliseconds (averaged over " & CStr(NumCallsOnePlusOne) & " calls)" & vbLf

22        InputData = Application.Evaluate("=RANDARRAY(" & VectorLength & ")")

          'Data transport tests. One Select Case per iteration - covering how to time the call, the
          'report text, and (for "identity" only) the correctness check - rather than testing
          'JuliaFunction against the same string literals in several separate If/ElseIf blocks.
23        For j = 1 To 4
24            JuliaFunction = Choose(j, "identity", "sum", "collect", "range")

25            Select Case JuliaFunction
                  Case "identity"
26                    t1 = ElapsedTime()
27                    For i = 1 To NumCallsVectors
28                        Res = JuliaCall(JuliaFunction, InputData)
29                    Next i
30                    t2 = ElapsedTime()
31                    If Not ArraysIdentical(Res, InputData) Then Throw "Ohoh, return from Julia function identity is not equal to its input"
32                    WhatWasExecuted = "JuliaCall(""identity"", vector of " & Format(VectorLength, "###,###") & " doubles)"
33                    Report = Report & "Two-way data transport test" & vbLf

34                Case "sum"
35                    t1 = ElapsedTime()
36                    For i = 1 To NumCallsVectors
37                        Res = JuliaCall(JuliaFunction, InputData)
38                    Next i
39                    t2 = ElapsedTime()
40                    WhatWasExecuted = "JuliaCall(""sum"", vector of " & Format(VectorLength, "###,###") & " doubles)"
41                    Report = Report & "One-way data transport test, Excel to Julia" & vbLf

42                Case "collect"
43                    t1 = ElapsedTime()
44                    For i = 1 To NumCallsVectors
45                        Res = JuliaEval("collect((1:" & VectorLength & ").*pi)")
46                    Next i
47                    t2 = ElapsedTime()
48                    WhatWasExecuted = "JuliaEval(""collect((1:" & CStr(VectorLength) & ").*pi)"")"
49                    Report = Report & "One-way data transport test, Julia to Excel" & vbLf

50                Case "range"
51                    t1 = ElapsedTime()
52                    For i = 1 To NumCallsVectors
53                        Res = JuliaEval("(1:" & VectorLength & ").*pi")
54                    Next i
55                    t2 = ElapsedTime()
56                    WhatWasExecuted = "JuliaEval(""(1:" & CStr(VectorLength) & ").*pi"")"
57                    Report = Report & "One-way data transport (AbstractRange), Julia to Excel" & vbLf
58            End Select

59            Report = Report & "Average time for " & WhatWasExecuted & " = " & _
                  CStr((t2 - t1) / NumCallsVectors) & " seconds (averaged over " & CStr(NumCallsVectors) & " calls)" & vbLf
60        Next j

61        Debug.Print "'" & Replace(Report, vbLf, vbLf & "'")
62        AppActivate Application.Caption
63        PerformanceTest = Report
64        Exit Function
ErrHandler:
65        PerformanceTest = ReThrow("PerformanceTest", Err, True)
End Function

'--------------------------------------------------
'Running method SerialisationPerformanceTest
'========================================================================================================================
'Running method SerialisationPerformanceTest
'Time now = 2026-08-16 15:00:28
'JuliaExcel Version = 138
'Computer = MSI
'Average time for SerialiseElement(vector of 100,000 Doubles) = 0.072152280001319 seconds (averaged over 10 calls)
'Average time for UnserialiseFromString(vector of 100,000 Doubles) = 0.102856119998614 seconds (averaged over 10 calls)
'========================================================================================================================
'Running method SerialisationPerformanceTest
'Time now = 2026-08-16 15:00:55
'JuliaExcel Version = 138
'Computer = MSI
'Average time for SerialiseElement(vector of 100,000 Doubles) = 7.19601280003553E-02 seconds (averaged over 50 calls)
'Average time for UnserialiseFromString(vector of 100,000 Doubles) = 0.222831700000097 seconds (averaged over 50 calls)
'========================================================================================================================
'Running method SerialisationPerformanceTest
'Time now = 2026-08-16 15:01:22
'JuliaExcel Version = 138
'Computer = MSI
'Average time for SerialiseElement(vector of 100,000 Doubles) = 7.22701920004329E-02 seconds (averaged over 50 calls)
'Average time for UnserialiseFromString(vector of 100,000 Doubles) = 0.22391951800033 seconds (averaged over 50 calls)
'========================================================================================================================

'SerialisationPerformanceTest (timed SerialiseElement/UnserialiseFromString round trip, general
'format framing, no HTTP) removed 2026-08-17: since TrySerialiseArrayAsV/decode_xl_array_v put pure-
'Double arrays through the fast "V" path automatically, this test's numbers no longer mean what the
'logged runs above them recorded (pre-"V"), and the comparison it existed for is now made explicitly,
'and more precisely, by VFormatDecodeSpeedTest below (decode) and by the encode side of
'TrySerialiseArrayAsV's own docstring measurements (modSerialise.bas).

Function VFormatDecodeSpeedTest() As String
          ' Isolates VBA-side decode cost only, comparing the general "*" array format against the
          ' new "V" format for the SAME underlying data - both wire strings are fetched once from
          ' Julia up front, so there's no HTTP/Julia-side noise inside the timed loop, and decoding
          ' is interleaved (one "*" decode, then one "V" decode, repeat) per the reasoning discussed
          ' elsewhere in this module: interleaving keeps the comparison fair even if something (heap
          ' pressure, thermal throttling) drifts over the course of the test.
          ' Returns the report as a String (as well as Debug.Print-ing it) so it can be captured
          ' programmatically, e.g. via Application.Run from outside VBA.
          Const NumCalls As Long = 50
          Const VectorLength As Long = 100000
          Dim GeneralString As String
          Dim i As Long
          Dim Report As String
          Dim ResGeneral As Variant
          Dim ResV As Variant
          Dim t1 As Double
          Dim t2 As Double
          Dim TotalGeneral As Double
          Dim TotalV As Double
          Dim VString As String

1         On Error GoTo ErrHandler
2         Debug.Print "'" & String(120, "=")
3         Debug.Print "'Running method VFormatDecodeSpeedTest"
4         Debug.Print "'Time now = " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
5         Debug.Print "'JuliaExcel Version = " & CStr(shAudit.Range("Headers").Cells(2, 1).Value)
6         Debug.Print "'Computer = " & Environ$("ComputerName")

7         JuliaEval "exit()"
8         JuliaLaunch , , gTestCommandOptions

          'Fetch both wire-format encodings of the same data directly from Julia, so the timed loop
          'below measures VBA-side decode cost only.
9         GeneralString = ThrowIfError(JuliaEvalVBA("JuliaExcel.encode_array_general(collect((1:" & VectorLength & ").*pi))"))
10        VString = ThrowIfError(JuliaEvalVBA("JuliaExcel.encode_for_xl(collect((1:" & VectorLength & ").*pi))"))

11        If Left$(VString, 1) <> "V" Then Throw "Expected a 'V'-format string but got a string starting '" & Left$(VString, 1) & "' - is the Julia session really running the local dev copy of JuliaExcel (pathof(JuliaExcel))?"

12        For i = 1 To NumCalls
13            t1 = ElapsedTime
14            ResGeneral = UnserialiseFromString(GeneralString, False, GetStringLengthLimit(), True)
15            t2 = ElapsedTime
16            TotalGeneral = TotalGeneral + (t2 - t1)

17            t1 = ElapsedTime
18            ResV = UnserialiseFromString(VString, False, GetStringLengthLimit(), True)
19            t2 = ElapsedTime
20            TotalV = TotalV + (t2 - t1)
21        Next i

22        If Not ArraysIdentical(ResGeneral, ResV) Then Throw "'*' and 'V' format decodes gave different results!"

23        Report = "Average decode time, '*' format (vector of " & Format(VectorLength, "###,###") & " Doubles) = " & _
              CStr(TotalGeneral / NumCalls) & " seconds (averaged over " & NumCalls & " calls, interleaved)" & vbLf & _
              "Average decode time, 'V' format (vector of " & Format(VectorLength, "###,###") & " Doubles) = " & _
              CStr(TotalV / NumCalls) & " seconds (averaged over " & NumCalls & " calls, interleaved)"
24        Debug.Print "'" & Replace(Report, vbLf, vbLf & "'")
25        VFormatDecodeSpeedTest = Report

26        Exit Function
ErrHandler:
27        VFormatDecodeSpeedTest = ReThrow("VFormatDecodeSpeedTest", Err, True)
End Function

Function VFormatEncodeSpeedTest() As String
          ' Diagnostic twin of VFormatDecodeSpeedTest, but for the encode direction: isolates
          ' VBA-side encode cost only (no HTTP, no Julia) by timing the real production
          ' TrySerialiseArrayAsV (modSerialise.bas) directly against a Variant() vector of Doubles
          ' from RANDARRAY, exactly as SerialiseElement receives from Range.Value2.
          ' Added 2026-08-18 to investigate a reported PerformanceTest regression in "sum"/"identity"
          ' (Excel -> Julia encode direction) while "collect" (decode direction) and the "1+1"
          ' latency test stayed flat. Root cause found: TrySerialiseArrayAsV's finiteness check was
          ' calling a separate Function (IsFiniteHex, ByVal String parameter) once per element -
          ' ~80ms of pure VBA function-call/BSTR-copy overhead for a 100,000-element array, roughly
          ' as much as the rest of the encoding combined. Fixed by inlining the check directly at
          ' each call site (see TrySerialiseArrayAsV's own docstring in modSerialise.bas) - keep this
          ' test around as an ongoing regression check for the encode side, mirroring
          ' VFormatDecodeSpeedTest's role on the decode side.
          Const NumCalls As Long = 50
          Const VectorLength As Long = 100000
          Dim EncodedV As String
          Dim i As Long
          Dim InputData As Variant
          Dim OK As Boolean
          Dim Report As String
          Dim t1 As Double
          Dim t2 As Double
          Dim Total As Double

1         On Error GoTo ErrHandler
2         Debug.Print "'" & String(120, "=")
3         Debug.Print "'Running method VFormatEncodeSpeedTest"
4         Debug.Print "'Time now = " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
5         Debug.Print "'JuliaExcel Version = " & CStr(shAudit.Range("Headers").Cells(2, 1).Value)
6         Debug.Print "'Computer = " & Environ$("ComputerName")

7         InputData = Application.Evaluate("=RANDARRAY(" & VectorLength & ")")

8         OK = TrySerialiseArrayAsV(InputData, EncodedV)
9         If Not OK Then Throw "TrySerialiseArrayAsV unexpectedly declined an all-Double RANDARRAY"
10        If Left$(EncodedV, 1) <> "V" Then Throw "Expected a 'V'-format string but got '" & Left$(EncodedV, 1) & "'"

11        For i = 1 To NumCalls
12            t1 = ElapsedTime
13            OK = TrySerialiseArrayAsV(InputData, EncodedV)
14            t2 = ElapsedTime
15            Total = Total + (t2 - t1)
16        Next i

17        Report = "Average encode time, 'V' format, VBA-side only, no HTTP (Variant() vector of " & _
              Format(VectorLength, "###,###") & " Doubles) = " & CStr(Total / NumCalls) & _
              " seconds (averaged over " & NumCalls & " calls)"
18        Debug.Print "'" & Replace(Report, vbLf, vbLf & "'")
19        VFormatEncodeSpeedTest = Report

20        Exit Function
ErrHandler:
21        VFormatEncodeSpeedTest = ReThrow("VFormatEncodeSpeedTest", Err, True)
End Function

'TryFastEncodeDoubleArrayAsV (a PROTOTYPE-only "V" encoder, used solely to measure whether a "V"
'encoder would be worth building) and VEncodeSpeedTest (which timed it against SerialiseElement)
'removed 2026-08-17: the question they existed to answer is settled - the real, NaN/Inf-safe, rank
'1-9 encoder is TrySerialiseArrayAsV (modSerialise.bas), shipped and in production use.
