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

                  Case "sum"
34                    t1 = ElapsedTime()
35                    For i = 1 To NumCallsVectors
36                        Res = JuliaCall(JuliaFunction, InputData)
37                    Next i
38                    t2 = ElapsedTime()
39                    WhatWasExecuted = "JuliaCall(""sum"", vector of " & Format(VectorLength, "###,###") & " doubles)"
40                    Report = Report & "One-way data transport test, Excel to Julia" & vbLf

                  Case "collect"
41                    t1 = ElapsedTime()
42                    For i = 1 To NumCallsVectors
43                        Res = JuliaEval("collect((1:" & VectorLength & ").*pi)")
44                    Next i
45                    t2 = ElapsedTime()
46                    WhatWasExecuted = "JuliaEval(""collect((1:" & CStr(VectorLength) & ").*pi)"")"
47                    Report = Report & "One-way data transport test, Julia to Excel" & vbLf

                  Case "range"
48                    t1 = ElapsedTime()
49                    For i = 1 To NumCallsVectors
50                        Res = JuliaEval("(1:" & VectorLength & ").*pi")
51                    Next i
52                    t2 = ElapsedTime()
53                    WhatWasExecuted = "JuliaEval(""(1:" & CStr(VectorLength) & ").*pi"")"
54                    Report = Report & "One-way data transport (AbstractRange), Julia to Excel" & vbLf
              End Select

55            Report = Report & "Average time for " & WhatWasExecuted & " = " & _
                  CStr((t2 - t1) / NumCallsVectors) & " seconds (averaged over " & CStr(NumCallsVectors) & " calls)" & vbLf
56        Next j

57        Debug.Print "'" & Replace(Report, vbLf, vbLf & "'")
58        AppActivate Application.Caption
59        PerformanceTest = Report
60        Exit Function
ErrHandler:
61        PerformanceTest = ReThrow("PerformanceTest", Err, True)
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


Sub SerialisationPerformanceTest()

          Const VectorLength As Long = 100000
          Const NumCalls As Long = 50
          Dim EncodedString As String
          Dim i As Long
          Dim InputData As Variant
          Dim Res As Variant
          Dim t1 As Double
          Dim t2 As Double

1         On Error GoTo ErrHandler
2         Debug.Print "'" & String(120, "=")
3         Debug.Print "'Running method SerialisationPerformanceTest"
4         Debug.Print "'Time now = " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
5         Debug.Print "'JuliaExcel Version = " & CStr(shAudit.Range("Headers").Cells(2, 1).Value)
6         Debug.Print "'Computer = " & Environ$("ComputerName")

7         InputData = Application.Evaluate("=RANDARRAY(" & VectorLength & ")")

          'Time SerialiseElement (VBA value -> wire format), in isolation - no HTTP call involved.
8         t1 = ElapsedTime
9         For i = 1 To NumCalls
10            EncodedString = SerialiseElement(InputData)
11        Next i
12        t2 = ElapsedTime
13        Debug.Print "'Average time for SerialiseElement(vector of " & Format(VectorLength, "###,###") & " Doubles) = " & _
              CStr((t2 - t1) / NumCalls) & " seconds (averaged over " & CStr(NumCalls) & " calls)"

          'Time UnserialiseFromString (wire format -> VBA value), using the string just produced as a
          'stand-in for what JuliaCall("identity", ...) would return - same shape and size.
14        t1 = ElapsedTime
15        For i = 1 To NumCalls
16            Res = UnserialiseFromString(EncodedString, False, GetStringLengthLimit(), True)
17        Next i
18        t2 = ElapsedTime
19        Debug.Print "'Average time for UnserialiseFromString(vector of " & Format(VectorLength, "###,###") & " Doubles) = " & _
              CStr((t2 - t1) / NumCalls) & " seconds (averaged over " & CStr(NumCalls) & " calls)"

20        If Not ArraysIdentical(Res, InputData) Then
21            Throw "Round trip through SerialiseElement/UnserialiseFromString did not return an identical array"
22        End If

23        Exit Sub
ErrHandler:
24        MsgBox ReThrow("SerialisationPerformanceTest", Err, True)
End Sub

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


' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TryFastEncodeDoubleArrayAsV
' Purpose    : PROTOTYPE ONLY - not wired into SerialiseElement (modSerialise.bas), which currently
'              has no "V"-format encoder at all (the Excel -> Julia direction only ever uses the
'              general "*" format). Measures whether a "V" encoder would be worth building, before
'              committing to one - see VEncodeSpeedTest below.
'              Single-pass, optimistic: checks VarType per element AS it appends hex to a buffer
'              array (joined once at the end via VBA.Join$, matching SerialiseElement's own
'              approach - a naive "Buf = Buf & ..." loop would risk O(n^2) string reallocation and
'              give a misleadingly slow result). Returns False (with EncodedV left unset) if any
'              element isn't a Double, so the caller can fall back to SerialiseElement - mirrors
'              the same optimistic-single-pass design as TryFastDecodeDoubleVector (formerly in the
'              now-deleted modUnserialiseExperimental.bas) on the decode side.
'              Does NOT check for NaN/Inf (unlike the real Julia-side V encoder) - deliberately
'              simplified since RANDARRAY-sourced benchmark data never contains them; a production
'              version would need that check added.
' -----------------------------------------------------------------------------------------------------------------------
Function TryFastEncodeDoubleArrayAsV(ByVal x As Variant, ByRef EncodedV As String) As Boolean
          Dim Chunks() As String
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim n As Long
          Dim NC As Long
          Dim NR As Long

1         TryFastEncodeDoubleArrayAsV = False

2         Select Case NumDimensions(x)
              Case 1
3                 n = UBound(x) - LBound(x) + 1
4                 If n = 0 Then Exit Function
5                 ReDim Chunks(1 To n)
6                 k = 1
7                 For i = LBound(x) To UBound(x)
8                     If VarType(x(i)) <> vbDouble Then Exit Function
9                     Chunks(k) = DoubleToHex(CDbl(x(i)))
10                    k = k + 1
11                Next i
12                EncodedV = "V1," & CStr(n) & ";" & VBA.Join$(Chunks, "")
13                TryFastEncodeDoubleArrayAsV = True

14            Case 2
15                NR = UBound(x, 1) - LBound(x, 1) + 1
16                NC = UBound(x, 2) - LBound(x, 2) + 1
17                If NR = 0 Or NC = 0 Then Exit Function
18                ReDim Chunks(1 To NR * NC)
19                k = 1
20                For j = LBound(x, 2) To UBound(x, 2)    ' column-major to match Julia
21                    For i = LBound(x, 1) To UBound(x, 1)
22                        If VarType(x(i, j)) <> vbDouble Then Exit Function
23                        Chunks(k) = DoubleToHex(CDbl(x(i, j)))
24                        k = k + 1
25                    Next i
26                Next j
27                EncodedV = "V2," & CStr(NR) & "," & CStr(NC) & ";" & VBA.Join$(Chunks, "")
28                TryFastEncodeDoubleArrayAsV = True
29        End Select
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : VEncodeSpeedTest
' Purpose    : Compares the current general-format SerialiseElement (modSerialise.bas) against the
'              prototype TryFastEncodeDoubleArrayAsV above, for a Variant() array of Doubles from
'              RANDARRAY via Application.Evaluate. Deliberately Variant(), not a genuinely-typed
'              Double() array: Range.Value2 (how real worksheet data actually arrives) is always
'              Variant(), even when every cell holds a number, so VarType has to be checked per
'              element rather than being free - this is the realistic, "worst case" scenario for a
'              fast encoder, not the easy case.
'              Confirms both encodings decode (via the existing production Case 86 'V' decoder) to
'              the same result before trusting the timing, and times each pair by interleaving
'              individual calls (call fn1 once, call fn2 once, repeat) rather than N calls to fn1
'              followed by N calls to fn2, per this module's established methodology (see
'              SpeedTestHexConversionLE's historical note, now removed but logged above).
' -----------------------------------------------------------------------------------------------------------------------
Function VEncodeSpeedTest() As String
          Const NumCalls As Long = 50
          Const VectorLength As Long = 100000
          Dim EncodedGeneral As String
          Dim EncodedV As String
          Dim i As Long
          Dim InputData As Variant
          Dim OK As Boolean
          Dim Report As String
          Dim t1 As Double
          Dim t2 As Double
          Dim TotalGeneral As Double
          Dim TotalV As Double

1         On Error GoTo ErrHandler
2         Debug.Print "'" & String(120, "=")
3         Debug.Print "'Running method VEncodeSpeedTest"
4         Debug.Print "'Time now = " & Format$(Now(), "yyyy-mm-dd hh:mm:ss")
5         Debug.Print "'JuliaExcel Version = " & CStr(shAudit.Range("Headers").Cells(2, 1).Value)
6         Debug.Print "'Computer = " & Environ$("ComputerName")

7         InputData = Application.Evaluate("=RANDARRAY(" & VectorLength & ")")

8         OK = TryFastEncodeDoubleArrayAsV(InputData, EncodedV)
9         If Not OK Then Throw "TryFastEncodeDoubleArrayAsV unexpectedly failed on an all-Double RANDARRAY"
10        EncodedGeneral = SerialiseElement(InputData)
11        If Not ArraysIdentical( _
              UnserialiseFromString(EncodedGeneral, False, GetStringLengthLimit(), True), _
              UnserialiseFromString(EncodedV, False, GetStringLengthLimit(), True)) Then
12            Throw "Prototype 'V'-encoded data does not decode to the same result as the general format - do not trust the timing below"
13        End If

14        For i = 1 To NumCalls
15            t1 = ElapsedTime
16            EncodedGeneral = SerialiseElement(InputData)
17            t2 = ElapsedTime
18            TotalGeneral = TotalGeneral + (t2 - t1)

19            t1 = ElapsedTime
20            OK = TryFastEncodeDoubleArrayAsV(InputData, EncodedV)
21            t2 = ElapsedTime
22            TotalV = TotalV + (t2 - t1)
23        Next i

24        Report = "Average encode time, general '*' format (Variant() vector of " & Format(VectorLength, "###,###") & " Doubles) = " & _
              CStr(TotalGeneral / NumCalls) & " seconds (averaged over " & NumCalls & " calls, interleaved)" & vbLf & _
              "Average encode time, prototype 'V' format (same data) = " & _
              CStr(TotalV / NumCalls) & " seconds (averaged over " & NumCalls & " calls, interleaved)"
25        Debug.Print "'" & Replace(Report, vbLf, vbLf & "'")
26        VEncodeSpeedTest = Report

27        Exit Function
ErrHandler:
28        VEncodeSpeedTest = ReThrow("VEncodeSpeedTest", Err, True)
End Function

Function GetADict() As Scripting.Dictionary
Dim out As New Scripting.Dictionary

out.Add "a", 1
out.Add "b", 2
out.Add "c", 3

Set GetADict = out

End Function


Function GetANumber()
GetANumber = 1

End Function




