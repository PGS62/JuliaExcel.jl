Attribute VB_Name = "modHexBulkPrototype"
' Copyright (c) 2021-2026 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme

' SCRATCH PROTOTYPE - not wired into production (TrySerialiseArrayAsV, modSerialise.bas, remains
' the real "V"-format encoder). Explores whether encoding a Double() array to the wire format's
' big-endian hex payload can be sped up by:
'   (a) one bulk RtlMoveMemory call to copy the array's raw bytes into a Byte() buffer, instead of
'       reinterpreting one Double at a time via LSet (as DoubleToHex, modUnserialise.bas, does) -
'       this also eliminates the per-element function-call overhead of invoking DoubleToHex N
'       times, which the IsFiniteHex investigation (see TrySerialiseArrayAsV's docstring,
'       modSerialise.bas) showed to be a genuinely dominant cost, not a negligible one;
'   (b) a bigger lookup table (65,536 entries, 2 bytes -> 4 hex chars) for the byte-to-hex step,
'       halving the number of lookups/concatenations versus the existing 256-entry (1 byte -> 2 hex
'       chars) table.
' A Double's 8 bytes are stored little-endian in memory on Windows; the wire format is big-endian
' (matching DoubleToHex), so each 8-byte group must be read/combined in reverse order.
' Run BulkHexBenchmark to compare timings (it verifies correctness first, via
' TestBulkHexCorrectness, before trusting any timing).
Option Explicit

#If VBA7 Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
#End If

Private HexByte(0 To 255) As String
Private HexByteBuilt As Boolean
Private HexPair(0 To 65535) As String
Private HexPairBuilt As Boolean

Private Sub BuildHexByteTable()
          Dim i As Long
1         If HexByteBuilt Then Exit Sub
2         For i = 0 To 255
3             HexByte(i) = Right$("0" & Hex$(i), 2)
4         Next i
5         HexByteBuilt = True
End Sub

Private Sub BuildHexPairTable()
          Dim i As Long
1         If HexPairBuilt Then Exit Sub
2         For i = 0 To 65535
3             HexPair(i) = Right$("000" & Hex$(i), 4)
4         Next i
5         HexPairBuilt = True
End Sub

' Baseline: the same technique as production TrySerialiseArrayAsV's Case 1 (modSerialise.bas), but
' against a genuinely-typed Double() array (no VarType checks) - isolates the memory/hex-conversion
' technique itself from the Variant-checking overhead, which is a separate, already-understood cost.
Function ReferenceEncode(ByRef x() As Double) As String
          Dim Chunks() As String
          Dim i As Long
          Dim k As Long
          Dim n As Long

1         n = UBound(x) - LBound(x) + 1
2         ReDim Chunks(1 To n)
3         k = 1
4         For i = LBound(x) To UBound(x)
5             Chunks(k) = DoubleToHex(x(i))
6             k = k + 1
7         Next i
8         ReferenceEncode = VBA.Join$(Chunks, "")
End Function

' Bulk variant A: one CopyMemory call, then a per-byte (256-entry table) hex lookup - isolates the
' saving from replacing N per-element LSet + function-call operations with one bulk memory copy,
' while keeping the same per-byte hex-lookup granularity as the reference/production approach.
Function BulkEncodeA(ByRef x() As Double) As String
          Dim base As Long
          Dim bytes() As Byte
          Dim Chunks() As String
          Dim i As Long
          Dim n As Long

1         BuildHexByteTable
2         n = UBound(x) - LBound(x) + 1
3         ReDim bytes(1 To n * 8)
4         CopyMemory bytes(1), x(LBound(x)), n * 8

5         ReDim Chunks(1 To n)
6         For i = 1 To n
7             base = (i - 1) * 8
8             Chunks(i) = HexByte(bytes(base + 8)) & HexByte(bytes(base + 7)) & HexByte(bytes(base + 6)) & HexByte(bytes(base + 5)) & _
                  HexByte(bytes(base + 4)) & HexByte(bytes(base + 3)) & HexByte(bytes(base + 2)) & HexByte(bytes(base + 1))
9         Next i
10        BulkEncodeA = VBA.Join$(Chunks, "")
End Function

' Bulk variant B: as BulkEncodeA, but combines bytes in pairs (big-endian order) and looks each
' pair up in a 65,536-entry table (2 bytes -> 4 hex chars), halving the number of lookups.
Function BulkEncodeB(ByRef x() As Double) As String
          Dim base As Long
          Dim bytes() As Byte
          Dim Chunks() As String
          Dim i As Long
          Dim n As Long

1         BuildHexPairTable
2         n = UBound(x) - LBound(x) + 1
3         ReDim bytes(1 To n * 8)
4         CopyMemory bytes(1), x(LBound(x)), n * 8

5         ReDim Chunks(1 To n)
6         For i = 1 To n
7             base = (i - 1) * 8
8             Chunks(i) = HexPair(CLng(bytes(base + 8)) * 256 + bytes(base + 7)) & _
                  HexPair(CLng(bytes(base + 6)) * 256 + bytes(base + 5)) & _
                  HexPair(CLng(bytes(base + 4)) * 256 + bytes(base + 3)) & _
                  HexPair(CLng(bytes(base + 2)) * 256 + bytes(base + 1))
9         Next i
10        BulkEncodeB = VBA.Join$(Chunks, "")
End Function

' Confirms both bulk variants produce byte-identical output to the reference (DoubleToHex-based)
' encoder, across a range of values (zero, negative, fractional, very large, very small) - must
' pass before any timing below is trustworthy.
Function TestBulkHexCorrectness() As Boolean
          Dim Expected As String
          Dim x(1 To 7) As Double

1         On Error GoTo ErrHandler
2         x(1) = 0#
3         x(2) = 1#
4         x(3) = -1#
5         x(4) = 3.14159265358979
6         x(5) = -2.5
7         x(6) = 1.79769313486231E+308    ' near Double max
8         x(7) = 4.94065645841247E-324    ' smallest positive subnormal Double

9         Expected = ReferenceEncode(x)
10        If BulkEncodeA(x) <> Expected Then Throw "BulkEncodeA mismatch"
11        If BulkEncodeB(x) <> Expected Then Throw "BulkEncodeB mismatch"

12        TestBulkHexCorrectness = True
13        Exit Function
ErrHandler:
14        Debug.Print "TestBulkHexCorrectness FAILED: " & Err.Description
15        TestBulkHexCorrectness = False
End Function

' Times ReferenceEncode, BulkEncodeA and BulkEncodeB against the same large Double() array,
' interleaved (one call to each per iteration) to keep the comparison fair. Confirms correctness
' first - a timing result is meaningless if the encoders don't agree.
Function BulkHexBenchmark() As String
          Const NumCalls As Long = 50
          Const VectorLength As Long = 100000
          Dim i As Long
          Dim Report As String
          Dim ResA As String
          Dim ResB As String
          Dim ResRef As String
          Dim t1 As Double
          Dim t2 As Double
          Dim TotalA As Double
          Dim TotalB As Double
          Dim TotalRef As Double
          Dim x() As Double

1         On Error GoTo ErrHandler
2         If Not TestBulkHexCorrectness() Then Throw "TestBulkHexCorrectness failed - not timing untrusted encoders"

3         ReDim x(1 To VectorLength)
4         For i = 1 To VectorLength
5             x(i) = Rnd() * 1000 - 500
6         Next i

7         ResRef = ReferenceEncode(x)
8         ResA = BulkEncodeA(x)
9         ResB = BulkEncodeB(x)
10        If ResA <> ResRef Then Throw "BulkEncodeA disagrees with ReferenceEncode on the benchmark array"
11        If ResB <> ResRef Then Throw "BulkEncodeB disagrees with ReferenceEncode on the benchmark array"

12        For i = 1 To NumCalls
13            t1 = ElapsedTime
14            ResRef = ReferenceEncode(x)
15            t2 = ElapsedTime
16            TotalRef = TotalRef + (t2 - t1)

17            t1 = ElapsedTime
18            ResA = BulkEncodeA(x)
19            t2 = ElapsedTime
20            TotalA = TotalA + (t2 - t1)

21            t1 = ElapsedTime
22            ResB = BulkEncodeB(x)
23            t2 = ElapsedTime
24            TotalB = TotalB + (t2 - t1)
25        Next i

26        Report = "Reference (LSet + DoubleToHex per element)        = " & CStr(TotalRef / NumCalls) & " s" & vbLf & _
              "BulkEncodeA (CopyMemory + 256-entry byte table)   = " & CStr(TotalA / NumCalls) & " s" & vbLf & _
              "BulkEncodeB (CopyMemory + 65536-entry pair table) = " & CStr(TotalB / NumCalls) & " s"
27        Debug.Print "'" & Report
28        BulkHexBenchmark = Report

29        Exit Function
ErrHandler:
30        BulkHexBenchmark = ReThrow("BulkHexBenchmark", Err, True)
End Function

' ------------------------------------------------------------------------------------------------
' Decode side (hex -> Double array) - the direction used for Julia -> Excel, e.g. the "V"-format
' branch of Unserialise (Case 86, modUnserialise.bas), which calls HexToDouble once per element.
' HexToDouble itself already parses each 16-hex-char chunk reasonably efficiently (two
' CLng("&H" & <8 chars>) calls, each consuming 4 bytes at once, not one call per byte) - so unlike
' the encode side, the opportunity here is narrower: eliminate the per-element function-call
' overhead and the N separate LSets, not the hex-parsing itself.
' ------------------------------------------------------------------------------------------------

' Baseline: the same technique as production Unserialise's Case 86 'V' decode (modUnserialise.bas),
' calling HexToDouble once per 16-hex-char chunk.
Function ReferenceDecode(ByVal Chars As String) As Double()
          Dim i As Long
          Dim n As Long
          Dim Result() As Double

1         n = Len(Chars) \ 16
2         ReDim Result(1 To n)
3         For i = 1 To n
4             Result(i) = HexToDouble(Mid$(Chars, (i - 1) * 16 + 1, 16))
5         Next i
6         ReferenceDecode = Result
End Function

' Bulk variant: parses each element's high/low 32-bit halves the same way HexToDouble does
' (CLng("&H" & <8 hex chars>)), but writes them straight into a Long() buffer inline - no per-
' element function call, no per-element LSet - then does ONE CopyMemory reinterpreting that whole
' buffer's bytes directly as a Double() array. Mirrors BulkEncodeA/B's "one bulk memory operation
' instead of N small ones", in reverse. A Double's low 32 bits sit first in memory (little-endian),
' so Raw() stores Lo then Hi for each element, in that order, to match.
Function BulkDecodeA(ByVal Chars As String) As Double()
          Dim base As Long
          Dim i As Long
          Dim n As Long
          Dim Raw() As Long
          Dim Result() As Double

1         n = Len(Chars) \ 16
2         ReDim Raw(1 To n * 2)
3         For i = 1 To n
4             base = (i - 1) * 16
5             Raw(2 * i - 1) = CLng("&H" & Mid$(Chars, base + 9, 8))   ' low 32 bits (last 8 hex chars)
6             Raw(2 * i) = CLng("&H" & Mid$(Chars, base + 1, 8))       ' high 32 bits (first 8 hex chars)
7         Next i

8         ReDim Result(1 To n)
9         CopyMemory Result(1), Raw(1), n * 8

10        BulkDecodeA = Result
End Function

' Confirms BulkDecodeA produces the same values as ReferenceDecode (which itself calls the trusted,
' production HexToDouble), for the same set of edge-case values used by TestBulkHexCorrectness
' above, round-tripped through the (already-verified) encode side to build correct wire hex.
Function TestBulkHexDecodeCorrectness() As Boolean
          Dim Chars As String
          Dim Expected() As Double
          Dim Got() As Double
          Dim i As Long
          Dim x(1 To 7) As Double

1         On Error GoTo ErrHandler
2         x(1) = 0#
3         x(2) = 1#
4         x(3) = -1#
5         x(4) = 3.14159265358979
6         x(5) = -2.5
7         x(6) = 1.79769313486231E+308    ' near Double max
8         x(7) = 4.94065645841247E-324    ' smallest positive subnormal Double

9         Chars = ReferenceEncode(x)
10        Expected = ReferenceDecode(Chars)
11        Got = BulkDecodeA(Chars)

12        For i = 1 To 7
13            If Got(i) <> Expected(i) Then Throw "BulkDecodeA mismatch at element " & i
14        Next i

15        TestBulkHexDecodeCorrectness = True
16        Exit Function
ErrHandler:
17        Debug.Print "TestBulkHexDecodeCorrectness FAILED: " & Err.Description
18        TestBulkHexDecodeCorrectness = False
End Function

' Times ReferenceDecode and BulkDecodeA against the same large wire-format hex string, interleaved,
' after confirming correctness on the actual benchmark data (not just the small edge-case set).
Function BulkHexDecodeBenchmark() As String
          Const NumCalls As Long = 50
          Const VectorLength As Long = 100000
          Dim Chars As String
          Dim i As Long
          Dim Report As String
          Dim ResA() As Double
          Dim ResRef() As Double
          Dim t1 As Double
          Dim t2 As Double
          Dim TotalA As Double
          Dim TotalRef As Double
          Dim x() As Double

1         On Error GoTo ErrHandler
2         If Not TestBulkHexDecodeCorrectness() Then Throw "TestBulkHexDecodeCorrectness failed - not timing untrusted decoders"

3         ReDim x(1 To VectorLength)
4         For i = 1 To VectorLength
5             x(i) = Rnd() * 1000 - 500
6         Next i
7         Chars = ReferenceEncode(x)

8         ResRef = ReferenceDecode(Chars)
9         ResA = BulkDecodeA(Chars)
10        For i = 1 To VectorLength
11            If ResA(i) <> ResRef(i) Then Throw "BulkDecodeA disagrees with ReferenceDecode at element " & i
12        Next i

13        For i = 1 To NumCalls
14            t1 = ElapsedTime
15            ResRef = ReferenceDecode(Chars)
16            t2 = ElapsedTime
17            TotalRef = TotalRef + (t2 - t1)

18            t1 = ElapsedTime
19            ResA = BulkDecodeA(Chars)
20            t2 = ElapsedTime
21            TotalA = TotalA + (t2 - t1)
22        Next i

23        Report = "Reference (HexToDouble per element)         = " & CStr(TotalRef / NumCalls) & " s" & vbLf & _
              "BulkDecodeA (inline parse + one CopyMemory) = " & CStr(TotalA / NumCalls) & " s"
24        Debug.Print "'" & Report
25        BulkHexDecodeBenchmark = Report

26        Exit Function
ErrHandler:
27        BulkHexDecodeBenchmark = ReThrow("BulkHexDecodeBenchmark", Err, True)
End Function

' ------------------------------------------------------------------------------------------------
' Two things the above prototypes gloss over, both stemming from the same fact: a Variant() array
' (what Range.Value2 always is, and what Unserialise's shared Ret() As Variant is declared as) is
' NOT a packed buffer of raw Doubles the way a genuinely-typed Double() array is - each slot is its
' own tagged Variant (16 bytes on 64-bit, not 8). So CopyMemory can't operate directly on either a
' Variant() source (encode) or a Variant() destination (decode) - both need a bridging step to/from
' a genuinely-typed Double() buffer, which is what the two tests below check the cost/feasibility of
' before any of this gets wired into TrySerialiseArrayAsV/Unserialise for real.
' ------------------------------------------------------------------------------------------------

' Does assigning a whole genuinely-typed Double() array to a Variant in one shot (v = MyDoubleArray)
' produce a working, correctly-typed array cheaply (a single SAFEARRAY-wrapping operation), or does
' VBA silently box every element into its own Variant, same as the current per-element
' "Ret(i) = HexToDouble(...)" loop already does? If the former, BulkDecodeA-style results can be
' handed back as Unserialise's return value directly, without ever populating a Variant() array
' element-by-element.
Function TestVariantArrayAssignEfficiency() As String
          Const NumCalls As Long = 50
          Const VectorLength As Long = 100000
          Dim c As Long
          Dim i As Long
          Dim Report As String
          Dim RetLoop() As Variant
          Dim RetWhole As Variant
          Dim SourceD() As Double
          Dim t1 As Double
          Dim t2 As Double
          Dim TotalLoop As Double
          Dim TotalWhole As Double

1         On Error GoTo ErrHandler
2         ReDim SourceD(1 To VectorLength)
3         For i = 1 To VectorLength
4             SourceD(i) = Rnd() * 1000 - 500
5         Next i

          'Correctness check: both approaches should give an array with matching values/VarType.
6         ReDim RetLoop(1 To VectorLength)
7         For i = 1 To VectorLength
8             RetLoop(i) = SourceD(i)
9         Next i
10        RetWhole = SourceD
11        If LBound(RetWhole) <> 1 Or UBound(RetWhole) <> VectorLength Then Throw "RetWhole has unexpected bounds"
12        If RetWhole(1) <> RetLoop(1) Or RetWhole(VectorLength) <> RetLoop(VectorLength) Then Throw "Value mismatch"
13        If VarType(RetLoop(1)) <> vbDouble Or VarType(RetWhole(1)) <> vbDouble Then Throw "Unexpected VarType"

14        For c = 1 To NumCalls
15            t1 = ElapsedTime
16            ReDim RetLoop(1 To VectorLength)
17            For i = 1 To VectorLength
18                RetLoop(i) = SourceD(i)
19            Next i
20            t2 = ElapsedTime
21            TotalLoop = TotalLoop + (t2 - t1)

22            t1 = ElapsedTime
23            RetWhole = SourceD
24            t2 = ElapsedTime
25            TotalWhole = TotalWhole + (t2 - t1)
26        Next c

27        Report = "Per-element assign into Variant() array (today's approach) = " & CStr(TotalLoop / NumCalls) & " s" & vbLf & _
              "Whole-array assign to a single Variant (v = MyDoubleArray)   = " & CStr(TotalWhole / NumCalls) & " s"
28        Debug.Print "'" & Report
29        TestVariantArrayAssignEfficiency = Report

30        Exit Function
ErrHandler:
31        TestVariantArrayAssignEfficiency = ReThrow("TestVariantArrayAssignEfficiency", Err, True)
End Function

' Note: assigning a Double() array to a variable declared Variant() (an array type, as opposed to a
' scalar Variant) is a COMPILE ERROR in VBA ("Can't assign to array") - confirmed empirically (it's
' what blocked Excel while this file still had that line in it). Unlike the scalar-Variant case
' above, there is no way to get the "whole-array, one cheap wrap" assignment behaviour into a
' variable that's explicitly typed as an array of Variants - which is exactly how Unserialise's
' shared Ret variable is declared (modUnserialise.bas). So wiring BulkDecodeA into production means
' either building the V-format branch's own local scalar Variant and returning that directly
' (bypassing Ret and the shared "Unserialise = Ret" line), or falling back to populating Ret
' element-by-element as today (in which case only the hex-parsing side of BulkDecodeA's saving
' applies, not this whole-array-assignment saving).
