Attribute VB_Name = "modSerialise"
' Copyright (c) 2021-2026 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme

Option Explicit
Option Private Module

#If VBA7 Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
#Else
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
#End If

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SerialiseElement
' Purpose    : Encode a single VBA value (scalar or array) into the JuliaExcel wire format.
'              Mirror of Unserialise in modUnserialise.bas, which has the full specification of
'              the wire format (type indicator characters, array/dictionary layout, examples).
'              Arrays are written column-major to match Julia's default array layout and
'              encode_for_xl.
' -----------------------------------------------------------------------------------------------------------------------
Public Function SerialiseElement(ByVal x As Variant) As String

          Dim d As Long
          Dim DictKey As Variant
          Dim Dims() As Long
          Dim DimStr() As String
          Dim Encoded() As String
          Dim EncodedV As String
          Dim i As Long
          Dim Idx() As Long
          Dim j As Long
          Dim k As Long
          Dim Lb() As Long
          Dim Lens() As String
          Dim n As Long
          Dim NC As Long
          Dim NR As Long
          Dim Rank As Long

1         On Error GoTo ErrHandler

2         If IsArray(x) Then
3             Select Case NumDimensions(x)
                  Case 1
4                     n = UBound(x) - LBound(x) + 1
5                     If n = 0 Then
6                         SerialiseElement = "*1,0;;"
7                         Exit Function
8                     End If
                      ' Try the compact "V" format (Case 86 in Unserialise, modUnserialise.bas)
                      ' before the general per-element encoding below - see TrySerialiseArrayAsV.
9                     If TrySerialiseArrayAsV(x, EncodedV) Then
10                        SerialiseElement = EncodedV
11                        Exit Function
12                    End If
13                    ReDim Encoded(1 To n)
14                    ReDim Lens(1 To n)
15                    k = 1
16                    For i = LBound(x) To UBound(x)
17                        Encoded(k) = SerialiseElement(x(i))
18                        Lens(k) = CStr(Len(Encoded(k)))
19                        k = k + 1
20                    Next i
21                    SerialiseElement = "*1," & CStr(n) & ";" & VBA.Join$(Lens, ",") & ",;" & VBA.Join$(Encoded, "")

22                Case 2
23                    NR = UBound(x, 1) - LBound(x, 1) + 1
24                    NC = UBound(x, 2) - LBound(x, 2) + 1
25                    If NR = 0 Or NC = 0 Then
26                        If NC = 1 Then
27                            SerialiseElement = "*1,0;;"
28                        Else
29                            SerialiseElement = "*2," & CStr(NR) & "," & CStr(NC) & ";;"
30                        End If
31                        Exit Function
32                    End If

                      ' As above: try the compact "V" format before the general encoding below.
33                    If TrySerialiseArrayAsV(x, EncodedV) Then
34                        SerialiseElement = EncodedV
35                        Exit Function
36                    End If

37                    ReDim Encoded(1 To NR * NC)
38                    ReDim Lens(1 To NR * NC)
39                    k = 1
40                    For j = LBound(x, 2) To UBound(x, 2)    ' column-major to match Julia
41                        For i = LBound(x, 1) To UBound(x, 1)
42                            Encoded(k) = SerialiseElement(x(i, j))
43                            Lens(k) = CStr(Len(Encoded(k)))
44                            k = k + 1
45                        Next i
46                    Next j
                      ' Nx1 -> 1D Vector (matches JuliaCallOld / README "single-column ranges arrive as vectors").
                      ' 1xN stays as 2D Matrix; 3D+ arrays are left as-is (no prior behaviour to replicate).
47                    If NC = 1 Then
48                        SerialiseElement = "*1," & CStr(NR) & ";" & VBA.Join$(Lens, ",") & ",;" & VBA.Join$(Encoded, "")
49                    Else
50                        SerialiseElement = "*2," & CStr(NR) & "," & CStr(NC) & ";" & VBA.Join$(Lens, ",") & ",;" & VBA.Join$(Encoded, "")
51                    End If

52                Case Else
53                    Rank = NumDimensions(x)
54                    If Rank > 9 Then Throw "Cannot serialise arrays with more than 9 dimensions"
55                    ReDim Dims(1 To Rank)
56                    ReDim Lb(1 To Rank)
57                    ReDim DimStr(1 To Rank)
58                    ReDim Idx(1 To Rank)
59                    n = 1
60                    For i = 1 To Rank
61                        Lb(i) = LBound(x, i)
62                        Dims(i) = UBound(x, i) - Lb(i) + 1
63                        DimStr(i) = CStr(Dims(i))
64                        n = n * Dims(i)
65                    Next i
66                    If n = 0 Then
67                        SerialiseElement = "*" & CStr(Rank) & "," & VBA.Join$(DimStr, ",") & ";;"
68                        Exit Function
69                    End If
                      ' As above: try the compact "V" format before the general encoding below.
70                    If TrySerialiseArrayAsV(x, EncodedV) Then
71                        SerialiseElement = EncodedV
72                        Exit Function
73                    End If
74                    ReDim Encoded(1 To n)
75                    ReDim Lens(1 To n)
76                    For i = 1 To Rank: Idx(i) = Lb(i): Next i
77                    k = 1
78                    Do
79                        Encoded(k) = SerialiseElement(GetAt(x, Idx))
80                        Lens(k) = CStr(Len(Encoded(k)))
81                        k = k + 1
82                        d = 1
83                        Do While d <= Rank
84                            Idx(d) = Idx(d) + 1
85                            If Idx(d) <= UBound(x, d) Then Exit Do
86                            Idx(d) = Lb(d)
87                            d = d + 1
88                        Loop
89                        If d > Rank Then Exit Do
90                    Loop
91                    SerialiseElement = "*" & CStr(Rank) & "," & VBA.Join$(DimStr, ",") & ";" & VBA.Join$(Lens, ",") & ",;" & VBA.Join$(Encoded, "")
92            End Select

93        Else
94            Select Case VarType(x)
                  Case vbDouble:   SerialiseElement = "#" & DoubleToHex(CDbl(x))
95                Case vbString:   SerialiseElement = Chr(163) & CStr(x)      ' Chr(163) = pound sterling sign
96                Case vbBoolean:  SerialiseElement = IIf(CBool(x), "T", "F")
97                Case vbEmpty:    SerialiseElement = "E"
98                Case vbNull:     SerialiseElement = "N"
99                Case vbInteger:  SerialiseElement = "%" & CStr(CInt(x))
100               Case vbLong:     SerialiseElement = "&" & CStr(CLng(x))
101               Case vbByte:     SerialiseElement = "B" & CStr(CByte(x))
102               Case vbSingle:   SerialiseElement = "S" & SingleToHex(CSng(x))
103               Case vbDate
                      ' CDbl of a VBA date gives the Excel serial number directly:
                      ' integer part = days since 1899-12-30, fractional part = time of day.
104                   If CDbl(x) = Int(CDbl(x)) Then
105                       SerialiseElement = "D" & CStr(CLng(CDbl(x)))         ' date only
106                   Else
107                       SerialiseElement = "G" & DoubleToHex(CDbl(x))        ' date + time
108                   End If
109               Case vbError
                      ' CStr(CVErr(n)) = "Error n"; extract the number after the space.
110                   SerialiseElement = "!" & Mid(CStr(x), InStr(CStr(x), " ") + 1)
111               Case vbObject
112                   If TypeName(x) = "Dictionary" Then
113                       n = x.Count
114                       If n = 0 Then
115                           SerialiseElement = "H0;;"
116                           Exit Function
117                       End If
118                       ReDim Encoded(1 To 2 * n)
119                       ReDim Lens(1 To 2 * n)
120                       k = 1
121                       For Each DictKey In x.Keys
122                           Encoded(k) = SerialiseElement(DictKey)
123                           Lens(k) = CStr(Len(Encoded(k)))
124                           k = k + 1
125                           Encoded(k) = SerialiseElement(x(DictKey))
126                           Lens(k) = CStr(Len(Encoded(k)))
127                           k = k + 1
128                       Next DictKey
129                       SerialiseElement = "H" & CStr(n) & ";" & VBA.Join$(Lens, ",") & ",;" & VBA.Join$(Encoded, "")
130                   Else
131                       Throw "Cannot serialise object of type " & TypeName(x)
132                   End If
#If Win64 Then
133               Case vbLongLong: SerialiseElement = "^" & CStr(x)
#End If
134               Case Else
135                   Throw "Cannot serialise VarType=" & CStr(VarType(x))
136           End Select
137       End If

138       Exit Function
ErrHandler:
139       ReThrow "SerialiseElement", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TrySerialiseArrayAsV
' Purpose    : Attempts the compact "V" wire-format encoding (see modUnserialise.bas's wire-format
'              spec comment and Case 86 'V' branch, and encode_for_xl(x::Array{Float64,N}) in
'              src/encode.jl, plus decode_xl_array_v in src/decode.jl which understands "V" strings
'              sent this direction) for a 1- to 9-dimensional array, before SerialiseElement falls
'              back to its own general "*" per-element encoding. Callers are expected to have
'              already handled the empty-array case (n/NR/NC = 0) - this function assumes at least
'              one element. Rank > 9 is not attempted here since it's already excluded by
'              SerialiseElement's own callers (its Case Else throws before reaching this call for
'              Rank > 9), matching VBA's own GetAt/ReDimVariantArray cap elsewhere in the codebase.
'              Single-pass and optimistic: checks VarType per element AS it appends hex to a buffer
'              array (joined once at the end via VBA.Join$, avoiding the O(n^2) reallocation risk of
'              repeated string concatenation) - returns False (with EncodedV left unset) the instant
'              any element isn't a Double, so the caller falls back to the general per-element
'              encoding unchanged. Mirrors the same optimistic single-pass design
'              TryFastDecodeDoubleVector (formerly in the now-deleted modUnserialiseExperimental.bas)
'              used on the decode side.
'              Rank 3-9 walks the array via GetAt/Idx() exactly as SerialiseElement's own Case Else
'              does for the general format, just without any per-element length lookup, since every
'              element is always exactly 16 hex characters.
'              Measured (via the now-removed prototype benchmark VEncodeSpeedTest, modPerformance.bas)
'              at roughly 40% faster than the general format even for the realistic worst case - a
'              Variant() array (as Range.Value2 always is, even when every cell holds a number)
'              rather than a genuinely-typed Double() array, so VarType must be checked per element
'              rather than being free.
'              No NaN/Inf check: an earlier version declined (fell back to general) for any element
'              that wasn't a finite Double, using a Function IsFiniteHex to test the hex-encoded
'              exponent bits (a real NaN-bit-pattern crash - see modUnserialise.bas's wire-format
'              history - had ruled out ever comparing the raw Double value, e.g. "v <> v", directly).
'              That check was removed 2026-08-18 once shown unnecessary: SerialiseElement's own
'              general per-scalar Double encoding ("#" & DoubleToHex(CDbl(x)), no NaN/Inf handling of
'              any kind) produces the exact same bit pattern the "V" path would, and Julia's decode
'              (decode_xl_array_v, or hex_to_float64 on the general path) reconstructs the identical
'              Float64 either way - so which path an Excel-sourced NaN/Inf Double takes never changed
'              what Julia received. That's different from the Julia -> Excel direction
'              (encode_for_xl(::Float64), src/encode.jl), where NaN/Inf genuinely must be translated
'              to Excel error values (Excel has no native representation for them) - that translation
'              only happens in the general per-scalar path there, which is why the Julia-side "V"
'              array encoder still excludes NaN/Inf and falls back to general. Removing the check
'              also removed the ~80ms/100,000-element cost of testing for it (see VFormatEncodeSpeedTest,
'              modPerformance.bas).
'              Per-element work here is now limited to the VarType check and copying the value into
'              Buf(), a genuinely-typed Double() array - the hex encoding itself happens once, in
'              bulk, via BulkHexOfDoubleArray below, rather than one DoubleToHex call per element.
'              See BulkHexOfDoubleArray's own docstring for why, and modHexBulkPrototype.bas (a
'              scratch, not-wired-in module) for the benchmark that measured it: ~30% faster than
'              calling DoubleToHex per element, for a 100,000-element array.
' -----------------------------------------------------------------------------------------------------------------------
Function TrySerialiseArrayAsV(ByVal x As Variant, ByRef EncodedV As String) As Boolean
          Dim Buf() As Double
          Dim Dims() As Long
          Dim DimStr() As String
          Dim El As Variant
          Dim i As Long
          Dim Idx() As Long
          Dim j As Long
          Dim k As Long
          Dim n As Long
          Dim NC As Long
          Dim NR As Long
          Dim q As Long
          Dim Rank As Long
          Dim Total As Long

1         On Error GoTo ErrHandler
2         TrySerialiseArrayAsV = False

3         Select Case NumDimensions(x)
              Case 1
4                 n = UBound(x) - LBound(x) + 1
5                 ReDim Buf(1 To n)
6                 k = 1
7                 For i = LBound(x) To UBound(x)
8                     If VarType(x(i)) <> vbDouble Then Exit Function
9                     Buf(k) = CDbl(x(i))
10                    k = k + 1
11                Next i
12                EncodedV = "V1," & CStr(n) & ";" & BulkHexOfDoubleArray(Buf)
13                TrySerialiseArrayAsV = True

14            Case 2
15                NR = UBound(x, 1) - LBound(x, 1) + 1
16                NC = UBound(x, 2) - LBound(x, 2) + 1
17                ReDim Buf(1 To NR * NC)
18                k = 1
19                For j = LBound(x, 2) To UBound(x, 2)    ' column-major to match Julia
20                    For i = LBound(x, 1) To UBound(x, 1)
21                        If VarType(x(i, j)) <> vbDouble Then Exit Function
22                        Buf(k) = CDbl(x(i, j))
23                        k = k + 1
24                    Next i
25                Next j
                  ' Nx1 -> 1D Vector, matching SerialiseElement's own Nx1 collapsing.
26                If NC = 1 Then
27                    EncodedV = "V1," & CStr(NR) & ";" & BulkHexOfDoubleArray(Buf)
28                Else
29                    EncodedV = "V2," & CStr(NR) & "," & CStr(NC) & ";" & BulkHexOfDoubleArray(Buf)
30                End If
31                TrySerialiseArrayAsV = True

              Case Else
32                Rank = NumDimensions(x)
33                ReDim Dims(1 To Rank)
34                ReDim DimStr(1 To Rank)
35                ReDim Idx(1 To Rank)
36                Total = 1
37                For q = 1 To Rank
38                    Dims(q) = UBound(x, q) - LBound(x, q) + 1
39                    DimStr(q) = CStr(Dims(q))
40                    Idx(q) = LBound(x, q)
41                    Total = Total * Dims(q)
42                Next q
43                ReDim Buf(1 To Total)
44                k = 1
45                Do
46                    El = GetAt(x, Idx)
47                    If VarType(El) <> vbDouble Then Exit Function
48                    Buf(k) = CDbl(El)
49                    k = k + 1
50                    q = 1
51                    Do While q <= Rank
52                        Idx(q) = Idx(q) + 1
53                        If Idx(q) <= UBound(x, q) Then Exit Do
54                        Idx(q) = LBound(x, q)
55                        q = q + 1
56                    Loop
57                    If q > Rank Then Exit Do
58                Loop
59                EncodedV = "V" & CStr(Rank) & "," & VBA.Join$(DimStr, ",") & ";" & BulkHexOfDoubleArray(Buf)
60                TrySerialiseArrayAsV = True
61        End Select

62        Exit Function
ErrHandler:
63        ReThrow "TrySerialiseArrayAsV", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : BulkHexOfDoubleArray
' Purpose    : Encodes a genuinely-typed Double() array as the "V" format's big-endian hex payload
'              (16 hex characters per element, no delimiters) - the same bit-for-bit encoding
'              DoubleToHex (modUnserialise.bas) produces per element, but for a whole array in one
'              pass. One bulk RtlMoveMemory ("CopyMemory") call copies the array's raw bytes into a
'              Byte() buffer, replacing N per-element LSet reinterpretations with a single memory
'              copy; each element's 8 bytes are then mapped through a static 256-entry lookup table
'              (built once, matching DoubleToHex's own table) to build its 16-character hex chunk.
'              This avoids the per-element function-call overhead of invoking DoubleToHex N times,
'              which a throwaway prototype (modHexBulkPrototype.bas, not wired into production, kept
'              purely as a record of the benchmark) measured as a genuinely dominant cost - not a
'              negligible one - worth roughly 30% for a 100,000-element array.
'              A Double's 8 bytes are little-endian in memory on Windows; the wire format is
'              big-endian (matching DoubleToHex), so each element's 8-byte group is read in reverse.
'              Callers are expected to pass a non-empty array (Buf must have at least one element).
'              Safety: CopyMemory is a raw memory copy - unlike an ordinary VBA error, a wrong length
'              argument here could corrupt memory or crash Excel outright, not just fail this call.
'              So immediately before calling it, both buffers' actual byte sizes are re-derived
'              independently from their own LBound/UBound (not by trusting the "n" arithmetic that
'              sized them) and compared; any mismatch throws a normal, catchable error instead of
'              proceeding. This is deliberately a hard failure, not a silent fallback to the old
'              per-element DoubleToHex loop: if this ever fires, something is genuinely wrong with
'              this function's own logic, and a fallback path that (if that logic is correct) never
'              executes in practice would itself be an untested, silently bit-rotting liability.
' -----------------------------------------------------------------------------------------------------------------------
Private Function BulkHexOfDoubleArray(ByRef Buf() As Double) As String
          Static HexByte(0 To 255) As String
          Static Initialized As Boolean
          Dim base As Long
          Dim BufBytes As Long
          Dim bytes() As Byte
          Dim BytesBytes As Long
          Dim Chunks() As String
          Dim i As Long
          Dim n As Long

1         If Not Initialized Then
2             For i = 0 To 255
3                 HexByte(i) = Right$("0" & Hex$(i), 2)
4             Next i
5             Initialized = True
6         End If

7         n = UBound(Buf) - LBound(Buf) + 1
8         If n <= 0 Then Throw "BulkHexOfDoubleArray requires a non-empty array"
9         ReDim bytes(1 To n * 8)

10        BufBytes = (UBound(Buf) - LBound(Buf) + 1) * 8
11        BytesBytes = (UBound(bytes) - LBound(bytes) + 1)
12        If BufBytes <> BytesBytes Then Throw "BulkHexOfDoubleArray: source is " & BufBytes & _
              " bytes but destination buffer is " & BytesBytes & " bytes - refusing to call CopyMemory"
13        CopyMemory bytes(1), Buf(LBound(Buf)), BytesBytes

14        ReDim Chunks(1 To n)
15        For i = 1 To n
16            base = (i - 1) * 8
17            Chunks(i) = HexByte(bytes(base + 8)) & HexByte(bytes(base + 7)) & HexByte(bytes(base + 6)) & HexByte(bytes(base + 5)) & _
                  HexByte(bytes(base + 4)) & HexByte(bytes(base + 3)) & HexByte(bytes(base + 2)) & HexByte(bytes(base + 1))
18        Next i

19        BulkHexOfDoubleArray = VBA.Join$(Chunks, "")
End Function

