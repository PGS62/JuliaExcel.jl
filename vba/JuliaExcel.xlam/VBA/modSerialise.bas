Attribute VB_Name = "modSerialise"
' Copyright (c) 2021-2026 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme

Option Explicit
Option Private Module

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
101               Case vbSingle:   SerialiseElement = "S" & SingleToHex(CSng(x))
102               Case vbDate
                      ' CDbl of a VBA date gives the Excel serial number directly:
                      ' integer part = days since 1899-12-30, fractional part = time of day.
103                   If CDbl(x) = Int(CDbl(x)) Then
104                       SerialiseElement = "D" & CStr(CLng(CDbl(x)))         ' date only
105                   Else
106                       SerialiseElement = "G" & DoubleToHex(CDbl(x))        ' date + time
107                   End If
108               Case vbError
                      ' CStr(CVErr(n)) = "Error n"; extract the number after the space.
109                   SerialiseElement = "!" & Mid(CStr(x), InStr(CStr(x), " ") + 1)
110               Case vbObject
111                   If TypeName(x) = "Dictionary" Then
112                       n = x.Count
113                       If n = 0 Then
114                           SerialiseElement = "H0;;"
115                           Exit Function
116                       End If
117                       ReDim Encoded(1 To 2 * n)
118                       ReDim Lens(1 To 2 * n)
119                       k = 1
120                       For Each DictKey In x.Keys
121                           Encoded(k) = SerialiseElement(DictKey)
122                           Lens(k) = CStr(Len(Encoded(k)))
123                           k = k + 1
124                           Encoded(k) = SerialiseElement(x(DictKey))
125                           Lens(k) = CStr(Len(Encoded(k)))
126                           k = k + 1
127                       Next DictKey
128                       SerialiseElement = "H" & CStr(n) & ";" & VBA.Join$(Lens, ",") & ",;" & VBA.Join$(Encoded, "")
129                   Else
130                       Throw "Cannot serialise object of type " & TypeName(x)
131                   End If
#If Win64 Then
132               Case vbLongLong: SerialiseElement = "^" & CStr(x)
#End If
133               Case Else
134                   Throw "Cannot serialise VarType=" & CStr(VarType(x))
135           End Select
136       End If

137       Exit Function
ErrHandler:
138       ReThrow "SerialiseElement", Err
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
'              Single-pass and optimistic: checks VarType and finiteness per element AS it appends
'              hex to a buffer array (joined once at the end via VBA.Join$, avoiding the O(n^2)
'              reallocation risk of repeated string concatenation) - returns False (with EncodedV
'              left unset) the instant any element isn't a finite Double, so the caller falls back
'              to the general per-element encoding unchanged. Mirrors the same optimistic
'              single-pass design TryFastDecodeDoubleVector (formerly in the now-deleted
'              modUnserialiseExperimental.bas) used on the decode side.
'              Rank 3-9 walks the array via GetAt/Idx() exactly as SerialiseElement's own Case Else
'              does for the general format, just without any per-element length lookup, since every
'              element is always exactly 16 hex characters.
'              Measured (via the now-removed prototype benchmark VEncodeSpeedTest, modPerformance.bas)
'              at roughly 40% faster than the general format even for the realistic worst case - a
'              Variant() array (as
'              Range.Value2 always is, even when every cell holds a number) rather than a
'              genuinely-typed Double() array, so VarType must be checked per element rather than
'              being free.
' -----------------------------------------------------------------------------------------------------------------------
' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TrySerialiseArrayAsV (finiteness check)
' Note       : The NaN/Inf-safe finiteness check (see the removed IsFiniteHex's historical note
'              below) is inlined directly at each of the three call sites below, rather than called
'              as a separate Function - measured (FunctionCallOverheadTest, modPerformance.bas,
'              2026-08-18) to cost ~80ms of pure VBA function-call/BSTR-copy overhead alone for a
'              100,000-element array, i.e. roughly as much as the rest of the encoding put together.
'              This was a real, measured regression (PerformanceTest's "sum"/"identity" cases going
'              from ~0.07s/0.15s to ~0.16s/0.25s - worse than the pre-"V" baseline) traced back to
'              the original fix for the NaN-crash incident: correct, but paid for as a separate
'              Function call on every element.
'              Historical note on WHY the check exists at all (previously in IsFiniteHex's own
'              docstring): True unless the 16-character big-endian hex string (from DoubleToHex,
'              modUnserialise.bas) represents NaN or +-Infinity - i.e. unless its 11-bit exponent
'              field is all-ones. Works purely on the hex STRING, never on the Double value itself:
'              VBA's own comparison operators (<>, >, <) raise a runtime "Overflow" error when
'              applied to a genuine NaN bit pattern, rather than returning a Boolean as IEEE-754
'              unordered-comparison semantics would - found the hard way (an early version of this
'              function used "v <> v", which crashed the VBA project badly enough to require
'              reverting the workbook from git). DoubleToHex itself is safe regardless of the bit
'              pattern - it only does byte reinterpretation via LSet, no comparisons - so routing
'              through it first and inspecting the resulting hex text avoids the dangerous operators
'              entirely. The exponent occupies the low 3 bits of the first hex digit (that digit's
'              top bit is the sign) plus all of the second and third hex digits - all-ones there
'              means the first digit is "7" or "F" and the next two are "FF".
' -----------------------------------------------------------------------------------------------------------------------
Function TrySerialiseArrayAsV(ByVal x As Variant, ByRef EncodedV As String) As Boolean
          Dim Chunks() As String
          Dim Dims() As Long
          Dim DimStr() As String
          Dim El As Variant
          Dim FirstDigit As String
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
5                 ReDim Chunks(1 To n)
6                 k = 1
7                 For i = LBound(x) To UBound(x)
8                     If VarType(x(i)) <> vbDouble Then Exit Function
9                     Chunks(k) = DoubleToHex(CDbl(x(i)))
10                    FirstDigit = Mid$(Chunks(k), 1, 1)
11                    If (FirstDigit = "7" Or FirstDigit = "F") And Mid$(Chunks(k), 2, 2) = "FF" Then Exit Function
12                    k = k + 1
13                Next i
14                EncodedV = "V1," & CStr(n) & ";" & VBA.Join$(Chunks, "")
15                TrySerialiseArrayAsV = True

16            Case 2
17                NR = UBound(x, 1) - LBound(x, 1) + 1
18                NC = UBound(x, 2) - LBound(x, 2) + 1
19                ReDim Chunks(1 To NR * NC)
20                k = 1
21                For j = LBound(x, 2) To UBound(x, 2)    ' column-major to match Julia
22                    For i = LBound(x, 1) To UBound(x, 1)
23                        If VarType(x(i, j)) <> vbDouble Then Exit Function
24                        Chunks(k) = DoubleToHex(CDbl(x(i, j)))
25                        FirstDigit = Mid$(Chunks(k), 1, 1)
26                        If (FirstDigit = "7" Or FirstDigit = "F") And Mid$(Chunks(k), 2, 2) = "FF" Then Exit Function
27                        k = k + 1
28                    Next i
29                Next j
                  ' Nx1 -> 1D Vector, matching SerialiseElement's own Nx1 collapsing.
30                If NC = 1 Then
31                    EncodedV = "V1," & CStr(NR) & ";" & VBA.Join$(Chunks, "")
32                Else
33                    EncodedV = "V2," & CStr(NR) & "," & CStr(NC) & ";" & VBA.Join$(Chunks, "")
34                End If
35                TrySerialiseArrayAsV = True

              Case Else
36                Rank = NumDimensions(x)
37                ReDim Dims(1 To Rank)
38                ReDim DimStr(1 To Rank)
39                ReDim Idx(1 To Rank)
40                Total = 1
41                For q = 1 To Rank
42                    Dims(q) = UBound(x, q) - LBound(x, q) + 1
43                    DimStr(q) = CStr(Dims(q))
44                    Idx(q) = LBound(x, q)
45                    Total = Total * Dims(q)
46                Next q
47                ReDim Chunks(1 To Total)
48                k = 1
49                Do
50                    El = GetAt(x, Idx)
51                    If VarType(El) <> vbDouble Then Exit Function
52                    Chunks(k) = DoubleToHex(CDbl(El))
53                    FirstDigit = Mid$(Chunks(k), 1, 1)
54                    If (FirstDigit = "7" Or FirstDigit = "F") And Mid$(Chunks(k), 2, 2) = "FF" Then Exit Function
55                    k = k + 1
56                    q = 1
57                    Do While q <= Rank
58                        Idx(q) = Idx(q) + 1
59                        If Idx(q) <= UBound(x, q) Then Exit Do
60                        Idx(q) = LBound(x, q)
61                        q = q + 1
62                    Loop
63                    If q > Rank Then Exit Do
64                Loop
65                EncodedV = "V" & CStr(Rank) & "," & VBA.Join$(DimStr, ",") & ";" & VBA.Join$(Chunks, "")
66                TrySerialiseArrayAsV = True
67        End Select

68        Exit Function
ErrHandler:
69        ReThrow "TrySerialiseArrayAsV", Err
End Function

