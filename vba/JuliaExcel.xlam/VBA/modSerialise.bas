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
70                    ReDim Encoded(1 To n)
71                    ReDim Lens(1 To n)
72                    For i = 1 To Rank: Idx(i) = Lb(i): Next i
73                    k = 1
74                    Do
75                        Encoded(k) = SerialiseElement(GetAt(x, Idx))
76                        Lens(k) = CStr(Len(Encoded(k)))
77                        k = k + 1
78                        d = 1
79                        Do While d <= Rank
80                            Idx(d) = Idx(d) + 1
81                            If Idx(d) <= UBound(x, d) Then Exit Do
82                            Idx(d) = Lb(d)
83                            d = d + 1
84                        Loop
85                        If d > Rank Then Exit Do
86                    Loop
87                    SerialiseElement = "*" & CStr(Rank) & "," & VBA.Join$(DimStr, ",") & ";" & VBA.Join$(Lens, ",") & ",;" & VBA.Join$(Encoded, "")
88            End Select

89        Else
90            Select Case VarType(x)
                  Case vbDouble:   SerialiseElement = "#" & DoubleToHex(CDbl(x))
91                Case vbString:   SerialiseElement = Chr(163) & CStr(x)      ' Chr(163) = pound sterling sign
92                Case vbBoolean:  SerialiseElement = IIf(CBool(x), "T", "F")
93                Case vbEmpty:    SerialiseElement = "E"
94                Case vbNull:     SerialiseElement = "N"
95                Case vbInteger:  SerialiseElement = "%" & CStr(CInt(x))
96                Case vbLong:     SerialiseElement = "&" & CStr(CLng(x))
97                Case vbSingle:   SerialiseElement = "S" & SingleToHex(CSng(x))
98                Case vbDate
                      ' CDbl of a VBA date gives the Excel serial number directly:
                      ' integer part = days since 1899-12-30, fractional part = time of day.
99                    If CDbl(x) = Int(CDbl(x)) Then
100                       SerialiseElement = "D" & CStr(CLng(CDbl(x)))         ' date only
101                   Else
102                       SerialiseElement = "G" & DoubleToHex(CDbl(x))        ' date + time
103                   End If
104               Case vbError
                      ' CStr(CVErr(n)) = "Error n"; extract the number after the space.
105                   SerialiseElement = "!" & Mid(CStr(x), InStr(CStr(x), " ") + 1)
106               Case vbObject
107                   If TypeName(x) = "Dictionary" Then
108                       n = x.Count
109                       If n = 0 Then
110                           SerialiseElement = "H0;;"
111                           Exit Function
112                       End If
113                       ReDim Encoded(1 To 2 * n)
114                       ReDim Lens(1 To 2 * n)
115                       k = 1
116                       For Each DictKey In x.Keys
117                           Encoded(k) = SerialiseElement(DictKey)
118                           Lens(k) = CStr(Len(Encoded(k)))
119                           k = k + 1
120                           Encoded(k) = SerialiseElement(x(DictKey))
121                           Lens(k) = CStr(Len(Encoded(k)))
122                           k = k + 1
123                       Next DictKey
124                       SerialiseElement = "H" & CStr(n) & ";" & VBA.Join$(Lens, ",") & ",;" & VBA.Join$(Encoded, "")
125                   Else
126                       Throw "Cannot serialise object of type " & TypeName(x)
127                   End If
#If Win64 Then
128               Case vbLongLong: SerialiseElement = "^" & CStr(x)
#End If
129               Case Else
130                   Throw "Cannot serialise VarType=" & CStr(VarType(x))
131           End Select
132       End If

133       Exit Function
ErrHandler:
134       ReThrow "SerialiseElement", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TrySerialiseArrayAsV
' Purpose    : Attempts the compact "V" wire-format encoding (see modUnserialise.bas's wire-format
'              spec comment and Case 86 'V' branch, and encode_for_xl(::Vector{Float64})/
'              (::Matrix{Float64}) in src/encode.jl, plus decode_xl_array_v in src/decode.jl which
'              understands "V" strings sent this direction) for a 1-D or 2-D array, before
'              SerialiseElement falls back to its own general "*" per-element encoding. Callers are
'              expected to have already handled the empty-array case (n/NR/NC = 0) - this function
'              assumes at least one element.
'              Single-pass and optimistic: checks VarType and finiteness per element AS it appends
'              hex to a buffer array (joined once at the end via VBA.Join$, avoiding the O(n^2)
'              reallocation risk of repeated string concatenation) - returns False (with EncodedV
'              left unset) the instant any element isn't a finite Double, so the caller falls back
'              to the general per-element encoding unchanged. Mirrors the same optimistic
'              single-pass design TryFastDecodeDoubleVector (formerly in the now-deleted
'              modUnserialiseExperimental.bas) used on the decode side.
'              Measured (VEncodeSpeedTest, modPerformance.bas) at roughly 40% faster than the
'              general format even for the realistic worst case - a Variant() array (as
'              Range.Value2 always is, even when every cell holds a number) rather than a
'              genuinely-typed Double() array, so VarType must be checked per element rather than
'              being free.
' -----------------------------------------------------------------------------------------------------------------------
Function TrySerialiseArrayAsV(ByVal x As Variant, ByRef EncodedV As String) As Boolean
          Dim Chunks() As String
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim n As Long
          Dim NC As Long
          Dim NR As Long

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
10                    If Not IsFiniteHex(Chunks(k)) Then Exit Function
11                    k = k + 1
12                Next i
13                EncodedV = "V1," & CStr(n) & ";" & VBA.Join$(Chunks, "")
14                TrySerialiseArrayAsV = True

15            Case 2
16                NR = UBound(x, 1) - LBound(x, 1) + 1
17                NC = UBound(x, 2) - LBound(x, 2) + 1
18                ReDim Chunks(1 To NR * NC)
19                k = 1
20                For j = LBound(x, 2) To UBound(x, 2)    ' column-major to match Julia
21                    For i = LBound(x, 1) To UBound(x, 1)
22                        If VarType(x(i, j)) <> vbDouble Then Exit Function
23                        Chunks(k) = DoubleToHex(CDbl(x(i, j)))
24                        If Not IsFiniteHex(Chunks(k)) Then Exit Function
25                        k = k + 1
26                    Next i
27                Next j
                  ' Nx1 -> 1D Vector, matching SerialiseElement's own Nx1 collapsing.
28                If NC = 1 Then
29                    EncodedV = "V1," & CStr(NR) & ";" & VBA.Join$(Chunks, "")
30                Else
31                    EncodedV = "V2," & CStr(NR) & "," & CStr(NC) & ";" & VBA.Join$(Chunks, "")
32                End If
33                TrySerialiseArrayAsV = True
34        End Select

35        Exit Function
ErrHandler:
36        ReThrow "TrySerialiseArrayAsV", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : IsFiniteHex
' Purpose    : True unless Hex16 (a 16-character big-endian hex string produced by DoubleToHex,
'              modUnserialise.bas) represents NaN or +-Infinity - i.e. unless its 11-bit exponent
'              field is all-ones. Works purely on the hex STRING, never on the Double value itself:
'              VBA's own comparison operators (<>, >, <) raise a runtime "Overflow" error when
'              applied to a genuine NaN bit pattern, rather than returning a Boolean as IEEE-754
'              unordered-comparison semantics would - found the hard way (an early version of
'              TrySerialiseArrayAsV used "v <> v", which crashed the VBA project badly enough to
'              require reverting the workbook from git). DoubleToHex itself is safe regardless of
'              the bit pattern - it only does byte reinterpretation via LSet, no comparisons - so
'              routing through it first and inspecting the resulting hex text avoids the dangerous
'              operators entirely.
'              The exponent occupies the low 3 bits of the first hex digit (that digit's top bit is
'              the sign) plus all of the second and third hex digits - all-ones there means the
'              first digit is "7" or "F" and the next two are "FF".
' -----------------------------------------------------------------------------------------------------------------------
Private Function IsFiniteHex(ByVal Hex16 As String) As Boolean
          Dim FirstDigit As String

1         FirstDigit = Mid$(Hex16, 1, 1)
2         IsFiniteHex = Not ((FirstDigit = "7" Or FirstDigit = "F") And Mid$(Hex16, 2, 2) = "FF")
End Function

