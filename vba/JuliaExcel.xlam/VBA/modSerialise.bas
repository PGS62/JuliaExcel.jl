Attribute VB_Name = "modSerialise"
' Copyright (c) 2021-2025 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme

Option Explicit
Option Private Module

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Serialise
' Purpose    : Encodes JuliaFunctionName and Arguments into the JuliaExcel wire format (a 1D
'              array where element 0 is the function name and elements 1..n are the serialised
'              arguments). The result is a string suitable for passing to SaveTextFile.
'              Offers a ParamArray call signature; see SerialiseArgs for the Variant equivalent.
'
' Example    : Serialise("sum", Range("A1:A100000")) -- encodes range values in wire format
' -----------------------------------------------------------------------------------------------------------------------
Function Serialise(JuliaFunctionName As String, ParamArray Arguments()) As String
          ' VBA does not allow a ParamArray to be forwarded to another function, so the outer
          ' array encoding is done here. SerialiseElement handles each value.
          Dim Arg As Variant
          Dim ContentsSection As String
          Dim i As Long
          Dim LengthsSection As String
          Dim NumArgs As Long
          Dim NumElements As Long
          Dim ThisEncoded As String

1         On Error GoTo ErrHandler

          ' Element 0: function name encoded as a string (Chr(163) = £ = string type indicator)
2         ThisEncoded = Chr(163) & JuliaFunctionName
3         LengthsSection = CStr(Len(ThisEncoded)) & ","
4         ContentsSection = ThisEncoded

          ' Elements 1..n: serialised arguments
5         NumArgs = IIf(UBound(Arguments) >= LBound(Arguments), _
              UBound(Arguments) - LBound(Arguments) + 1, 0)
6         NumElements = 1 + NumArgs
7         For i = 0 To NumArgs - 1
8             Arg = Arguments(LBound(Arguments) + i)
9             If TypeName(Arg) = "Range" Then Arg = Arg.Value2
10            ThisEncoded = SerialiseElement(Arg)
11            LengthsSection = LengthsSection & CStr(Len(ThisEncoded)) & ","
12            ContentsSection = ContentsSection & ThisEncoded
13        Next i

14        Serialise = "*1," & CStr(NumElements) & ";" & LengthsSection & ";" & ContentsSection

15        Exit Function
ErrHandler:
16        ReThrow "Serialise", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SerialiseArgs
' Purpose    : Variant-argument equivalent of Serialise. Takes Arguments as a plain Variant
'              array rather than ParamArray, for VBA callers that have already assembled their
'              arguments into an array.
'
'              Wire format (same as Julia's encode_for_xl / VBA's Unserialise):
'                Double     -> "#" + 16 hex chars (IEEE-754 bit pattern via DoubleToHex)
'                Single     -> "S" + 8 hex chars
'                String     -> Chr(163) + content          (Chr(163) = pound sign = £)
'                Boolean    -> "T" or "F"
'                Long       -> "&" + decimal
'                Integer    -> "%" + decimal
'                LongLong   -> "^" + decimal               (64-bit only)
'                Date       -> "D" + excel serial (date-only) or "G" + 16 hex (datetime)
'                Empty      -> "E"
'                Null       -> "N"
'                Error      -> "!" + error number
'                Array 1D   -> "*1,N;<len1>,<len2>,...,;<elements>"        (column-major)
'                Array 2D   -> "*2,NR,NC;<len1>,...,;<elements>"           (column-major)
'                Dictionary -> "H<count>;<k1_len>,<v1_len>,...,;<k1><v1>..." (key-value pairs)
'
'              Lengths in the lengths section use Len(), which counts UTF-16 code units and
'              thus matches Julia's xl_length (supplementary chars each count as 2).
' -----------------------------------------------------------------------------------------------------------------------
Function SerialiseArgs(JuliaFunctionName As String, Arguments As Variant) As String

          Dim Arg As Variant
          Dim ContentsSection As String
          Dim i As Long
          Dim LengthsSection As String
          Dim NumArgs As Long
          Dim NumElements As Long
          Dim ThisEncoded As String

1         On Error GoTo ErrHandler

2         If IsArray(Arguments) Then
3             NumArgs = IIf(UBound(Arguments) >= LBound(Arguments), _
                  UBound(Arguments) - LBound(Arguments) + 1, 0)
4         Else
              ' Scalar passed directly -- treat as a single argument
5             NumArgs = 1
6         End If
7         NumElements = 1 + NumArgs

          ' Element 0: function name, encoded as a string (Chr(163) = £ = string type indicator)
8         ThisEncoded = Chr(163) & JuliaFunctionName
9         LengthsSection = CStr(Len(ThisEncoded)) & ","
10        ContentsSection = ThisEncoded

          ' Elements 1..n: serialised arguments
11        If NumArgs = 1 And Not IsArray(Arguments) Then
              ' Scalar passed directly to SerialiseArgs (not via ParamArray)
12            Arg = Arguments
13            If TypeName(Arg) = "Range" Then Arg = Arg.Value2
14            ThisEncoded = SerialiseElement(Arg)
15            LengthsSection = LengthsSection & CStr(Len(ThisEncoded)) & ","
16            ContentsSection = ContentsSection & ThisEncoded
17        Else
18            For i = 0 To NumArgs - 1
19                Arg = Arguments(LBound(Arguments) + i)
20                If TypeName(Arg) = "Range" Then Arg = Arg.Value2
21                ThisEncoded = SerialiseElement(Arg)
22                LengthsSection = LengthsSection & CStr(Len(ThisEncoded)) & ","
23                ContentsSection = ContentsSection & ThisEncoded
24            Next i
25        End If

26        SerialiseArgs = "*1," & CStr(NumElements) & ";" & LengthsSection & ";" & ContentsSection

27        Exit Function
ErrHandler:
28        ReThrow "SerialiseArgs", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SerialiseElement
' Purpose    : Encode a single VBA value (scalar or array) into the JuliaExcel wire format.
'              Mirror of Unserialise in modSerialise.bas. Arrays are written column-major to
'              match Julia's default array layout and encode_for_xl.
' -----------------------------------------------------------------------------------------------------------------------
Public Function SerialiseElement(ByVal x As Variant) As String

          Dim d As Long
          Dim DictKey As Variant
          Dim Dims() As Long
          Dim DimStr() As String
          Dim Encoded() As String
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
9                     ReDim Encoded(1 To n)
10                    ReDim Lens(1 To n)
11                    k = 1
12                    For i = LBound(x) To UBound(x)
13                        Encoded(k) = SerialiseElement(x(i))
14                        Lens(k) = CStr(Len(Encoded(k)))
15                        k = k + 1
16                    Next i
17                    SerialiseElement = "*1," & CStr(n) & ";" & VBA.Join$(Lens, ",") & ",;" & VBA.Join$(Encoded, "")

18                Case 2
19                    NR = UBound(x, 1) - LBound(x, 1) + 1
20                    NC = UBound(x, 2) - LBound(x, 2) + 1
21                    If NR = 0 Or NC = 0 Then
22                        If NC = 1 Then
23                            SerialiseElement = "*1,0;;"
24                        Else
25                            SerialiseElement = "*2," & CStr(NR) & "," & CStr(NC) & ";;"
26                        End If
27                        Exit Function
28                    End If

29                    ReDim Encoded(1 To NR * NC)
30                    ReDim Lens(1 To NR * NC)
31                    k = 1
32                    For j = LBound(x, 2) To UBound(x, 2)    ' column-major to match Julia
33                        For i = LBound(x, 1) To UBound(x, 1)
34                            Encoded(k) = SerialiseElement(x(i, j))
35                            Lens(k) = CStr(Len(Encoded(k)))
36                            k = k + 1
37                        Next i
38                    Next j
                      ' Nx1 -> 1D Vector (matches JuliaCallOld / README "single-column ranges arrive as vectors").
                      ' 1xN stays as 2D Matrix; 3D+ arrays are left as-is (no prior behaviour to replicate).
39                    If NC = 1 Then
40                        SerialiseElement = "*1," & CStr(NR) & ";" & VBA.Join$(Lens, ",") & ",;" & VBA.Join$(Encoded, "")
41                    Else
42                        SerialiseElement = "*2," & CStr(NR) & "," & CStr(NC) & ";" & VBA.Join$(Lens, ",") & ",;" & VBA.Join$(Encoded, "")
43                    End If

44                Case Else
45                    Rank = NumDimensions(x)
46                    If Rank > 9 Then Throw "Cannot serialise arrays with more than 9 dimensions"
47                    ReDim Dims(1 To Rank)
48                    ReDim Lb(1 To Rank)
49                    ReDim DimStr(1 To Rank)
50                    ReDim Idx(1 To Rank)
51                    n = 1
52                    For i = 1 To Rank
53                        Lb(i) = LBound(x, i)
54                        Dims(i) = UBound(x, i) - Lb(i) + 1
55                        DimStr(i) = CStr(Dims(i))
56                        n = n * Dims(i)
57                    Next i
58                    If n = 0 Then
59                        SerialiseElement = "*" & CStr(Rank) & "," & VBA.Join$(DimStr, ",") & ";;"
60                        Exit Function
61                    End If
62                    ReDim Encoded(1 To n)
63                    ReDim Lens(1 To n)
64                    For i = 1 To Rank: Idx(i) = Lb(i): Next i
65                    k = 1
66                    Do
67                        Encoded(k) = SerialiseElement(GetAt(x, Idx))
68                        Lens(k) = CStr(Len(Encoded(k)))
69                        k = k + 1
70                        d = 1
71                        Do While d <= Rank
72                            Idx(d) = Idx(d) + 1
73                            If Idx(d) <= UBound(x, d) Then Exit Do
74                            Idx(d) = Lb(d)
75                            d = d + 1
76                        Loop
77                        If d > Rank Then Exit Do
78                    Loop
79                    SerialiseElement = "*" & CStr(Rank) & "," & VBA.Join$(DimStr, ",") & ";" & VBA.Join$(Lens, ",") & ",;" & VBA.Join$(Encoded, "")
80            End Select

81        Else
82            Select Case VarType(x)
                    Case vbDouble:   SerialiseElement = "#" & DoubleToHex(CDbl(x))
83                  Case vbString:   SerialiseElement = Chr(163) & CStr(x)      ' Chr(163) = £
84                  Case vbBoolean:  SerialiseElement = IIf(CBool(x), "T", "F")
85                  Case vbEmpty:    SerialiseElement = "E"
86                  Case vbNull:     SerialiseElement = "N"
87                  Case vbInteger:  SerialiseElement = "%" & CStr(CInt(x))
88                  Case vbLong:     SerialiseElement = "&" & CStr(CLng(x))
89                  Case vbSingle:   SerialiseElement = "S" & SingleToHex(CSng(x))
90                  Case vbDate
                      ' CDbl of a VBA date gives the Excel serial number directly:
                      ' integer part = days since 1899-12-30, fractional part = time of day.
91                      If CDbl(x) = Int(CDbl(x)) Then
92                          SerialiseElement = "D" & CStr(CLng(CDbl(x)))         ' date only
93                      Else
94                          SerialiseElement = "G" & DoubleToHex(CDbl(x))        ' date + time
95                      End If
96                  Case vbError
                      ' CStr(CVErr(n)) = "Error n"; extract the number after the space.
97                      SerialiseElement = "!" & Mid(CStr(x), InStr(CStr(x), " ") + 1)
98                  Case vbObject
99                      If TypeName(x) = "Dictionary" Then
100                         n = x.Count
101                         If n = 0 Then
102                             SerialiseElement = "H0;;"
103                             Exit Function
104                         End If
105                         ReDim Encoded(1 To 2 * n)
106                         ReDim Lens(1 To 2 * n)
107                         k = 1
108                         For Each DictKey In x.Keys
109                             Encoded(k) = SerialiseElement(DictKey)
110                             Lens(k) = CStr(Len(Encoded(k)))
111                             k = k + 1
112                             Encoded(k) = SerialiseElement(x(DictKey))
113                             Lens(k) = CStr(Len(Encoded(k)))
114                             k = k + 1
115                         Next DictKey
116                         SerialiseElement = "H" & CStr(n) & ";" & VBA.Join$(Lens, ",") & ",;" & VBA.Join$(Encoded, "")
117                     Else
118                         Throw "Cannot serialise object of type " & TypeName(x)
119                     End If
#If Win64 Then
120                 Case vbLongLong: SerialiseElement = "^" & CStr(x)
#End If
121                 Case Else
122                     Throw "Cannot serialise VarType=" & CStr(VarType(x))
123           End Select
124       End If

125       Exit Function
ErrHandler:
126       ReThrow "SerialiseElement", Err
End Function

