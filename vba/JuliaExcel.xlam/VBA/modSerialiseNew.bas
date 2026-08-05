Attribute VB_Name = "modSerialiseNew"
' Copyright (c) 2021-2025 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme

Option Explicit
Option Private Module

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Serialise
' Purpose    : Convenience wrapper for SerialiseArgs allowing a ParamArray call signature. Encodes
'              JuliaFunctionName and Arguments into the JuliaExcel wire format (a 1D array where
'              element 0 is the function name and elements 1..n are the serialised arguments).
'              The result is a string suitable for passing to SaveTextFile.
'
'              SerialiseArgs is the implementation; Serialise is a thin ParamArray wrapper.
'              JuliaCallNew calls SerialiseArgs directly to forward its own ParamArray.
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
' Purpose    : Implementation of Serialise. Takes Arguments as a plain Variant (an array) rather
'              than ParamArray, so that JuliaCallNew can forward its ParamArray directly:
'                  SerialiseArgs(fn, Args)   -- Args is JuliaCallNew's ParamArray
'
'              Wire format (same as Julia's encode_for_xl / VBA's Unserialise):
'                Double   -> "#" + 16 hex chars (IEEE-754 bit pattern via DoubleToHex)
'                Single   -> "S" + 8 hex chars
'                String   -> Chr(163) + content          (Chr(163) = pound sign = £)
'                Boolean  -> "T" or "F"
'                Long     -> "&" + decimal
'                Integer  -> "%" + decimal
'                LongLong -> "^" + decimal               (64-bit only)
'                Date     -> "D" + excel serial (date-only) or "G" + 16 hex (datetime)
'                Empty    -> "E"
'                Null     -> "N"
'                Error    -> "!" + error number
'                Array 1D -> "*1,N;<len1>,<len2>,...,;<elements>"  (column-major)
'                Array 2D -> "*2,NR,NC;<len1>,...,;<elements>"     (column-major)
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

          Dim Encoded() As String
          Dim i As Long
          Dim j As Long
          Dim k As Long
          Dim Lens() As String
          Dim n As Long
          Dim DictKey As Variant
          Dim NC As Long
          Dim NR As Long
          Dim ThisEncoded As String

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
                          SerialiseElement = "*2," & CStr(NR) & "," & CStr(NC) & ";;"
                          Exit Function
                      End If

22                    ReDim Encoded(1 To NR * NC)
23                    ReDim Lens(1 To NR * NC)
24                    k = 1
25                    For j = LBound(x, 2) To UBound(x, 2)    ' column-major to match Julia
26                        For i = LBound(x, 1) To UBound(x, 1)
27                            Encoded(k) = SerialiseElement(x(i, j))
28                            Lens(k) = CStr(Len(Encoded(k)))
29                            k = k + 1
30                        Next i
31                    Next j
32                    SerialiseElement = "*2," & CStr(NR) & "," & CStr(NC) & ";" & VBA.Join$(Lens, ",") & ",;" & VBA.Join$(Encoded, "")

33                Case Else
34                    Throw "Cannot serialise arrays with more than 2 dimensions"
35            End Select

36        Else
37            Select Case VarType(x)
                    Case vbDouble:   SerialiseElement = "#" & DoubleToHex(CDbl(x))
38                  Case vbString:   SerialiseElement = Chr(163) & CStr(x)      ' Chr(163) = £
39                  Case vbBoolean:  SerialiseElement = IIf(CBool(x), "T", "F")
40                  Case vbEmpty:    SerialiseElement = "E"
41                  Case vbNull:     SerialiseElement = "N"
42                  Case vbInteger:  SerialiseElement = "%" & CStr(CInt(x))
43                  Case vbLong:     SerialiseElement = "&" & CStr(CLng(x))
44                  Case vbSingle:   SerialiseElement = "S" & SingleToHex(CSng(x))
45                  Case vbDate
                      ' CDbl of a VBA date gives the Excel serial number directly:
                      ' integer part = days since 1899-12-30, fractional part = time of day.
46                      If CDbl(x) = Int(CDbl(x)) Then
47                          SerialiseElement = "D" & CStr(CLng(CDbl(x)))         ' date only
48                      Else
49                          SerialiseElement = "G" & DoubleToHex(CDbl(x))        ' date + time
50                      End If
51                  Case vbError
                      ' CStr(CVErr(n)) = "Error n"; extract the number after the space.
52                      SerialiseElement = "!" & Mid(CStr(x), InStr(CStr(x), " ") + 1)
53                  Case vbObject
54                      If TypeName(x) = "Dictionary" Then
62                          n = x.Count
63                          If n = 0 Then
64                              SerialiseElement = "H0;;"
65                              Exit Function
66                          End If
67                          ReDim Encoded(1 To 2 * n)
68                          ReDim Lens(1 To 2 * n)
69                          k = 1
70                          For Each DictKey In x.Keys
71                              Encoded(k) = SerialiseElement(DictKey)
72                              Lens(k) = CStr(Len(Encoded(k)))
73                              k = k + 1
74                              Encoded(k) = SerialiseElement(x(DictKey))
75                              Lens(k) = CStr(Len(Encoded(k)))
76                              k = k + 1
77                          Next DictKey
78                          SerialiseElement = "H" & CStr(n) & ";" & VBA.Join$(Lens, ",") & ",;" & VBA.Join$(Encoded, "")
79                      Else
80                          Throw "Cannot serialise object of type " & TypeName(x)
81                      End If
#If Win64 Then
55                  Case vbLongLong: SerialiseElement = "^" & CStr(x)
#End If
56                  Case Else
57                      Throw "Cannot serialise VarType=" & CStr(VarType(x))
58            End Select
59        End If

60        Exit Function
ErrHandler:
61        ReThrow "SerialiseElement", Err
End Function
