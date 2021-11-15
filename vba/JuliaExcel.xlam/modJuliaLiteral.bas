Attribute VB_Name = "modJuliaLiteral"
Option Explicit
Option Private Module

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MakeJuliaLiteral
' Purpose    : Convert an array into a string which Julia will parse as x.
'
' Examples:
' In VBA immediate window:
' ?MakeJuliaLiteral(Array(1#, 2#, 3#),False)
' [1.0 2.0 3.0]
'
' In Julia REPL:
' julia> [1.0,2.0,3.0]
' 3-element Vector{Float64}:
'  1.0
'  2.0
'  3.0

' In VBA immediate window:
' ?MakeJuliaLiteral(Array(1#, 2#, 3#),True)
' [1.0,2.0,3.0]
'
' In Julia REPL:
'julia> [1.0 2.0 3.0]
'1×3 Matrix{Float64}:
' 1.0  2.0  3.0

'Handles nested arrays:
' ?Print MakeJuliaLiteral(Array(1#, 2#, Array(3#, 4#)),False)
' Any[1.0,2.0,[3.0,4.0]]
' -----------------------------------------------------------------------------------------------------------------------
Function MakeJuliaLiteral(x As Variant)
          Dim Res As String

1         On Error GoTo ErrHandler
2         Select Case VarType(x)

              Case vbString
3                 Res = x
4                 If InStr(x, "\") > 0 Then
5                     Res = Replace(Res, "\", "\\")
6                 End If
7                 If InStr(x, vbCr) > 0 Then
8                     Res = Replace(Res, vbCr, "\r")
9                 End If
10                If InStr(x, vbLf) > 0 Then
11                    Res = Replace(Res, vbLf, "\n")
12                End If
13                If InStr(x, "$") > 0 Then
14                    Res = Replace(Res, "$", "\$")
15                End If
16                If InStr(x, """") > 0 Then
17                    Res = Replace(Res, """", "\""")
18                End If
19                MakeJuliaLiteral = """" & Res & """"
20                Exit Function
21            Case vbDouble
22                Res = CStr(x)
23                If InStr(Res, ".") = 0 Then
24                    If InStr(Res, "E") = 0 Then
25                        Res = Res + ".0"
26                    End If
27                End If
28                MakeJuliaLiteral = Res
29                Exit Function
30            Case vbLong, vbInteger
31                MakeJuliaLiteral = CStr(x)
32                Exit Function
33            Case vbBoolean
34                MakeJuliaLiteral = IIf(x, "true", "false")
35                Exit Function
36            Case vbEmpty
37                MakeJuliaLiteral = "missing"
38                Exit Function
39            Case vbDate
40                If CDbl(x) = CLng(x) Then
41                    MakeJuliaLiteral = "Date(""" & Format(x, "yyyy-mm-dd") & """)"
42                Else
43                    MakeJuliaLiteral = "DateTime(""" & VBA.Format$(x, "yyyy-mm-ddThh:mm:ss.000") & """)"
44                End If
45                Exit Function
46            Case Is >= vbArray
                  Dim AllSameType As Boolean
                  Dim FirstType As Long
                  Dim i As Long
                  Dim j As Long
                  Dim onerow() As String
                  Dim Tmp() As String
          
47                On Error GoTo ErrHandler
48                If TypeName(x) = "Range" Then
49                    x = x.Value2
50                End If

51                Select Case NumDimensions(x)
                      Case 1
52                        ReDim Tmp(LBound(x) To UBound(x))
53                        FirstType = VarType(x(LBound(x)))
54                        AllSameType = True
55                        For i = LBound(x) To UBound(x)
56                            Tmp(i) = MakeJuliaLiteral(x(i))
57                            If AllSameType Then
58                                If VarType(x(i)) <> FirstType Then
59                                    AllSameType = False
60                                End If
61                            End If
62                        Next i
63                        MakeJuliaLiteral = IIf(AllSameType, "[", "Any[") & VBA.Join$(Tmp, ",") & "]"
64                    Case 2
65                        ReDim onerow(LBound(x, 2) To UBound(x, 2))
66                        ReDim Tmp(LBound(x, 1) To UBound(x, 1))
67                        FirstType = VarType(x(LBound(x, 1), LBound(x, 2)))
68                        AllSameType = True
69                        For i = LBound(x, 1) To UBound(x, 1)
70                            For j = LBound(x, 2) To UBound(x, 2)
71                                onerow(j) = MakeJuliaLiteral(x(i, j))
72                                If AllSameType Then
73                                    If VarType(x(i, j)) <> FirstType Then
74                                        AllSameType = False
75                                    End If
76                                End If
77                            Next j
78                            Tmp(i) = VBA.Join$(onerow, " ")
79                        Next i

80                        MakeJuliaLiteral = IIf(AllSameType, "[", "Any[") & VBA.Join$(Tmp, ";") & "]"
81                    Case Else
82                        Throw "case more than two dimensions not handled" 'In VBA there's no way to handle arrays with arbitrary number of dimensions. Easy in Julia!
83                End Select

84            Case Else
85                Throw "Variable of type " + TypeName(x) + " is not handled"
86        End Select

87        Exit Function
ErrHandler:
88        Throw "#MakeJuliaLiteral (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function
