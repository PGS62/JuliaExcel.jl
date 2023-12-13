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
          Dim k As Long

1         On Error GoTo ErrHandler
2         Select Case VarType(x)

              Case vbString
3                 Res = x
                  'Must do this substitution first
4                 If InStr(x, "\") > 0 Then
5                     Res = Replace(Res, "\", "\\")
6                 End If
                  'The conversions in the two loops below are needed to avoid an error: _
                  Base.Meta.ParseError("unbalanced bidirectional formatting in string literal") _
                  'Julia's "caution" in relation to these characters is a defence against "Trojan Source" attacks.
                  'https://github.com/JuliaLang/julia/pull/42918
                  'https://trojansource.codes/
7                 For k = 8234 To 8238
8                     If InStr(x, ChrW(k)) Then
9                         Res = Replace(Res, ChrW(k), "\u" & LCase(Hex(k)))
10                    End If
11                Next k
12                For k = 8294 To 8297
13                    If InStr(x, ChrW(k)) Then
14                        Res = Replace(Res, ChrW(k), "\u" & LCase(Hex(k)))
15                    End If
16                Next k
17                If InStr(x, vbCr) > 0 Then
18                    Res = Replace(Res, vbCr, "\r")
19                End If
20                If InStr(x, vbLf) > 0 Then
21                    Res = Replace(Res, vbLf, "\n")
22                End If
23                If InStr(x, "$") > 0 Then
24                    Res = Replace(Res, "$", "\$")
25                End If
26                If InStr(x, """") > 0 Then
27                    Res = Replace(Res, """", "\""")
28                End If
29                MakeJuliaLiteral = """" & Res & """"
30                Exit Function
31            Case vbDouble
32                Res = CStr(x)
33                If InStr(Res, ".") = 0 Then
34                    If InStr(Res, "E") = 0 Then
35                        Res = Res + ".0"
36                    End If
37                End If
38                MakeJuliaLiteral = Res
39                Exit Function
40            Case vbLong, vbInteger
41                MakeJuliaLiteral = CStr(x)
42                Exit Function
43            Case vbBoolean
44                MakeJuliaLiteral = IIf(x, "true", "false")
45                Exit Function
46            Case vbEmpty
47                MakeJuliaLiteral = "missing"
48                Exit Function
49            Case vbDate
50                If CDbl(x) = CLng(x) Then
51                    MakeJuliaLiteral = "Date(""" & Format(x, "yyyy-mm-dd") & """)"
52                Else
53                    MakeJuliaLiteral = "DateTime(""" & VBA.Format$(x, "yyyy-mm-ddThh:mm:ss.000") & """)"
54                End If
55                Exit Function
56            Case Is >= vbArray
                  Dim AllSameType As Boolean
                  Dim FirstType As Long
                  Dim i As Long
                  Dim j As Long
                  Dim onerow() As String
                  Dim Tmp() As String
          
57                On Error GoTo ErrHandler
58                If TypeName(x) = "Range" Then
59                    x = x.Value2
60                End If

61                Select Case NumDimensions(x)
                      Case 1
62                        ReDim Tmp(LBound(x) To UBound(x))
63                        FirstType = VarType(x(LBound(x)))
64                        AllSameType = True
65                        For i = LBound(x) To UBound(x)
66                            Tmp(i) = MakeJuliaLiteral(x(i))
67                            If AllSameType Then
68                                If VarType(x(i)) <> FirstType Then
69                                    AllSameType = False
70                                End If
71                            End If
72                        Next i
73                        MakeJuliaLiteral = IIf(AllSameType, "[", "Any[") & VBA.Join$(Tmp, ",") & "]"
74                    Case 2
75                        ReDim onerow(LBound(x, 2) To UBound(x, 2))
76                        ReDim Tmp(LBound(x, 1) To UBound(x, 1))
77                        FirstType = VarType(x(LBound(x, 1), LBound(x, 2)))
78                        AllSameType = True
79                        For i = LBound(x, 1) To UBound(x, 1)
80                            For j = LBound(x, 2) To UBound(x, 2)
81                                onerow(j) = MakeJuliaLiteral(x(i, j))
82                                If AllSameType Then
83                                    If VarType(x(i, j)) <> FirstType Then
84                                        AllSameType = False
85                                    End If
86                                End If
87                            Next j
88                            Tmp(i) = VBA.Join$(onerow, " ")
89                        Next i

90                        MakeJuliaLiteral = IIf(AllSameType, "[", "Any[") & VBA.Join$(Tmp, ";") & "]"
91                    Case Else
92                        Throw "case more than two dimensions not handled" 'In VBA there's no way to handle arrays with arbitrary number of dimensions. Easy in Julia!
93                End Select

94            Case Else
95                Throw "Variable of type " + TypeName(x) + " is not handled"
96        End Select

97        Exit Function
ErrHandler:
98        Throw "#MakeJuliaLiteral (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function


