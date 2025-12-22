Attribute VB_Name = "modJuliaLiteral"
' Copyright (c) 2021-2025 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme

Option Explicit
Option Private Module

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MakeJuliaLiteral
' Purpose    : Convert an variant x into a string which Julia will parse as x.
'
' Examples:
' In VBA immediate window:
' ?MakeJuliaLiteral(Array(1#, 2#, 3#))
' [htd("3FF0000000000000"),htd("4000000000000000"),htd("4008000000000000")]
'
' In Julia REPL:
'julia> using JuliaExcel

'julia> [htd("3FF0000000000000"),htd("4000000000000000"),htd("4008000000000000")]
'3-element Vector{Float64}:
' 1.0
' 2.0
' 3.0
' -----------------------------------------------------------------------------------------------------------------------
Function MakeJuliaLiteral(x As Variant)
          Dim k As Long
          Dim Res As String

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
                  'Avoid loss of precision by representing x as its IEEE-754 bit pattern. _
                   Also avoids having to worry about whether the decimal separator is point or comma.
32                MakeJuliaLiteral = "htd(""" & DoubleToHex(x) & """)"
33                Exit Function
34            Case vbLong, vbInteger
35                MakeJuliaLiteral = CStr(x)
36                Exit Function
37            Case vbBoolean
38                MakeJuliaLiteral = IIf(x, "true", "false")
39                Exit Function
40            Case vbEmpty
41                MakeJuliaLiteral = "missing"
42                Exit Function
43            Case vbDate
44                If CDbl(x) = CLng(x) Then
45                    MakeJuliaLiteral = "Date(""" & Format(x, "yyyy-mm-dd") & """)"
46                Else
47                    MakeJuliaLiteral = "DateTime(""" & VBA.Format$(x, "yyyy-mm-ddThh:mm:ss.000") & """)"
48                End If
49                Exit Function
50            Case vbSingle
                  'Avoid loss of precision by representing x as its IEEE-754 bit pattern. _
                   Also avoids having to worry about whether the decimal separator is point or comma.
51                MakeJuliaLiteral = "hts(""" & SingleToHex(x) & """)"
52                Exit Function
53            Case Is >= vbArray
                  Dim AllSameType As Boolean
                  Dim FirstType As Long
                  Dim i As Long
                  Dim j As Long
                  Dim onerow() As String
                  Dim Tmp() As String
          
54                On Error GoTo ErrHandler
55                If TypeName(x) = "Range" Then
56                    x = x.Value2
57                End If

58                Select Case NumDimensions(x)
                      Case 1
59                        ReDim Tmp(LBound(x) To UBound(x))
60                        FirstType = VarType(x(LBound(x)))
61                        AllSameType = True
62                        For i = LBound(x) To UBound(x)
63                            Tmp(i) = MakeJuliaLiteral(x(i))
64                            If AllSameType Then
65                                If VarType(x(i)) <> FirstType Then
66                                    AllSameType = False
67                                End If
68                            End If
69                        Next i
70                        MakeJuliaLiteral = IIf(AllSameType, "[", "Any[") & VBA.Join$(Tmp, ",") & "]"
71                    Case 2
72                        ReDim onerow(LBound(x, 2) To UBound(x, 2))
73                        ReDim Tmp(LBound(x, 1) To UBound(x, 1))
74                        FirstType = VarType(x(LBound(x, 1), LBound(x, 2)))
75                        AllSameType = True
76                        For i = LBound(x, 1) To UBound(x, 1)
77                            For j = LBound(x, 2) To UBound(x, 2)
78                                onerow(j) = MakeJuliaLiteral(x(i, j))
79                                If AllSameType Then
80                                    If VarType(x(i, j)) <> FirstType Then
81                                        AllSameType = False
82                                    End If
83                                End If
84                            Next j
85                            Tmp(i) = VBA.Join$(onerow, " ")
86                        Next i

87                        MakeJuliaLiteral = IIf(AllSameType, "[", "Any[") & VBA.Join$(Tmp, ";") & "]"
88                    Case Else
89                        Throw "case more than two dimensions not handled" 'In VBA there's no way to handle arrays with arbitrary number of dimensions. Easy in Julia!
90                End Select

91            Case Else
92                Throw "Variable of type " + TypeName(x) + " is not handled"
93        End Select

94        Exit Function
ErrHandler:
95        ReThrow "MakeJuliaLiteral", Err
End Function



