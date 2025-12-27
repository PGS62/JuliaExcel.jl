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
                  Dim rank As Long

54                On Error GoTo ErrHandler
55                If TypeName(x) = "Range" Then
56                    x = x.Value2
57                End If

58                rank = NumDimensions(x)

59                Select Case rank
                      Case 1
60                        ReDim Tmp(LBound(x) To UBound(x))
61                        FirstType = VarType(x(LBound(x)))
62                        AllSameType = True
63                        For i = LBound(x) To UBound(x)
64                            Tmp(i) = MakeJuliaLiteral(x(i))
65                            If AllSameType Then
66                                If VarType(x(i)) <> FirstType Then
67                                    AllSameType = False
68                                End If
69                            End If
70                        Next i
71                        MakeJuliaLiteral = IIf(AllSameType, "[", "Any[") & VBA.Join$(Tmp, ",") & "]"

72                    Case 2
73                        ReDim onerow(LBound(x, 2) To UBound(x, 2))
74                        ReDim Tmp(LBound(x, 1) To UBound(x, 1))
75                        FirstType = VarType(x(LBound(x, 1), LBound(x, 2)))
76                        AllSameType = True
77                        For i = LBound(x, 1) To UBound(x, 1)
78                            For j = LBound(x, 2) To UBound(x, 2)
79                                onerow(j) = MakeJuliaLiteral(x(i, j))
80                                If AllSameType Then
81                                    If VarType(x(i, j)) <> FirstType Then
82                                        AllSameType = False
83                                    End If
84                                End If
85                            Next j
86                            Tmp(i) = VBA.Join$(onerow, " ")
87                        Next i
88                        MakeJuliaLiteral = IIf(AllSameType, "[", "Any[") & VBA.Join$(Tmp, ";") & "]"

89                    Case Else
                          ' rank >= 3: flatten (column-major) and reshape
                          Dim elems() As String
                          Dim Dims() As Long
                          Dim dimStr() As String
                          Dim nElts As Long

                          ' Flatten elements and check homogeneity
90                        nElts = FlattenArrayElements(x, elems, AllSameType, FirstType)

                          ' Build dimension lengths vector (ignores lower bounds)
91                        ReDim Dims(1 To rank)
92                        For k = 1 To rank
93                            Dims(k) = UBound(x, k) - LBound(x, k) + 1
94                        Next k

                          ' Vector literal: homogeneous -> [..], heterogeneous -> Any[..]
                          Dim vecLit As String
95                        vecLit = IIf(AllSameType, "[" & VBA.Join(elems, ",") & "]", _
                              "Any[" & VBA.Join(elems, ",") & "]")

                          ' dims to comma-separated string
96                        ReDim dimStr(1 To rank)
97                        For k = 1 To rank
98                            dimStr(k) = CStr(Dims(k))
99                        Next k

100                       MakeJuliaLiteral = "reshape(" & vecLit & "," & VBA.Join(dimStr, ",") & ")"
101               End Select

102           Case Else
103               Throw "Variable of type " + TypeName(x) + " is not handled"
104       End Select

105       Exit Function
ErrHandler:
106       ReThrow "MakeJuliaLiteral", Err
End Function

' Flatten any VBA array (rank >= 1) to a vector of Julia element literals.
' Traversal order: column-major (dim 1 varies fastest).
' Returns number of elements; also sets AllSameType and FirstType.
Private Function FlattenArrayElements(ByRef A As Variant, _
          ByRef elems() As String, _
          ByRef AllSameType As Boolean, _
          ByRef FirstType As Long) As Long
1         Dim n As Long: n = NumDimensions(A)
2         If n <= 0 Then
3             ReDim elems(1 To 0)
4             AllSameType = True
5             FirstType = vbEmpty
6             FlattenArrayElements = 0
7             Exit Function
8         End If

          Dim Lb() As Long, ub() As Long, idx() As Long
          Dim d As Long, total As Long

9         ReDim Lb(1 To n)
10        ReDim ub(1 To n)
11        ReDim idx(1 To n)

12        total = 1
13        For d = 1 To n
14            Lb(d) = LBound(A, d)
15            ub(d) = UBound(A, d)
16            idx(d) = Lb(d)
17            total = total * (ub(d) - Lb(d) + 1)
18        Next d

19        ReDim elems(1 To total)

          ' Initialize homogeneity checks from the first element
          Dim v As Variant
20        v = GetAt(A, idx)
21        FirstType = VarType(v)
22        AllSameType = True

23        Dim count As Long: count = 0
24        Do
25            count = count + 1
26            v = GetAt(A, idx)
27            elems(count) = MakeJuliaLiteral(v)
28            If AllSameType Then
29                If VarType(v) <> FirstType Then AllSameType = False
30            End If

              ' Increment indices: dim 1 fastest (column-major)
31            d = 1
32            Do While d <= n
33                idx(d) = idx(d) + 1
34                If idx(d) <= ub(d) Then Exit Do
35                idx(d) = Lb(d)
36                d = d + 1
37            Loop
38            If d > n Then Exit Do
39        Loop

40        FlattenArrayElements = total
End Function


