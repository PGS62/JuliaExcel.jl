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
' Example 1 - Array of Longs in VBA yields a literal that's interpreted in Julia as an array of Int64s.
'
'VBA immediate window:
'?MakeJuliaLiteral(Array(1, 2, 3))
'[1,2,3]
'
'Julia REPL:
'julia> [1,2,3]
'3-element Vector{Int64}:
'1
'2
'3
'
' Example 2 - Array of Doubles in VBA yields a literal that's interpreted in Julia as an Array of
' Float64s, but note the requirement for JuliaExcel so that the function htd is in scope.
'
'VBA immediate window:
'?MakeJuliaLiteral(Array(Sqr(2), Sqr(3), Sqr(5)))
'[htd("3FF6A09E667F3BCD"),htd("3FFBB67AE8584CAA"),htd("4001E3779B97F4A8")]
'
'Julia REPL:
'julia> using JuliaExcel
'
'julia> [htd("3FF6A09E667F3BCD"),htd("3FFBB67AE8584CAA"),htd("4001E3779B97F4A8")]
'3-element Vector{Float64}:
' 1.4142135623730951
' 1.7320508075688772
' 2.23606797749979
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
34            Case vbLongLong
35                MakeJuliaLiteral = CStr(x)
36                Exit Function
37            Case vbLong
38                MakeJuliaLiteral = "Int32(" & CStr(x) & ")"
39                Exit Function
40            Case vbInteger
41                MakeJuliaLiteral = "Int16(" & CStr(x) & ")"
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
56            Case vbSingle
57                MakeJuliaLiteral = "hts(""" & SingleToHex(x) & """)"
58                Exit Function
59            Case Is >= vbArray

                  Dim AllSameType As Boolean
                  Dim FirstType As Long
                  Dim i As Long
                  Dim j As Long
                  Dim OneRow() As String
                  Dim Rank As Long
                  Dim Tmp() As String

60                On Error GoTo ErrHandler
61                If TypeName(x) = "Range" Then
62                    x = x.Value2
63                End If

64                Rank = NumDimensions(x)

65                Select Case Rank
                      Case 1
66                        ReDim Tmp(LBound(x) To UBound(x))
67                        FirstType = VarType(x(LBound(x)))
68                        AllSameType = True
69                        For i = LBound(x) To UBound(x)
70                            Tmp(i) = MakeJuliaLiteral(x(i))
71                            If AllSameType Then
72                                If VarType(x(i)) <> FirstType Then
73                                    AllSameType = False
74                                End If
75                            End If
76                        Next i
77                        MakeJuliaLiteral = IIf(AllSameType, "[", "Any[") & VBA.Join$(Tmp, ",") & "]"

78                    Case 2
79                        ReDim OneRow(LBound(x, 2) To UBound(x, 2))
80                        ReDim Tmp(LBound(x, 1) To UBound(x, 1))
81                        FirstType = VarType(x(LBound(x, 1), LBound(x, 2)))
82                        AllSameType = True
83                        For i = LBound(x, 1) To UBound(x, 1)
84                            For j = LBound(x, 2) To UBound(x, 2)
85                                OneRow(j) = MakeJuliaLiteral(x(i, j))
86                                If AllSameType Then
87                                    If VarType(x(i, j)) <> FirstType Then
88                                        AllSameType = False
89                                    End If
90                                End If
91                            Next j
92                            Tmp(i) = VBA.Join$(OneRow, " ")
93                        Next i
94                        MakeJuliaLiteral = IIf(AllSameType, "[", "Any[") & VBA.Join$(Tmp, ";") & "]"

95                    Case Else
                          ' rank >= 3: flatten (column-major) and reshape
                          Dim Dims() As Long
                          Dim DimStr() As String
                          Dim Elems() As String
                          Dim NElts As Long

                          ' Flatten elements and check homogeneity
96                        NElts = FlattenArrayElements(x, Elems, AllSameType, FirstType)

                          ' Build dimension lengths vector (ignores lower bounds)
97                        ReDim Dims(1 To Rank)
98                        For k = 1 To Rank
99                            Dims(k) = UBound(x, k) - LBound(x, k) + 1
100                       Next k

                          ' Vector literal: homogeneous -> [..], heterogeneous -> Any[..]
                          Dim VecLit As String
101                       VecLit = IIf(AllSameType, "[" & VBA.Join(Elems, ",") & "]", _
                              "Any[" & VBA.Join(Elems, ",") & "]")

                          ' dims to comma-separated string
102                       ReDim DimStr(1 To Rank)
103                       For k = 1 To Rank
104                           DimStr(k) = CStr(Dims(k))
105                       Next k

106                       MakeJuliaLiteral = "reshape(" & VecLit & "," & VBA.Join(DimStr, ",") & ")"
107               End Select
108           Case vbObject
109               If TypeName(x) = "Dictionary" Then
                      Dim Key As Variant
                      Dim Tokens() As String
                      Dim v As Variant
110                   ReDim Tokens(1 To x.Count)
111                   k = 1
112                   For Each Key In x.Keys
113                       Tokens(k) = MakeJuliaLiteral(Key) & " => " & MakeJuliaLiteral(x(Key))
114                       k = k + 1
115                   Next Key

116                   MakeJuliaLiteral = "Dict(" & VBA.Join(Tokens, ",") & ")"
117               Else
118                   Throw "Variable of type " + TypeName(x) + " is not handled"
119               End If
120           Case Else
121               Throw "Variable of type " + TypeName(x) + " is not handled"
122       End Select

123       Exit Function
ErrHandler:
124       ReThrow "MakeJuliaLiteral", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FlattenArrayElements
' Purpose    : Flatten any VBA array (rank >= 1) to a vector of Julia element literals.
'              Traversal order: column-major (dim 1 varies fastest).
'              Returns number of elements; also sets AllSameType and FirstType.
' -----------------------------------------------------------------------------------------------------------------------
Private Function FlattenArrayElements(ByRef A As Variant, ByRef Elems() As String, ByRef AllSameType As Boolean, _
          ByRef FirstType As Long) As Long
          
          Dim Count As Long
          Dim d As Long
          Dim Idx() As Long
          Dim Lb() As Long
          Dim n As Long
          Dim Total As Long
          Dim Ub() As Long
          Dim v As Variant
          
1         On Error GoTo ErrHandler
2         n = NumDimensions(A)
3         If n <= 0 Then
4             ReDim Elems(1 To 0)
5             AllSameType = True
6             FirstType = vbEmpty
7             FlattenArrayElements = 0
8             Exit Function
9         End If

10        ReDim Lb(1 To n)
11        ReDim Ub(1 To n)
12        ReDim Idx(1 To n)

13        Total = 1
14        For d = 1 To n
15            Lb(d) = LBound(A, d)
16            Ub(d) = UBound(A, d)
17            Idx(d) = Lb(d)
18            Total = Total * (Ub(d) - Lb(d) + 1)
19        Next d

20        ReDim Elems(1 To Total)

          ' Initialize homogeneity checks from the first element
21        v = GetAt(A, Idx)
22        FirstType = VarType(v)
23        AllSameType = True

24        Count = 0
25        Do
26            Count = Count + 1
27            v = GetAt(A, Idx)
28            Elems(Count) = MakeJuliaLiteral(v)
29            If AllSameType Then
30                If VarType(v) <> FirstType Then AllSameType = False
31            End If

              ' Increment indices: dim 1 fastest (column-major)
32            d = 1
33            Do While d <= n
34                Idx(d) = Idx(d) + 1
35                If Idx(d) <= Ub(d) Then Exit Do
36                Idx(d) = Lb(d)
37                d = d + 1
38            Loop
39            If d > n Then Exit Do
40        Loop

41        FlattenArrayElements = Total

42        Exit Function
ErrHandler:
43        ReThrow "FlattenArrayElements", Err
End Function

