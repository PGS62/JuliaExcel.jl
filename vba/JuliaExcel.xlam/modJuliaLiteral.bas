Attribute VB_Name = "modJuliaLiteral"
Option Explicit
Option Private Module

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MakeJuliaLiteral
' Purpose    : Convert an array into a string which julia will parse as x:
' Parameters :
'  x         : The variable in Excel or VBA.
'  OneDtoTwoD: If TRUE, then if x is a one-dimensional array, the return will be a string that will be parsed by Julia
'              to a two-dimensional array with one row. Otherwise, if FALSE, the return will be a string that will be
'              parsed to a one-dimensional array.
'
'              Unfortunately, the "best, most natural" value for OneDtoTwoD is different when calling from the worksheet
'              versus when calling from VBA:
'
'           1) From the worksheet: Excel treats 1-dimensional arrays as as if they were 2-dimensional arrays with a single
'              row, in the following senses:
'           a) If you call a VBA UDF from a worksheet formula, and the UDF returns a 1-d array then the formula spills
'              to a range with a single row.
'           b) If you pass a 2-d array with a single row to a VBA UDF then the variable appears in VBA as a 1-dimensional
'              array. This holds for string literals such as {1,2,3,4} or a call to (say) SEQUENCE(1,3) or to a reference
'              to a range with only one row.
'
'              But we almost certainly want a 2-d array with one row (or range with one row) in Excel to "arrive in Julia"
'              as two dimensional. That means that, in the context of worksheet formulas, one-dimensional value for the
'              argument x must result in a string that Julia will parse to a two-dimensional array with a single row. And
'              that's easy to arrange - just use space character as the delimiter, Julia parses [1 2 3] as 2-d with one row.

'              By contrast, VBA supports 1-d arrays "properly" so when calling into Julia from VBA we would expect 1-d arrays to
'              "arrive in Julia" with one dimension
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
Function MakeJuliaLiteral(x As Variant, OneDtoTwoD As Boolean)
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
56                            Tmp(i) = MakeJuliaLiteral(x(i), OneDtoTwoD)
57                            If AllSameType Then
58                                If VarType(x(i)) <> FirstType Then
59                                    AllSameType = False
60                                End If
61                            End If
62                        Next i
63                        MakeJuliaLiteral = IIf(AllSameType, "[", "Any[") & VBA.Join$(Tmp, IIf(OneDtoTwoD, " ", ",")) & "]"
64                    Case 2
65                        ReDim onerow(LBound(x, 2) To UBound(x, 2))
66                        ReDim Tmp(LBound(x, 1) To UBound(x, 1))
67                        FirstType = VarType(x(LBound(x, 1), LBound(x, 2)))
68                        AllSameType = True
69                        For i = LBound(x, 1) To UBound(x, 1)
70                            For j = LBound(x, 2) To UBound(x, 2)
71                                onerow(j) = MakeJuliaLiteral(x(i, j), OneDtoTwoD)
72                                If AllSameType Then
73                                    If VarType(x(i, j)) <> FirstType Then
74                                        AllSameType = False
75                                    End If
76                                End If
77                            Next j
78                            Tmp(i) = VBA.Join$(onerow, " ")
79                        Next i

80                        MakeJuliaLiteral = IIf(AllSameType, "[", "Any[") & VBA.Join$(Tmp, ";") & "]"
                          'One column case is tricky, could change this code when using Julia 1.7
                          'https://discourse.julialang.org/t/show-versus-parse-and-arrays-with-2-dimensions-but-only-one-column/70142/2
81                        If UBound(x, 2) = LBound(x, 2) Then
                              Dim NR As Long
82                            NR = UBound(x, 1) - LBound(x, 1) + 1
83                            MakeJuliaLiteral = "reshape(" & MakeJuliaLiteral & "," & CStr(NR) & ",1)"
84                        End If
85                    Case Else
86                        Throw "case more than two dimensions not handled" 'In VBA there's no way to handle arrays with arbitrary number of dimensions. Easy in Julia!
87                End Select


88            Case Else
89                Throw "Variable of type " + TypeName(x) + " is not handled"
90        End Select

91        Exit Function
ErrHandler:
92        Throw "#MakeJuliaLiteral (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

