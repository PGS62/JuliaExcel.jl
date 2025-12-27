Attribute VB_Name = "modPrivate"
Option Explicit
Option Private Module

Function GetAt(ByRef A As Variant, ByRef idx() As Long) As Variant
1         On Error GoTo ErrHandler
2         Select Case UBound(idx)
              Case 1: GetAt = A(idx(1))
3             Case 2: GetAt = A(idx(1), idx(2))
4             Case 3: GetAt = A(idx(1), idx(2), idx(3))
5             Case 4: GetAt = A(idx(1), idx(2), idx(3), idx(4))
6             Case 5: GetAt = A(idx(1), idx(2), idx(3), idx(4), idx(5))
7             Case 6: GetAt = A(idx(1), idx(2), idx(3), idx(4), idx(5), idx(6))
8             Case 7: GetAt = A(idx(1), idx(2), idx(3), idx(4), idx(5), idx(6), idx(7))
9             Case 8: GetAt = A(idx(1), idx(2), idx(3), idx(4), idx(5), idx(6), idx(7), idx(8))
10            Case Else
11                Throw "Rank > 8 not supported"
12        End Select

13        Exit Function
ErrHandler:
14        ReThrow "GetAt", Err
End Function

Function SetAt(ByRef A As Variant, ByRef idx() As Long, x As Variant) As Variant
1         On Error GoTo ErrHandler
2         Select Case UBound(idx)
              Case 1: A(idx(1)) = x
3             Case 2: A(idx(1), idx(2)) = x
4             Case 3: A(idx(1), idx(2), idx(3)) = x
5             Case 4: A(idx(1), idx(2), idx(3), idx(4)) = x
6             Case 5: A(idx(1), idx(2), idx(3), idx(4), idx(5)) = x
7             Case 6: A(idx(1), idx(2), idx(3), idx(4), idx(5), idx(6)) = x
8             Case 7: A(idx(1), idx(2), idx(3), idx(4), idx(5), idx(6), idx(7)) = x
9             Case 8: A(idx(1), idx(2), idx(3), idx(4), idx(5), idx(6), idx(7), idx(8)) = x
10            Case Else
11                Throw "Rank > 8 not supported"
12        End Select

13        Exit Function
ErrHandler:
14        ReThrow "SetAt", Err
End Function

Function ArrayGetLinear(arr As Variant, idx As Long) As Variant
          Dim Dims As Long
          Dim Sizes() As Long, Lb() As Long
          Dim i As Long, t As Long, Offset As Long
          Dim Indices() As Long
          
1         On Error GoTo ErrHandler
2         Dims = NumDimensions(arr)
3         ReDim Sizes(1 To Dims)
4         ReDim Lb(1 To Dims)
5         ReDim Indices(1 To Dims)
          
          ' Collect bounds and sizes
6         For i = 1 To Dims
7             Lb(i) = LBound(arr, i)
8             Sizes(i) = UBound(arr, i) - Lb(i) + 1
9         Next i
          
          ' Convert linear index to multidimensional indices (column-major)
10        t = idx - 1
11        For i = 1 To Dims
12            Offset = t Mod Sizes(i)
13            Indices(i) = Lb(i) + Offset
14            t = t \ Sizes(i)
15        Next i
             
16        ArrayGetLinear = GetAt(arr, Indices)

17        Exit Function
ErrHandler:
18        ReThrow "ArrayGetLinear", Err
End Function


Sub ArraySetLinear(arr As Variant, idx As Long, element As Variant)
          Dim Dims As Long
          Dim Sizes() As Long, Lb() As Long
          Dim i As Long, t As Long, Offset As Long
          Dim Indices() As Long
          
1         On Error GoTo ErrHandler
2         Dims = NumDimensions(arr)
3         ReDim Sizes(1 To Dims)
4         ReDim Lb(1 To Dims)
5         ReDim Indices(1 To Dims)
          
          ' Collect bounds and sizes
6         For i = 1 To Dims
7             Lb(i) = LBound(arr, i)
8             Sizes(i) = UBound(arr, i) - Lb(i) + 1
9         Next i
          
          ' Convert linear index to multidimensional indices (column-major)
10        t = idx - 1
11        For i = 1 To Dims
12            Offset = t Mod Sizes(i)
13            Indices(i) = Lb(i) + Offset
14            t = t \ Sizes(i)
15        Next i
             
16        SetAt arr, Indices, element

17        Exit Sub
ErrHandler:
18        ReThrow "ArraySetLinear", Err
End Sub

