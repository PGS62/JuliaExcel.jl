Attribute VB_Name = "modPrivate"
' Copyright (c) 2021-2025 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme
Option Explicit
Option Private Module

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetAt
' Purpose    : Returns the element in array Arr at position Idx.
' -----------------------------------------------------------------------------------------------------------------------
Function GetAt(ByRef Arr As Variant, ByRef Idx() As Long) As Variant
1         On Error GoTo ErrHandler
2         Select Case UBound(Idx)
              Case 1: GetAt = Arr(Idx(1))
3             Case 2: GetAt = Arr(Idx(1), Idx(2))
4             Case 3: GetAt = Arr(Idx(1), Idx(2), Idx(3))
5             Case 4: GetAt = Arr(Idx(1), Idx(2), Idx(3), Idx(4))
6             Case 5: GetAt = Arr(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5))
7             Case 6: GetAt = Arr(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5), Idx(6))
8             Case 7: GetAt = Arr(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5), Idx(6), Idx(7))
9             Case 8: GetAt = Arr(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5), Idx(6), Idx(7), Idx(8))
10            Case Else
11                Throw "Rank > 8 not supported"
12        End Select

13        Exit Function
ErrHandler:
14        ReThrow "GetAt", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SetAt
' Purpose    : Sets the element at position Idx in array Arr to be Value.
' -----------------------------------------------------------------------------------------------------------------------
Function SetAt(ByRef Arr As Variant, ByRef Idx() As Long, Value As Variant) As Variant
1         On Error GoTo ErrHandler
2         Select Case UBound(Idx)
              Case 1: Arr(Idx(1)) = Value
3             Case 2: Arr(Idx(1), Idx(2)) = Value
4             Case 3: Arr(Idx(1), Idx(2), Idx(3)) = Value
5             Case 4: Arr(Idx(1), Idx(2), Idx(3), Idx(4)) = Value
6             Case 5: Arr(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5)) = Value
7             Case 6: Arr(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5), Idx(6)) = Value
8             Case 7: Arr(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5), Idx(6), Idx(7)) = Value
9             Case 8: Arr(Idx(1), Idx(2), Idx(3), Idx(4), Idx(5), Idx(6), Idx(7), Idx(8)) = Value
10            Case Else
11                Throw "Rank > 8 not supported"
12        End Select

13        Exit Function
ErrHandler:
14        ReThrow "SetAt", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetAtLinear
' Purpose    : Returns the element of a multi-dimensional VBA array Arr addressed by a single linear index Idx.
'              The mapping from Idx to (i1, i2, ..., in) mimics Julia's linear indexing semantics:
'                - Column-major order (first dimension varies fastest, last varies slowest)
'                - 1-based linear index (Idx = 1 corresponds to (LBound(Arr,1), LBound(Arr,2), ..., LBound(Arr,n)))
' -----------------------------------------------------------------------------------------------------------------------
Function GetAtLinear(Arr As Variant, Idx As Long) As Variant
          Dim Dims As Long
          Dim i As Long
          Dim Indices() As Long
          Dim Lb() As Long
          Dim Offset As Long
          Dim Sizes() As Long
          Dim t As Long
          
1         On Error GoTo ErrHandler
2         Dims = NumDimensions(Arr)
3         ReDim Sizes(1 To Dims)
4         ReDim Lb(1 To Dims)
5         ReDim Indices(1 To Dims)
          
6         For i = 1 To Dims
7             Lb(i) = LBound(Arr, i)
8             Sizes(i) = UBound(Arr, i) - Lb(i) + 1
9         Next i
          
10        t = Idx - 1
11        For i = 1 To Dims
12            Offset = t Mod Sizes(i)
13            Indices(i) = Lb(i) + Offset
14            t = t \ Sizes(i)
15        Next i
             
16        GetAtLinear = GetAt(Arr, Indices)

17        Exit Function
ErrHandler:
18        ReThrow "GetAtLinear", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SetAtLinear
' Purpose    : Sets the element of a multi-dimensional VBA array Arr addressed by a single linear index Idx to have value Value.
'              The mapping from Idx to (i1, i2, ..., in) mimics Julia's linear indexing semantics:
'                - Column-major order (first dimension varies fastest, last varies slowest)
'                - 1-based linear index (Idx = 1 corresponds to (LBound(Arr,1), LBound(Arr,2), ..., LBound(Arr,n)))
' -----------------------------------------------------------------------------------------------------------------------
Sub SetAtLinear(Arr As Variant, Idx As Long, Value As Variant)
          Dim Dims As Long
          Dim i As Long
          Dim Indices() As Long
          Dim Lb() As Long
          Dim Offset As Long
          Dim Sizes() As Long
          Dim t As Long
          
1         On Error GoTo ErrHandler
2         Dims = NumDimensions(Arr)
3         ReDim Sizes(1 To Dims)
4         ReDim Lb(1 To Dims)
5         ReDim Indices(1 To Dims)
          
6         For i = 1 To Dims
7             Lb(i) = LBound(Arr, i)
8             Sizes(i) = UBound(Arr, i) - Lb(i) + 1
9         Next i
          
10        t = Idx - 1
11        For i = 1 To Dims
12            Offset = t Mod Sizes(i)
13            Indices(i) = Lb(i) + Offset
14            t = t \ Sizes(i)
15        Next i
             
16        SetAt Arr, Indices, Value

17        Exit Sub
ErrHandler:
18        ReThrow "SetAtLinear", Err
End Sub

