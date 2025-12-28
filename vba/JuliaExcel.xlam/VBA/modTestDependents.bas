Attribute VB_Name = "modTestDependents"
' Copyright (c) 2021-2025 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme

' =========================================================================================
' Module: ArrayCompare, written by Copilot 22 Dec 2025, with amendments by Philip Swannell
' Prompt was as follows:
' Please provide a VBA function that tests if two arrays are identical, i.e. same number
' of dimensions, same length of dimensions, same contents. Assume elements are singletons
' i.e. not arrays or objects
'
' Methods in this module should be called only from modTest.
' =========================================================================================
Option Explicit
Option Private Module

' Public entry point.
' Returns True iff both arrays are:
'   - arrays (both initialized or both uninitialized in the same way),
'   - same number of dimensions,
'   - same bounds per dimension (LBound/UBound),
'   - and every corresponding element compares equal (scalars only).
Function ArraysIdentical(ByVal A As Variant, ByVal B As Variant) As Boolean
          ' --- Quick identity / trivial cases ---
1         If Not IsArray(A) Or Not IsArray(B) Then
2             ArraysIdentical = False
3             Exit Function
4         End If

          Dim InitA As Boolean
          Dim InitB As Boolean
5         InitA = IsArrayInitialized(A)
6         InitB = IsArrayInitialized(B)

7         If InitA Xor InitB Then
8             ArraysIdentical = False
9             Exit Function
10        End If

          ' Both uninitialized: consider identical
11        If Not InitA And Not InitB Then
12            ArraysIdentical = True
13            Exit Function
14        End If

          ' --- Dimension checks ---
          Dim nA As Long
          Dim nB As Long
15        nA = NumDimensions(A)
16        nB = NumDimensions(B)
17        If nA <> nB Then
18            ArraysIdentical = False
19            Exit Function
20        End If

          Dim d As Long
21        For d = 1 To nA
22            If LBound(A, d) <> LBound(B, d) Then
23                ArraysIdentical = False
24                Exit Function
25            End If
26            If UBound(A, d) <> UBound(B, d) Then
27                ArraysIdentical = False
28                Exit Function
29            End If
30        Next d

          ' --- Element-wise comparison (rank-agnostic) ---
31        If nA = 0 Then
32            ArraysIdentical = True
33            Exit Function
34        End If

35        ArraysIdentical = WalkAndCompare(A, B, nA)
End Function

' Returns True iff two scalar values are equal under the following rules:
' - String: binary compare (case-sensitive, invariant)
' - Numeric/Boolean/Date/Currency: = comparison
' - Empty equals Empty; Null equals Null
' - Otherwise, falls back to "=" (will raise if invalid types) together with a check that x and y are of the same type
Function ScalarsEqual(ByVal x As Variant, ByVal y As Variant) As Boolean
          ' Handle Empty and Null explicitly
1         If IsEmpty(x) Then
2             ScalarsEqual = IsEmpty(y)
3             Exit Function
4         ElseIf IsEmpty(y) Then
5             ScalarsEqual = False
6             Exit Function
7         End If

8         If IsNull(x) Then
9             ScalarsEqual = IsNull(y)
10            Exit Function
11        ElseIf IsNull(y) Then
12            ScalarsEqual = False
13            Exit Function
14        End If

          ' Strings: enforce binary comparison (independent of Option Compare)
15        If VarType(x) = vbString And VarType(y) = vbString Then
16            ScalarsEqual = (StrComp(CStr(x), CStr(y), vbBinaryCompare) = 0)
17            Exit Function
18        End If

          ' Default: numeric/boolean/date/currency, etc.
19        On Error GoTo NotComparable
20        ScalarsEqual = (x = y) And (VarType(x) = VarType(y))
21        Exit Function

NotComparable:
          ' If types are incomparable, they are not equal.
22        ScalarsEqual = False
End Function

' Iterate all indices of an array of rank n and compare elements.
' This avoids hard-coding nested loops.
Function WalkAndCompare(ByRef A As Variant, ByRef B As Variant, ByVal n As Long) As Boolean
          Dim d As Long
          Dim Idx() As Long
          Dim Lb() As Long
          Dim Ub() As Long

1         ReDim Idx(1 To n)
2         ReDim Lb(1 To n)
3         ReDim Ub(1 To n)

4         For d = 1 To n
5             Lb(d) = LBound(A, d)
6             Ub(d) = UBound(A, d)
7             Idx(d) = Lb(d)
8         Next d

9         Do
              ' Compare A(idx...) and B(idx...)
10            If Not ScalarsEqual(GetAt(A, Idx), GetAt(B, Idx)) Then
11                WalkAndCompare = False
12                Exit Function
13            End If

              ' Increment last dimension, then carry to higher dimensions as needed
14            d = n
15            Do While d >= 1
16                Idx(d) = Idx(d) + 1
17                If Idx(d) <= Ub(d) Then Exit Do
18                Idx(d) = Lb(d)
19                d = d - 1
20            Loop
21            If d = 0 Then Exit Do    ' completed all iterations
22        Loop

23        WalkAndCompare = True
End Function

' Detect whether a Variant array is actually allocated (has bounds).
Function IsArrayInitialized(ByVal A As Variant) As Boolean
1         On Error GoTo NotInit
2         If Not IsArray(A) Then Exit Function
          Dim Lb As Long
3         Lb = LBound(A, 1)  ' Will error if not initialized
4         IsArrayInitialized = True
5         Exit Function
NotInit:
6         IsArrayInitialized = False
End Function
