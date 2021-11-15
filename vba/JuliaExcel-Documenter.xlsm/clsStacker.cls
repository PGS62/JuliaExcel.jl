VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsStacker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
' -----------------------------------------------------------------------------------------------------------------------
' Module    : clsStacker
' Author    : Philip Swannell
' Date      : 19-Jun-2013
' Purpose   : Class to implement "accumulation by stacking" of arrays of data. Class has
'             methods: StackData and Report. Restriction: Every call to Stack must pass in an
'             array with the same number of columns as previous calls - that number is initialised
'             in the first call to Stack.
'Data passed to StackData can be singleton (Long, Double, Boolean, String, Date etc) or 1-d array or 2-d array either 1-based or 0-based.
'             Return from Report is always a two dimensional array (1-based)
' 16-May-2019 Added methods Stack0D, Stack1D and Stack2D. Faster for stacking data when we know dimensionality.

'Complexity:
'             Because the Redim Preserve operation can only change the number of columns
'             of a two dimensional array, we have to transpose the data in both the StackData
'             and Report methods. ReportAsTranspose avoids doing a Transpose operation and is therefore a little quicker.
' -----------------------------------------------------------------------------------------------------------------------
Option Explicit
Private m_TheArray() As Variant
Private m_NumCols As Long
Private m_NumRows As Long
Private m_LastWrittenRow As Long

Public Function Report()
          Dim ArrayToReport() As Variant
          Dim i As Long
          Dim j As Long
1         On Error GoTo ErrHandler

2         If m_LastWrittenRow = 0 Then
3             Report = "#Nothing to report!"
4         Else
5             ReDim ArrayToReport(1 To m_LastWrittenRow, 1 To m_NumCols)
6             For i = 1 To m_LastWrittenRow
7                 For j = 1 To m_NumCols
8                     ArrayToReport(i, j) = m_TheArray(j, i)
9                 Next j
10            Next i
11            Report = ArrayToReport
12        End If

13        Exit Function
ErrHandler:
14        Throw "#clsStacker.Report (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Public Function NumRows()
1         NumRows = m_LastWrittenRow
End Function
' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ReportInTranspose
' Author    : Philip Swannell
' Date      : 19-Jun-2013
' Purpose   : Returns the data in transpose form, which is actually quicker to do...
' -----------------------------------------------------------------------------------------------------------------------
Public Function ReportInTranspose()
1         On Error GoTo ErrHandler

2         If m_LastWrittenRow = 0 Then
3             Report = "#Nothing to report!"
4         Else
5             ReDim Preserve m_TheArray(1 To m_NumCols, 1 To m_LastWrittenRow)
6             ReportInTranspose = m_TheArray
7         End If

8         Exit Function
ErrHandler:
9         Throw "#ReportInTranspose (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function Max(x As Long, y As Long)
1         If x > y Then
2             Max = x
3         Else
4             Max = y
5         End If
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : StackData
' Author     : Philip Swannell
' Date       : 16-May-2019
' Purpose    : Use to Stack data of either 0,1 or 2 dimensions
' -----------------------------------------------------------------------------------------------------------------------
Public Sub StackData(DataToStack As Variant)
1         On Error GoTo ErrHandler
2         Select Case NumDimensions(DataToStack)
              Case 0
3                 Stack0D DataToStack
4             Case 1
5                 Stack1D DataToStack
6             Case 2
7                 Stack2D DataToStack
8             Case Else
9                 Throw "Too many dimensions to stack!"
10        End Select
11        Exit Sub
ErrHandler:
12        Throw "#clsStacker.StackData (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Stack0D
' Author     : Philip Swannell
' Date       : 16-May-2019
' Purpose    : Version of StackData for use when we know we are stacking a singleton
' -----------------------------------------------------------------------------------------------------------------------
Public Sub Stack0D(DataToStack As Variant)

1         On Error GoTo ErrHandler

2         If m_NumCols = 0 Then m_NumCols = 1

3         If 1 <> m_NumCols Then Throw "Inconsistent number of columns in DataToStack"
4         If m_LastWrittenRow + 1 > m_NumRows Then
              'Remember that we have to accumulate in the transpose
5             ReDim Preserve m_TheArray(1 To m_NumCols, 1 To Max(2 * m_NumRows, m_NumRows + 1))
6             m_NumRows = UBound(m_TheArray, 2)
7         End If

8         m_TheArray(1, m_LastWrittenRow + 1) = DataToStack
9         m_LastWrittenRow = m_LastWrittenRow + 1

10        Exit Sub
ErrHandler:
11        Throw "#clsStacker.Stack0D (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Stack1D
' Author     : Philip Swannell
' Date       : 16-May-2019
' Purpose    : Version of StackData for use when we know we are stacking 1-dimensional array
' -----------------------------------------------------------------------------------------------------------------------
Public Sub Stack1D(DataToStack As Variant)
          Dim Base As Long
          Dim j As Long
          Dim NumCols As Long

1         On Error GoTo ErrHandler

2         Base = LBound(DataToStack)
3         NumCols = UBound(DataToStack) - Base + 1

4         If m_NumCols = 0 Then m_NumCols = NumCols

5         If NumCols <> m_NumCols Then Throw "Inconsistent number of columns in DataToStack"
6         If m_LastWrittenRow + 1 > m_NumRows Then
              'Remember that we have to accumulate in the transpose
7             ReDim Preserve m_TheArray(1 To m_NumCols, 1 To Max(2 * m_NumRows, m_NumRows + 1))
8             m_NumRows = UBound(m_TheArray, 2)
9         End If

10        For j = 1 To NumCols
11            m_TheArray(j, m_LastWrittenRow + 1) = DataToStack(j + Base - 1)
12        Next j
13        m_LastWrittenRow = m_LastWrittenRow + 1

14        Exit Sub
ErrHandler:
15        Throw "#clsStacker.Stack1D (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Stack2D
' Author     : Philip Swannell
' Date       : 16-May-2019
' Purpose    : Version of StackData for use when we know we are stacking 2-dimensional array
' -----------------------------------------------------------------------------------------------------------------------
Public Sub Stack2D(DataToStack As Variant)
          Dim Base As Long
          Dim i As Long
          Dim j As Long
          Dim NumCols As Long
          Dim NumRows

1         On Error GoTo ErrHandler

2         Base = LBound(DataToStack, 1)
3         NumRows = UBound(DataToStack, 1) - Base + 1
4         NumCols = UBound(DataToStack, 2) - Base + 1

5         If m_NumCols = 0 Then m_NumCols = NumCols

6         If NumCols <> m_NumCols Then Throw "Inconsistent number of columns in DataToStack"
7         If m_LastWrittenRow + NumRows > m_NumRows Then
              'Remember that we have to accumulate in the transpose
8             ReDim Preserve m_TheArray(1 To m_NumCols, 1 To Max(2 * m_NumRows, m_NumRows + NumRows))
9             m_NumRows = UBound(m_TheArray, 2)
10        End If

11        For i = 1 To NumRows
12            For j = 1 To NumCols
13                m_TheArray(j, m_LastWrittenRow + i) = DataToStack(i + Base - 1, j + Base - 1)
14            Next j
15        Next i
16        m_LastWrittenRow = m_LastWrittenRow + NumRows

17        Exit Sub
ErrHandler:
18        Throw "#clsStacker.Stack2D (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
