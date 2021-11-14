Attribute VB_Name = "modRegister"
Option Explicit
Option Private Module

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RegisterAll
' Purpose    : Register functions with the Excel function wizard, taking the information form the Intellisense sheet
'              that is also parsed by Excel.DNA Intellisense add-in.
' -----------------------------------------------------------------------------------------------------------------------
Sub RegisterAll()

          Dim ArgDescs() As String
          Dim c As Range
          Dim Description As String
          Dim FunctionName As String
          Dim i As Long
          Dim NumArgs
          Dim OldSaveStatus As Boolean
          Dim rngArgsAndArgDescs As Range
          Dim rngFunctions As Range
          
1         On Error GoTo ErrHandler
2         OldSaveStatus = ThisWorkbook.Saved
3         Application.ScreenUpdating = False
          'Without setting .IsAddin to False, I see errors:
          '"Cannot edit a macro on a hidden workbook. Unhide the workbook using the Unhide command."
          'Not ideal, setting IsAddin to False causes screen flicker.
4         ThisWorkbook.IsAddin = False

5         With shIntellisense
6             Set rngFunctions = .Range(.Cells(2, 1), .Cells(1, 1).End(xlDown))
7         End With

8         For Each c In rngFunctions.Cells
9             FunctionName = c.Value
10            Description = c.Offset(0, 1).Value
        
11            If IsEmpty(c.Offset(, 3).Value) Then
12                NumArgs = 0
13            Else
14                Set rngArgsAndArgDescs = Range(c.Offset(, 3), c.Offset(, 3).End(xlToRight))
15                NumArgs = rngArgsAndArgDescs.Columns.Count / 2
16                ReDim ArgDescs(1 To NumArgs)
17                For i = 1 To NumArgs
18                    ArgDescs(i) = rngArgsAndArgDescs.Cells(1, i * 2 - 1).Value
19                Next i
20            End If

21            If NumArgs = 0 Then
22                MacroOptions FunctionName, Description
23            Else
24                MacroOptions FunctionName, Description, ArgDescs
25            End If
26        Next c
27        ThisWorkbook.IsAddin = True
28        ThisWorkbook.Saved = OldSaveStatus


29        Exit Sub
ErrHandler:
30        Debug.Print "#RegisterAll (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Function MacroOptions(FunctionName As String, Description As String, Optional ArgDescs As Variant)
    Application.MacroOptions FunctionName, Description, , , , , gPackageName, , , , ArgDescs
    Exit Function
ErrHandler:
    Debug.Print "Warning from " + gPackageName + ": Registration of function " & FunctionName & " failed with error: " + Err.Description
End Function

