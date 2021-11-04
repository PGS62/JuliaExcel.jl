Attribute VB_Name = "modIntellisense"
Option Explicit

Sub CreateIntellisenseWorkbook()

          Dim wb As Workbook
          Dim targetsheet As Worksheet
          Dim FnName As String
          Dim SourceRange As Range
          Dim i As Long
          Dim j As Long

1         On Error GoTo ErrHandler
2         Set wb = Application.Workbooks.Add()
3         Set targetsheet = wb.Worksheets(1)

4         targetsheet.Name = "_Intellisense_"

5         targetsheet.Cells(1, 1).value = "FunctionInfo"
6         targetsheet.Cells(1, 2).value = "'1.0"

7         For i = 1 To 2
8             FnName = Choose(i, "CSVRead", "CSVWrite")
9             Set SourceRange = ThisWorkbook.Worksheets("Help").Range(FnName & "Args")
10            targetsheet.Cells(1 + i, 1) = FnName
11            targetsheet.Cells(1 + i, 2) = SourceRange.Cells(1, 1).Offset(-2, 1).value
12            For j = 1 To SourceRange.Rows.Count
13                targetsheet.Cells(1 + i, 2 * (1 + j)).value = SourceRange.Cells(j, 1).value
14                targetsheet.Cells(1 + i, 1 + 2 * (1 + j)).value = SourceRange.Cells(j, 2).value
15            Next j
16        Next i

17        With targetsheet.UsedRange
18            .Columns.ColumnWidth = 40
19            .WrapText = True
20            .VerticalAlignment = xlVAlignCenter
21            .Columns.AutoFit
22        End With
23        Application.DisplayAlerts = False
24        wb.SaveAs ThisWorkbook.Path & "\VBA-CSV-Intellisense.xlsx", xlOpenXMLWorkbook
25        wb.Close False

26        Exit Sub
ErrHandler:
27        Throw "#CreateIntellisenseWorkbook (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub
