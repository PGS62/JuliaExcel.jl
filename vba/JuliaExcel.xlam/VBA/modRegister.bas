Attribute VB_Name = "modRegister"
' Copyright (c) 2021-2026 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme

Option Explicit
Option Private Module

Public Sub RegisterFunctions()
1         On Error GoTo ErrHandler
2         RegisterJuliaExcelFunctionsWithFunctionWizard
3         On Error Resume Next
4         AddIns("Excel-DNA IntelliSense Host").Installed = False
5         AddIns("Excel-DNA IntelliSense Host").Installed = True
6         Exit Sub
ErrHandler:
7         MsgBox "#RegisterFunctions (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RegisterJuliaExcelFunctionsWithFunctionWizard
' Purpose    : Register functions with the Excel function wizard, taking the information form the Intellisense sheet
'              that is also parsed by Excel.DNA Intellisense add-in.
'              This method does not need to be run at "Load Time", but at "add-in creation time"
' -----------------------------------------------------------------------------------------------------------------------
Private Sub RegisterJuliaExcelFunctionsWithFunctionWizard()

          Dim ArgDescs() As String
          Dim c As Range
          Dim Description As String
          Dim FunctionName As String
          Dim i As Long
          Dim NumArgs As Variant
          Dim OldIsAddinStatus As Boolean
          Dim OldSaveStatus As Boolean
          Dim rngArgsAndArgDescs As Range
          Dim rngFunctions As Range
          
1         On Error GoTo ErrHandler
2         OldSaveStatus = ThisWorkbook.Saved
3         OldIsAddinStatus = ThisWorkbook.IsAddin
          'Without setting .IsAddin to False, I see errors:
          '"Cannot edit a macro on a hidden workbook. Unhide the workbook using the Unhide command."
          'Not ideal, setting IsAddin to False causes screen flicker.
4         If OldIsAddinStatus Then
5             Application.ScreenUpdating = False
6             ThisWorkbook.IsAddin = False
7         End If

8         With shIntellisense
9             Set rngFunctions = .Range(.Cells(2, 1), .Cells(1, 1).End(xlDown))
10        End With

11        For Each c In rngFunctions.Cells
12            FunctionName = c.Value
13            Description = c.Offset(0, 1).Value
        
14            If IsEmpty(c.Offset(, 3).Value) Then
15                NumArgs = 0
16            Else
17                Set rngArgsAndArgDescs = Range(c.Offset(, 3), c.Offset(, 3).End(xlToRight))
18                NumArgs = rngArgsAndArgDescs.Columns.Count / 2
19                ReDim ArgDescs(1 To NumArgs)
20                For i = 1 To NumArgs
21                    ArgDescs(i) = Trim255(rngArgsAndArgDescs.Cells(1, i * 2).Value)
22                Next i
23            End If

24            If NumArgs = 0 Then
25                Application.MacroOptions FunctionName, Trim255(Description)
26            Else
27                Application.MacroOptions FunctionName, Trim255(Description), , , , , , "JuliaExcel", , , ArgDescs
28            End If
29        Next c
30        If OldIsAddinStatus Then
31            ThisWorkbook.IsAddin = True
32            ThisWorkbook.Saved = OldSaveStatus
33        End If

34        Exit Sub
ErrHandler:
35        Debug.Print "#RegisterJuliaExcelFunctionsWithFunctionWizard (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Trim255
' Purpose    : Truncates help text (from the _Intellisense worksheet) to a length the old-fashioned
'              Excel Function Wizard accepts, and replaces any non-ASCII character with an ASCII
'              equivalent (or, failing that, an underscore) - the Function Wizard's registration
'              mechanism is ASCII-only, and text typed or pasted into a worksheet cell (e.g. from
'              Word or Outlook) commonly picks up "smart" typographic characters - curly quotes,
'              en/em dashes, an ellipsis - that would otherwise corrupt the exported .bas source.
' -----------------------------------------------------------------------------------------------------------------------
Private Function Trim255(Text As String) As String
          Dim i As Long
          Dim Res As String

          'Common "smart" typographic substitutions, handled before the generic per-character
          'fallback below so they read naturally rather than as underscores.
1         Res = Text
2         Res = Replace(Res, ChrW(8216), "'")       'left single quotation mark
3         Res = Replace(Res, ChrW(8217), "'")       'right single quotation mark
4         Res = Replace(Res, ChrW(8220), Chr(34))   'left double quotation mark
5         Res = Replace(Res, ChrW(8221), Chr(34))   'right double quotation mark
6         Res = Replace(Res, ChrW(8211), "-")       'en dash
7         Res = Replace(Res, ChrW(8212), "-")       'em dash
8         Res = Replace(Res, ChrW(8230), "...")     'ellipsis

          'Anything else non-ASCII becomes an underscore.
9         For i = 1 To Len(Res)
10            If AscW(Mid$(Res, i, 1)) > 127 Then Mid$(Res, i, 1) = "_"
11        Next i

12        If Len(Res) < 255 Then
13            Trim255 = Res
14        Else
15            Trim255 = Left$(Res, 252) & "..."
16        End If
End Function

' Confirms Trim255's non-ASCII handling: common "smart" typographic characters get sensible ASCII
' equivalents, anything else non-ASCII becomes an underscore, and the 255-character truncation still
' works correctly on the result. Trim255 previously just truncated (see git history) - the ellipsis
' character it used to append there was itself non-ASCII, which corrupted the exported .bas source
' once it made its way into a registered function/argument description via the _Intellisense sheet.
' Lives here (not modTest.bas) because Trim255 is Private to this module.
Function TestTrim255() As Boolean
          Dim OK As Boolean

1         On Error GoTo ErrHandler
2         OK = True

          'Smart quotes, dashes and an ellipsis each get a sensible ASCII equivalent.
3         OK = OK And Trim255(ChrW(8216) & "a" & ChrW(8217)) = "'a'"
4         OK = OK And Trim255(ChrW(8220) & "a" & ChrW(8221)) = Chr(34) & "a" & Chr(34)
5         OK = OK And Trim255("a" & ChrW(8211) & "b") = "a-b"
6         OK = OK And Trim255("a" & ChrW(8212) & "b") = "a-b"
7         OK = OK And Trim255("a" & ChrW(8230) & "b") = "a...b"

          'Anything else non-ASCII becomes an underscore (infinity symbol, not specifically handled).
8         OK = OK And Trim255("a" & ChrW(8734) & "b") = "a_b"

          'A plain ASCII string under 255 characters passes through unchanged.
9         OK = OK And Trim255("hello") = "hello"

          'Truncation to 255 characters (as "...") still works, on already-sanitised text.
10        OK = OK And Len(Trim255(String(300, "x"))) = 255
11        OK = OK And Right$(Trim255(String(300, "x")), 3) = "..."

12        TestTrim255 = OK
13        Exit Function
ErrHandler:
14        PrintTwice ReThrow("TestTrim255", Err, True)
15        TestTrim255 = False
End Function
