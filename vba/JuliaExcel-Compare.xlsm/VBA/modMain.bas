Attribute VB_Name = "modMain"
' Copyright (c) 2026 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme

Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RunOneTest
' Author     : Philip Swannell
' Date       : 20-Aug-2026
' Purpose    : Compare execution time for two formulas, one using JuliaExcel, the other using JuliaInExcel. Since JuliaInXL
'              does not seem to be callable from VBA, we test by pasting formulas into the worksheets JuliaExcel and JuliaInXL
' Parameters :
'  Description    : Narrative giving the purpose of the test - e.g. Latency
'  VectorLength   : The length of the vector passed to the JInXL/JE function
'  JEFormula      : The text of the formula using JuliaExcel's JuliaCall or JuliaEval
'  JInXLFormula   : The text of the formula using JuliaInXL's jlcall or jleval. To indicate that the call must be a Ctrl+Shift+Enter array formula, the first and last characters should be { and }.
'  NumCalls       : The number of times that the formula is evaluatd, with the average time recorded.
'  OneCallPerCell : If TRUE then rather than a single array formula, the test is applied by pasting the same formula to VectorLength cells. Suitable for latency tests.
'  ResultCellJE   : The cell to which the average execution time for the JuliaExcel function to be evaluated is written.
'  ResultCellJInXL: The cell to which the average execution time for the JuliaInXL function to be evaluated is written.
'  RatioCell      : The cell to which the ration of the two evaluation times is written
' -----------------------------------------------------------------------------------------------------------------------
Sub RunOneTest(Description As String, VectorLength As Long, ByVal JEFormula As String, ByVal JInXLFormula As String, _
          NumCalls As Long, OneCallPerCell As Boolean, ResultCellJE As Range, ResultCellJInXL As Range, RatioCell As Range)

          Dim InputDataAddress As String
          Dim NR As Long
          Dim TargetInputData As Range
          Dim TargetRangeJE As Range
          Dim TargetRangeJInXL As Range
          Dim UseInputData As Boolean
          
1         On Error GoTo ErrHandler

2         UseInputData = InStr(JEFormula, "InputData") > 0 Or _
              InStr(JInXLFormula, "InputData") > 0

3         JEFormula = Replace(JEFormula, "VectorLength", CStr(VectorLength))
4         JInXLFormula = Replace(JInXLFormula, "VectorLength", CStr(VectorLength))

5         If UseInputData Then
6             If IsInCollection(shInputData, "InputData") Then
7                 NR = shInputData.Range("InputData").Rows.Count
8                 If NR < VectorLength Then
9                     Set TargetInputData = shInputData.Range("InputData").Offset(NR).Resize(VectorLength - NR)
10                    TargetInputData.Value = Application.WorksheetFunction.RandArray(VectorLength - NR)
11                    shInputData.Names.Add "InputData", shInputData.Range("InputData").Resize(VectorLength)
12                End If
13            Else
14                Set TargetInputData = shInputData.Cells(1, 1).Resize(VectorLength)
15                TargetInputData.Value = Application.WorksheetFunction.RandArray(VectorLength)
16                shInputData.Names.Add "InputData", TargetInputData
17            End If
18        End If

19        If UseInputData Then
20            InputDataAddress = shInputData.Name & "!" & Replace(shInputData.Range("Inputdata").Resize(VectorLength).Address, "$", "")
21            JEFormula = Replace(JEFormula, "InputData", InputDataAddress)
22            JInXLFormula = Replace(JInXLFormula, "InputData", InputDataAddress)
23        End If

24        Set TargetRangeJE = FindTargetRange(shJuliaExcel, VectorLength)
25        Set TargetRangeJInXL = FindTargetRange(shJuliaInXL, VectorLength)

26        If JEFormula <> "" Then
27            LogMessage "Testing " & Description & " for JuliaExcel"
28            ResultCellJE.Value = PasteAndTimeFormula(TargetRangeJE, JEFormula, OneCallPerCell, NumCalls)
29        Else
30            ResultCellJE.ClearContents
31        End If

32        If JInXLFormula <> "" Then
33            LogMessage "Testing " & Description & " for JuliaInXL"
34            ResultCellJInXL.Value = PasteAndTimeFormula(TargetRangeJInXL, JInXLFormula, OneCallPerCell, NumCalls)
35        Else
36            ResultCellJInXL.ClearContents
37        End If

38        RatioCell.ClearContents
39        On Error Resume Next
40        RatioCell.Value = ResultCellJInXL.Value / ResultCellJE.Value
41        On Error GoTo ErrHandler

42        Exit Sub
ErrHandler:
43        ReThrow "RunOneTest", Err
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : RunManyTests
' Author     : Philip Swannell
' Date       : 20-Aug-2026
' Purpose    : Loop around the rows of range Main.TestSpecWithHeaders, calling RunOneTest for each row.
' Parameters :
'  FromRow:  Which row to start from, defaults to 2 (since the range has a header row).
'  ToRow  :  Which row to finish at, defaults to the last row.
' -----------------------------------------------------------------------------------------------------------------------
Sub RunManyTests(Optional FromRow As Long = 2, Optional ToRow As Long)

          Const cn_Description = 1
          Const cn_VectorLength = 2
          Const cn_JuliaExcelFormula = 3
          Const cn_JuliaInXLFormula = 4
          Const cn_NumCalls = 5
          Const cn_OneCallPerCell = 6
          Const cn_JuliaExcel_s = 7
          Const cn_JuliaInXL_s = 8
          Const cn_JuliaInXLOverJuliaExcel = 9
          
          Dim i As Long
          Dim OrigCalculation As Long
          Dim OrigProtect As Boolean

1         On Error GoTo ErrHandler
2         Application.ScreenUpdating = False
3         OrigCalculation = Application.Calculation
4         OrigProtect = shMain.ProtectContents

5         Application.Calculation = xlCalculationManual
6         shMain.Unprotect

7         shMain.Calculate
8         TestInstallation

9         If FromRow = 0 Then FromRow = 2
10        If FromRow < 2 Then Throw "FromRow must be at least 2"

11        ClearOutSheets

12        With shMain.Range("TestSpecWithHeaders")

13            If ToRow = 0 Then ToRow = .Rows.Count
14            If ToRow > .Rows.Count Then Throw "ToRow must be no greater then the number of rows in range " & shMain.Name & "!TestSpecWithHeaders"

15            For i = FromRow To ToRow

16                RunOneTest .Cells(i, cn_Description).Value, _
                      .Cells(i, cn_VectorLength).Value, _
                      .Cells(i, cn_JuliaExcelFormula).Value, _
                      .Cells(i, cn_JuliaInXLFormula).Value, _
                      .Cells(i, cn_NumCalls).Value, _
                      .Cells(i, cn_OneCallPerCell).Value, _
                      .Cells(i, cn_JuliaExcel_s), _
                      .Cells(i, cn_JuliaInXL_s), _
                      .Cells(i, cn_JuliaInXLOverJuliaExcel)
17            Next i

18        End With

19        AddAllResultsRangeNameToSheet shJuliaExcel
20        AddAllResultsRangeNameToSheet shJuliaInXL
21        AlignColumnWidths
22        shMain.Range("Results_Identical?").Calculate
23        shMain.Range("MaximumAbsoluteDifference").Calculate
24        shMain.Range("LogBase2_MaxAbsDiff").Calculate

25        LogMessage False
26        Application.Calculation = OrigCalculation
27        shMain.Protect , , OrigProtect
28        Exit Sub
ErrHandler:
29        MsgBox ReThrow("RunManyTests", Err, True), vbOKOnly + vbCritical, "Run Many Tests"
30        LogMessage False
31        Application.Calculation = OrigCalculation
32        shMain.Protect , , OrigProtect
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : PasteAndTimeFormula
' Author     : Philip Swannell
' Date       : 20-Aug-2026
' Purpose    : Pastes a single array formula to a range and returns the average calculation time over NumCalls calculations.
' Parameters :
'  TargetRange   : The range pasted to
'  Formula       : The formula, including the initial = character.
'  OneCallPerCell: If True then the same frmula is pasted to each cell in TargetRange, otherwise when Formula is wrapped
'                  by { and } a single CSE array formula is entered to the range, otherwise a the formula is entered to
'                  the top left cell of the range and assumed to spill to the entire range
'  NumCalls      : The number of recalculation over which we average.
' -----------------------------------------------------------------------------------------------------------------------
Function PasteAndTimeFormula(TargetRange As Range, Formula As String, OneCallPerCell As Boolean, NumCalls As Long) As Double
          Dim i As Long
          Dim t As Double
          Dim t1 As Double
          Dim t2 As Double

1         On Error GoTo ErrHandler

          Dim BaseMessage As String
2         BaseMessage = Application.StatusBar

3         If OneCallPerCell Then
4             TargetRange.Formula2 = Formula
5         ElseIf Left(Formula, 1) = "{" Then
6             If Right(Formula, 1) <> "}" Then Throw "Formulas that start with '{' must end with '}', but got '" & Formula & "'"
7             TargetRange.FormulaArray = Mid$(Formula, 2, Len(Formula) - 2)
8         Else
9             TargetRange.Cells(1, 1).Formula2 = Formula
10        End If

11        For i = 1 To NumCalls
12            LogMessage BaseMessage & " " & CStr(i) & "/" & CStr(NumCalls)
13            TargetRange.Dirty
14            t1 = ElapsedTime()
15            TargetRange.Calculate
16            t2 = ElapsedTime()
17            t = t + t2 - t1
18        Next i
19        LogMessage BaseMessage

20        PasteAndTimeFormula = t / NumCalls

21        If OneCallPerCell Then
22            PasteAndTimeFormula = PasteAndTimeFormula / TargetRange.Cells.CountLarge
23        End If
24        With TargetRange.Cells(0, 1)
25            .Value = "'" & Formula
26            .Columns.AutoFit
27        End With

28        TargetRange.Cells(-1, 1).Value = PasteAndTimeFormula

29        Exit Function
ErrHandler:
30        ReThrow "PasteAndTimeFormula", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FindTargetRange
' Author     : Philip Swannell
' Date       : 20-Aug-2026
' Purpose    : Returns a range on the worksheet ws which is available to paste a formula who's return has 1 column and
'              VectorLength rows.
' -----------------------------------------------------------------------------------------------------------------------
Function FindTargetRange(ws As Worksheet, VectorLength As Long) As Range

          Dim Out As Range
          Dim i As Long

1         On Error GoTo ErrHandler
2         Set Out = ws.Cells(3, 16384).End(xlToLeft)
3         While Not IsEmpty(Out.Value)
4             Set Out = Out.Offset(, 1)
5         Wend
6         Set Out = Out.Resize(VectorLength)
          'For safety, Out should be blank already!
7         While Not IsRangeBlank(Out)
8             Out = Out.Offset(, 1)
9             i = i + 1
10            If i > 1000 Then Throw "Unexpected error in method FindTargetRange"
11        Wend

12        Set FindTargetRange = Out

13        Exit Function
ErrHandler:
14        ReThrow "FindTargetRange", Err
End Function

Sub TestInstallation()
1         On Error GoTo ErrHandler
2         With shMain.Range("JuliaExcelTestCell")
3             .Dirty
4             .Calculate
5             If VarType(.Value) <> vbDouble Then Throw CStr(.Value)
6         End With

7         With shMain.Range("JuliaInXLTestCell")
8             .Dirty
9             .Calculate
10            If VarType(.Value) <> vbDouble Then Throw CStr(.Value)
11        End With

12        Exit Sub
ErrHandler:
13        ReThrow "TestInstallation", Err
End Sub

Sub LogMessage(Message As Variant)
1     Application.StatusBar = Message
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ClearOutSheet
' Author     : Philip Swannell
' Date       : 20-Aug-2026
' Purpose    : Delete all names and contents from a worksheet
' -----------------------------------------------------------------------------------------------------------------------
Sub ClearOutSheet(ws As Worksheet)
          Dim n As Name
          Dim UR As Range

1         On Error GoTo ErrHandler

2         If ws.ProtectContents Then ws.Unprotect

3         For Each n In ws.Names
4             n.Delete
5         Next

6         ws.UsedRange.EntireColumn.Delete

7         Set UR = ws.UsedRange

8         Exit Sub
ErrHandler:
9         ReThrow "ClearOutSheet", Err
End Sub

Sub ClearOutSheets()
1         On Error GoTo ErrHandler

2         ClearOutSheet shInputData
3         ClearOutSheet shJuliaExcel
4         ClearOutSheet shJuliaInXL
5         shMain.Activate

6         Exit Sub
ErrHandler:
7         ReThrow "ClearOutSheets", Err
End Sub

Sub AddAllResultsRangeNameToSheet(ws As Worksheet)
          Dim AllResults As Range
          Dim TopLeft As Range
          Dim BottomRight As Range

1         On Error GoTo ErrHandler
2         Set TopLeft = ws.Cells(3, 1)
3         With ws.UsedRange
4             Set BottomRight = .Cells(.Rows.Count, .Columns.Count)
5         End With
6         Set AllResults = Range(TopLeft, BottomRight)
7         ws.Names.Add "AllResults", AllResults
8         Exit Sub
ErrHandler:
9         ReThrow "AddAllResultsRangeNameToSheet", Err
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : AlignColumnWidths
' Author     : Philip Swannell
' Date       : 20-Aug-2026
' Purpose    : Make the column widths the same so that when flipping between the two sheets via Ctrl+PageUp/Ctrl+PageDown
'              one sees viual confirmation of identical results
' -----------------------------------------------------------------------------------------------------------------------
Sub AlignColumnWidths()
          Dim JEColWidth As Double
          Dim JInXLColWidth As Double
          Dim i As Long

1         On Error GoTo ErrHandler
2         For i = 1 To shJuliaExcel.UsedRange.Columns.Count
3             JEColWidth = shJuliaExcel.Cells(1, i).ColumnWidth
4             JInXLColWidth = shJuliaInXL.Cells(1, i).ColumnWidth
5             If JEColWidth > JInXLColWidth Then
6                 shJuliaInXL.Cells(1, i).ColumnWidth = JEColWidth
7             Else
8                 shJuliaExcel.Cells(1, i).ColumnWidth = JInXLColWidth
9             End If
10        Next i

11        Exit Sub
ErrHandler:
12        ReThrow "AlignColumnWidths", Err
End Sub

