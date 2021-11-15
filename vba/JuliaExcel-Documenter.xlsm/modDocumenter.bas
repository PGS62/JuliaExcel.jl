Attribute VB_Name = "modDocumenter"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GrabFromIntelliSenseSheet
' Purpose    : The methods in this module all get their data from the _IntelliSense_ worksheet of JuliaExcel.xlam
'              and this method is a shared data grabber for them.
' Parameters :
'  FunctionName               :
'  Description                :
'  ArgNames                   :
'  ArgDescs                   :
'  AllFunctions               :
'  rngFunctionsAndDescriptions:
' -----------------------------------------------------------------------------------------------------------------------
Private Sub GrabFromIntelliSenseSheet(FunctionName As String, ByRef Description As String, _
          ByRef ArgNames() As Variant, ByRef ArgDescs() As Variant, ByRef AllFunctions() As Variant, _
          Optional ByRef rngFunctionsAndDescriptions As Range)

          Dim c As Range
          Dim FunctionNameCell As Range
          Dim i As Long
          Dim N As Long
          Dim TheTable As Range
          Dim wb As Workbook

          Const SourceBookName = "JuliaExcel.xlam"
          Const SourceSheetName = "_IntelliSense_"
          Dim rngArgsAndDescs As Range
          Dim SourceCell As Range
          Dim SourceRange As Range

1         On Error GoTo ErrHandler

2         If Not IsInCollection(Application.Workbooks, SourceBookName) Then
3             Throw "Workbook '" & SourceBookName & "' must be open, but it's not"
4         End If
5         If Not IsInCollection(Application.Workbooks(SourceBookName).Worksheets, SourceSheetName) Then
6             Throw "Workbook '" & SourceBookName & "' must have a sheet '" + SourceSheetName & "' but it does not"
7         End If

8         With Application.Workbooks(SourceBookName).Worksheets(SourceSheetName).Cells(1, 1)
9             Set SourceRange = Range(.Offset(1), .End(xlDown))
10        End With

11        ReDim AllFunctions(1 To SourceRange.Cells.Count)
12        For i = 1 To SourceRange.Cells.Count
13            AllFunctions(i) = SourceRange.Cells(i, 1).Value
14        Next i
            
          Set rngFunctionsAndDescriptions = SourceRange.Resize(, 2)

15        If FunctionName = "" Then
16            Exit Sub
17        End If

18        For Each c In SourceRange.Cells
19            If c.Value = FunctionName Then
20                Set SourceCell = c
21                Exit For
22            End If
23        Next c
24        If SourceCell Is Nothing Then Throw "Cannot find function '" + FunctionName + "' listed on sheet '" + SourceSheetName + "' of workbook '" & SourceBookName + "'"

25        Description = c.Offset(0, 1).Value

26        Set rngArgsAndDescs = Range(c.Offset(0, 3), c.Offset(0, 3).End(xlToRight))
27        N = rngArgsAndDescs.Cells.Count
28        ReDim ArgNames(1 To N / 2)
29        ReDim ArgDescs(1 To N / 2)
30        For i = 1 To N / 2
31            ArgNames(i) = rngArgsAndDescs.Cells(1, 2 * i - 1).Value
32            ArgDescs(i) = rngArgsAndDescs.Cells(1, 2 * i).Value
33        Next i

34        Exit Sub
ErrHandler:
35        Throw "#GrabFromIntelliSenseSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Function MarkdownForSummaryOfFunctions()

          Dim AllFunctions() As Variant
          Dim ArgDescs()
          Dim ArgNames()
          Dim i As Long
          Dim rngFunctionsAndDescriptions As Range
          Dim STK As clsStacker
          Dim ThisHelp As Variant

1         On Error GoTo ErrHandler
2         On Error GoTo ErrHandler
3         GrabFromIntelliSenseSheet "", "", ArgNames, ArgDescs, AllFunctions, rngFunctionsAndDescriptions

4         MarkdownForSummaryOfFunctions = AddPipes(rngFunctionsAndDescriptions.Value)

5         Exit Function
ErrHandler:
6         MarkdownForSummaryOfFunctions = "#MarkdownForSummaryOfFunctions (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function AddPipes(Data)
          Dim i As Long
          Dim j As Long
          Dim NC As Long
          Dim NR As Long
          Dim Result() As String
1         If TypeName(Data) = "Range" Then Data = Data.Value

2         NR = sNRows(Data)
3         NC = sNCols(Data)

4         ReDim Result(1 To NR + 2, 1 To 1)
5         Result(1, 1) = "|Name|Description|"
6         Result(2, 1) = "|----|-----------|"
7         For i = 1 To NR
8             Result(i + 2, 1) = "|"
9             For j = 1 To NC
10                If j = 1 Then
11                    Result(i + 2, 1) = Result(i + 2, 1) + "[" + Data(i, j) + "](#" + LCase(Data(i, j)) + ")|"
12                Else
13                    Result(i + 2, 1) = Result(i + 2, 1) + Data(i, j) + "|"
14                End If
15            Next j
16        Next i
17        AddPipes = Result

End Function

Function HelpForVBEAll(SourceFile As String)
          Dim AllFunctions() As Variant
          Dim ArgDescs()
          Dim ArgNames()
          Dim i As Long
          Dim STK As clsStacker
          Dim ThisHelp As Variant

1         On Error GoTo ErrHandler
2         GrabFromIntelliSenseSheet "", "", ArgNames, ArgDescs, AllFunctions

3         Set STK = CreateStacker()

4         For i = LBound(AllFunctions) To UBound(AllFunctions)
5             ThisHelp = HelpForVBE2(CStr(AllFunctions(i)))
6             STK.StackData ThisHelp
7             STK.StackData ""
8             STK.StackData ""
9         Next i

10        HelpForVBEAll = STK.Report

11        Exit Function
ErrHandler:
12        HelpForVBEAll = "#HelpForVBEAll (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function HelpForVBE2(FunctionName As String, Optional ExtraHelp As String, Optional Author As String, Optional DateWritten As Long)

          Dim AllFunctions() As Variant
          Dim ArgDescs() As Variant
          Dim ArgNames() As Variant
          Dim Description As String

1         On Error GoTo ErrHandler
2         GrabFromIntelliSenseSheet FunctionName, Description, ArgNames, ArgDescs, AllFunctions

3         HelpForVBE2 = HelpForVBE(FunctionName, Description, ArgNames, ArgDescs, ExtraHelp, Author, DateWritten)
          
4         Exit Function
ErrHandler:
5         HelpForVBE2 = "#HelpForVBE2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : HelpForVBE
' Purpose    : Generate a header to paste into the VBE. The header generated will be consistent with the registration
'              created by calling CodeToRegister.
' -----------------------------------------------------------------------------------------------------------------------
Function HelpForVBE(FunctionName As String, FunctionDescription As String, ArgNames, ArgDescriptions, _
          Optional ExtraHelp As String, Optional Author As String, Optional DateWritten As Long)

          Dim Hlp As String
          Dim i As Long
          Dim NumArgs As Long
          Dim Spacers As String
          
1         On Error GoTo ErrHandler
          
2         Hlp = Hlp & "' " & String(119, "-") & vbLf
3         Hlp = Hlp & "' Procedure : " & FunctionName & vbLf
4         If Len(Author) > 0 Then
5             Hlp = Hlp & "' Author    : " & Author & "" & vbLf
6         End If
7         If DateWritten <> 0 Then
8             Hlp = Hlp & "' Date      : " & Format$(DateWritten, "dd-mmm-yyyy") & vbLf
9         End If

10        Hlp = Hlp & "' Purpose   :" & InsertBreaks(FunctionDescription, Len("Len(ArgNames(i))")) & vbLf
11        Hlp = Hlp & "' Arguments" & vbLf

12        NumArgs = UBound(ArgNames) - LBound(ArgNames) + 1
13        For i = 1 To NumArgs
14            Hlp = Hlp & "' " & ArgNames(i)
15            If Len(ArgNames(i)) < 10 Then
16                Spacers = String(10 - Len(ArgNames(i)), " ")
17            Else
18                Spacers = ""
19            End If
              
20            Hlp = Hlp & Spacers
21            Hlp = Hlp & ":" & InsertBreaks(ArgDescriptions(i), Len(ArgNames(i)) + Len(Spacers) + 2) + vbLf
22        Next
23        If Len(ExtraHelp) > 0 Then
24            Do While (Left$(ExtraHelp, 1) = vbLf Or Left$(ExtraHelp, 1) = vbCr)
25                ExtraHelp = Right$(ExtraHelp, Len(ExtraHelp) - 1)
26            Loop
27            Hlp = Hlp & ("'" & vbLf)
28            Hlp = Hlp & "' Notes     :"
29            Hlp = Hlp & InsertBreaks(ExtraHelp)
30            Hlp = Hlp & vbLf
31        End If
32        Hlp = Hlp & "' " & String(119, "-")
33        HelpForVBE = Application.WorksheetFunction.Transpose(VBA.Split(Hlp, vbLf))
34        Exit Function
ErrHandler:
35        HelpForVBE = "#HelpForVBE (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function MarkdownHelpAll(SourceFile As String, Optional Replacements)
          Dim AllFunctions() As Variant
          Dim ArgDescs()
          Dim ArgNames()
          Dim i As Long
          Dim STK As clsStacker
          Dim ThisHelp As Variant

1         On Error GoTo ErrHandler
2         GrabFromIntelliSenseSheet "", "", ArgNames, ArgDescs, AllFunctions

3         Set STK = CreateStacker()

4         For i = LBound(AllFunctions) To UBound(AllFunctions)
5             ThisHelp = MarkdownHelp2(SourceFile, CStr(AllFunctions(i)), Replacements)
6             STK.StackData ThisHelp
7             STK.StackData ""
8             STK.StackData ""
9         Next i

10        MarkdownHelpAll = STK.Report

11        Exit Function
ErrHandler:
12        MarkdownHelpAll = "#MarkdownHelpAll (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function MarkdownHelp2(SourceFile As String, FunctionName As String, Optional Replacements)

          Dim AllFunctions() As Variant
          Dim ArgDescs() As Variant
          Dim ArgNames() As Variant
          Dim Description As String

1         On Error GoTo ErrHandler
2         GrabFromIntelliSenseSheet FunctionName, Description, ArgNames, ArgDescs, AllFunctions

3         MarkdownHelp2 = MarkdownHelp(SourceFile, FunctionName, Description, ArgNames, ArgDescs, Replacements, AllFunctions)
          
4         Exit Function
ErrHandler:
5         MarkdownHelp2 = "#MarkdownHelp2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MarkdownHelp
' Purpose    : Formats the help as a markdown table.
' -----------------------------------------------------------------------------------------------------------------------
Function MarkdownHelp(SourceFile As String, FunctionName As String, ByVal FunctionDescription As String, _
          ArgNames, ArgDescriptions, Replacements, AllFunctions() As Variant)

          Dim Declaration As String
          Dim Hlp As String
          Dim i As Long
          Dim j As Long
          Dim LeftString As String
          Dim RightString As String
          Dim SourceCode As String
          Dim StringsToEncloseInBackTicks
          Dim ThisArgDescription As String

1         On Error GoTo ErrHandler

2         SourceCode = RawFileContents(SourceFile)
3         SourceCode = Replace(SourceCode, vbCrLf, vbLf)

4         LeftString = "Public Function " & FunctionName & "("

5         If InStr(SourceCode, LeftString) = 0 Then
6             LeftString = "Private Function " & FunctionName & "("
7             If InStr(SourceCode, LeftString) = 0 Then
8                 LeftString = "Function " & FunctionName & "("
9                 If InStr(SourceCode, LeftString) = 0 Then Throw "Cannot find function declaration in SourceFile"
10            End If
11        End If

12        Declaration = StringBetweenStrings(SourceCode, LeftString, ")", True, True)

          Dim MatchPoint As Long
          Dim NextChars As String

          'Bodge ParamArray() confuses my cheap and chearful language parsing
13        If Mid(Declaration, Len(Declaration) - 1) = "()" Then
14            Declaration = StringBetweenStrings(SourceCode, Declaration, ")", True, True)
15        End If

          'Bodge - get the "As VarType"
16        MatchPoint = InStr(SourceCode, Declaration)
17        NextChars = Mid$(SourceCode, MatchPoint + Len(Declaration), 100)
18        If Left$(NextChars, 4) = " As " Then
19            NextChars = StringBetweenStrings(NextChars, " As ", vbLf, True, False)
20            NextChars = " " & Trim(NextChars)
21            Declaration = Declaration & NextChars
22        End If

23        Hlp = "#### _" & FunctionName & "_" & vbLf

24        StringsToEncloseInBackTicks = VBA.Split(VBA.Join(ArgNames, ",") & "," & VBA.Join(AllFunctions, ","), ",")

25        For j = LBound(StringsToEncloseInBackTicks) To UBound(StringsToEncloseInBackTicks)
26            FunctionDescription = sRegExReplace(FunctionDescription, "\b" & StringsToEncloseInBackTicks(j) & "\b", "`" & StringsToEncloseInBackTicks(j) & "`", True)
27        Next j

28        Hlp = Hlp & FunctionDescription & vbLf
          
29        Hlp = Hlp & "```vba" & vbLf & _
              Declaration & vbLf & _
              "```" & vbLf & vbLf & _
              "|Argument|Description|" & vbLf & _
              "|:-------|:----------|"
          
30        For i = LBound(ArgNames) To UBound(ArgNames)
31            ThisArgDescription = ArgDescriptions(i)
32            For j = LBound(StringsToEncloseInBackTicks) To UBound(StringsToEncloseInBackTicks)
33                ThisArgDescription = sRegExReplace(ThisArgDescription, "\b" & StringsToEncloseInBackTicks(j) & "\b", "`" & StringsToEncloseInBackTicks(j) & "`", True)
34            Next j
35            ThisArgDescription = Replace(ThisArgDescription, vbCrLf, vbLf)
36            ThisArgDescription = Replace(ThisArgDescription, vbCr, vbLf)
37            ThisArgDescription = Replace(ThisArgDescription, vbLf, "<br/>")
38            Hlp = Hlp & vbLf & "|`" & ArgNames(i) & "`|" & ThisArgDescription & "|"
39        Next i

40        If Not IsMissing(Replacements) Then
41            If TypeName(Replacements) = "Range" Then
42                Replacements = Replacements.Value
43            End If
44            For i = 1 To sNRows(Replacements)
45                Hlp = Replace(Hlp, Replacements(i, 1), Replacements(i, 2))
46            Next i
47        End If

48        MarkdownHelp = Application.Transpose(VBA.Split(Hlp, vbLf))

49        Exit Function
ErrHandler:
50        MarkdownHelp = "#MarkdownHelp (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Private Function InsertBreaks(ByVal TheString As String, Optional FirstRowShorterBy As Long)

          Const FirstTab = 0
          Const NextTabs = 13
          Const Width = 106
          Dim DoNewLine As Boolean
          Dim i As Long
          Dim LineLength As Long
          Dim Res As String
          Dim Words
          Dim WordsNLB
          
1         On Error GoTo ErrHandler
          
2         If InStr(TheString, " ") = 0 Then
3             InsertBreaks = TheString
4             Exit Function
5         End If
          
6         TheString = Replace(TheString, vbLf, vbLf + " ")
7         TheString = Replace(TheString, vbLf + "  ", vbLf + " ")
          
8         Res = String(FirstTab, " ")
9         LineLength = FirstTab + FirstRowShorterBy

10        Words = VBA.Split(TheString, " ")
11        WordsNLB = Words
12        For i = LBound(Words) To UBound(Words)
13            WordsNLB(i) = Replace(WordsNLB(i), vbLf, vbNullString)
14        Next

15        For i = LBound(Words) To UBound(Words)
16            DoNewLine = LineLength + Len(WordsNLB(i)) > Width
17            If i > 1 Then
18                If InStr(Words(i - 1), vbLf) > 0 Then
19                    DoNewLine = True
20                End If
21            End If

22            If DoNewLine Then
23                Res = Res + vbLf + "'" + String(NextTabs, " ") + WordsNLB(i)
24                LineLength = 1 + NextTabs + Len(WordsNLB(i))
25            Else
26                Res = Res + " " + WordsNLB(i)
27                LineLength = LineLength + 1 + Len(WordsNLB(i))
28            End If
29        Next
30        InsertBreaks = Res

31        Exit Function
ErrHandler:
32        Throw "#InsertBreaks: " & Err.Description & "!"
End Function

