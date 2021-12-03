Attribute VB_Name = "modDocumenter"
Option Explicit

Sub testGFIS()
Dim Description As String, ArgNames(), ArgDescs(), AllFunctions(), rngFunctionsAndDescriptions As Range

    On Error GoTo ErrHandler
    GrabFromIntelliSenseSheet "JuliaResultFile", Description, ArgNames, ArgDescs, AllFunctions, rngFunctionsAndDescriptions



    Exit Sub
ErrHandler:
    MsgBox "#testGFIS (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub



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
          ByRef ArgNames As Variant, ByRef ArgDescs As Variant, ByRef AllFunctions() As Variant, _
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
            
15        Set rngFunctionsAndDescriptions = SourceRange.Resize(, 2)

16        If FunctionName = "" Then
17            Exit Sub
18        End If

19        For Each c In SourceRange.Cells
20            If c.Value = FunctionName Then
21                Set SourceCell = c
22                Exit For
23            End If
24        Next c
25        If SourceCell Is Nothing Then Throw "Cannot find function '" + FunctionName + "' listed on sheet '" + SourceSheetName + "' of workbook '" & SourceBookName + "'"

26        Description = c.Offset(0, 1).Value
27        If IsEmpty(c.Offset(0, 3)) Then
28            Set rngArgsAndDescs = c.Offset(0, 3).Resize(, 2)

29        Else


30            Set rngArgsAndDescs = Range(c.Offset(0, 3), c.Offset(0, 3).End(xlToRight))
31    End If
32            N = rngArgsAndDescs.Cells.Count
33            ReDim ArgNames(1 To N / 2)
34            ReDim ArgDescs(1 To N / 2)
35            For i = 1 To N / 2
36                ArgNames(i) = rngArgsAndDescs.Cells(1, 2 * i - 1).Value
37                ArgDescs(i) = rngArgsAndDescs.Cells(1, 2 * i).Value
38            Next i

40        Exit Sub
ErrHandler:
41        Throw "#GrabFromIntelliSenseSheet (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Function MarkdownForSummaryOfFunctions()

          Dim AllFunctions() As Variant
          Dim ArgDescs()
          Dim ArgNames()
          Dim i As Long
          Dim j As Long
          Dim rngFunctionsAndDescriptions As Range
          Dim STK As clsStacker
          Dim ThisHelp As Variant
          Dim Table As Variant

1         On Error GoTo ErrHandler
2         GrabFromIntelliSenseSheet "", "", ArgNames, ArgDescs, AllFunctions, rngFunctionsAndDescriptions

3         Table = rngFunctionsAndDescriptions.Value
          
4         For i = 1 To UBound(Table, 1)
5             For j = 1 To UBound(AllFunctions)
6                 Table(i, 2) = sRegExReplace(Table(i, 2), "\b" & AllFunctions(j) & "\b", "`" & AllFunctions(j) & "`")
7             Next j
8         Next i

9         MarkdownForSummaryOfFunctions = AddPipes(Table)

10        Exit Function
ErrHandler:
11        MarkdownForSummaryOfFunctions = "#MarkdownForSummaryOfFunctions (line " & CStr(Erl) + "): " & Err.Description & "!"
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
11        NumArgs = UBound(ArgNames) - LBound(ArgNames) + 1
12        If NumArgs = 1 Then
13            If ArgNames(1) = "" Then
14                NumArgs = 0
15            End If
16        End If


17        If NumArgs > 0 Then

18            Hlp = Hlp & "' Arguments" & vbLf


19            For i = 1 To NumArgs
20                Hlp = Hlp & "' " & ArgNames(i)
21                If Len(ArgNames(i)) < 10 Then
22                    Spacers = String(10 - Len(ArgNames(i)), " ")
23                Else
24                    Spacers = ""
25                End If
              
26                Hlp = Hlp & Spacers
27                Hlp = Hlp & ":" & InsertBreaks(ArgDescriptions(i), Len(ArgNames(i)) + Len(Spacers) + 2) + vbLf
28            Next
29        End If
30        If Len(ExtraHelp) > 0 Then
31            Do While (Left$(ExtraHelp, 1) = vbLf Or Left$(ExtraHelp, 1) = vbCr)
32                ExtraHelp = Right$(ExtraHelp, Len(ExtraHelp) - 1)
33            Loop
34            Hlp = Hlp & ("'" & vbLf)
35            Hlp = Hlp & "' Notes     :"
36            Hlp = Hlp & InsertBreaks(ExtraHelp)
37            Hlp = Hlp & vbLf
38        End If
39        Hlp = Hlp & "' " & String(119, "-")
40        HelpForVBE = Application.WorksheetFunction.Transpose(VBA.Split(Hlp, vbLf))
41        Exit Function
ErrHandler:
42        HelpForVBE = "#HelpForVBE (line " & CStr(Erl) + "): " & Err.Description & "!"
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
          ByVal ArgNames, ByVal ArgDescriptions, Replacements, AllFunctions() As Variant)

          Dim Declaration As String
          Dim Hlp As String
          Dim i As Long
          Dim j As Long
          Dim LeftString As String
          Dim RightString As String
          Dim SourceCode As String
          Dim StringsToEncloseInBackTicks
          Dim ThisArgDescription As String
          Dim NumArgs As Long

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
13        If Right(Declaration, Len(FunctionName) + 2) <> FunctionName + "()" Then
14            If Mid(Declaration, Len(Declaration) - 1) = "()" Then
15                Declaration = StringBetweenStrings(SourceCode, Declaration, ")", True, True)
16            End If
17        End If

          'Bodge - get the "As VarType"
18        MatchPoint = InStr(SourceCode, Declaration)
19        NextChars = Mid$(SourceCode, MatchPoint + Len(Declaration), 100)
20        If Left$(NextChars, 4) = " As " Then
21            NextChars = StringBetweenStrings(NextChars, " As ", vbLf, True, False)
22            NextChars = " " & Trim(NextChars)
23            Declaration = Declaration & NextChars
24        End If

25        Hlp = "### `" & FunctionName & "`" & vbLf

26        NumArgs = UBound(ArgNames)
27        If NumArgs = 1 Then
28            If ArgNames(1) = "" Then
29                NumArgs = 0
30            End If
31        End If

32        If NumArgs > 0 Then
33            StringsToEncloseInBackTicks = VBA.Split(VBA.Join(ArgNames, ",") & "," & VBA.Join(AllFunctions, ","), ",")
34        Else
35            StringsToEncloseInBackTicks = AllFunctions
36        End If

37        For j = LBound(StringsToEncloseInBackTicks) To UBound(StringsToEncloseInBackTicks)
38            FunctionDescription = sRegExReplace(FunctionDescription, "\b" & StringsToEncloseInBackTicks(j) & "\b", "`" & StringsToEncloseInBackTicks(j) & "`", True)
39        Next j

40        Hlp = Hlp & FunctionDescription & vbLf

42            Hlp = Hlp & "```vba" & vbLf & _
                  Declaration & vbLf & _
                  "```"
                  
41        If NumArgs > 0 Then
                  
                  Hlp = Hlp & vbLf & vbLf & "|Argument|Description|" & vbLf & _
                  "|:-------|:----------|"
          
43            For i = LBound(ArgNames) To UBound(ArgNames)
44                ThisArgDescription = ArgDescriptions(i)
45                For j = LBound(StringsToEncloseInBackTicks) To UBound(StringsToEncloseInBackTicks)
46                    ThisArgDescription = sRegExReplace(ThisArgDescription, "\b" & StringsToEncloseInBackTicks(j) & "\b", "`" & StringsToEncloseInBackTicks(j) & "`", True)
47                Next j
48                ThisArgDescription = Replace(ThisArgDescription, vbCrLf, vbLf)
49                ThisArgDescription = Replace(ThisArgDescription, vbCr, vbLf)
50                ThisArgDescription = Replace(ThisArgDescription, vbLf, "<br/>")
51                Hlp = Hlp & vbLf & "|`" & ArgNames(i) & "`|" & ThisArgDescription & "|"
52            Next i
53        End If
54        If Not IsMissing(Replacements) Then
55            If TypeName(Replacements) = "Range" Then
56                Replacements = Replacements.Value
57            End If
58            For i = 1 To sNRows(Replacements)
59                Hlp = Replace(Hlp, Replacements(i, 1), Replacements(i, 2))
60            Next i
61        End If

62        MarkdownHelp = Application.Transpose(VBA.Split(Hlp, vbLf))

63        Exit Function
ErrHandler:
64        MarkdownHelp = "#MarkdownHelp (line " & CStr(Erl) + "): " & Err.Description & "!"
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

