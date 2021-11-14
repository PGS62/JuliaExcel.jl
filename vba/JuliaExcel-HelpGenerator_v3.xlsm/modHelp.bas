Attribute VB_Name = "modHelp"
Option Explicit

Function AddPipes(Data)
          Dim Result() As String
          Dim i As Long, j As Long, NR As Long, NC As Long
1         If TypeName(Data) = "Range" Then Data = Data.Value

2         NR = sNRows(Data)
3         NC = sNCols(Data)

4         ReDim Result(1 To NR, 1 To 1)

5         For i = 1 To NR
6             Result(i, 1) = "|"
7             For j = 1 To NC
8                 If j = 1 Then
9                     Result(i, 1) = Result(i, 1) + "[" + Data(i, j) + "](#" + LCase(Data(i, j)) + ")|"
10                Else
11                    Result(i, 1) = Result(i, 1) + Data(i, j) + "|"
12                End If
13            Next j
14        Next i

15        AddPipes = Result


End Function


Sub GrabArgsFromTable(FunctionName As String, IntellisenseTable As Range, ByRef Description As String, ByRef ArgNames() As Variant, ByRef ArgDescs() As Variant)

          Dim TheTable As Range
          Dim FunctionNameCell As Range
          Dim c As Range
          Dim i As Long

1         On Error GoTo ErrHandler
2         Set TheTable = IntellisenseTable.CurrentRegion

3         For Each c In TheTable.Columns(1).Cells
4             If c.Value = FunctionName Then
5                 Set FunctionNameCell = c
6                 Description = c.Offset(0, 1).Value
7                 Exit For
8             End If
9         Next

10        ReDim ArgDescs(1 To 1, 1 To 1)
          
11        ArgDescs(1, 1) = FunctionNameCell.Offset(0, 4).Value
          
12        For i = 6 To 100 Step 2
13            If Not IsEmpty(FunctionNameCell.Offset(0, i).Value) Then
14                ReDim Preserve ArgDescs(1 To 1, 1 To UBound(ArgDescs, 2) + 1)
15                ArgDescs(1, UBound(ArgDescs, 2)) = FunctionNameCell.Offset(0, i).Value
16            Else
17                Exit For
18            End If
19        Next i
20        If UBound(ArgDescs, 2) > 1 Then
21            ArgDescs = Application.WorksheetFunction.Transpose(ArgDescs)
22        End If




23        ReDim ArgNames(1 To 1, 1 To 1)
          
24        ArgNames(1, 1) = FunctionNameCell.Offset(0, 3).Value
          
25        For i = 5 To 101 Step 2
26            If Not IsEmpty(FunctionNameCell.Offset(0, i).Value) Then
27                ReDim Preserve ArgNames(1 To 1, 1 To UBound(ArgNames, 2) + 1)
28                ArgNames(1, UBound(ArgNames, 2)) = FunctionNameCell.Offset(0, i).Value
29            Else
30                Exit For
31            End If
32        Next i
33        If UBound(ArgNames, 2) > 1 Then
34            ArgNames = Application.WorksheetFunction.Transpose(ArgNames)
35        End If




36        Exit Sub
ErrHandler:
37        Throw "#GrabArgsFromTable (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

Function CodeToRegister2(FunctionName As String, IntellisenseTable As Range)
Dim Description As String
Dim ArgNames() As Variant
Dim ArgDescs() As Variant

    On Error GoTo ErrHandler
    GrabArgsFromTable FunctionName, IntellisenseTable, Description, ArgNames, ArgDescs

CodeToRegister2 = CodeToRegister(FunctionName, Description, ArgDescs)
    
    Exit Function
ErrHandler:
    CodeToRegister2 = "#CodeToRegister2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function HelpForVBE2(FunctionName As String, IntellisenseTable As Range, Optional ExtraHelp As String, Optional Author As String, Optional DateWritten As Long)

          Dim Description As String
          Dim ArgNames() As Variant
          Dim ArgDescs() As Variant

1         On Error GoTo ErrHandler
2         GrabArgsFromTable FunctionName, IntellisenseTable, Description, ArgNames, ArgDescs

3         HelpForVBE2 = HelpForVBE(FunctionName, Description, ArgNames, ArgDescs, ExtraHelp, Author, DateWritten)
          
4         Exit Function
ErrHandler:
5         HelpForVBE2 = "#HelpForVBE2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function MarkdownHelp2(SourceFile As String, FunctionName As String, IntellisenseTable As Range, Optional Replacements)

          Dim Description As String
          Dim ArgNames() As Variant
          Dim ArgDescs() As Variant

1         On Error GoTo ErrHandler
2         GrabArgsFromTable FunctionName, IntellisenseTable, Description, ArgNames, ArgDescs

3         MarkdownHelp2 = MarkdownHelp(SourceFile, FunctionName, Description, ArgNames, ArgDescs, Replacements)
          
4         Exit Function
ErrHandler:
5         MarkdownHelp2 = "#MarkdownHelp2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function





' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CodeToRegister
' Purpose    : Generate VBA code to register a function.
' -----------------------------------------------------------------------------------------------------------------------
Function CodeToRegister(FunctionName, Description As String, ArgDescs)

          Const DQ = """"
          Dim code As String
          Dim i As Long
          
1         On Error GoTo ErrHandler
2         If TypeName(ArgDescs) = "Range" Then ArgDescs = ArgDescs.Value
3         If VarType(ArgDescs) < vbArray Then
4             ReDim Temp(1 To 1, 1 To 1): Temp(1, 1) = ArgDescs: ArgDescs = Temp
5         End If

6         If Len(Description) > 255 Then Throw "Description " + CStr(i) + " is of length " + CStr(Len(Description)) + " but must be of length 255 or less."
          
7         code = code & "' " & String(119, "-") & vbLf
8         code = code & "' Procedure  : Register" & FunctionName & vbLf
9         code = code & "' Purpose    : Register the function " & FunctionName & " with the Excel function wizard, to be called from the WorkBook_Open" & vbLf
10        code = code & "'              event." & vbLf
11        code = code & "' " & String(119, "-") & vbLf

12        code = code & "Private Sub Register" + FunctionName + "()" + vbLf
13        code = code + "    Const Description As String = " + InsertBreaksInStringLiteral(DQ + Replace(Description, DQ, DQ + DQ) + DQ, 34) + vbLf
14        code = code + "    Dim " + "ArgDescs() As String" + vbLf + vbLf
15        code = code + "    On Error GoTo ErrHandler" + vbLf + vbLf
          
16        code = code + "    ReDim " + "ArgDescs(" + CStr(LBound(ArgDescs, 1)) + " To " & CStr(UBound(ArgDescs, 1)) + ")" + vbLf

17        For i = LBound(ArgDescs, 1) To UBound(ArgDescs, 1)
18            If Len(ArgDescs(i, 1)) > 255 Then Throw "ArgDescs element " + CStr(i) + " is of length " + CStr(Len(ArgDescs(i, 1))) + " but must be of length 255 or less."
19            code = code + "    " + "ArgDescs(" & CStr(i) & ") = " & InsertBreaksInStringLiteral(DQ + Replace(ArgDescs(i, 1), DQ, DQ + DQ) + DQ, IIf(i < 10, 18, 19)) + vbLf
20        Next i

21        code = code + "    Application.MacroOptions """ + FunctionName + """, Description, , , , , , , , , ArgDescs" + vbLf
          
22        code = code + "    Exit Sub" & vbLf + vbLf
          
23        code = code + "ErrHandler:" + vbLf
24        code = code + "    Debug.Print ""Warning: Registration of function " + FunctionName + " failed with error: "" + Err.Description" + vbLf
25        code = code + "End Sub"

26        CodeToRegister = Application.WorksheetFunction.Transpose(VBA.Split(code, vbLf))


27        Exit Function
ErrHandler:
28        CodeToRegister = "#CodeToRegister (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function InsertBreaksInStringLiteral(ByVal TheString As String, Optional FirstRowShorterBy As Long)

    Const FirstTab = 0
    Dim NextTabs As Long
    Const Width = 114
    Dim DoNewLine As Boolean
    Dim i As Long
    Dim LineLength As Long
    Dim Res As String
    Dim Words
    Dim WordsNLB
    
    On Error GoTo ErrHandler
    
    NextTabs = FirstRowShorterBy
    
    If InStr(TheString, " ") = 0 Then
        InsertBreaksInStringLiteral = TheString
        Exit Function
    End If
    
    Res = String(FirstTab, " ")
    LineLength = FirstTab + FirstRowShorterBy

    Words = VBA.Split(TheString, " ")
    WordsNLB = Words

    For i = LBound(Words) To UBound(Words)
        DoNewLine = LineLength + Len(WordsNLB(i)) > Width

        If DoNewLine Then
            Res = Res + " "" & _" + vbLf + String(NextTabs, " ") + """" + WordsNLB(i)
            LineLength = 1 + NextTabs + Len(WordsNLB(i))
        Else
            Res = Res + " " + WordsNLB(i)
            LineLength = LineLength + 1 + Len(WordsNLB(i))
        End If
    Next
    InsertBreaksInStringLiteral = Trim(Res)

    Exit Function
ErrHandler:
    Throw "#InsertBreaksInStringLiteral: " & Err.Description & "!"
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
          
1             On Error GoTo ErrHandler

2         If VarType(ArgNames) < vbArray Then
3             ReDim Temp(1 To 1, 1 To 1): Temp(1, 1) = ArgNames: ArgNames = Temp
4         End If

          
5         Hlp = Hlp & "' " & String(119, "-") & vbLf
6         Hlp = Hlp & "' Procedure : " & FunctionName & vbLf
7         If Len(Author) > 0 Then
8             Hlp = Hlp & "' Author    : " & Author & "" & vbLf
9         End If
10        If DateWritten <> 0 Then
11            Hlp = Hlp & "' Date      : " & Format$(DateWritten, "dd-mmm-yyyy") & vbLf
12        End If

13        Hlp = Hlp & "' Purpose   :" & InsertBreaks(FunctionDescription, Len("Len(ArgNames(i, 1))")) & vbLf
14        Hlp = Hlp & "' Arguments" & vbLf

15        If TypeName(ArgNames) = "Range" Then ArgNames = ArgNames.Value

16        NumArgs = UBound(ArgNames, 1) - LBound(ArgNames, 1) + 1
17        For i = 1 To NumArgs
              '    If InStr(ArgNames(i, 1), "EOL") > 0 Then Stop
18            Hlp = Hlp & "' " & ArgNames(i, 1)
19            If Len(ArgNames(i, 1)) < 10 Then
20                Spacers = String(10 - Len(ArgNames(i, 1)), " ")
21            Else
22                Spacers = ""
23            End If
              
24            Hlp = Hlp & Spacers
25            Hlp = Hlp & ":" & InsertBreaks(ArgDescriptions(i, 1), Len(ArgNames(i, 1)) + Len(Spacers) + 2) + vbLf
26        Next
27        If Len(ExtraHelp) > 0 Then
28            Do While (Left$(ExtraHelp, 1) = vbLf Or Left$(ExtraHelp, 1) = vbCr)
29                ExtraHelp = Right$(ExtraHelp, Len(ExtraHelp) - 1)
30            Loop
31            Hlp = Hlp & ("'" & vbLf)
32            Hlp = Hlp & "' Notes     :"
33            Hlp = Hlp & InsertBreaks(ExtraHelp)
34            Hlp = Hlp & vbLf
35        End If
36        Hlp = Hlp & "' " & String(119, "-")
37        HelpForVBE = Application.WorksheetFunction.Transpose(VBA.Split(Hlp, vbLf))
38        Exit Function
ErrHandler:
39        HelpForVBE = "#HelpForVBE (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function InsertBreaks(ByVal TheString As String, Optional FirstRowShorterBy As Long)

        Const FirstTab = 0
        Const NextTabs = 13
        Const Width = 106
        Dim DoNewLine As Boolean
        Dim i As Long
        Dim LineLength As Long
        Dim Res As String
        Dim Words
        Dim WordsNLB
    
        On Error GoTo ErrHandler
    
        If InStr(TheString, " ") = 0 Then
              InsertBreaks = TheString
              Exit Function
        End If
    
        TheString = Replace(TheString, vbLf, vbLf + " ")
        TheString = Replace(TheString, vbLf + "  ", vbLf + " ")
    
        Res = String(FirstTab, " ")
        LineLength = FirstTab + FirstRowShorterBy

        Words = VBA.Split(TheString, " ")
        WordsNLB = Words
        For i = LBound(Words) To UBound(Words)
              WordsNLB(i) = Replace(WordsNLB(i), vbLf, vbNullString)
        Next

        For i = LBound(Words) To UBound(Words)
              DoNewLine = LineLength + Len(WordsNLB(i)) > Width
              If i > 1 Then
                  If InStr(Words(i - 1), vbLf) > 0 Then
                      DoNewLine = True
                  End If
              End If

              If DoNewLine Then
                  Res = Res + vbLf + "'" + String(NextTabs, " ") + WordsNLB(i)
                  LineLength = 1 + NextTabs + Len(WordsNLB(i))
              Else
                  Res = Res + " " + WordsNLB(i)
                  LineLength = LineLength + 1 + Len(WordsNLB(i))
              End If
        Next
        InsertBreaks = Res

        Exit Function
ErrHandler:
        Throw "#InsertBreaks: " & Err.Description & "!"
End Function


' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : MarkdownHelp
' Purpose    : Formats the help as a markdown table.
' -----------------------------------------------------------------------------------------------------------------------
Function MarkdownHelp(SourceFile As String, FunctionName As String, ByVal FunctionDescription As String, _
          ArgNames, ArgDescriptions, Optional Replacements)

          Dim SourceCode As String
          Dim Declaration As String
          Dim LeftString As String, RightString As String
          Dim Hlp As String
          Dim i As Long, j As Long
          Dim ThisArgDescription As String
          Dim StringsToEncloseInBackTicks

1         On Error GoTo ErrHandler

2         If VarType(ArgNames) < vbArray Then
3             ReDim Temp(1 To 1, 1 To 1): Temp(1, 1) = ArgNames: ArgNames = Temp
4         End If

5         SourceCode = RawFileContents(SourceFile)
6         SourceCode = Replace(SourceCode, vbCrLf, vbLf)

7         LeftString = "Public Function " & FunctionName & "("

8         If InStr(SourceCode, LeftString) = 0 Then
9             LeftString = "Private Function " & FunctionName & "("
10            If InStr(SourceCode, LeftString) = 0 Then
11                LeftString = "Function " & FunctionName & "("
12                If InStr(SourceCode, LeftString) = 0 Then Throw "Cannot find function declaration in SourceFile"
13            End If
14        End If

15        Declaration = StringBetweenStrings(SourceCode, LeftString, ")", True, True)

          Dim NextChars As String
          Dim MatchPoint As Long

          'Bodge ParamArray() confuses my cheap and chearful language parsing
16        If Mid(Declaration, Len(Declaration) - 1) = "()" Then
17            Declaration = StringBetweenStrings(SourceCode, Declaration, ")", True, True)
18        End If

          'Bodge - get the "As VarType"
19        MatchPoint = InStr(SourceCode, Declaration)
20        NextChars = Mid$(SourceCode, MatchPoint + Len(Declaration), 100)
21        If Left$(NextChars, 4) = " As " Then
22            NextChars = StringBetweenStrings(NextChars, " As ", vbLf, True, False)
23            NextChars = " " & Trim(NextChars)
24            Declaration = Declaration & NextChars
25        End If

26        Hlp = "#### _" & FunctionName & "_" & vbLf

          'TODO Make this an argument
27        StringsToEncloseInBackTicks = VStack(ArgNames, "JuliaLaunch", "JuliaInclude", "JuliaEval", "JuliaCall", "JuliaCall2", "JuliaSetVar")

28        For j = 1 To sNRows(StringsToEncloseInBackTicks)
29            FunctionDescription = sRegExReplace(FunctionDescription, "\b" & StringsToEncloseInBackTicks(j, 1) & "\b", "`" & StringsToEncloseInBackTicks(j, 1) & "`", True)
30        Next j

31        Hlp = Hlp & FunctionDescription & vbLf
          
32        Hlp = Hlp & "```vba" & vbLf & _
              Declaration & vbLf & _
              "```" & vbLf & vbLf & _
              "|Argument|Description|" & vbLf & _
              "|:-------|:----------|"
          
33        For i = 1 To sNRows(ArgNames)
34            ThisArgDescription = ArgDescriptions(i, 1)
35            For j = 1 To sNRows(StringsToEncloseInBackTicks)
36                ThisArgDescription = sRegExReplace(ThisArgDescription, "\b" & StringsToEncloseInBackTicks(j, 1) & "\b", "`" & StringsToEncloseInBackTicks(j, 1) & "`", True)
37            Next j
38            ThisArgDescription = Replace(ThisArgDescription, vbCrLf, vbLf)
39            ThisArgDescription = Replace(ThisArgDescription, vbCr, vbLf)
40            ThisArgDescription = Replace(ThisArgDescription, vbLf, "<br/>")
41            Hlp = Hlp & vbLf & "|`" & ArgNames(i, 1) & "`|" & ThisArgDescription & "|"
42        Next i

43        If Not IsMissing(Replacements) Then
44            If TypeName(Replacements) = "Range" Then
45                Replacements = Replacements.Value
46            End If
47            For i = 1 To sNRows(Replacements)
48                Hlp = Replace(Hlp, Replacements(i, 1), Replacements(i, 2))
49            Next i
50        End If

51        MarkdownHelp = Application.Transpose(VBA.Split(Hlp, vbLf))

52        Exit Function
ErrHandler:
53        MarkdownHelp = "#MarkdownHelp (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function


Function RawFileContents(FileName As String)
    Dim FSO As New FileSystemObject, F As Scripting.File, T As Scripting.TextStream
    On Error GoTo ErrHandler
    Set F = FSO.GetFile(FileName)
    Set T = F.OpenAsTextStream()
    RawFileContents = T.ReadAll
    T.Close

    Exit Function
ErrHandler:
   ' Throw "#RawFileContents (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function



