Attribute VB_Name = "modCSVHelp"
Option Explicit

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CodeToRegister
' Purpose    : Generate VBA code to register a function.
' -----------------------------------------------------------------------------------------------------------------------
Function CodeToRegister(FunctionName, Description As String, ArgDescs)

          Const DQ = """"
          Dim code As String
          Dim i As Long
          
1         On Error GoTo ErrHandler
2         If TypeName(ArgDescs) = "Range" Then ArgDescs = ArgDescs.value
3         If VarType(ArgDescs) < vbArray Then
4             ReDim Temp(1 To 1, 1 To 1): Temp(1, 1) = ArgDescs: ArgDescs = Temp
5         End If

6         If Len(Description) > 255 Then Throw "Description " + CStr(i) + " is of length " + CStr(Len(Description)) + " but must be of length 255 or less."
          
7         code = code & "' " & String(119, "-") & vbLf
8         code = code & "' Procedure  : Register" & FunctionName & vbLf
9         code = code & "' Purpose    : Register the function " & FunctionName & " with the Excel function wizard. Suggest this function is called from a" & vbLf
10        code = code & "'              WorkBook_Open event." & vbLf
11        code = code & "' " & String(119, "-") & vbLf

12        code = code & "Public Sub Register" + FunctionName + "()" + vbLf
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

15        If TypeName(ArgNames) = "Range" Then ArgNames = ArgNames.value

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

          'Bodge - get the "As VarType"
          Dim NextChars As String
          Dim matchPoint As Long
16        matchPoint = InStr(SourceCode, Declaration)
17        NextChars = Mid$(SourceCode, matchPoint + Len(Declaration), 100)
18        If Left$(NextChars, 4) = " As " Then
19            NextChars = StringBetweenStrings(NextChars, " As ", vbLf, True, False)
20            NextChars = " " & Trim(NextChars)
21            Declaration = Declaration & NextChars
22        End If

23        Hlp = "#### _" & FunctionName & "_" & vbLf

24        StringsToEncloseInBackTicks = VStack(ArgNames, "JuliaLaunch", "JuliaEval", "JuliaCall", "JuliaInclude", "#", "!")

25        For j = 1 To sNRows(StringsToEncloseInBackTicks)
26            FunctionDescription = sRegExReplace(FunctionDescription, "\b" & StringsToEncloseInBackTicks(j, 1) & "\b", "`" & StringsToEncloseInBackTicks(j, 1) & "`", True)
27        Next j

28        Hlp = Hlp & FunctionDescription & vbLf
          
29        Hlp = Hlp & "```vba" & vbLf & _
              Declaration & vbLf & _
              "```" & vbLf & vbLf & _
              "|Argument|Description|" & vbLf & _
              "|:-------|:----------|"
          
30        For i = 1 To sNRows(ArgNames)
31            ThisArgDescription = ArgDescriptions(i, 1)
32            For j = 1 To sNRows(StringsToEncloseInBackTicks)
33                ThisArgDescription = sRegExReplace(ThisArgDescription, "\b" & StringsToEncloseInBackTicks(j, 1) & "\b", "`" & StringsToEncloseInBackTicks(j, 1) & "`", True)
34            Next j
35            ThisArgDescription = Replace(ThisArgDescription, vbCrLf, vbLf)
36            ThisArgDescription = Replace(ThisArgDescription, vbCr, vbLf)
37            ThisArgDescription = Replace(ThisArgDescription, vbLf, "<br/>")
38            Hlp = Hlp & vbLf & "|`" & ArgNames(i, 1) & "`|" & ThisArgDescription & "|"
39        Next i

40        If Not IsMissing(Replacements) Then
41            If TypeName(Replacements) = "Range" Then
42                Replacements = Replacements.value
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



