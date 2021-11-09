Attribute VB_Name = "modMain"
' Copyright (c) 2021 - Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme

Option Explicit
#If VBA7 And Win64 Then
    Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
    Public Declare PtrSafe Function IsWindow Lib "user32" (ByVal hwnd As LongPtr) As Long
#Else
    Declare Function GetCurrentProcessId Lib "kernel32" () As Long
    Public Declare Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
    Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
#End If

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : JuliaLaunch
' Purpose   : Launches a local Julia session which "listens" to the current Excel session and
'             responds to calls to JuliaEval etc..
' Arguments
' MinimiseWindow: If TRUE, then the Julia session window is minimised, if FALSE (the default) then the
'             window is sized normally.
' JuliaExe  : The location of julia.exe. If omitted, then the function searches for julia.exe, first on the
'             path and then at the default locations for Julia installation on Windows, taking the most
'             recently installed version if more than one is available.
' -----------------------------------------------------------------------------------------------------------------------
Function JuliaLaunch(Optional MinimiseWindow As Boolean, Optional ByVal JuliaExe As String)
Attribute JuliaLaunch.VB_Description = "Launches a local Julia session which ""listens"" to the current Excel session and responds to calls to JuliaEval etc.."
Attribute JuliaLaunch.VB_ProcData.VB_Invoke_Func = " \n14"

          Const PackageName As String = "JuliaExcel"
          Dim Command As String
          Dim ErrorCode As Long
          Dim ErrorFile As String
          Dim FlagFile As String
          Dim HwndJulia As LongPtr
          Dim LoadFile As String
          Dim LoadFileContents As String
          Dim PID As Long
          Dim WindowPartialTitle As String
          Dim WindowTitle As String
          Dim wsh As WshShell

1         On Error GoTo ErrHandler
2         If JuliaExe = "" Then
3             JuliaExe = JuliaLocation()
4         Else
5             If LCase(Right(JuliaExe, 10)) <> "\julia.exe" Then
6                 Throw "Argument JuliaExe has been provided but is not the full path to a file with name julia.exe"
7             ElseIf Not FileExists(JuliaExe) Then
8                 Throw "Cannot find file '" + JuliaExe + "'"
9             End If
10        End If

11        PID = GetCurrentProcessId
12        WindowPartialTitle = "serving Excel PID " & CStr(PID)
13        GetHandleFromPartialCaption HwndJulia, WindowPartialTitle

14        If HwndJulia <> 0 Then
15            WindowTitle = WindowTitleFromHandle(HwndJulia)
16            JuliaLaunch = "Julia is already running in window """ & WindowTitle & """"
17            Exit Function
18        End If

19        FlagFile = LocalTemp() & "\JuliaExcelFlag_" & CStr(GetCurrentProcessId()) & ".txt"
20        ErrorFile = LocalTemp() & "\JuliaExcelLoadError_" & CStr(GetCurrentProcessId()) & ".txt"
21        If FileExists(ErrorFile) Then Kill ErrorFile
          
22        SaveTextFile FlagFile, "", TristateFalse
23        LoadFile = LocalTemp() & "\JuliaExcelStartUp_" & CStr(GetCurrentProcessId()) & ".jl"
              
24        LoadFileContents = _
              "try" & vbLf & _
              "    #println(""Executing $(@__FILE__)"")" & vbLf & _
              "    using " & PackageName & vbLf & _
              "    using Dates" & vbLf & _
              "    global const xlpid = " & CStr(GetCurrentProcessId) & vbLf & _
              "    " & PackageName & ".settitle()" & vbLf & _
              "    println(""Julia $VERSION, using JuliaExcel to serve Excel running as process ID " & CStr(GetCurrentProcessId) & """)" & vbLf & _
              "    rm(""" & Replace(FlagFile, "\", "/") & """)" & vbLf & _
              "catch e" & vbLf & _
              "    theerror = ""$e""" & vbLf & _
              "    @error theerror " & vbLf & _
              "    errorfile = """ & Replace(ErrorFile, "\", "/") & """" & vbLf & _
              "    io = open(errorfile, ""w"")" & vbLf & _
              "    write(io,theerror)" & vbLf & _
              "    close(io)" & vbLf & _
              "    rm(""" & Replace(FlagFile, "\", "/") & """)" & vbLf & _
              "    #exit()" & vbLf & _
              "end"

25        SaveTextFile LoadFile, LoadFileContents, TristateFalse
        
26        Set wsh = New WshShell
27        Command = JuliaExe & " --banner=no --load """ & LoadFile & """"
28        ErrorCode = wsh.Run(Command, IIf(MinimiseWindow, vbMinimizedFocus, vbNormalNoFocus), False)
29        If ErrorCode <> 0 Then
30            Throw "Command '" + Command + "' failed with error code " + CStr(ErrorCode)
31        End If
          
32        While FileExists(FlagFile)
33            Sleep 10
34        Wend
35        CleanLocalTemp
36        If FileExists(ErrorFile) Then
37            Throw "Julia launched but encountered an error when executing '" & LoadFile & "' the error was: " & ReadTextFile(ErrorFile, TristateFalse)
38        End If
          
39        GetHandleFromPartialCaption HwndJulia, WindowPartialTitle
40        WindowTitle = WindowTitleFromHandle(HwndJulia)
          
41        JuliaLaunch = "Julia launched OK" ' in window """ & WindowTitle & """"

42        Exit Function
ErrHandler:
43        JuliaLaunch = "#JuliaLaunch (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : JuliaLocation
' Purpose    : Returns the location of the Julia executable. First looks at the path, and if not found looks at the
'              locations to which Julia is (by default) installed. If more than one version is found then returns the
'              most recently installed.
' -----------------------------------------------------------------------------------------------------------------------
Private Function JuliaLocation()

          Dim ChildFolder As Scripting.Folder
          Dim ChosenExe As String
          Dim CreatedDate As Double
          Dim ErrString As String
          Dim ExeFile As String
          Dim Folder As String
          Dim FSO As New FileSystemObject
          Dim i As Long
          Dim ParentFolder As Scripting.Folder
          Dim ParentFolderName As String
          Dim Path As String
          Dim Paths() As String
          Dim ThisCreatedDate As Double

1         On Error GoTo ErrHandler

          'First search on PATH
2         Path = Environ("PATH")
3         Paths = VBA.Split(Path, ";")
4         For i = LBound(Paths) To UBound(Paths)
5             Folder = Paths(i)
6             If Right(Folder, 1) <> "\" Then Folder = Folder + "\"
7             ExeFile = Folder + "julia.exe"
8             If FileExists(ExeFile) Then
9                 JuliaLocation = ExeFile
10                Exit Function
11            End If
12        Next i

          'If not found on path, search in the locations to which the windows installer installs
          'julia (if the user accepts defaults) and choose the most recently installed

13        ParentFolderName = Environ("LOCALAPPDATA") & "\Programs"
14        Set ParentFolder = FSO.GetFolder(ParentFolderName)

15        For Each ChildFolder In ParentFolder.SubFolders
16            If Left(ChildFolder.Name, 5) = "Julia" Then
17                ExeFile = ParentFolder & "\" & ChildFolder.Name & "\bin\julia.exe"
18                If FileExists(ExeFile) Then
19                    ThisCreatedDate = ChildFolder.DateCreated
20                    If ThisCreatedDate > CreatedDate Then
21                        CreatedDate = ThisCreatedDate
22                        ChosenExe = ExeFile
23                    End If
24                End If
25            End If
26        Next
          
27        If ChosenExe = "" Then
28            ErrString = "Julia executable not found, after looking on the path and then in sub-folders of " + _
                  ParentFolderName + " which is the default location for Julia on Windows"
29            Throw ErrString
30        Else
31            JuliaLocation = ChosenExe
32        End If

33        Exit Function
ErrHandler:
34        Throw "#JuliaLocation (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : JuliaEval
' Purpose   : Evaluate a Julia expression and return the result to Excel or VBA.
' Arguments
' JuliaExpression: Any valid Julia code, as a string. Can also be a one-column range to evaluate multiple
'             Julia statements.
' PrecedentCell: Provides control over worksheet calculation dependency. Enter a cell or range that must be
'             calculated before JuliaEval is executed.
' -----------------------------------------------------------------------------------------------------------------------
Function JuliaEval(ByVal JuliaExpression As Variant, Optional PrecedentCell As Range)
Attribute JuliaEval.VB_Description = "Evaluate a Julia expression and return the result to Excel or VBA."
Attribute JuliaEval.VB_ProcData.VB_Invoke_Func = " \n33"
          
          Dim ExpressionFile As String
          Dim FlagFile As String
          Dim ResultFile As String
          Dim strJuliaExpression As String
          Dim Tmp As String
          Dim WindowTitle As String
          Static HwndJulia As LongPtr
          Static JuliaExe As String
          Static PID As Long

1         On Error GoTo ErrHandler

2         strJuliaExpression = ConcatenateExpressions(JuliaExpression)

3         If JuliaExe = "" Then
4             JuliaExe = JuliaLocation()
5         End If
6         If PID = 0 Then
7             PID = GetCurrentProcessId
8         End If
            
9         If HwndJulia = 0 Or IsWindow(HwndJulia) = 0 Then
10            WindowTitle = "serving Excel PID " & CStr(PID)
11            GetHandleFromPartialCaption HwndJulia, WindowTitle
12        End If

13        If HwndJulia = 0 Or IsWindow(HwndJulia) = 0 Then
14            JuliaEval = "#Please call JuliaLaunch before calling JuliaEval or JuliaCall!"
15            Exit Function
16        End If
          
17        Tmp = LocalTemp()
          
18        FlagFile = Tmp & "\JuliaExcelFlag_" & CStr(PID) & ".txt"
19        ResultFile = Tmp & "\JuliaExcelResult_" & CStr(PID) & ".txt"
20        ExpressionFile = Tmp & "\JuliaExcelExpression_" & CStr(PID) & ".txt"

21        SaveTextFile FlagFile, "", TristateTrue
22        SaveTextFile ExpressionFile, strJuliaExpression, TristateTrue
          
23        PostMessageToJulia HwndJulia

24        Do While FileExists(FlagFile)
25            Sleep 1
26            If IsWindow(HwndJulia) = 0 Then
27                JuliaEval = "#The expression evaluated caused Julia to shut down!"
28                Exit Function
29            End If
30        Loop

31        JuliaEval = UnserialiseFromFile(ResultFile)

32        Exit Function
ErrHandler:
33        JuliaEval = "#JuliaEval (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ConcatenateExpressions
' Purpose    : It's convenient to be able to pass in a multi-line expression, which we first concatenate with semi-colon
'              delimiter before passing to Julia for evaluation
' -----------------------------------------------------------------------------------------------------------------------
Private Function ConcatenateExpressions(JuliaExpression As Variant) As String
          Dim i As Long
          Dim NC As Long
          Dim Tmp() As String
1         On Error GoTo ErrHandler
2         If TypeName(JuliaExpression) = "Range" Then
3             JuliaExpression = JuliaExpression.Value
4         End If
5         Select Case NumDimensions(JuliaExpression)
              Case 0
6                 ConcatenateExpressions = CStr(JuliaExpression)
7             Case 1
8                 ConcatenateExpressions = VBA.Join(JuliaExpression, ";")
9             Case 2
10                NC = UBound(JuliaExpression, 2) - LBound(JuliaExpression, 1) + 1
11                If NC > 1 Then Throw "When passed as an array or a Range, JuliaExpression should have only one column, but got " + CStr(NC) + " columns"
12                ReDim Tmp(LBound(JuliaExpression, 1) To UBound(JuliaExpression, 1))
13                For i = LBound(Tmp) To UBound(Tmp)
14                    Tmp(i) = JuliaExpression(i, LBound(JuliaExpression, 2))
15                Next
16                ConcatenateExpressions = VBA.Join(Tmp, ";")
17            Case Else
18                Throw "Too many dimensions in JuliaExpression"
19        End Select
20        Exit Function
ErrHandler:
21        Throw "#ConcatenateExpressions (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : JuliaSetVar
' Purpose   : Set a global variable in the Julia process.
' Arguments
' VariableName: The name of the variable to be set. Must follow Julia's rules for allowed variable names.
' RefersTo  : An Excel range (from which the .Value2 property is read) or more generally a number, string,
'             Boolean, Empty or array of such types. When called from VBA, nested arrays are supported.
' PrecedentCell: Provides control over worksheet calculation dependency. Enter a cell or range that must be
'             calculated before JuliaSetVar is executed.
' -----------------------------------------------------------------------------------------------------------------------
Function JuliaSetVar(VariableName As String, RefersTo As Variant, Optional PrecedentCell As Range)
Attribute JuliaSetVar.VB_Description = "Set a global variable in the Julia process."
Attribute JuliaSetVar.VB_ProcData.VB_Invoke_Func = " \n33"
1         On Error GoTo ErrHandler
2         JuliaSetVar = JuliaCall("JuliaExcel.setvar", VariableName, RefersTo)

3         Exit Function
ErrHandler:
4         JuliaSetVar = "#JuliaSetVar (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : JuliaCall
' Purpose   : Call a named Julia function, passing in data from the worksheet or from VBA.
' Arguments
' JuliaFunction: The name of a Julia function that's defined in the Julia session, perhaps as a result of
'             prior calls to JuliaInclude.
' Args...   : Zero or more arguments, which may be Excel ranges or variables in VBA code.
' -----------------------------------------------------------------------------------------------------------------------
Function JuliaCall(JuliaFunction As String, ParamArray Args())
Attribute JuliaCall.VB_Description = "Call a named Julia function, passing in data from the worksheet or from VBA."
Attribute JuliaCall.VB_ProcData.VB_Invoke_Func = " \n33"
          Dim Expression As String
          Dim i As Long
          Dim Tmp() As String

1         On Error GoTo ErrHandler
2         If UBound(Args) >= LBound(Args) Then
3             ReDim Tmp(LBound(Args) To UBound(Args))

4             For i = LBound(Args) To UBound(Args)
5                 If TypeName(Args(i)) = "Range" Then Args(i) = Args(i).Value2
6                 Tmp(i) = MakeJuliaLiteral(Args(i))
7             Next i
8             Expression = JuliaFunction & "(" & VBA.Join$(Tmp, ",") & ")"
9         Else
10            Expression = JuliaFunction & "()"
11        End If

12        JuliaCall = JuliaEval(Expression)

13        Exit Function
ErrHandler:
14        JuliaCall = "#JuliaCall (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : JuliaCall2
' Purpose   : Call a named Julia function, passing in data from the worksheet or from VBA, with
'             control of worksheet calculation dependency.
' Arguments
' JuliaFunction: The name of a Julia function that's available in the Main module of the running Julia
'             session.
' PrecedentCell: Provides control over worksheet calculation dependency. Enter a cell or range that must be
'             calculated before JuliaCall2 is executed.
'
' Note the unpleasant repetition of the code of JuliaCall, but ParamArray is tricky to work with, and I couldn't figure
' out a way to have JuliaCall2 be a wrapper to JuliaCall.
' -----------------------------------------------------------------------------------------------------------------------
Function JuliaCall2(JuliaFunction As String, PrecedentCell As Range, ParamArray Args())
Attribute JuliaCall2.VB_Description = "Call a named Julia function, passing in data from the worksheet or from VBA, with control of worksheet calculation dependency."
Attribute JuliaCall2.VB_ProcData.VB_Invoke_Func = " \n33"
          Dim Expression As String
          Dim i As Long
          Dim Tmp() As String

1         On Error GoTo ErrHandler
2         If UBound(Args) >= LBound(Args) Then
3             ReDim Tmp(LBound(Args) To UBound(Args))
4             For i = LBound(Args) To UBound(Args)
5                 If TypeName(Args(i)) = "Range" Then Args(i) = Args(i).Value2
6                 Tmp(i) = MakeJuliaLiteral(Args(i))
7             Next i
8             Expression = JuliaFunction & "(" & VBA.Join$(Tmp, ",") & ")"
9         Else
10            Expression = JuliaFunction & "()"
11        End If

12        JuliaCall2 = JuliaEval(Expression)

13        Exit Function
ErrHandler:
14        JuliaCall2 = "#JuliaCall2 (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : JuliaInclude
' Purpose   : Load a Julia source file into the Julia process, to make additional functions available
'             via JuliaEval and JuliaCall.
' Arguments
' FileName  : The full name of the file to be included.
' PrecedentCell: Provides control over worksheet calculation dependency. Enter a cell or range that must be
'             calculated before JuliaInclude is executed.
' -----------------------------------------------------------------------------------------------------------------------
Function JuliaInclude(FileName As String, Optional PrecedentCell As Range)
Attribute JuliaInclude.VB_Description = "Load a Julia source file into the Julia process, with the likely intention of making additional functions available via JuliaEval and JuliaCall."
Attribute JuliaInclude.VB_ProcData.VB_Invoke_Func = " \n33"
1         JuliaInclude = JuliaCall("JuliaExcel.include", Replace(FileName, "\", "/"))
End Function

'05-Nov-2021 16:18:37        DESKTOP-0VD2AF0
'Expression = Fill("xxx", 1000, 1000)
'Average time in JuliaEval    1.47189380999916
'06-Nov-2021 12:28:58        PHILIP-LAPTOP
'Expression = Fill("xxx", 1000, 1000)
'Average time in JuliaEval    1.9295860900078
Private Sub SpeedTest()

          Const Expression As String = "fill(""xxx"",1000,1000)"
          Const NumCalls = 10
          Dim i As Long
          Dim Res
          Dim t1 As Double
          Dim t2 As Double

1         JuliaLaunch
2         t1 = ElapsedTime
3         For i = 1 To NumCalls
4             Res = JuliaEval(Expression)
5         Next i
6         t2 = ElapsedTime

7         Debug.Print Format(Now(), "dd-mmm-yyyy hh:mm:ss"), Environ("ComputerName")
8         Debug.Print "Expression = " & Expression
9         Debug.Print "Average time in JuliaEval", (t2 - t1) / NumCalls

End Sub


