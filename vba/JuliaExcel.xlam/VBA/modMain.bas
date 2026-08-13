Attribute VB_Name = "modMain"
' Copyright (c) 2021-2026 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme

Option Explicit
#If VBA7 And Win64 Then
Private Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
Private Declare PtrSafe Function IsWindow Lib "USER32" (ByVal hWnd As LongPtr) As Long
#Else
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
Private Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
#End If

Public Const gPackageName As String = "JuliaExcel"

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : JuliaIsRunning
' Purpose   : Returns TRUE if an instance of Julia is running and "listening" to the current Excel
'             session, or FALSE otherwise.
' -----------------------------------------------------------------------------------------------------------------------
Public Function JuliaIsRunning() As Boolean

          Dim HwndJulia As LongPtr
          Dim WindowPartialTitle As String

1         On Error GoTo ErrHandler
2         WindowPartialTitle = "serving Excel PID " & CStr(GetCurrentProcessId) 'Must be in synch with Julia function JuliaExcel.settitle
3         GetHandleFromPartialCaption HwndJulia, WindowPartialTitle
4         JuliaIsRunning = HwndJulia <> 0

5         Exit Function
ErrHandler:
6         JuliaIsRunning = ReThrow("JuliaIsRunning", Err, True)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : JuliaLaunch
' Purpose   : Launches a local Julia session which "listens" to the current Excel session and responds
'             to calls to JuliaEval etc..
' Arguments
' UseLinux  : TRUE to run Julia as a Linux process under Windows Subsystem for Linux; FALSE (the default) to
'             run as a Windows process.
' MinimiseWindow: If TRUE, then the Julia session window is minimised; if FALSE (the default) then the
'             window is sized normally.
' CommandLineOptions: Command line options set when launching Julia.
'             Example : `--threads=auto --banner=no`.
'             https://docs.julialang.org/en/v1/manual/command-line-options/
' Packages  : Packages to load, which must be available in the default Julia environment (or environment set
'             via the `--project` command line option). Delimit multiple packages with commas.
' BashStatements: Relevant only when UseLinux is TRUE. Bash statements executed prior to launching Julia,
'             which can be used to set environment variables. Example `export
'             JULIA_PKG_DEVDIR=/mnt/c/Projects`. Delimit multiple statements with the line feed character.
' TimeOut   : The number of seconds to wait for Julia to fully start (including any package
'             precompilation) before JuliaLaunch gives up waiting and returns an informational
'             message rather than an error - Julia is not killed, and calling JuliaLaunch or
'             JuliaEval again once it has finished starting will work normally. A separate, much
'             shorter internal check (the lesser of TimeOut and 5 seconds) detects a genuine launch
'             failure, e.g. from mal-formed CommandLineOptions, and reports that as an error
'             immediately. TimeOut is optional and defaults to 30.
' -----------------------------------------------------------------------------------------------------------------------
Public Function JuliaLaunch(Optional UseLinux As Boolean, Optional MinimiseWindow As Boolean, _
          Optional ByVal CommandLineOptions As String, Optional ByVal Packages As String, _
          Optional ByVal BashStatements As String, Optional TimeOut As Long = 30)

          Const WSLExecutable = "C:\Windows\System32\wsl.exe"
          Dim Command As String
          Dim CommsFolderX As String
          Dim ErrDescription As String
          Dim ErrorFile As String
          Dim ErrorFileX As String
          Dim ExistingPort As Long
          Dim FlagFileX As String
          Dim HwndJulia As LongPtr
          Dim IsListening As Boolean
          Dim JuliaExe As String
          Dim LaunchFile As String
          Dim LaunchFileContents As String
          Dim LaunchFileNecessary As Boolean
          Dim LaunchFileX As String
          Dim LoadFile As String
          Dim LoadFileContents As String
          Dim LoadFileX As String
          Dim PID As Long
          Dim UserSuppliedCommandLineOptions As String
          Dim usingStatements As String
          Dim WindowPartialTitle As String
          Dim WindowTitle As String
          Dim wsh As WshShell

1         On Error GoTo ErrHandler

2         If IsFunctionWizardActive() Then
3             JuliaLaunch = "#Disabled in Function Wizard!"
4             Exit Function
5         End If

6         JuliaExe = JuliaExeLocation(UseLinux)

7         UserSuppliedCommandLineOptions = CommandLineOptions
8         If InStr(CommandLineOptions, "-L") > 0 Or InStr(CommandLineOptions, "--load ") > 0 Then
9             Throw "CommandLineOptions cannot include the -L or --load options. Instead use JuliaLaunch without that option and then use JuliaCall(""include"",""path_to_file"")"
10        ElseIf InStr(CommandLineOptions, "-i") = 0 Then
              'It's convenient if Julia package OhMyREPL works correctly in the REPL, but that requires the -i (interactive) command line option.
              'https://github.com/KristofferC/OhMyREPL.jl/issues/271
11            CommandLineOptions = Trim$(CommandLineOptions) & " -i"
12        End If

13        PID = GetCurrentProcessId
14        WindowPartialTitle = "serving Excel PID " & CStr(PID) 'Must be in synch with Julia function JuliaExcel.settitle
15        GetHandleFromPartialCaption HwndJulia, WindowPartialTitle

16        If HwndJulia <> 0 Then
17            WindowTitle = WindowTitleFromHandle(HwndJulia)
18            On Error Resume Next
19            ExistingPort = GetJuliaPort()
20            On Error GoTo ErrHandler
21            If ExistingPort > 0 Then
22                IsListening = JuliaIsListening(ExistingPort)
23            End If
24            If IsListening Then
25                JuliaLaunch = "Julia is already running in window """ & WindowTitle & """"
26                Exit Function
27            Else
28                Throw "A Julia session titled """ & WindowTitle & """ is already running for this Excel session, but it is not responding to HTTP requests. Switch to that window: if it's sitting at a ""julia>"" prompt, its HTTP server has likely crashed - close the window and call JuliaLaunch again. If it's busy running code, either wait for it to finish, or press Ctrl+C to interrupt it, then call JuliaLaunch again."
29            End If
30        End If

          'Now we are not exiting early set JuliaPort to zero so that we can test for the connection having been correctly established.
31        SetJuliaPort 0

32        ErrorFile = LocalTemp() & "\LoadError_" & CStr(GetCurrentProcessId()) & ".txt"
33        If FileExists(ErrorFile) Then Kill ErrorFile

34        SaveTextFile JuliaFlagFile, "", TristateFalse
35        LoadFile = LocalTemp() & "\StartUp_" & CStr(GetCurrentProcessId()) & ".jl"

36        If UseLinux Then
37            If Not FileExists(WSLExecutable) Then
38                Throw "Cannot find the WSL executable at '" + WSLExecutable + "'. Check if the file exists and whether read and execute permissions are set user '" & Environ$("USERNAME") & "'"
39            End If

40            ErrorFileX = WSLAddress(ErrorFile)
41            FlagFileX = WSLAddress(JuliaFlagFile())
42            CommsFolderX = WSLAddress(LocalTemp())
43            LoadFileX = WSLAddress(LoadFile)
44            If BashStatements <> "" Then
45                LaunchFileNecessary = True
46                BashStatements = BashStatements & vbLf
47                LaunchFile = LocalTemp & "\launchjulia.sh"
48                LaunchFileX = WSLAddress(LaunchFile)
49                LaunchFileContents = _
                      "#!/bin/bash" & vbLf & _
                      BashStatements & _
                      JuliaExe & " " & Trim(CommandLineOptions) & " --load """ & LoadFileX & """"
50                SaveTextFile LaunchFile, LaunchFileContents, TristateFalse
51            End If
52        Else
53            FlagFileX = Replace(JuliaFlagFile(), "\", "/")
54            CommsFolderX = Replace(LocalTemp(), "\", "/")
55            ErrorFileX = Replace(ErrorFile, "\", "/")
56            LoadFileX = Replace(LoadFile, "\", "/")
57        End If

58        If UseLinux Then
59            If LaunchFileNecessary Then
60                Command = "wsl """ & LaunchFileX & """ && exit"
61            Else
62                Command = "wsl " & JuliaExe & " " & Trim(CommandLineOptions) & " --load """ & LoadFileX & """"
63            End If
64        Else
65            Command = """" & JuliaExe & """" & " " & Trim(CommandLineOptions) & " --load """ & LoadFileX & """"
66        End If

          Dim LiteralCommand As String
67        LiteralCommand = MakeJuliaLiteral(Command)
68        LiteralCommand = Mid(LiteralCommand, 2, Len(LiteralCommand) - 2)

          Dim i As Long
          Dim PackagesArray() As String

          'PGS 8 Dec 2021. It's important to make using JuliaExcel be the last "using" statement as I believe that helps avoid "world-age" problems.
69        If Packages = "" Then
70            Packages = "Dates," & gPackageName
71        Else
72            Packages = "Dates," & Packages & "," & gPackageName
73        End If
74        PackagesArray = VBA.Split(Packages, ",")

75        For i = LBound(PackagesArray) To UBound(PackagesArray)
76            Select Case PackagesArray(i)
                  Case Else
77                    usingStatements = usingStatements & _
                          "    println(""using " & Trim(PackagesArray(i)) & """)" & vbLf & _
                          "    using " & Trim(PackagesArray(i)) & vbLf
78            End Select
79        Next

80        LoadFileContents = _
              "try" & vbLf & _
              usingStatements & _
              "    setxlpid(" & CStr(GetCurrentProcessId) & ")" & vbLf & _
              "    JuliaExcel.setcommsfolder(""" & CommsFolderX & """)" & vbLf & _
              "    println(""Julia $VERSION, using " & gPackageName & " to serve Excel running as process ID " & GetCurrentProcessId() & "."")" & vbLf & _
              "    println(""Julia launched with command: " & LiteralCommand & " "")" & vbLf & _
              "    JuliaExcel.start_server()" & vbLf & _
              "    rm(""" & FlagFileX & """)" & vbLf & _
              "catch e" & vbLf & _
              "    theerror = ""$e""" & vbLf & _
              "    @error theerror " & vbLf & _
              "    errorfile = """ & ErrorFileX & """" & vbLf & _
              "    io = open(errorfile, ""w"")" & vbLf & _
              "    write(io,theerror)" & vbLf & _
              "    close(io)" & vbLf & _
              "    rm(""" & FlagFileX & """)" & vbLf & _
              "end"

81        SaveTextFile LoadFile, LoadFileContents, TristateFalse

82        Set wsh = New WshShell

          Dim NumBefore As Long
          Dim StartTime As Double
83        StartTime = ElapsedTime()
          'The title Julia gives its console window before settitle() customises it - which only
          'happens once the "using" statements (and any package precompilation they trigger) have
          'finished. Matching on this generic caption lets us detect "the process launched" without
          'waiting for precompilation to complete.
          Const GenericJuliaCaption As String = "Julia"
          Dim LaunchDetected As Boolean
          Dim LaunchDetectionSecs As Double
84        LaunchDetectionSecs = IIf(TimeOut < 5, TimeOut, 5)
85        NumBefore = NumWindowsWithCaption(GenericJuliaCaption)

86        wsh.Run Command, IIf(MinimiseWindow, vbMinimizedFocus, vbNormalNoFocus), False
          'Unfortunately, if the CommandLineOptions are invalid, Julia's window can appear briefly and
          'then close again as the process dies - and either way, the call to wsh.Run does not throw an
          'error. Work-around is to track whether a window whose caption contains "Julia" has appeared
          '(without depending on settitle() having run yet, so this also covers a launch that's just
          'slow, e.g. because of package precompilation) and, if it later disappears again while
          'JuliaFlagFile still exists, treat that as a launch failure rather than continuing to wait.
87        Do While FileExists(JuliaFlagFile)
88            Sleep 50
89            If NumWindowsWithCaption(GenericJuliaCaption) > NumBefore Then
90                LaunchDetected = True
91                If ElapsedTime() - StartTime > TimeOut Then
                      'Julia's window is present (so the process did launch) but it has not yet reported
                      'success - most likely it's still precompiling packages. That's not a failure: once
                      'it finishes, GetJuliaPort will recover the real port on the next call, so just let
                      'the user know to try again rather than reporting an error.
92                    JuliaLaunch = "Julia has not signalled it's ready after " & CStr(TimeOut) & " seconds - it may still be precompiling packages. Once you see ""JuliaExcel HTTP server listening"" printed in its window, call JuliaLaunch (or JuliaEval) again."
93                    Exit Function
94                End If
95            ElseIf LaunchDetected Then
                  'The window we detected has since disappeared, but Julia's startup script never
                  'signalled completion (JuliaFlagFile still exists) - the process must have died before
                  'finishing startup, typically because CommandLineOptions was invalid.
96                ErrDescription = "Julia's console window closed before start-up finished."
97                If UserSuppliedCommandLineOptions <> "" Then
98                    ErrDescription = ErrDescription & " Check the CommandLineOptions are valid (https://docs.julialang.org/en/v1/manual/command-line-options/)"
99                End If
100               Throw ErrDescription
101           ElseIf ElapsedTime() - StartTime > LaunchDetectionSecs Then
102               ErrDescription = "Julia failed to launch after " & CStr(LaunchDetectionSecs) & " seconds."
103               If UserSuppliedCommandLineOptions <> "" Then
104                   ErrDescription = ErrDescription & " Check the CommandLineOptions are valid (https://docs.julialang.org/en/v1/manual/command-line-options/)"
105               End If
106               Throw ErrDescription
107           End If
108       Loop
          Dim PortFile As String
          Dim PortStr As String
109       PortFile = LocalTemp() & "\Port_" & CStr(PID) & ".txt"
110       PortStr = ""
111       On Error Resume Next
112       PortStr = ReadTextFile(PortFile, TristateFalse)
113       On Error GoTo ErrHandler
114       If IsNumeric(PortStr) Then
115           If CLng(PortStr) > 0 Then
116               SetJuliaPort CLng(PortStr)
117           End If
118       End If
119       CleanLocalTemp

120       If FileExists(ErrorFile) Then
121           Throw "Julia launched but encountered an error when executing '" & LoadFile & "' the error was: " & ReadTextFile(ErrorFile, TristateFalse)
122       End If
123       If GetJuliaPort() = 0 Then
124           Throw "Failed to establish connection between Julia and Excel. Is JuliaExcel installed correctly. See https://github.com/PGS62/JuliaExcel.jl#installation"
125       End If

126       GetHandleFromPartialCaption HwndJulia, WindowPartialTitle
127       WindowTitle = WindowTitleFromHandle(HwndJulia)

128       JuliaLaunch = "Julia launched in window """ & WindowTitle & """"

129       Exit Function
ErrHandler:
130       JuliaLaunch = ReThrow("JuliaLaunch", Err, True)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : JuliaEval_LowLevel
' Purpose    : Evaluate a Julia expression, exposing more arguments than we should show to the user.
' Parameters :
'  JuliaExpression      :
'  AllowNested          : Should the function throw an error if it detects that the return from Julia cannot be displayed
'                         in a worksheet, for example if it's a dictionary or an array of arrays.
'                         Should be False when calling from a worksheet since Excel would otherwise display a single
'                         "#VALUE!" with no hint as to what caused the problem.
'  StringLengthLimit    : The longest string allowed in (an element of) the return from Julia. If exceeded the function
'                         throws an intelligible error. When calling from the worksheet, should be set to the return from
'                         GetStringLengthLimit, which returns either 255 or 32767 according to the Excel version.
'  JuliaVectorToXLColumn: Should a return from Julia that's a vector (array with one dimension) be unserialised as a two
'                         dimensional array? Should be True when calling from a worksheet, or False when calling from VBA.
'                         In both cases round tripping will work correctly.
' -----------------------------------------------------------------------------------------------------------------------
Private Function JuliaEval_LowLevel(ByVal JuliaExpression As Variant, _
          Optional AllowNested As Boolean, Optional StringLengthLimit As Long, _
          Optional JuliaVectorToXLColumn As Boolean = True)

          Dim strJuliaExpression As String

1         On Error GoTo ErrHandler
2         If GetJuliaPort() = 0 Then
3             JuliaEval_LowLevel = "#Please call JuliaLaunch before calling JuliaEval or JuliaCall!"
4             Exit Function
5         End If
6         strJuliaExpression = ConcatenateExpressions(JuliaExpression)
7         Assign JuliaEval_LowLevel, UnserialiseFromString(JuliaHttpPost(strJuliaExpression), AllowNested, StringLengthLimit, JuliaVectorToXLColumn)
8         Exit Function
ErrHandler:
9         ReThrow "JuliaEval_LowLevel", Err
End Function

Sub PreciseSleep(Milliseconds As Double)
          Dim StartTime As Double
1         StartTime = ElapsedTime()
2         Do Until ((ElapsedTime() - StartTime) > Milliseconds / 1000) Or (ElapsedTime() < StartTime)
3         Loop
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : JuliaEval
' Purpose   : Evaluate a Julia expression and return the result to an Excel worksheet.
' Arguments
' JuliaExpression: Any valid Julia code, as a string. Can also be a one-column range to evaluate multiple
'             Julia statements.
' -----------------------------------------------------------------------------------------------------------------------
Public Function JuliaEval(ByVal JuliaExpression As Variant)
1         On Error GoTo ErrHandler
          
2         If IsFunctionWizardActive() Then
3             JuliaEval = "#Disabled in Function Wizard!"
4             Exit Function
5         End If

6         Assign JuliaEval, JuliaEval_LowLevel(JuliaExpression, False, GetStringLengthLimit(), True)

7         Exit Function
ErrHandler:
8         JuliaEval = ReThrow("JuliaEval", Err, True)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : JuliaEvalVBA
' Purpose   : Evaluate a Julia expression from VBA . Differs from JuliaCall in handling of 1-dimensional
'             arrays, and strings longer than 32,767 characters. May return data of types that cannot be
'             displayed on a worksheet, such as a dictionary or an array of arrays.
' Arguments
' JuliaExpression: Any valid Julia code, as a string. Can also be a one-column range to evaluate multiple
'             Julia statements.
' -----------------------------------------------------------------------------------------------------------------------
Public Function JuliaEvalVBA(ByVal JuliaExpression As Variant)
1         On Error GoTo ErrHandler
2         Assign JuliaEvalVBA, JuliaEval_LowLevel(JuliaExpression, AllowNested:=True, StringLengthLimit:=0, JuliaVectorToXLColumn:=False)
3         Exit Function
ErrHandler:
4         ReThrow "JuliaEvalVBA", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : JuliaSetVar
' Purpose   : Set a global variable in the Julia process.
' Arguments
' VariableName: The name of the variable to be set. Must follow Julia's rules for allowed variable names.
' RefersTo  : An Excel range (from which the .Value2 property is read) or more generally a number, string,
'             Boolean, Empty or array of such types. When called from VBA, nested arrays are supported.
' -----------------------------------------------------------------------------------------------------------------------
Public Function JuliaSetVar(VariableName As String, RefersTo As Variant)
1         On Error GoTo ErrHandler
2         JuliaSetVar = JuliaCall(gPackageName & ".setvar", VariableName, RefersTo)

3         Exit Function
ErrHandler:
4         JuliaSetVar = ReThrow("JuliaSetVar", Err, True)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : JuliaInclude
' Purpose   : Load a Julia source file into the Julia process, to make additional functions available
'             via JuliaEval and JuliaCall.
' Arguments
' FileName  : The full name of the file to be included.
' -----------------------------------------------------------------------------------------------------------------------
Public Function JuliaInclude(FileName As String)
1         If IsFunctionWizardActive() Then
2             JuliaInclude = "#Disabled in Function Wizard!"
3             Exit Function
4         End If
5         JuliaInclude = JuliaCall(gPackageName & ".include", Replace(FileName, "\", "/"))
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : JuliaFlagFile
' Purpose   : Returns the name of a sentinel file. The file is created (by VBA code) at the same time as
'             the expression file and deleted (by Julia code) when Julia execution has finished.
' -----------------------------------------------------------------------------------------------------------------------
Private Function JuliaFlagFile() As String
          Static FlagFile As String
1         If FlagFile = "" Then
2             FlagFile = LocalTemp() & "\Flag_" & CStr(GetCurrentProcessId()) & ".txt"
3         End If
4         JuliaFlagFile = FlagFile
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : JuliaExeLocation
' Purpose    : Returns the location of the Julia executable. First looks at the path, and if not found looks at the
'              locations to which Julia is (by default) installed. If more than one version is found then returns the
'              most recently installed.
' -----------------------------------------------------------------------------------------------------------------------
Private Function JuliaExeLocation(Optional UseLinux As Boolean)

          Dim ChildFolder As Scripting.Folder
          Dim ChosenExe As String
          Dim CreatedDate As Double
          Dim ErrString As String
          Dim ExeFile As String
          Dim Folder As String
          Dim fso As New FileSystemObject
          Dim i As Long
          Dim ParentFolder As Scripting.Folder
          Dim ParentFolderName As String
          Dim Path As String
          Dim Paths() As String
          Dim ThisCreatedDate As Double

1         On Error GoTo ErrHandler
          
2         If UseLinux Then
3             JuliaExeLocation = "julia"
4             Exit Function
5         End If
          
          'First search on PATH
6         Path = Environ("PATH")
7         Paths = VBA.Split(Path, ";")
8         For i = LBound(Paths) To UBound(Paths)
9             Folder = Paths(i)
10            If Right(Folder, 1) <> "\" Then Folder = Folder + "\"
11            ExeFile = Folder + "julia.exe"
12            If FileExists(ExeFile) Then
13                JuliaExeLocation = ExeFile
14                Exit Function
15            End If
16        Next i

          'If not found on path, search in the locations to which the windows installer installs
          'julia (if the user accepts defaults) and choose the most recently installed
17        For i = 1 To 2
18            If i = 1 Then
19                ParentFolderName = Environ("LOCALAPPDATA") & "\Programs"
20            Else
21                ParentFolderName = Environ("LOCALAPPDATA")
22            End If
23            Set ParentFolder = fso.GetFolder(ParentFolderName)
24            For Each ChildFolder In ParentFolder.SubFolders
25                If Left(ChildFolder.Name, 5) = "Julia" Then
26                    ExeFile = ParentFolder & "\" & ChildFolder.Name & "\bin\julia.exe"
27                    If FileExists(ExeFile) Then
28                        ThisCreatedDate = ChildFolder.DateCreated
29                        If ThisCreatedDate > CreatedDate Then
30                            CreatedDate = ThisCreatedDate
31                            ChosenExe = ExeFile
32                        End If
33                    End If
34                End If
35            Next
36        Next i
          
37        If ChosenExe = "" Then
38            ErrString = "Julia executable not found, after looking on the path and in folders to which Julia " & _
                  "is typically installed on Windows. When installing Julia check the ""Add Julia to Path"" option."
39            Throw ErrString
40        Else
41            JuliaExeLocation = ChosenExe
42        End If

43        Exit Function
ErrHandler:
44        ReThrow "JuliaExeLocation", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Assign
' Purpose    : Assign b to a whether or not b is an object.
' -----------------------------------------------------------------------------------------------------------------------
Sub Assign(ByRef A, B)
1         If IsObject(B) Then
2             Set A = B
3         Else
4             Let A = B
5         End If
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ThrowIfError
' Purpose   : In the event of an error, methods intended to be callable from spreadsheets
'             return an error string (starts with "#", ends with "!"). ThrowIfError allows such
'             methods to be used from VBA code while keeping error handling robust
'             MyVariable = ThrowIfError(MyFunctionThatReturnsAStringIfAnErrorHappens(...))
' -----------------------------------------------------------------------------------------------------------------------
Function ThrowIfError(Data As Variant)
1         ThrowIfError = Data
2         If VarType(Data) = vbString Then
3             If Left$(Data, 1) = "#" Then
4                 If Right$(Data, 1) = "!" Then
5                     Throw CStr(Data)
6                 End If
7             End If
8         End If
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
21        ReThrow "ConcatenateExpressions", Err
End Function

'--------------------------------------------------
'05-Nov-2021 16:18:37        DESKTOP-0VD2AF0
'Expression = fill("xxx", 1000, 1000)
'Average time in JuliaEval    1.47189380999916
'--------------------------------------------------
'06-Nov-2021 12:28:58        PHILIP-LAPTOP
'Expression = fill("xxx", 1000, 1000)
'Average time in JuliaEval    1.9295860900078
'--------------------------------------------------
'30-Nov-2021 10:13:30        PHILIP-LAPTOP
'Expression = fill("xxx", 1000, 1000)
'Average time in JuliaEval    2.82354638000252  <--- Mmm, why the slowdown since 6-Nov version? Use of Assign?
'--------------------------------------------------
'01-Dec-2021 10:30:10       DESKTOP-0VD2AF0
'Expression = fill("xxx",1000,1000)
'Average time in JuliaEval   2.25666286000051   <-- also seeing slowdown on PC in the office
'--------------------------------------------------
'20-Sep-2023 16:34:52       DESKTOP-HSGAM5S
'Expression = fill("xxx",1000,1000)
'Average time in JuliaEval   1.42395350000006  <-- higher spec PC
'--------------------------------------------------
'29-Oct-2025 18:40:16       PHILIP-HPZ1
'Expression = fill("xxx",1000,1000)
'Average time in JuliaEval   2.66512269999985           Averaged over 10 calls
'--------------------------------------------------
'22-Dec-2025 15:57:14       MSI
'Expression = fill("xxx",1000,1000)
'Average time in JuliaEval   1.7418744300001            Averaged over 20 calls
'--------------------------------------------------
Private Sub SpeedTest()

          Const Expression As String = "fill(""xxx"",1000,1000)"
          Const UseLinux As Boolean = False
          Const NumCalls = 20
          Dim i As Long
          Dim Res As Variant
          Dim t1 As Double
          Dim t2 As Double

1         JuliaLaunch UseLinux
2         t1 = ElapsedTime
3         For i = 1 To NumCalls
4             Res = JuliaEval(Expression)
5         Next i
6         t2 = ElapsedTime

7         Debug.Print "'" & Format(Now(), "dd-mmm-yyyy hh:mm:ss"), Environ("ComputerName")
8         Debug.Print "'Expression = " & Expression
9         Debug.Print "'Average time in JuliaEval", (t2 - t1) / NumCalls, "Averaged over " & CStr(NumCalls) & " calls"
10        Debug.Print "'--------------------------------------------------"
End Sub

'--------------------------------------------------
'29-Oct-2025 18:37:22       PHILIP-HPZ1
'Expression = 1+1
'Average time in JuliaEval   6.16188229999898E-03       Averaged over 1000 calls
'--------------------------------------------------
'22-Dec-2025 15:58:22       MSI
'Expression = 1+1
'Average time in JuliaEval   0.015039338300001          Averaged over 1000 calls
'--------------------------------------------------
' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : JuliaCall_LowLevel
' Purpose    : Shared dispatch for JuliaCall and JuliaCallVBA. POSTs EncodedArgs (a 1D array in the
'              JuliaExcel wire format whose first element is a function name and remaining elements
'              are its arguments) to the /call HTTP endpoint, handled by Julia's srv_call_inner.
'              IsFromWorksheet controls result handling in the same way as for JuliaEval_LowLevel:
'                False -> AllowNested=True, no string-length limit, vectors stay 1D  (VBA caller)
'                True  -> AllowNested=False, GetStringLengthLimit(), vectors become columns (worksheet)
'              Note: cannot accept ParamArray here - VBA forbids using a ParamArray parameter as
'              an argument in any call, so the encoding loop is duplicated in each public wrapper.
' -----------------------------------------------------------------------------------------------------------------------
Private Function JuliaCall_LowLevel(EncodedArgs As String, IsFromWorksheet As Boolean)
1         On Error GoTo ErrHandler
2         If GetJuliaPort() = 0 Then
3             JuliaCall_LowLevel = "#Please call JuliaLaunch before calling JuliaEval or JuliaCall!"
4             Exit Function
5         End If
6         If IsFromWorksheet Then
7             Assign JuliaCall_LowLevel, UnserialiseFromString(JuliaHttpPost(EncodedArgs, "/call"), False, GetStringLengthLimit(), True)
8         Else
9             Assign JuliaCall_LowLevel, UnserialiseFromString(JuliaHttpPost(EncodedArgs, "/call"), True, 0, False)
10        End If
11        Exit Function
ErrHandler:
12        ReThrow "JuliaCall_LowLevel", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : JuliaCall
' Purpose   : Call a named Julia function from a worksheet, passing data in the JuliaExcel wire
'             format. Bypasses Meta.parse of large array literals (~1s for 100K doubles in
'             JuliaCallOld). Returns an error string for results that cannot be displayed on a
'             worksheet (nested arrays, dictionaries, overlong strings). See JuliaCallVBA for
'             the VBA equivalent which lifts those restrictions.
' Arguments
' JuliaFunction: The name of a Julia function visible from the Julia REPL.
' Args...   : Zero or more arguments. Each may be a number, string, Boolean, empty cell, array
'             or Range. Ranges are expanded to their .Value2 before encoding.
' -----------------------------------------------------------------------------------------------------------------------
Public Function JuliaCall(JuliaFunction As String, ParamArray Args())
          Dim Arg As Variant
          Dim ContentsSection As String
          Dim i As Long
          Dim LengthsSection As String
          Dim NumArgs As Long
          Dim NumElements As Long
          Dim ThisEncoded As String

1         On Error GoTo ErrHandler

2         If IsFunctionWizardActive() Then
3             JuliaCall = "#Disabled in Function Wizard!"
4             Exit Function
5         End If

6         ThisEncoded = Chr(163) & JuliaFunction
7         LengthsSection = CStr(Len(ThisEncoded)) & ","
8         ContentsSection = ThisEncoded
9         NumArgs = IIf(UBound(Args) >= LBound(Args), UBound(Args) - LBound(Args) + 1, 0)
10        NumElements = 1 + NumArgs
11        For i = 0 To NumArgs - 1
12            Assign Arg, Args(LBound(Args) + i)
13            If TypeName(Arg) = "Range" Then Arg = Arg.Value2
14            ThisEncoded = SerialiseElement(Arg)
15            LengthsSection = LengthsSection & CStr(Len(ThisEncoded)) & ","
16            ContentsSection = ContentsSection & ThisEncoded
17        Next i
18        Assign JuliaCall, JuliaCall_LowLevel("*1," & CStr(NumElements) & ";" & LengthsSection & ";" & ContentsSection, IsFromWorksheet:=True)

19        Exit Function
ErrHandler:
20        JuliaCall = ReThrow("JuliaCall", Err, True)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : JuliaCallVBA
' Purpose   : Call a named Julia function from VBA, passing data in the JuliaExcel wire format.
'             Differs from JuliaCall in handling of 1-dimensional arrays and strings longer
'             than 32,767 characters. May return data not displayable on a worksheet (e.g. a
'             dictionary or array of arrays).
' Arguments
' JuliaFunction: The name of a Julia function visible from the Julia REPL.
' Args...   : Zero or more arguments. Each may be a number, string, Boolean, empty cell, array
'             or Range. Ranges are expanded to their .Value2 before encoding.
' -----------------------------------------------------------------------------------------------------------------------
Public Function JuliaCallVBA(JuliaFunction As String, ParamArray Args())
          Dim Arg As Variant
          Dim ContentsSection As String
          Dim i As Long
          Dim LengthsSection As String
          Dim NumArgs As Long
          Dim NumElements As Long
          Dim ThisEncoded As String

1         On Error GoTo ErrHandler

2         ThisEncoded = Chr(163) & JuliaFunction
3         LengthsSection = CStr(Len(ThisEncoded)) & ","
4         ContentsSection = ThisEncoded
5         NumArgs = IIf(UBound(Args) >= LBound(Args), UBound(Args) - LBound(Args) + 1, 0)
6         NumElements = 1 + NumArgs
7         For i = 0 To NumArgs - 1
8             Assign Arg, Args(LBound(Args) + i)
9             If TypeName(Arg) = "Range" Then Arg = Arg.Value2
10            ThisEncoded = SerialiseElement(Arg)
11            LengthsSection = LengthsSection & CStr(Len(ThisEncoded)) & ","
12            ContentsSection = ContentsSection & ThisEncoded
13        Next i
14        Assign JuliaCallVBA, JuliaCall_LowLevel("*1," & CStr(NumElements) & ";" & LengthsSection & ";" & ContentsSection, IsFromWorksheet:=False)

15        Exit Function
ErrHandler:
16        ReThrow "JuliaCallVBA", Err
End Function

Private Sub SpeedTest2()

          Const Expression As String = "1+1"
          Const UseLinux As Boolean = False
          Const NumCalls = 1000
          Dim i As Long
          Dim Res As Variant
          Dim t1 As Double
          Dim t2 As Double

1         JuliaLaunch UseLinux
2         t1 = ElapsedTime
3         For i = 1 To NumCalls
4             Res = JuliaEval(Expression)
5             If Res <> 2 Then Stop
6         Next i
7         t2 = ElapsedTime

8         Debug.Print "'" & Format(Now(), "dd-mmm-yyyy hh:mm:ss"), Environ("ComputerName")
9         Debug.Print "'Expression = " & Expression
10        Debug.Print "'Average time in JuliaEval", (t2 - t1) / NumCalls, "Averaged over " & CStr(NumCalls) & " calls"
11        Debug.Print "'--------------------------------------------------"
End Sub



