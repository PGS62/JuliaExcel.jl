Attribute VB_Name = "modMain"
' Copyright (c) 2021 - Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaVBA.jl#readme

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
' Procedure  : JuliaLaunch
' Purpose    : Launches Julia, ready to "serve" current instance of Excel.
' -----------------------------------------------------------------------------------------------------------------------
Function JuliaLaunch(Optional MinimiseWindow As Boolean, Optional ByVal JuliaExe As String)

          Const PackageName As String = "JuliaVBA"
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

19        FlagFile = LocalTemp() & "\JuliaVBAFlag_" & CStr(GetCurrentProcessId()) & ".txt"
20        ErrorFile = LocalTemp() & "\JuliaVBALoadError_" & CStr(GetCurrentProcessId()) & ".txt"
21        If FileExists(ErrorFile) Then Kill ErrorFile
          
22        SaveTextFile FlagFile, "", TristateFalse
23        LoadFile = LocalTemp() & "\JuliaVBAStartUp_" & CStr(GetCurrentProcessId()) & ".jl"
              
24        LoadFileContents = _
              "try" & vbLf & _
              "    #println(""Executing $(@__FILE__)"")" & vbLf & _
              "    using " & PackageName & vbLf & _
              "    using Dates" & vbLf & _
              "    global const xlpid = " & CStr(GetCurrentProcessId) & vbLf & _
              "    " & PackageName & ".settitle()" & vbLf & _
              "    println(""Julia $VERSION, using JuliaVBA to serve Excel running as process ID " & CStr(GetCurrentProcessId) & """)" & vbLf & _
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
          
41        JuliaLaunch = "Julia launched in window """ & WindowTitle & """"

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
' Procedure  : JuliaEval
' Purpose    : Evaluate arbitrary Julia code, returning the result to VBA.
' Parameters :
'  JuliaExpression : Some Julia code such as "1+1" or "collect(1:100)", may be a one-column array for multi-line code.
'  PrecedentCell   : Offers control of the calculation order when calling from worksheet. Pass in a cell which must be
'                    calculated before the JuliaExpression is evaluated.
' -----------------------------------------------------------------------------------------------------------------------
Function JuliaEval(ByVal JuliaExpression As Variant, Optional PrecedentCell As Range)
          
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
14            Throw "Cannot find instance of Julia serving this instance of Excel (PID " & CStr(PID) & "). Please call function JuliaLaunch"
15        End If
          
16        Tmp = LocalTemp()
          
17        FlagFile = Tmp & "\JuliaVBAFlag_" & CStr(PID) & ".txt"
18        ResultFile = Tmp & "\JuliaVBAResult_" & CStr(PID) & ".txt"
19        ExpressionFile = Tmp & "\JuliaVBAExpression_" & CStr(PID) & ".txt"

20        SaveTextFile FlagFile, "", TristateTrue
21        SaveTextFile ExpressionFile, strJuliaExpression, TristateTrue
          
22        PostMessageToJulia HwndJulia

23        Do While FileExists(FlagFile)
24            Sleep 1
25            If IsWindow(HwndJulia) = 0 Then Throw "The expression evaluated caused Julia to shut down"
26        Loop

27        JuliaEval = UnserialiseFromFile(ResultFile)

28        Exit Function
ErrHandler:
29        JuliaEval = "#JuliaEval (line " & CStr(Erl) + "): " & Err.Description & "!"
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
' Procedure  : JuliaSetVar
' Purpose    : Sets a variable with global scope in the Julia session.
' Parameters :
'  VariableName : The name of the variable in Julia. An error will result if this contains disallowed characters such as spaces.
'  RefersTo     : A value to which the Julia variable will refer
'  PrecedentCell: Optional and allows control of the calculation order when calling from a worksheet. Pass in a cell
'                 which must be calculated prior to execution.
' -----------------------------------------------------------------------------------------------------------------------
Function JuliaSetVar(VariableName As String, RefersTo As Variant, Optional PrecedentCell As Range)
1         On Error GoTo ErrHandler
2         JuliaSetVar = JuliaCall("JuliaVBA.setvar", VariableName, RefersTo)

3         Exit Function
ErrHandler:
4         JuliaSetVar = "#JuliaSetVar (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : JuliaCall
' Purpose    : Call a Julia function.
' Parameters :
'  JuliaFunction: The name of the julia function to call, can be suffixed with a dot for broadcasting behaviour.
'  Args        : The arguments to the function
' -----------------------------------------------------------------------------------------------------------------------
Function JuliaCall(JuliaFunction As String, ParamArray Args())
          Dim Expression As String
          Dim i As Long
          Dim Tmp() As String

1         On Error GoTo ErrHandler
2         ReDim Tmp(LBound(Args) To UBound(Args))

3         For i = LBound(Args) To UBound(Args)
4             Tmp(i) = ToJuliaLiteral(Args(i))
5         Next i

6         Expression = JuliaFunction & "(" & VBA.Join$(Tmp, ",") & ")"

7         JuliaCall = JuliaEval(Expression)

8         Exit Function
ErrHandler:
9         JuliaCall = "#JuliaCall (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : JuliaCall2
' Purpose    : JuliaCall, but with control of calculation order.
' -----------------------------------------------------------------------------------------------------------------------
Function JuliaCall2(JuliaFunction As String, PrecedentCell As Range, ParamArray Args())
1         JuliaCall2 = JuliaCall(JuliaFunction, Args)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ToJuliaLiteral
' Purpose    : Convert an array into a string which julia will parse as the equivalent to the passed in x. e.g:
'
' In VBA immediate window:
' ?ToJuliaLiteral(Array(1#, 2#, 3#))
' [1.0,2.0,3.0]
'
' In Julia REPL:
' julia> [1.0,2.0,3.0]
' 3-element Vector{Float64}:
'  1.0
'  2.0
'  3.0
' -----------------------------------------------------------------------------------------------------------------------
Private Function ToJuliaLiteral(ByVal x As Variant)
          Dim AllSameType As Boolean
          Dim FirstType As Long
          Dim i As Long
          Dim j As Long
          Dim onerow() As String
          Dim Tmp() As String
          
1         On Error GoTo ErrHandler
2         If TypeName(x) = "Range" Then
3             x = x.Value2
4         End If

5         Select Case NumDimensions(x)
              Case 0
6                 ToJuliaLiteral = SingletonToJuliaLiteral(x)
7             Case 1
8                 ReDim Tmp(LBound(x) To UBound(x))
9                 FirstType = VarType(x(LBound(x)))
10                AllSameType = True
11                For i = LBound(x) To UBound(x)
12                    Tmp(i) = SingletonToJuliaLiteral(x(i))
13                    If AllSameType Then
14                        If VarType(x(i)) <> FirstType Then
15                            AllSameType = False
16                        End If
17                    End If
18                Next i
19                ToJuliaLiteral = IIf(AllSameType, "[", "Any[") & VBA.Join$(Tmp, ",") & "]"
20            Case 2
21                ReDim onerow(LBound(x, 2) To UBound(x, 2))
22                ReDim Tmp(LBound(x, 1) To UBound(x, 1))
23                FirstType = VarType(x(LBound(x, 1), LBound(x, 2)))
24                AllSameType = True
25                For i = LBound(x, 1) To UBound(x, 1)
26                    For j = LBound(x, 2) To UBound(x, 2)
27                        onerow(j) = SingletonToJuliaLiteral(x(i, j))
28                        If AllSameType Then
29                            If VarType(x(i, j)) <> FirstType Then
30                                AllSameType = False
31                            End If
32                        End If
33                    Next j
34                    Tmp(i) = VBA.Join$(onerow, " ")
35                Next i

36                ToJuliaLiteral = IIf(AllSameType, "[", "Any[") & VBA.Join$(Tmp, ";") & "]"
                  'One column case is tricky, could change this code when using Julia 1.7
                  'https://discourse.julialang.org/t/show-versus-parse-and-arrays-with-2-dimensions-but-only-one-column/70142/2
37                If UBound(x, 2) = LBound(x, 2) Then
                      Dim NR As Long
38                    NR = UBound(x, 1) - LBound(x, 1) + 1
39                    ToJuliaLiteral = "reshape(" & ToJuliaLiteral & "," & CStr(NR) & ",1)"
40                End If
41            Case Else
42                Throw "case more than two dimensions not handled" 'In VBA there's no way to handle arrays with arbitrary number of dimensions. Easy in Julia!
43        End Select

44        Exit Function
ErrHandler:
45        Throw "#ToJuliaLiteral (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SingletonToJuliaLiteral
' Purpose    : Convert a singleton into a string which julia will parse as the equivalent to the passed in x.
' -----------------------------------------------------------------------------------------------------------------------
Private Function SingletonToJuliaLiteral(x As Variant)
          Dim res As String

1         On Error GoTo ErrHandler
2         Select Case VarType(x)

              Case vbString
3                 res = x
4                 If InStr(x, "\") > 0 Then
5                     res = Replace(res, "\", "\\")
6                 End If
7                 If InStr(x, vbCr) > 0 Then
8                     res = Replace(res, vbCr, "\r")
9                 End If
10                If InStr(x, vbLf) > 0 Then
11                    res = Replace(res, vbLf, "\n")
12                End If
13                If InStr(x, "$") > 0 Then
14                    res = Replace(res, "$", "\$")
15                End If
16                If InStr(x, """") > 0 Then
17                    res = Replace(res, """", "\""")
18                End If
19                SingletonToJuliaLiteral = """" & res & """"
20                Exit Function
21            Case vbDouble
22                res = CStr(x)
23                If InStr(res, ".") = 0 Then
24                    If InStr(res, "E") = 0 Then
25                        res = res + ".0"
26                    End If
27                End If
28                SingletonToJuliaLiteral = res
29                Exit Function
30            Case vbLong, vbInteger
31                SingletonToJuliaLiteral = CStr(x)
32                Exit Function
33            Case vbBoolean
34                SingletonToJuliaLiteral = IIf(x, "true", "false")
35                Exit Function
36            Case vbEmpty
37                SingletonToJuliaLiteral = "missing"
38                Exit Function
39            Case vbDate
40                If CDbl(x) = CLng(x) Then
41                    SingletonToJuliaLiteral = "Date(""" & Format(x, "yyyy-mm-dd") & """)"
42                Else
43                    SingletonToJuliaLiteral = "DateTime(""" & VBA.Format$(x, "yyyy-mm-ddThh:mm:ss.000") & """)"
44                End If
45                Exit Function
46            Case Else
47                Throw "Variable of type " + TypeName(x) + " is not handled"
48        End Select

49        Exit Function
ErrHandler:
50        Throw "#SingletonToJuliaLiteral (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function JuliaInclude(FileName As String)
1         JuliaInclude = JuliaCall("JuliaVBA.include", Replace(FileName, "\", "/"))
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
          Dim res
          Dim t1 As Double
          Dim t2 As Double

1         JuliaLaunch
2         t1 = ElapsedTime
3         For i = 1 To NumCalls
4             res = JuliaEval(Expression)
5         Next i
6         t2 = ElapsedTime

7         Debug.Print Format(Now(), "dd-mmm-yyyy hh:mm:ss"), Environ("ComputerName")
8         Debug.Print "Expression = " & Expression
9         Debug.Print "Average time in JuliaEval", (t2 - t1) / NumCalls

End Sub


