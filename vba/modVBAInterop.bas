Attribute VB_Name = "modVBAInterop"
Option Explicit
Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
Public Declare PtrSafe Function IsWindow Lib "user32" (ByVal hwnd As LongPtr) As Long

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : JuliaEval
' Author     : Philip Swannell
' Date       : 18-Oct-2021
' Purpose    : Evaluate some julia code, returning the result to VBA.
' Parameters :
'  JuliaCode     : Some julia code such as "1+1" or "collect(1:100)"
' -----------------------------------------------------------------------------------------------------------------------
Function JuliaEval(ByVal JuliaCode As String)
          
          Dim ExpressionFile As String
          Dim FlagFile As String
          Dim HeaderRow As Variant
          Dim NumDims As String
          Dim Res As Variant
          Dim ResultFile As String
          Dim Tmp As String
          Dim WindowTitle As String
          Static HwndJulia As LongPtr
          Static JuliaExe As String
          Static PID As Long

1         On Error GoTo ErrHandler

2         If JuliaExe = "" Then
3             JuliaExe = DefaultJuliaExe()
4         End If
5         If PID = 0 Then
6             PID = GetCurrentProcessId
7         End If
            
8         If HwndJulia = 0 Or IsWindow(HwndJulia) = 0 Then
9             WindowTitle = "serving Excel PID " & CStr(PID)
10            GetHandleFromPartialCaption HwndJulia, WindowTitle
11        End If

12        If HwndJulia = 0 Or IsWindow(HwndJulia) = 0 Then
13            Throw "Cannot find instance of Julia serving this instance of Excel (PID " & CStr(PID) & "). Please call function JuliaLaunch"
14        End If
          
15        Tmp = LocalTemp()
          
16        FlagFile = Tmp & "\VBAInteropFlag_" & CStr(PID) & ".txt"
17        ResultFile = Tmp & "\VBAInteropResult_" & CStr(PID) & ".csv"
18        ExpressionFile = Tmp & "\VBAInteropExpression_" & CStr(PID) & ".txt"
19        SaveTextFile FlagFile, "", TristateTrue
20        SaveTextFile ExpressionFile, JuliaCode, TristateTrue
          
21        SendMessageToJulia HwndJulia

22        Do While FileExists(FlagFile)
23            Sleep 1
24            If IsWindow(HwndJulia) = 0 Then Throw "The expression evaluated caused Julia to shut down"
25        Loop

26        Res = CSVRead(ResultFile, True, ",", , "ISO", , , 1, , , , , , , , , "UTF-8", , HeaderRow)

27        NumDims = StringBetweenStrings(CStr(HeaderRow(1, 1)), "NumDims=", "|")
28        If NumDims = "0" Then
29            Res = Res(1, 1)
30        ElseIf NumDims = "1" Then
31            Res = TwoDColTo1D(Res)
32        End If
33        JuliaEval = Res

34        Exit Function
ErrHandler:
35        JuliaEval = "#JuliaEval (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : JuliaLaunch
' Author     : Philip Swannell
' Date       : 14-Oct-2021
' Purpose    : Launches Julia, ready to "serve" current instance of Excel.
' -----------------------------------------------------------------------------------------------------------------------
Function JuliaLaunch(Optional Minimised As Boolean)

          Const PackageName As String = "VBAInterop"
          Dim Command As String
          Dim ErrorCode As Long
          Dim ErrorFile As String
          Dim FlagFile As String
          Dim HwndJulia As LongPtr
          Dim LoadFile As String
          Dim LoadFileContents As String
          Dim WindowPartialTitle As String
          Dim WindowTitle As String
          Dim wsh As WshShell
          Dim JuliaExe As String
          Dim PID As Long

1         On Error GoTo ErrHandler

2         JuliaExe = DefaultJuliaExe()
3         PID = GetCurrentProcessId
4         WindowPartialTitle = "serving Excel PID " & CStr(PID)
5         GetHandleFromPartialCaption HwndJulia, WindowPartialTitle

6         If HwndJulia <> 0 Then
7             WindowTitle = WindowTitleFromHandle(HwndJulia)
8             JuliaLaunch = "Julia is already running with title """ & WindowTitle & """"
9             Exit Function
10        End If

11        FlagFile = LocalTemp() & "\VBAInteropFlag_" & CStr(GetCurrentProcessId()) & ".txt"
12        ErrorFile = LocalTemp() & "\VBAInteropLoadError_" & CStr(GetCurrentProcessId()) & ".txt"
13        If FileExists(ErrorFile) Then Kill ErrorFile
          
14        SaveTextFile FlagFile, "", TristateFalse
15        LoadFile = LocalTemp() & "\VBAInteropStartUp_" & CStr(GetCurrentProcessId()) & ".jl"
              
16        LoadFileContents = _
              "try" & vbLf & _
              "    #println(""Executing $(@__FILE__)"")" & vbLf & _
              "    using " & PackageName & vbLf & _
              "    using Dates" & vbLf & _
              "    global const xlpid = " & CStr(GetCurrentProcessId) & vbLf & _
              "    " & PackageName & ".settitle()" & vbLf & _
              "    println(""Julia $VERSION, using VBAInterop to serve Excel running as process ID " & CStr(GetCurrentProcessId) & """)" & vbLf & _
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

17        SaveTextFile LoadFile, LoadFileContents, TristateFalse
        
18        Set wsh = New WshShell
19        Command = JuliaExe & " --banner=no --load """ & LoadFile & """"
20        ErrorCode = wsh.Run(Command, IIf(Minimised, vbMinimizedFocus, vbNormalNoFocus), False)
21        If ErrorCode <> 0 Then
22            Throw "Command '" + Command + "' failed with error code " + CStr(ErrorCode)
23        End If
          
24        While FileExists(FlagFile)
25            Sleep 10
26        Wend
27        CleanLocalTemp
28        If FileExists(ErrorFile) Then
29            Throw "Julia launched but encountered an error when executing '" & LoadFile & "' the error was: " & ReadAllFromTextFile(ErrorFile, TristateFalse)
30        End If
          
31        GetHandleFromPartialCaption HwndJulia, WindowPartialTitle
32        WindowTitle = WindowTitleFromHandle(HwndJulia)
          
33        JuliaLaunch = "Julia launched with title """ & WindowTitle & """"

34        Exit Function
ErrHandler:
35        JuliaLaunch = "#JuliaLaunch (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : DefaultJuliaExe
' Author     : Philip Swannell
' Date       : 14-Oct-2021
' Purpose    : Returns the location of the Julia executable. Assumes the user has installed to the default location
'              suggested by the installer program, and if more than one version of Julia is installed, it returns the
'              most-recently installed version - likely, but not necessarily, the version with the highest version number.
' -----------------------------------------------------------------------------------------------------------------------
Private Function DefaultJuliaExe()

          Dim ChildFolder As Scripting.Folder
          Dim ChosenExe As String
          Dim CreatedDate As Double
          Dim ErrString As String
          Dim ExeFile As String
          Dim FSO As New FileSystemObject
          Dim ParentFolder As Scripting.Folder
          Dim ParentFolderName As String
          Dim ThisCreatedDate As Double

1         On Error GoTo ErrHandler
2         ParentFolderName = Environ("LOCALAPPDATA") & "\Programs"
3         Set ParentFolder = FSO.GetFolder(ParentFolderName)

4         For Each ChildFolder In ParentFolder.SubFolders
5             If Left(ChildFolder.Name, 5) = "Julia" Then
6                 ExeFile = ParentFolder & "\" & ChildFolder.Name & "\bin\julia.exe"
7                 If FileExists(ExeFile) Then
8                     ThisCreatedDate = ChildFolder.DateCreated
9                     If ThisCreatedDate > CreatedDate Then
10                        CreatedDate = ThisCreatedDate
11                        ChosenExe = ExeFile
12                    End If
13                End If
14            End If
15        Next
          
16        If ChosenExe = "" Then
              ErrString = "Julia executable not found, after looking in sub-folders of " + ParentFolderName + " which is the default location for Julia on Windows"
18            Throw ErrString
19        Else
20            DefaultJuliaExe = ChosenExe
21        End If

22        Exit Function
ErrHandler:
23        Throw "#DefaultJuliaExe (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : JuliaCall
' Author     : Philip Swannell
' Date       : 19-Oct-2021
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
' Procedure  : ToJuliaLiteral
' Author     : Philip Swannell
' Date       : 19-Oct-2021
' Purpose    : Convert an array into a string which julia will parse as the equivalent to the passed in x. e.g:
'
' In VBA immediate window:
' ?ToJuliaLiteral(Array(1#, 2#, 3#))
' [1.0,2.0,3.0]
'
' In Julia REPL
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
3             x = x.value
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
42                Throw "case more than two dimensions not handled"
43        End Select

44        Exit Function
ErrHandler:
45        Throw "#ToJuliaLiteral (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SingletonToJuliaLiteral
' Author     : Philip Swannell
' Date       : 20-Oct-2021
' Purpose    : Convert a singleton into a string which julia will parse as the equivalent to the passed in x.
' -----------------------------------------------------------------------------------------------------------------------
Private Function SingletonToJuliaLiteral(x As Variant)
          Dim Res As String

1         On Error GoTo ErrHandler
2         Select Case VarType(x)

              Case vbString
3                 Res = x
4                 If InStr(x, "\") > 0 Then
5                     Res = Replace(Res, "\", "\\")
6                 End If
7                 If InStr(x, vbCr) > 0 Then
8                     Res = Replace(Res, vbCr, "\r")
9                 End If
10                If InStr(x, vbLf) > 0 Then
11                    Res = Replace(Res, vbLf, "\n")
12                End If
13                If InStr(x, "$") > 0 Then
14                    Res = Replace(Res, "$", "\$")
15                End If
16                If InStr(x, """") > 0 Then
17                    Res = Replace(Res, """", "\""")
18                End If
19                SingletonToJuliaLiteral = """" & Res & """"
20                Exit Function
21            Case vbDouble
22                Res = CStr(x)
23                If InStr(Res, ".") = 0 Then
24                    If InStr(Res, "E") = 0 Then
25                        Res = Res + ".0"
26                    End If
27                End If
28                SingletonToJuliaLiteral = Res
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


