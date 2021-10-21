Attribute VB_Name = "modVBAInterop"
Option Explicit
Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hwnd As LongPtr) As Long

Sub TimeIt()
    Dim Expression As String
    Dim i As Long
    Dim Result As Variant

    Expression = "1+1"
    
    For i = 1 To 100
        Result = JuliaEval(Expression)
    Next

End Sub

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
13            WindowTitle = "serving Excel PID " & CStr(PID)
14            LaunchJulia
15            GetHandleFromPartialCaption HwndJulia, WindowTitle
16            If HwndJulia = 0 Then
17                Throw "Unexpected error"
18            End If
19        End If
          
20        Tmp = LocalTemp()
          
21        FlagFile = Tmp & "\VBAInteropFlag_" & CStr(PID) & ".txt"
22        ResultFile = Tmp & "\VBAInteropResult_" & CStr(PID) & ".csv"
23        ExpressionFile = Tmp & "\VBAInteropExpression_" & CStr(PID) & ".txt"
24        SaveTextFile FlagFile, "", TristateTrue
25        SaveTextFile ExpressionFile, JuliaCode, TristateTrue
          
26        SendMessageToJulia HwndJulia

27        Do While FileExists(FlagFile)
28            Sleep 1
29        Loop

30        Res = CSVRead(ResultFile, True, ",", , "ISO", , , 1, , , , , , , , , "UTF-8", , HeaderRow)

31        NumDims = StringBetweenStrings(CStr(HeaderRow(1, 1)), "NumDims=", "|")
32        If NumDims = "0" Then
33            Res = Res(1, 1)
34        ElseIf NumDims = "1" Then
35            Res = TwoDColTo1D(Res)
36        End If
37        JuliaEval = Res

38        Exit Function
ErrHandler:
39        JuliaEval = "#JuliaEval (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : LaunchJulia
' Author     : Philip Swannell
' Date       : 14-Oct-2021
' Purpose    : Launches Julia, and loads and instantiates the VBAInterop package. The instantiate step might be overkill,
'              as it really needs to be done only at install-time, but useful when developing
' -----------------------------------------------------------------------------------------------------------------------
Private Sub LaunchJulia()

          Dim Command As String
          Dim ErrorCode As Long
          Dim FlagFile As String
          Dim JuliaExe As String
          Dim LoadFile As String
          Dim LoadFileContents As String
          Dim PackageLocation As String
          Dim PackageLocationUnix As String
          Dim PackageName As String
          Dim ProjectFile As String
          Dim wsh As WshShell

1         On Error GoTo ErrHandler
2         JuliaExe = DefaultJuliaExe()
3         PackageLocation = "c:\Projects\VBAInterop"
4         PackageLocationUnix = Replace(PackageLocation, "\", "/")
5         PackageName = Mid(PackageLocation, InStrRev(PackageLocation, "\") + 1)
6         If Not FolderExists(PackageLocation) Then Throw "Cannot find folder '" + PackageLocation + "'"
7         ProjectFile = PackageLocation & "\Project.toml"
8         If Not FileExists(ProjectFile) Then Throw "Cannot find file '" + ProjectFile + "'"

10        FlagFile = LocalTemp() & "\VBAInteropFlag_" & CStr(GetCurrentProcessId()) & ".txt"
11        SaveTextFile FlagFile, "", TristateFalse

12        LoadFile = LocalTemp() & "\VBAInteropStartUp_" & CStr(GetCurrentProcessId()) & ".jl"

          'TODO change this code once I release VBAInterop as a package on GitHub, even if private... should not need to know the location of the package on the c:drive...
13        LoadFileContents = "println(""Executing " & Replace(LoadFile, "\", "/") & """)" & vbLf & _
              "using Revise" & vbLf & _
              "cd(""" & PackageLocationUnix & """)" & vbLf & _
              "using Pkg" & vbLf & _
              "Pkg.activate(""" & PackageLocationUnix & """)" & vbLf & _
              "Pkg.instantiate()" & vbLf & _
              "using " & PackageName & vbLf & _
              "using Dates" & vbLf & _
              "const xlpid = " & CStr(GetCurrentProcessId) & vbLf & _
              "@show xlpid" & vbLf & _
              PackageName & ".settitle()" & vbLf & _
              "println(""Julia $VERSION, using VBAInterop to serve Excel running as process ID " & CStr(GetCurrentProcessId) & """)" & vbLf & _
              "rm(""" & Replace(FlagFile, "\", "/") & """)"

15        SaveTextFile LoadFile, LoadFileContents, TristateFalse
        
16        Set wsh = New WshShell
17        Command = JuliaExe & " --load """ & LoadFile & """"
18        ErrorCode = wsh.Run(Command, vbMinimizedNoFocus, False)
19        If ErrorCode <> 0 Then
20            Throw "Command '" + Command + "' failed with error code " + CStr(ErrorCode)
21        End If
          
23        While FileExists(FlagFile)
24            DoEvents
25        Wend
26        CleanLocalTemp

27        Exit Sub
ErrHandler:
28        Throw "#LaunchJulia (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

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


