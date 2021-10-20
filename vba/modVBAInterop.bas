Attribute VB_Name = "modVBAInterop"
Option Explicit
Declare PtrSafe Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal Milliseconds As Long)
Private Declare PtrSafe Function IsWindow Lib "user32" (ByVal hwnd As LongPtr) As Long

Sub TimeIt()
    Dim Expression As String
    Dim Result As Variant
    Dim i As Long

    Expression = "1+1"
    tic
    For i = 1 To 100
        Result = JuliaEval(Expression)
    Next
    toc "JuliaEval"

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
          Dim res As Variant
          Dim res1d() As Variant
          Dim ResultFile As String
          Dim tmp As String
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
          
20        tmp = LocalTemp()
          
21        FlagFile = tmp & "\VBAInteropFlag_" & CStr(PID) & ".txt"
22        ResultFile = tmp & "\VBAInteropResult_" & CStr(PID) & ".csv"
23        ExpressionFile = tmp & "\VBAInteropExpression_" & CStr(PID) & ".txt"
24        SaveTextFile FlagFile, ""
25        SaveTextFile ExpressionFile, JuliaCode
          
26        SendMessageToJulia HwndJulia

27        Do While sFileExists(FlagFile)
28            Sleep 1
29        Loop

30        res = sCSVRead(ResultFile, True, ",", , "ISO", , , 1, , , , , , , , , "UTF-8", , HeaderRow)

31        NumDims = sStringBetweenStrings(HeaderRow(1, 1), "NumDims=", "|")
32        If NumDims = "0" Then 'TODO cope with 1-dimensional case.
33            res = res(1, 1)
34        ElseIf NumDims = "1" Then
35            res = TwoDColTo1D(res)
36        End If
37        JuliaEval = res

38        Exit Function
ErrHandler:
39        JuliaEval = "#JuliaEval (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function TwoDColTo1D(x As Variant)
          Dim i As Long
          Dim j As Long
1         j = LBound(x, 2)
          Dim res() As Variant
2         ReDim res(LBound(x, 1) To UBound(x, 1))
3         For i = LBound(x, 1) To UBound(x, 1)
4             res(i) = x(i, j)
5         Next i
6         TwoDColTo1D = res
End Function

Function SaveTextFile(FileName As String, Contents As String)
          Dim FSO As New Scripting.FileSystemObject
          Dim ts As Scripting.TextStream
1         On Error GoTo ErrHandler
2         Set ts = FSO.OpenTextFile(FileName, ForWriting, True, TristateTrue)
3         ts.Write Contents
4         ts.Close
5         Exit Function
ErrHandler:
6         Throw "#SaveTextFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : LaunchJulia
' Author     : Philip Swannell
' Date       : 14-Oct-2021
' Purpose    : Launches Julia, and loads and instantiates the VBAInterop package. The instantiate step might be overkill,
'              as it really needs to be done only at install-time, but useful when developing
' -----------------------------------------------------------------------------------------------------------------------
Sub LaunchJulia()

          Dim FlagFile As String
          Dim JuliaExe As String
          Dim LoadFile As String
          Dim LoadFileContents As String
          Dim PackageLocation As String
          Dim PackageLocationUnix As String
          Dim PackageName As String
          Dim ProjectFile As String
          Dim Command As String
          Dim wsh As WshShell
          Dim ErrorCode As Long

1         On Error GoTo ErrHandler
2         JuliaExe = DefaultJuliaExe()
3         PackageLocation = "c:\Projects\VBAInterop"
4         PackageLocationUnix = Replace(PackageLocation, "\", "/")
5         PackageName = sSplitPath(PackageLocation)
6         If Not sFolderExists(PackageLocation) Then Throw "Cannot find folder '" + PackageLocation + "'"
7         ProjectFile = sJoinPath(PackageLocation, "Project.toml")
8         If Not sFileExists(ProjectFile) Then Throw "Cannot find file '" + ProjectFile + "'"

9         ThrowIfError sCreateFolder(LocalTemp())

10        FlagFile = LocalTemp() & "\VBAInteropFlag_" & CStr(GetCurrentProcessId()) & ".txt"
11        SaveTextFile FlagFile, ""

12        LoadFile = LocalTemp() & "\VBAInteroploadfile.jl"

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
              PackageName & ".settitle()" & vbLf & _
              "println(""Julia $VERSION, using VBAInterop to serve Excel running as process ID " & CStr(GetCurrentProcessId) & """)" & vbLf & _
              "rm(""" & Replace(FlagFile, "\", "/") & """)"

15        ThrowIfError sFileSave(LoadFile, LoadFileContents, "")
        
16        Set wsh = New WshShell
17        Command = JuliaExe & " --load """ & LoadFile & """"
18        ErrorCode = wsh.Run(Command, vbNormalFocus, False)
19        If ErrorCode <> 0 Then
20            Throw "Command '" + Command + "' failed with error code " + CStr(ErrorCode)
21        End If
          
23        While sFileExists(FlagFile)
24            DoEvents
25        Wend
26        CleanLocalTemp

27        Exit Sub
ErrHandler:
28        Throw "#LaunchJulia (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : LocalTemp
' Author     : Philip Swannell
' Date       : 14-Oct-2021
' Purpose    : Return a writable directory for saving results files to be communicated to Julia.
' -----------------------------------------------------------------------------------------------------------------------
Function LocalTemp()
          Static res As String
1         On Error GoTo ErrHandler
2         If res <> "" Then
3             LocalTemp = res
4             Exit Function
5         End If
6         res = sJoinPath(Environ("TEMP"), "VBAInterop")
7         ThrowIfError sCreateFolder(res)
8         LocalTemp = res
9         Exit Function
ErrHandler:
10        Throw "#LocalTemp (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CleanLocalTemp
' Author     : Philip Swannell
' Date       : 20-Oct-2021
' Purpose    : Clean out files in the LocalTemp folder that have not been accessed for more than
'              DeleteFilesOlderThan days.
' -----------------------------------------------------------------------------------------------------------------------
Sub CleanLocalTemp()
          Const DeleteFilesOlderThan As Double = 3
          Dim FSO As New Scripting.FileSystemObject
          Dim Fld As Scripting.Folder
          Dim f As Scripting.File
1         On Error GoTo ErrHandler
2         Set Fld = FSO.GetFolder(LocalTemp)
3         For Each f In Fld.Files
4             If Left(f.Name, 10) = "VBAInterop" Then
5                 If (Now() - f.DateLastAccessed) > DeleteFilesOlderThan Then
6                     f.Delete
7                 End If
8             End If
9         Next
10        Exit Sub
ErrHandler:
11        Throw "#CleanLocalTemp (line " & CStr(Erl) + "): " & Err.Description & "!"
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

          Dim ParentFolder As String
          Dim JuliaFolders
          Dim ErrString
          Dim i As Long
          Dim ExeFile As String

1         On Error GoTo ErrHandler
2         ParentFolder = sJoinPath(Environ("LOCALAPPDATA"), "Programs")

          'Folders are returned in descending order of creation date
3         JuliaFolders = sDirList(ParentFolder, False, , "FCV", "D", , "*Julia*")
4         If sIsErrorString(JuliaFolders) Then
5             ErrString = "By default, Julia is installed in a sub-folder of '" + ParentFolder + "' but no folder starting with 'Julia' exists under that folder"
6             Throw ErrString
7         End If

8         For i = 1 To sNRows(JuliaFolders)
9             ExeFile = sJoinPath(JuliaFolders(i, 1), "bin\julia.exe")
10            If sFileExists(ExeFile) Then
11                DefaultJuliaExe = ExeFile
12                Exit Function
13            End If
14        Next i

15        ErrString = "By default, Julia is installed in a sub-folder of '" + ParentFolder + "' but cannot find a sub-sub-folder containg Julia.exe"
16        Throw ErrString

17        Exit Function
ErrHandler:
18        Throw "#DefaultJuliaExe (line " & CStr(Erl) + "): " & Err.Description & "!"
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
          Dim tmp() As String
          Dim i As Long
          Dim Expression As String

1         On Error GoTo ErrHandler
2         ReDim tmp(LBound(Args) To UBound(Args))

3         For i = LBound(Args) To UBound(Args)
4             tmp(i) = ToJuliaLiteral(Args(i))
5         Next i

6         Expression = JuliaFunction & "(" & VBA.Join$(tmp, ",") & ")"

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
          Dim tmp() As String
          Dim onerow() As String
          Dim i As Long
          Dim j As Long
          Dim AllSameType As Boolean
          Dim FirstType As Long
          
1         On Error GoTo ErrHandler
2         If TypeName(x) = "Range" Then
3             x = x.Value
4         End If

5         Select Case NumDimensions(x)
              Case 0
6                 ToJuliaLiteral = SingletonToJuliaLiteral(x)
7             Case 1
8                 ReDim tmp(LBound(x) To UBound(x))
9                 FirstType = VarType(x(LBound(x)))
10                AllSameType = True
11                For i = LBound(x) To UBound(x)
12                    tmp(i) = SingletonToJuliaLiteral(x(i))
13                    If AllSameType Then
14                        If VarType(x(i)) <> FirstType Then
15                            AllSameType = False
16                        End If
17                    End If
18                Next i
19                ToJuliaLiteral = IIf(AllSameType, "[", "Any[") & VBA.Join$(tmp, ",") & "]"
20            Case 2
21                ReDim onerow(LBound(x, 2) To UBound(x, 2))
22                ReDim tmp(LBound(x, 1) To UBound(x, 1))
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
34                    tmp(i) = VBA.Join$(onerow, " ")
35                Next i
36                ToJuliaLiteral = IIf(AllSameType, "[", "Any[") & VBA.Join$(tmp, ";") & "]"
37            Case Else
38                Throw "case more than two dimensions not handled"
39        End Select

40        Exit Function
ErrHandler:
41        Throw "#ToJuliaLiteral (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SingletonToJuliaLiteral
' Author     : Philip Swannell
' Date       : 20-Oct-2021
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

