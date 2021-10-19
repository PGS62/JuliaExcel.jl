Attribute VB_Name = "modVBAInterop"
Option Explicit
Private m_WshShell As Object

Public Declare PtrSafe Function GetForegroundWindow Lib "user32" () As LongPtr
Private Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal HWnd As LongPtr, ByVal lpString As String, ByVal cch As Long) As Long

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ActiveWindowTitle
' Author     : Philip Swannell
' Date       : 18-Oct-2021
' Purpose    : Returns the window caption of the active application.
' -----------------------------------------------------------------------------------------------------------------------
Function ActiveWindowTitle() As String
          Dim WinText As String
          Dim HWnd As LongLong
          Dim L As LongLong
1         On Error GoTo ErrHandler
2         HWnd = GetForegroundWindow()
3         WinText = String(255, vbNullChar)
4         L = GetWindowText(HWnd, WinText, 255)
5         ActiveWindowTitle = Left(WinText, InStr(1, WinText, vbNullChar) - 1)

6         Exit Function
ErrHandler:
7         Throw "#ActiveWindowTitle (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : Sendkeys
' Author     : Philip Swannell
' Date       : 18-Nov-2020
' Purpose    : Alternative to Application.SendKeys that has the advantage of not messing with the NUMLOCK and CAPSLOCK states
' But sendkeys is a proper PITA. Does not work as expected if Shift key is depressed for example.
' -----------------------------------------------------------------------------------------------------------------------
Public Sub Sendkeys(text As Variant, Optional wait As Boolean = False)
1         If m_WshShell Is Nothing Then Set m_WshShell = CreateObject("wscript.shell")
2         m_WshShell.Sendkeys CStr(text), wait
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : JuliaEval
' Author     : Philip Swannell
' Date       : 18-Oct-2021
' Purpose    : Evaluate some julia code, returning the result to VBA.
' Parameters :
'  JuliaCode     : Some julia code such as "1+1" or "collect(1:100)"
'  KeepExcelActive: Alas, the current implementation involves SendKeys (arrgh!) and thus it's necessary to make the
'                  Julia REPL the active application in order to send the keystrokes and then activate Excel again.
'                  Switching applications is slow. If KeepExcelActive is passed as False then we don't activate Excel
'                  to save time
' -----------------------------------------------------------------------------------------------------------------------
Function JuliaEval(ByVal JuliaCode As String, Optional KeepExcelActive As Boolean = True)
          
          Const PackageLocation As String = "c:/Projects/VBAInterop"
          Dim EN As Long
          Dim ExpressionFile As String
          Dim FlagFile As String
          Dim FSO As New Scripting.FileSystemObject
          Dim HeaderRow As Variant
          Dim NumDims As String
          Dim res As Variant
          Dim ResultFile As String
          Dim ts As Scripting.TextStream
          Static JuliaExe As String
          Static Rocket As String

1         On Error GoTo ErrHandler

2         If JuliaExe = "" Then
3             JuliaExe = DefaultJuliaExe()
4         End If
            
5         On Error Resume Next
          
6         If ActiveWindowTitle() <> JuliaExe Then
7             AppActivate JuliaExe
8         End If

9         EN = Err.Number
10        On Error GoTo ErrHandler
11        If EN <> 0 Then
12            LaunchJulia PackageLocation
13        End If

          Dim tmp As String
          
14        tmp = LocalTemp()
          
15        FlagFile = tmp & "\VBAInteropFlag.txt"
16        ResultFile = tmp & "\VBAInteropResult.csv"
17        ExpressionFile = tmp & "\VBAInteropExpression.txt"
18        Set ts = FSO.OpenTextFile(FlagFile, ForWriting, True, TristateTrue)
19        ts.Write ""
20        ts.Close
21        Set ts = FSO.OpenTextFile(ExpressionFile, ForWriting, True, TristateTrue)
22        ts.Write JuliaCode
23        ts.Close
          
24        Sendkeys "{ESC}{BACKSPACE}x{(}{)}~"

25        If KeepExcelActive Then
26            AppActivate Application.Caption
27        End If

28        Do While sFileExists(FlagFile)
29            DoEvents
30        Loop

31        res = sCSVRead(ResultFile, True, ",", , "ISO", , , 1, , , , , , , , , "UTF-8", , HeaderRow)

32        NumDims = sStringBetweenStrings(HeaderRow(1, 1), "NumDims=", "|")
33        If NumDims = "0" Then
34            res = res(1, 1)
35        End If
36        JuliaEval = res

37        Exit Function
ErrHandler:
38        JuliaEval = "#JuliaEval (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

'https://docs.julialang.org/en/v1/manual/getting-started/#man-getting-started
' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : LaunchJulia
' Author     : Philip Swannell
' Date       : 14-Oct-2021
' Purpose    : Launches Julia, and loads and instantiates the passed-in package. The instantiate step might be overkill,
'              as it really needs to be done only at install-time...
' -----------------------------------------------------------------------------------------------------------------------
Sub LaunchJulia(Optional ByVal PackageLocation As String)

          Dim LoadFile As String
          Dim LoadFileContents As String
          Dim startTime As Double
          Dim ProjectFile As String
          Dim PackageName As String
          Dim WithPackage As Boolean
          Dim JuliaExe As String

1         On Error GoTo ErrHandler
2         JuliaExe = DefaultJuliaExe()

3         WithPackage = PackageLocation <> ""
          
4         If WithPackage Then
              'Ensure Windows-style
5             PackageLocation = Replace(PackageLocation, "/", "\")
6             If Right(PackageLocation, 1) = "\" Then
7                 PackageLocation = Left(PackageLocation, Len(PackageLocation) - 1)
8             End If
9             PackageName = sSplitPath(PackageLocation)
10            If Not sFolderExists(PackageLocation) Then Throw "Cannot find folder '" + ProjectFile + "'"
11            ProjectFile = sJoinPath(PackageLocation, "Project.toml")
12            If Not sFileExists(ProjectFile) Then Throw "Cannot find file '" + ProjectFile + "'"
              'Flip to Unix style
13            PackageLocation = Replace(sSplitPath(ProjectFile, False), "\", "/")
14        End If

15        ThrowIfError sCreateFolder(LocalTemp())
16        LoadFile = LocalTemp() & "loadfile.jl"

17        LoadFileContents = "@info(""Executing " & Replace(LoadFile, "\", "/") & """)" & vbLf & _
              "@show using Revise"

          'TODO change this code once I release VBAInterop as a package on GitHub, even if private... should not need to know the location of the package on the c:drive...
18        If WithPackage Then
19            LoadFileContents = LoadFileContents & vbLf & _
                  "cd(""" + PackageLocation + """)" + vbLf + _
                  "using Pkg" + vbLf + _
                  "Pkg.activate(""" + PackageLocation + """)" & vbLf & _
                  "Pkg.instantiate()" + vbLf + _
                  "@show using " + PackageName + vbLf + _
                  "@show using Dates" + vbLf
20        End If

21        ThrowIfError sFileSave(LoadFile, LoadFileContents, "")

22        ExecuteCommand JuliaExe, " --load """ + LoadFile + """", False, vbNormalFocus
23        startTime = sElapsedTime
24        While sElapsedTime() < startTime + 2
25            DoEvents
26        Wend
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
Function LocalTemp(Optional Refresh As Boolean = False)
          Static res As String
1         On Error GoTo ErrHandler
2         If Not Refresh And res <> "" Then
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
' Procedure  : DefaultJuliaExe
' Author     : Philip Swannell
' Date       : 14-Oct-2021
' Purpose    : Returns the location of the Julia executable. Assumes the user has installed to the default location
'              suggested by the installer program, and if more than one version of Julia is installed, it returns the
'              most-recently installed version - likely, but not necessarily, the version with the highest version number.
' -----------------------------------------------------------------------------------------------------------------------
Function DefaultJuliaExe()

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



Sub vfserf()
Debug.Print ToJuliaLiteral(Array(1#, 2#, 3#))
End Sub
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
Function ToJuliaLiteral(ByVal x As Variant)
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

