Attribute VB_Name = "modUtils"
Option Explicit
Option Private Module

#If VBA7 And Win64 Then
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
#Else
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
#End If

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ElapsedTime
' Author    : Philip Swannell
' Date      : 16-Jun-2013
' Purpose   : Retrieves the current value of the performance counter, which is a high resolution (<1us)
'             time stamp that can be used for time-interval measurements.
'
'             See http://msdn.microsoft.com/en-us/library/windows/desktop/ms644904(v=vs.85).aspx
' -----------------------------------------------------------------------------------------------------------------------
Function ElapsedTime() As Double
          Dim a As Currency
          Dim b As Currency
1         On Error GoTo ErrHandler

2         QueryPerformanceCounter a
3         QueryPerformanceFrequency b
4         ElapsedTime = a / b

5         Exit Function
ErrHandler:
6         Throw "#ElapsedTime (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FileExists
' Author     : Philip Swannell
' Date       : 21-Oct-2021
' Purpose    : Does a file exit?
' -----------------------------------------------------------------------------------------------------------------------
Function FileExists(FileName As String) As Boolean
          Dim F As Scripting.File
          Static FSO As Scripting.FileSystemObject
1         On Error GoTo ErrHandler
2         If FSO Is Nothing Then Set FSO = New FileSystemObject
3         Set F = FSO.GetFile(FileName)
4         FileExists = True
5         Exit Function
ErrHandler:
6         FileExists = False
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FolderExists
' Author     : Philip Swannell
' Date       : 21-Oct-2021
' Purpose    : Does a folder exist?
' Parameters :
'  FolderPath: full path to folder, may or may not be terminated with backslash
' -----------------------------------------------------------------------------------------------------------------------
Function FolderExists(ByVal FolderPath As String) As Boolean
          Dim F As Scripting.Folder
          Dim FSO As Scripting.FileSystemObject
1         On Error GoTo ErrHandler
2         Set FSO = New FileSystemObject
3         Set F = FSO.GetFolder(FolderPath)
4         FolderExists = True
5         Exit Function
ErrHandler:
6         FolderExists = False
End Function

Function SaveTextFile(FileName As String, Contents As String, Format As TriState)
          Dim FSO As New Scripting.FileSystemObject
          Dim ts As Scripting.TextStream
1         On Error GoTo ErrHandler
2         Set ts = FSO.OpenTextFile(FileName, ForWriting, True, Format)
3         ts.Write Contents
4         ts.Close
5         SaveTextFile = FileName
6         Exit Function
ErrHandler:
7         Throw "#SaveTextFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function ReadAllFromTextFile(FileName As String, Format As TriState)
          Dim FSO As New Scripting.FileSystemObject
          Dim ts As Scripting.TextStream
1         On Error GoTo ErrHandler
2         Set ts = FSO.OpenTextFile(FileName, ForReading, , Format)
3         ReadAllFromTextFile = ts.ReadAll
4         ts.Close
5         Exit Function
ErrHandler:
6         Throw "#ReadAllFromTextFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : LocalTemp
' Author     : Philip Swannell
' Date       : 14-Oct-2021
' Purpose    : Return a writable directory for saving results files to be communicated to Julia.
' -----------------------------------------------------------------------------------------------------------------------
Function LocalTemp()
          Static Res As String
          Const FolderName = "VBAInterop"
1         On Error GoTo ErrHandler
2         If Res <> "" Then
3             LocalTemp = Res
4             Exit Function
5         End If

6         If Not FolderExists(Environ("TEMP") & "\" & FolderName) Then
              Dim F As Scripting.Folder
              Dim FSO As New FileSystemObject
7             Set F = FSO.GetFolder(Environ("TEMP"))
8             F.SubFolders.Add FolderName
9         End If
10        Res = Environ("TEMP") & "\" & FolderName

          'Arrgh Environ("Temp") still (2021) uses 8.3 convention if user name > 8 characters, patch up.
11        If InStr(Res, "~") > 0 Then
              Dim ThirdPart As String, Parts() As String
12            Parts = VBA.Split(Res, "\")
13            If UBound(Parts) > 2 Then
14                ThirdPart = Parts(2)
15                If InStr(ThirdPart, "~") > 0 Then
16                    If UCase(Left(ThirdPart, 5)) = UCase(Left(Environ("username"), 5)) Then
17                        Parts(2) = Environ("username")
18                        If FolderExists(VBA.Join(Parts, "\")) Then
19                            Res = VBA.Join(Parts, "\")
20                        End If
21                    End If
22                End If
23            End If
24        End If

25        LocalTemp = Res
26        Exit Function
ErrHandler:
27        Throw "#LocalTemp (line " & CStr(Erl) + "): " & Err.Description & "!"
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
          Dim F As Scripting.File
          Dim Fld As Scripting.Folder
          Dim FSO As New Scripting.FileSystemObject
1         On Error GoTo ErrHandler
2         Set Fld = FSO.GetFolder(LocalTemp)
3         For Each F In Fld.Files
4             If Left(F.Name, 10) = "VBAInterop" Then
5                 If (Now() - F.DateLastAccessed) > DeleteFilesOlderThan Then
6                     F.Delete
7                 End If
8             End If
9         Next
10        Exit Sub
ErrHandler:
11        Throw "#CleanLocalTemp (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : NumDimensions
' Author    : Philip Swannell
' Date      : 16-Jun-2013
' Purpose   : Returns the number of dimensions in an array variable, or 0 if the variable
'             is not an array.
' -----------------------------------------------------------------------------------------------------------------------
Function NumDimensions(x As Variant) As Long
          Dim i As Long
          Dim y As Long
1         If Not IsArray(x) Then
2             NumDimensions = 0
3             Exit Function
4         Else
5             On Error GoTo ExitPoint
6             i = 1
7             Do While True
8                 y = LBound(x, i)
9                 i = i + 1
10            Loop
11        End If
ExitPoint:
12        NumDimensions = i - 1
End Function

Sub Throw(ByVal ErrorString As String)
          '"Out of stack space" errors can lead to enormous error strings, _
           but Excel cannot handle strings longer than 32767, so just take the right part...
1         If Len(ErrorString) > 32000 Then
2             Err.Raise vbObjectError + 1, , Left$(ErrorString, 1) & Right$(ErrorString, 31999)
3         Else
4             Err.Raise vbObjectError + 1, , Right$(ErrorString, 32000)
5         End If
End Sub

Sub MakeMeAnAddin()
1         ThisWorkbook.IsAddin = True
End Sub


