Attribute VB_Name = "modUtils"
' Copyright (c) 2021-2025 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme

Option Explicit
Option Private Module

#If VBA7 And Win64 Then
    Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
    Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
        Private Declare PtrSafe Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
#Else
    Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
    Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
    Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
#End If

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetFullTempPath
' Purpose    : Gets the location of the temporary folder. Works even when the username is longer than 8 characters, which
'              may not be the case for Environ("Temp").
' -----------------------------------------------------------------------------------------------------------------------
Function GetFullTempPath() As String
    Dim Buffer As String * 260
    Dim Length As Long
    Length = GetTempPath(260, Buffer)
    GetFullTempPath = Left$(Buffer, Length)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ElapsedTime
' Purpose   : Retrieves the current value of the performance counter, which is a high resolution (<1us)
'             time stamp that can be used for time-interval measurements.
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
' Purpose    : Does a file exist?
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
' Purpose    : Does a folder exist?
' Parameters :
'  FolderPath: Full path to folder, may or may not be terminated with backslash
' -----------------------------------------------------------------------------------------------------------------------
Function FolderExists(ByVal FolderPath As String) As Boolean
          Dim F As Scripting.Folder
          Static FSO As Scripting.FileSystemObject
1         On Error GoTo ErrHandler
2         If FSO Is Nothing Then Set FSO = New FileSystemObject
3         Set F = FSO.GetFolder(FolderPath)
4         FolderExists = True
5         Exit Function
ErrHandler:
6         FolderExists = False
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SaveTextFile
' Purpose    : Save a text file to disk. Retries up to 10 times, with 25 millisecond delay between tries.
'  Format  : TriStateTrue for UTF-16, TriStateFalse for ascii
' -----------------------------------------------------------------------------------------------------------------------
Function SaveTextFile(FileName As String, Contents As String, Format As TriState) As String

          Const DelayMs As Long = 25
          Const MaxRetries As Integer = 10
          Dim Attempts As Integer
          Dim ts As Scripting.TextStream
          Static FSO As Scripting.FileSystemObject

1         If FSO Is Nothing Then Set FSO = New Scripting.FileSystemObject

2         On Error GoTo ErrHandler
3         For Attempts = 1 To MaxRetries
4             On Error Resume Next
5             Set ts = FSO.OpenTextFile(FileName, ForWriting, True, Format)
6             If Err.Number = 0 Then Exit For
7             On Error GoTo ErrHandler
8             DoEvents
9             PreciseSleep DelayMs
10        Next Attempts

11        If ts Is Nothing Then Throw "Failed to open file '" & FileName & "'after " & CStr(MaxRetries) & " attempts."

12        With ts
13            .Write Contents
14            .Close
15        End With

16        SaveTextFile = FileName
17        Exit Function

ErrHandler:
18        Throw "SaveTextFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ReadTextFile
' Purpose    : Returns the contents of a text file.
'  Format  : TriStateTrue for UTF-16, TriStateFalse for ascii
' -----------------------------------------------------------------------------------------------------------------------
Function ReadTextFile(FileName As String, Format As TriState)
          Dim FSO As New Scripting.FileSystemObject
          Dim ts As Scripting.TextStream
1         On Error GoTo ErrHandler
2         Set ts = FSO.OpenTextFile(FileName, ForReading, , Format)
3         ReadTextFile = ts.ReadAll
4         ts.Close
5         Exit Function
ErrHandler:
6         Throw "#ReadTextFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : WSLAddress
' Author     : Philip Swannell
' Date       : 06-Dec-2021
' Purpose    : Convert the (Windows) address of a file into the address which references that file from within Windows
'              subsystem for Linux. e.g. WSLAddress("c:\Temp\foo.txt") = "/mnt/c/temp/foo.tmp"
' -----------------------------------------------------------------------------------------------------------------------
Function WSLAddress(WindowsAddress As String)
1         On Error GoTo ErrHandler
2         Select Case Mid(WindowsAddress, 2, 2)
           Case ":/", ":\"
3             WSLAddress = "/mnt/" & LCase(Left(WindowsAddress, 1)) & Replace(Mid(WindowsAddress, 3), "\", "/")
4         Case Else
5             Throw "WindowsAddress must start with characters ""x:\"" for some drive-letter x"
6         End Select
7         Exit Function
ErrHandler:
8         Throw "#WSLAddress (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : LocalTemp
' Purpose    : Return a writable directory for saving results files to be communicated to Julia.
' -----------------------------------------------------------------------------------------------------------------------
Function LocalTemp()
          
          Const SubFolderName = "@" & gPackageName
          Dim F As Scripting.Folder
          Dim FSO As New FileSystemObject
          Dim Parent As String
          Static Res As String

1         On Error GoTo ErrHandler

2         If Res <> "" Then
3             LocalTemp = Res
4             Exit Function
5         End If
6         Parent = GetFullTempPath()
7         If Right(Parent, 1) <> "\" Then
8             Parent = Parent & "\"
9         End If
10        If Not FolderExists(Parent & SubFolderName) Then
11            Set F = FSO.GetFolder(Parent)
12            F.SubFolders.Add SubFolderName
13        End If
14        Res = Parent & SubFolderName

15        LocalTemp = Res
16        Exit Function
ErrHandler:
17        Throw "#LocalTemp (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CleanLocalTemp
' Purpose    : Clean out files in the LocalTemp folder that have not been accessed for more than
'              DeleteFilesOlderThan days.
' -----------------------------------------------------------------------------------------------------------------------
Sub CleanLocalTemp()
          Const DeleteFilesOlderThan As Double = 3
          Dim F As Scripting.File
          Dim Fld As Scripting.Folder
          Dim FSO As New Scripting.FileSystemObject
1         On Error GoTo ErrHandler
2         Set Fld = FSO.GetFolder(LocalTemp())
3         For Each F In Fld.Files
4             If (Now() - F.DateLastAccessed) > DeleteFilesOlderThan Then
5                 F.Delete
6             End If
7         Next
8         Exit Sub
ErrHandler:
9         Throw "#CleanLocalTemp (line " & CStr(Erl) + "): " & Err.Description & "!"
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : NumDimensions
' Purpose   : Returns the number of dimensions in x, or 0 if x is not an array or is an uninitialised array.
' -----------------------------------------------------------------------------------------------------------------------
Function NumDimensions(x As Variant) As Long
          Dim i As Long
          Dim Lbnd As Long
1         On Error GoTo ErrHandler
2         Do
3             i = i + 1
4             Lbnd = LBound(x, i)
5         Loop
6         Exit Function
ErrHandler:
7         NumDimensions = i - 1
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

'Called from "Menu..." button on sheet Audit.
Sub MenuButton()
1         On Error GoTo ErrHandler
2         Application.Run "SolumAddin.xlam!AuditMenuForAddin"
3         Exit Sub
ErrHandler:
4         MsgBox "#MenuButton (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical
End Sub


