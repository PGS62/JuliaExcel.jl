Attribute VB_Name = "modUtils"
' Copyright (c) 2021-2026 Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme

Option Explicit
Option Private Module

#If VBA7 And Win64 Then
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare PtrSafe Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As LongPtr
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As LongPtr) As Long
#Else
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
#End If

Private Const PROCESS_QUERY_LIMITED_INFORMATION As Long = &H1000

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetFullTempPath
' Purpose    : Gets the location of the temporary folder. Works even when the username is longer than 8 characters, which
'              may not be the case for Environ("Temp").
' -----------------------------------------------------------------------------------------------------------------------
Function GetFullTempPath() As String
          Dim Buffer As String * 260
          Dim Length As Long
1         Length = GetTempPath(260, Buffer)
2         GetFullTempPath = Left$(Buffer, Length)
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : ElapsedTime
' Purpose   : Retrieves the current value of the performance counter, which is a high resolution (<1us)
'             time stamp that can be used for time-interval measurements.
' -----------------------------------------------------------------------------------------------------------------------
Function ElapsedTime() As Double
          Dim A As Currency
          Dim B As Currency
1         On Error GoTo ErrHandler

2         QueryPerformanceCounter A
3         QueryPerformanceFrequency B
4         ElapsedTime = A / B

5         Exit Function
ErrHandler:
6         ReThrow "ElapsedTime", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : FileExists
' Purpose    : Does a file exist?
' -----------------------------------------------------------------------------------------------------------------------
Function FileExists(FileName As String) As Boolean
          Dim F As Scripting.File
          Static fso As Scripting.FileSystemObject
1         On Error GoTo ErrHandler
2         If fso Is Nothing Then Set fso = New FileSystemObject
3         Set F = fso.GetFile(FileName)
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
          Static fso As Scripting.FileSystemObject
1         On Error GoTo ErrHandler
2         If fso Is Nothing Then Set fso = New FileSystemObject
3         Set F = fso.GetFolder(FolderPath)
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
          Dim TS As Scripting.TextStream
          Static fso As Scripting.FileSystemObject

1         On Error GoTo ErrHandler
2         If fso Is Nothing Then Set fso = New Scripting.FileSystemObject

3         For Attempts = 1 To MaxRetries
4             On Error Resume Next
5             Set TS = fso.OpenTextFile(FileName, ForWriting, True, Format)
6             If Err.Number = 0 Then Exit For
7             On Error GoTo ErrHandler
8             DoEvents
9             PreciseSleep DelayMs
10        Next Attempts

11        If TS Is Nothing Then Throw "Failed to open file '" & FileName & "'after " & CStr(MaxRetries) & " attempts."

12        With TS
13            .Write Contents
14            .Close
15        End With

16        SaveTextFile = FileName
17        Exit Function

18        Exit Function
ErrHandler:
19        ReThrow "SaveTextFile", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : ReadTextFile
' Purpose    : Returns the contents of a text file.
'  Format  : TriStateTrue for UTF-16, TriStateFalse for ascii
' -----------------------------------------------------------------------------------------------------------------------
Function ReadTextFile(FileName As String, Format As TriState)
          Dim fso As New Scripting.FileSystemObject
          Dim TS As Scripting.TextStream
1         On Error GoTo ErrHandler
2         Set TS = fso.OpenTextFile(FileName, ForReading, , Format)
3         ReadTextFile = TS.ReadAll
4         TS.Close
5         Exit Function
ErrHandler:
6         ReThrow "ReadTextFile", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : WSLAddress
' Purpose    : Convert the (Windows) address of a file into the address which references that file from within Windows
'              subsystem for Linux. e.g. WSLAddress("c:\Temp\foo.txt") = "/mnt/c/temp/foo.tmp"
' -----------------------------------------------------------------------------------------------------------------------
Function WSLAddress(WindowsAddress As String)
1         On Error GoTo ErrHandler
2         Select Case Mid(WindowsAddress, 2, 2)
              Case ":/", ":\"
3                 WSLAddress = "/mnt/" & LCase(Left(WindowsAddress, 1)) & Replace(Mid(WindowsAddress, 3), "\", "/")
4             Case Else
5                 Throw "WindowsAddress must start with characters ""x:\"" for some drive-letter x"
6         End Select
7         Exit Function
ErrHandler:
8         ReThrow "WSLAddress", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : LocalTemp
' Purpose    : Return a writable directory for saving results files to be communicated to Julia.
' -----------------------------------------------------------------------------------------------------------------------
Function LocalTemp()
          
          Const SubFolderName = "@" & gPackageName
          Dim F As Scripting.Folder
          Dim fso As New FileSystemObject
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
11            Set F = fso.GetFolder(Parent)
12            F.SubFolders.Add SubFolderName
13        End If
14        Res = Parent & SubFolderName

15        LocalTemp = Res
16        Exit Function
ErrHandler:
17        ReThrow "LocalTemp", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TryExtractPidFromFilename
' Purpose    : Extracts the process id from a file name of the form "<Prefix>_<PID>.<ext>" (e.g.
'              "Port_60176.txt"), as used for the files JuliaLaunch writes to LocalTemp. Returns
'              False (leaving PID unchanged) if FileName doesn't match that pattern.
' -----------------------------------------------------------------------------------------------------------------------
Function TryExtractPidFromFilename(ByVal FileName As String, ByRef PID As Long) As Boolean
          Dim DotPos As Long
          Dim PidStr As String
          Dim UnderscorePos As Long
1         On Error GoTo ErrHandler
2         DotPos = InStrRev(FileName, ".")
3         If DotPos = 0 Then Exit Function
4         UnderscorePos = InStrRev(FileName, "_", DotPos - 1)
5         If UnderscorePos = 0 Then Exit Function
6         PidStr = Mid$(FileName, UnderscorePos + 1, DotPos - UnderscorePos - 1)
7         If Not IsNumeric(PidStr) Then Exit Function
8         PID = CLng(PidStr)
9         TryExtractPidFromFilename = True
10        Exit Function
ErrHandler:
11        ReThrow "TryExtractPidFromFilename", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : IsProcessRunning
' Purpose    : Returns True if a process with the given id is currently running.
' -----------------------------------------------------------------------------------------------------------------------
Function IsProcessRunning(ByVal PID As Long) As Boolean
          Dim hProcess As LongPtr
1         On Error GoTo ErrHandler
2         hProcess = OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, 0, PID)
3         IsProcessRunning = (hProcess <> 0)
4         If hProcess <> 0 Then CloseHandle hProcess
5         Exit Function
ErrHandler:
6         ReThrow "IsProcessRunning", Err
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : CleanLocalTemp
' Purpose    : Clean out files in the LocalTemp folder whose name encodes the process id of an Excel
'              session (e.g. "Port_60176.txt") that is no longer running. Previously deleted files
'              based on age (not accessed for 3+ days), but NTFS disables last-access-time updates
'              by default, so that timestamp may not reflect real usage at all - and in any case, an
'              Excel session (and its attached Julia session) can legitimately stay open for well
'              over 3 days, so age was never a reliable proxy for "no longer needed". A file whose
'              name doesn't match the "<Prefix>_<PID>.<ext>" pattern is left alone rather than
'              guessed at.
' -----------------------------------------------------------------------------------------------------------------------
Sub CleanLocalTemp()
          Dim F As Scripting.File
          Dim Fld As Scripting.Folder
          Dim fso As New Scripting.FileSystemObject
          Dim PID As Long
1         On Error GoTo ErrHandler
2         Set Fld = fso.GetFolder(LocalTemp())
3         For Each F In Fld.Files
4             If TryExtractPidFromFilename(F.Name, PID) Then
5                 If Not IsProcessRunning(PID) Then
6                     F.Delete
7                 End If
8             End If
9         Next
10        Exit Sub
ErrHandler:
11        ReThrow "CleanLocalTemp", Err
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
2         Application.Run "SolumAddin.xlam!AuditMenu"
3         Exit Sub
ErrHandler:
4         MsgBox ReThrow("MenuButton", Err, True), vbCritical
End Sub
