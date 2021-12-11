Attribute VB_Name = "modUtils"
' Copyright (c) 2021 - Philip Swannell
' License MIT (https://opensource.org/licenses/MIT)
' Document: https://github.com/PGS62/JuliaExcel.jl#readme

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
' Purpose   : Retrieves the current value of the performance counter, which is a high resolution (<1us)
'             time stamp that can be used for time-interval measurements.
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

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SaveTextFile
' Purpose    : Save a text file to disk.
'  Format  : TriStateTrue for UTF-16, TriStateFalse for ascii
' -----------------------------------------------------------------------------------------------------------------------
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
          Static Res As String
1         On Error GoTo ErrHandler
2         If Res <> "" Then
3             LocalTemp = Res
4             Exit Function
5         End If

          Const SubFolderName = "@" & gPackageName

6         If Not FolderExists(Environ("TEMP") & "\" & SubFolderName) Then
              Dim F As Scripting.Folder
              Dim FSO As New FileSystemObject
7             Set F = FSO.GetFolder(Environ("TEMP"))
8             F.SubFolders.Add SubFolderName
9         End If
10        Res = Environ("TEMP") & "\" & SubFolderName

          'Arrgh. Windows 11 now uses long-form user names when creating folders under c:\Users but the _
          TEMP environment variable still has the 8.3 contraction of the user name. So fix, with caution.
          
          Dim Res2 As String
          Dim Tokenised
11        Tokenised = VBA.Split(Res, "\")
12        If UBound(Tokenised) > 2 Then
13            If InStr(Tokenised(2), "~") > 0 Then
14                If LCase(Tokenised(2)) <> LCase(Environ("username")) Then
15                    Tokenised(2) = Environ("username")
16                    Res2 = VBA.Join(Tokenised, "\")
17                    If FolderExists(Res2) Then
18                        Res = Res2
19                    End If
20                End If
21            End If
22        End If

23        LocalTemp = Res
24        Exit Function
ErrHandler:
25        Throw "#LocalTemp (line " & CStr(Erl) + "): " & Err.Description & "!"
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

'Called from "Menu..." button on sheet Audit.
Sub MenuButton()
1         On Error GoTo ErrHandler
2         Application.Run "SolumAddin.xlam!AuditMenu"
3         Exit Sub
ErrHandler:
4         MsgBox "#MenuButton (line " & CStr(Erl) + "): " & Err.Description & "!", vbCritical
End Sub


