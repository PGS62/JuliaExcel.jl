Attribute VB_Name = "modUtils"
Option Explicit
Option Private Module

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

' -----------------------------------------------------------------------------------------------------------------------
' Procedure : StringBetweenStrings
' Author    : Philip Swannell
' Purpose   : The function returns the substring of the input TheString which lies between LeftString
'             and RightString.
' Arguments
' TheString : The input string to be searched.
' LeftString: The returned string will start immediately after the first occurrence of LeftString in
'             TheString. If LeftString is not found or is the null string or missing, then
'             the return will start at the first character of TheString.
' RightString: The return will stop immediately before the first subsequent occurrence of RightString. If
'             such occurrrence is not found or if RightString is the null string or
'             missing, then the return will stop at the last character of TheString.
' IncludeLeftString: If TRUE, then if LeftString appears in TheString, the return will include LeftString. This
'             argument is optional and defaults to FALSE.
' IncludeRightString: If TRUE, then if RightString appears in TheString (and appears after the first occurance
'             of LeftString) then the return will include RightString. This argument is
'             optional and defaults to FALSE.
' -----------------------------------------------------------------------------------------------------------------------
Function StringBetweenStrings(TheString As String, LeftString As String, RightString As String, Optional IncludeLeftString As Boolean, Optional IncludeRightString As Boolean)
          Dim MatchPoint1 As Long        ' the position of the first character to return
          Dim MatchPoint2 As Long        ' the position of the last character to return
          Dim FoundLeft As Boolean
          Dim FoundRight As Boolean
          
1         On Error GoTo ErrHandler
          
2         If LeftString = vbNullString Then
3             MatchPoint1 = 0
4         Else
5             MatchPoint1 = InStr(1, TheString, LeftString, vbTextCompare)
6         End If

7         If MatchPoint1 = 0 Then
8             FoundLeft = False
9             MatchPoint1 = 1
10        Else
11            FoundLeft = True
12        End If

13        If RightString = vbNullString Then
14            MatchPoint2 = 0
15        ElseIf FoundLeft Then
16            MatchPoint2 = InStr(MatchPoint1 + Len(LeftString), TheString, RightString, vbTextCompare)
17        Else
18            MatchPoint2 = InStr(1, TheString, RightString, vbTextCompare)
19        End If

20        If MatchPoint2 = 0 Then
21            FoundRight = False
22            MatchPoint2 = Len(TheString)
23        Else
24            FoundRight = True
25            MatchPoint2 = MatchPoint2 - 1
26        End If

27        If Not IncludeLeftString Then
28            If FoundLeft Then
29                MatchPoint1 = MatchPoint1 + Len(LeftString)
30            End If
31        End If

32        If IncludeRightString Then
33            If FoundRight Then
34                MatchPoint2 = MatchPoint2 + Len(RightString)
35            End If
36        End If

37        StringBetweenStrings = Mid$(TheString, MatchPoint1, MatchPoint2 - MatchPoint1 + 1)

38        Exit Function
ErrHandler:
39        StringBetweenStrings = "#StringBetweenStrings (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : TwoDColTo1D
' Author     : Philip Swannell
' Date       : 01-Nov-2021
' Purpose    : Convert a two dimensional array with a single column into a 1-dimensional array.
' -----------------------------------------------------------------------------------------------------------------------
Function TwoDColTo1D(x As Variant)
          Dim i As Long
          Dim j As Long
1         j = LBound(x, 2)
          Dim Res() As Variant
2         ReDim Res(LBound(x, 1) To UBound(x, 1))
3         For i = LBound(x, 1) To UBound(x, 1)
4             Res(i) = x(i, j)
5         Next i
6         TwoDColTo1D = Res
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : SaveTextFile
' Author     : Philip Swannell
' Date       : 01-Nov-2021
' Purpose    :
' Parameters :
'  FileName:
'  Contents:
'  Format  :
' -----------------------------------------------------------------------------------------------------------------------
Function SaveTextFile(FileName As String, Contents As String, Format As TriState)
          Dim FSO As New Scripting.FileSystemObject
          Dim ts As Scripting.TextStream
1         On Error GoTo ErrHandler
2         Set ts = FSO.OpenTextFile(FileName, ForWriting, True, Format)
3         ts.Write Contents
4         ts.Close
5         Exit Function
ErrHandler:
6         Throw "#SaveTextFile (line " & CStr(Erl) + "): " & Err.Description & "!"
End Function

Function ReadAllFromTextFile(FileName As String, Format As TriState)
          Dim FSO As New Scripting.FileSystemObject
          Dim ts As Scripting.TextStream
1         On Error GoTo ErrHandler
2         On Error GoTo ErrHandler
3         Set ts = FSO.OpenTextFile(FileName, ForReading, , Format)
4         ReadAllFromTextFile = ts.ReadAll
5         ts.Close
6         Exit Function
ErrHandler:
7         Throw "#ReadAllFromTextFile (line " & CStr(Erl) + "): " & Err.Description & "!"
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

11        LocalTemp = Res
12        Exit Function
ErrHandler:
13        Throw "#LocalTemp (line " & CStr(Erl) + "): " & Err.Description & "!"
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
           but we cannot handle strings longer than 32767, so just take the right part...
2         If Len(ErrorString) > 32000 Then
3             Err.Raise vbObjectError + 1, , Left$(ErrorString, 1) & Right$(ErrorString, 31999)
4         Else
5             Err.Raise vbObjectError + 1, , Right$(ErrorString, 32000)
6         End If
End Sub


