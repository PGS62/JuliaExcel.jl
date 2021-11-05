' Installer for VBAInterop.xlam
' Philip Swannell 4 Nov 2021

'To Debug this file, install visual studio set up for debugging (https://www.codeproject.com/Tips/864659/How-to-Debug-Visual-Basic-Script-with-Visual-Studi)
'Then run from a command prompt (in the appropriate folder)
'cscript.exe /x Install.vbs

Option Explicit

Const AddInNames = "VBAInterop.xlam"

Dim gErrorsEncountered
Dim myWS, AddinsDest, MsgBoxTitle, MsgBoxTitleBad, AltStartupPath, AltStartupAlreadyDefined

Function IsProcessRunning(strComputer, strProcess)
    Dim Process, strObject
    IsProcessRunning = False
    strObject = "winmgmts://" & strComputer
    For Each Process In GetObject(strObject).InstancesOf("win32_process")
        If UCase(Process.Name) = UCase(strProcess) Then
            IsProcessRunning = True
            Exit Function
        End If
    Next
End Function

Function CheckProcess(TheProcessName)
    Dim exc, result
    exc = IsProcessRunning(".", TheProcessName)
    If (exc = True) Then
        result = MsgBox(TheProcessName & " is running. Please close it and then click OK to continue.", vbOKOnly + vbExclamation, MsgBoxTitle)
        exc = IsProcessRunning(".", TheProcessName)
        If (exc = True) Then
            result = MsgBox(TheProcessName & " is still running. Please close the program and restart the installation." + vbLf + vbLf + _
                "Can't see " & TheProcessName & "?" & vbLf & "Use Windows Task Manager to check for a ""ghost"" process.", vbOKOnly + vbExclamation, MsgBoxTitle)

            WScript.Quit
        End If
    End If
End Function

Function FolderExists(TheFolderName)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    FolderExists = fso.FolderExists(TheFolderName)
End Function

Function DeleteFolder(TheFolderName)
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    f = fso.DeleteFolder(TheFolderName)
    If Err.Number <> 0 Then
        gErrorsEncountered = True
        MsgBox "Failed to delete folder '" & TheFolderName & "'" & vbLf & _
            Err.Description, vbExclamation, MsgBoxTitleBad
    End If
End Function

Function DeleteFile(FileName)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    fso.GetFile(FileName).Delete
    If Err.Number <> 0 Then
        gErrorsEncountered = True
        MsgBox "Failed to delete file '" & FileName & "'" & vbLf & _
            Err.Description, vbExclamation, MsgBoxTitleBad
    End If
End Function

Function FileExists(FileName)
    Dim fso, f
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set f = fso.GetFile(FileName)
    On Error GoTo 0
    FileExists = TypeName(f) <> "Empty"
    Exit Function
End Function

'Pass FileNames as a string, comma-delimited for multiple files
Function CopyNamedFiles(ByVal TheSourceFolder, ByVal TheDestinationFolder, ByVal FileNames, ThrowErrorIfNoSourceFile)
    Dim fso
    Dim FileNamesArray, i, ErrorMessage
    Set fso = CreateObject("Scripting.FileSystemObject")
    If (Right(TheSourceFolder, 1) <> "\") Then
        TheSourceFolder = TheSourceFolder & "\"
    End If
    If (Right(TheDestinationFolder, 1) <> "\") Then
        TheDestinationFolder = TheDestinationFolder & "\"
    End If

    FileNamesArray = Split(FileNames, ",")
    For i = LBound(FileNamesArray) To UBound(FileNamesArray)
        If Not (FileExists(TheSourceFolder & FileNamesArray(i))) Then
            If ThrowErrorIfNoSourceFile Then
                gErrorsEncountered = True
                ErrorMessage = "Cannot find file: " & TheSourceFolder & FileNamesArray(i)
                MsgBox ErrorMessage, vbOKOnly + vbExclamation, MsgBoxTitleBad
            End If
        Else
            On Error Resume Next
            fso.CopyFile TheSourceFolder & FileNamesArray(i), TheDestinationFolder & FileNamesArray(i), True
            If Err.Number <> 0 Then
                gErrorsEncountered = True
                ErrorMessage = "Failed to copy from: " & TheSourceFolder & FileNamesArray(i) & vbLf & _
                    "to: " & TheDestinationFolder & FileNamesArray(i) & vbLf & _
                    "Error: " & Err.Description
                MsgBox ErrorMessage, vbOKOnly + vbExclamation, MsgBoxTitleBad
            End If
        End If
    Next
End Function

Function ForceFolderToExist(TheFolderName)
    If FolderExists(TheFolderName) = False Then
        Dim fso
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder (TheFolderName)
    End If
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetAltStartupPath
' Author    : Philip Swannell
' Date      : Nov-2017
' Purpose   : Gets the AltStartupPath, by looking in the Registry
'             There is some chance that this returns the wrong result - e.g. on a PC
'             where Office 16.0 was previously installed (leaving data in the Registry)
'             but the version of Office used is Office 15.0 - For example the "Bloomberg PC" in Solum's offices
'---------------------------------------------------------------------------------------
Function GetAltStartupPath() 'App)
    GetAltStartupPath = RegistryRead("HKEY_CURRENT_USER\Software\Microsoft\Office\" & OfficeVersion(1) & "\Excel\Options\AltStartup", "Not found")
End Function

'---------------------------------------------------------------------------------------
' Procedure : SetAltStartupPath
' Author    : Philip Swannell
' Date      : Nov-2017
' Purpose   : Sets the AltStartupPath, by looking in the Registry. See caution for GetAltStartupPath
'---------------------------------------------------------------------------------------
Function SetAltStartupPath(Path) '(App,Path)
    RegistryWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\" & OfficeVersion(1) & "\Excel\Options\AltStartup", Path
End Function

'---------------------------------------------------------------------------------------
' Procedure : RegistryRead
' Author    : Philip Swannell
' Date      : 30-Nov-2017
' Purpose   : Read a value from the Registry
'---------------------------------------------------------------------------------------
Function RegistryRead(RegKey, DefaultValue)
    RegistryRead = DefaultValue
    Set myWS = CreateObject("WScript.Shell")
    On Error Resume Next
    RegistryRead = myWS.RegRead(RegKey)        ' See https://msdn.microsoft.com/en-us/library/x05fawxd(v=vs.84).aspx
    Exit Function
End Function

'---------------------------------------------------------------------------------------
' Procedure : RegistryWrite
' Author    : Philip Swannell
' Date      : 30-Nov-2017
' Purpose   : Write to the Registry
'---------------------------------------------------------------------------------------
Function RegistryWrite(RegKey, NewValue)
    Dim myWS
    Set myWS = CreateObject("WScript.Shell")
    myWS.RegWrite RegKey, NewValue, "REG_SZ"        'See https://msdn.microsoft.com/en-us/library/yfdfhz1b(v=vs.84).aspx
End Function

'---------------------------------------------------------------------------------------
' Procedure : sRegistryKeyExists
' Author    : Philip Swannell
' Date      : 25-Apr-2016
' Purpose   : Returns True or False according to whether a RegKey exists in the Registry
'---------------------------------------------------------------------------------------
Function RegistryKeyExists(RegKey)
    Dim myWS, Res
    Res = Empty
    Set myWS = CreateObject("WScript.Shell")
    On Error Resume Next
    Res = myWS.RegRead(RegKey)
    On Error GoTo 0
    RegistryKeyExists = Not (IsEmpty(Res))
End Function

Function RegistryDelete(RegKey)
    Set myWS = CreateObject("WScript.Shell")
    myWS.regDelete RegKey
End Function

Function OfficeVersion(NumDecimalsAfterPoint)
    Dim i, RegKey
    For i = 20 To 11 Step -1
        RegKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & FormatNumber(i, 1) & "\Excel\"
        If RegistryKeyExists(RegKey) Then
            OfficeVersion = FormatNumber(i, NumDecimalsAfterPoint)
            Exit Function
        End If
    Next
    OfficeVersion = "Office Not found"
End Function

'Apparently VBScript has no in-line if. So create one.
Function IIf( expr, truepart, falsepart )
   IIf = falsepart
   If expr Then IIf = truepart
End Function

'***********************************************************************************************************************************************
'Effective start of this VBScript. Note elevating to admin as per http://www.winhelponline.com/blog/vbscripts-and-uac-elevation/
'although, by design we put files in places where admim shouldn't be required
'***********************************************************************************************************************************************
If WScript.Arguments.length = 0 Then
   Dim objShell, ThisFileName
   Set objShell = CreateObject("Shell.Application")
   'Pass a bogus argument, say [ uac]
   ThisFileName = WScript.ScriptFullName
   objShell.ShellExecute "wscript.exe", Chr(34) & _
      ThisFileName & Chr(34) & " uac", "", "runas", 1
Else
    Set myWS = CreateObject("WScript.Shell")
    
    MsgBoxTitle = "Install VBAInterop"
    MsgBoxTitleBad = "Install VBAInterop - Error Encountered"

    gErrorsEncountered = False
    CheckProcess "Excel.exe"

    AddinsDest = "C:\ProgramData\VBAInterop\Addins\"

    AltStartupPath = GetAltStartupPath()
    AltStartupAlreadyDefined = True
    If AltStartupPath = "" Or AltStartupPath = "Not found" Then
        AltStartupAlreadyDefined = False
        SetAltStartupPath Left(AddinsDest, Len(AddinsDest) - 1)
    End If
    'If the user already has an AltStartUp path set then we use that location...
    AddinsDest = GetAltStartupPath() & "\"

    Dim AddinsSource
    AddinsSource = WScript.ScriptFullName
    AddinsSource = Left(AddinsSource, InStrRev(AddinsSource, "\") - 1)
    AddinsSource = Left(AddinsSource, InStrRev(AddinsSource, "\"))
    AddinsSource = AddinsSource & "workbooks\"

    if OfficeVersion(0) = "Office Not found" Then
        MsgBox "Installation cannot proceed because no version of Microsoft Office has been detected opn this PC.",vbCritical,MsgBoxTitleBad
        WScript.Quit
    End If

    Dim Prompt
    Prompt = "This will install VBAInterop.xlsm by copying it from " & vbLf & vblf & _
        AddinsSource & vbLf & vbLf & _
        "To Excel's AltStartup location " & iif(AltStartupAlreadyDefined,"which is at:","which has been set to:") & vbLf & AddinsDest & vbLf & vbLf & _
        "Do you wish to continue?"
    Dim result

    result = MsgBox(Prompt, vbYesNo + vbQuestion, MsgBoxTitle)
    if result <> vbYes Then WScript.Quit

    ForceFolderToExist AddinsDest

    'Copy files
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    CopyNamedFiles AddinsSource , AddinsDest, AddInNames,True

    If gErrorsEncountered Then
        Prompt = "The install script has finished, but errors were encountered, which may mean the software will not work correctly."
        MsgBox Prompt, vbOKOnly + vbCritical, MsgBoxTitleBad
    Else
        Prompt = "VBAInterop is installed, and its functions such as JuliaEval and JuliaCall will be available the next time you start Excel."
        MsgBox Prompt, vbOKOnly + vbInformation, MsgBoxTitle
    End If

    WScript.Quit
End If