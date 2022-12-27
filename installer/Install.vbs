' Installer for JuliaExcel.xlam
' Philip Swannell 19 Dec 2021

'To Debug this file, install visual studio set up for debugging 
'Then run from a command prompt (in the appropriate folder)
'cscript.exe /x Install.vbs
'https://www.codeproject.com/Tips/864659/How-to-Debug-Visual-Basic-Script-with-Visual-Studi

Option Explicit

Const AddinName = "JuliaExcel.xlam"
Const Website = "https://github.com/PGS62/JuliaExcel.jl"
Const GIFRecordingFlagFile = "C:\Temp\RecordingGIF.tmp"
Const MsgBoxTitle = "Install JuliaExcel"
Const MsgBoxTitleBad = "Install JuliaExcel - Error Encountered"
'Const AddinsDest = "C:\ProgramData\JuliaExcel\"  'Would need Admin rights to write to
Const AddinsDest = "C:\Users\Public\JuliaExcel\"  'Does not need admin rights
Const ElevateToAdmin = False 'Since writing to c:\Users\Public does not need admin rights

' Putting the add-in in the same folder for all users has both advantages and 
' disadvantages:
' Advantage: Avoid "Excel Link Hell" caused by the fact that workbooks store the 
'     absolute address of files to which they link (unless the file is in the same 
'     folder). Causes endless problems when two users share a workbook.
' Disadvantage: Two different users of the same PC would share copy of the add-in and 
'     thus be forced to use the same version of the add-in, though they don't both have 
'     to have the addin installed since that's controlled via the registry, which _is_ 
'     user specific.

Dim gErrorsEncountered
Dim GIFRecordingMode
Dim AddinsSource
Dim IntellisenseSource
Dim IntellisenseName
Dim InstallIntellisense
Dim Prompt
Dim myWS

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

'---------------------------------------------------------------------------------------
' Procedure : CheckExcel
' Purpose   : Invite user to shut down Excel, exits once the user does so or quits the
'             script if they decline.
'---------------------------------------------------------------------------------------
Function CheckExcel()
    Const ProcessName = "Excel.exe"
    Const FriendlyName = "Microsoft Excel"
    Dim exc, result
    exc = IsProcessRunning(".", ProcessName)
    If (exc = True) Then
        result = MsgBox(FriendlyName & " is running. Please close it and then click OK to continue.", _
        vbOKOnly + vbExclamation, MsgBoxTitle)
        While True      
            exc = IsProcessRunning(".", ProcessName)
            If (exc = True) Then
                result = MsgBox(FriendlyName & " is still running. Please close it and then click OK to continue, or click Cancel to quit." & vbLf & vbLf & _
                 "Can't see " & FriendlyName & "?" & vbLf & "Use Windows Task Manager to check if " & FriendlyName & _
                 " is running as a ""background process"", and if so use the right-click menu to ""End task"" the process.", vbOKCancel + vbExclamation, MsgBoxTitle)
                If result <> vbOK Then
                    WScript.Quit
                End If
            Else
                Exit Function
            End If
        Wend
    End If
End Function

Function FolderExists(TheFolderName)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    FolderExists = fso.FolderExists(TheFolderName)
End Function

Function FolderIsWritable(FolderPath)
          Dim FName
          Dim fso
          Dim Counter
          Dim EN
          Dim T

         If (Right(FolderPath, 1) <> "\") Then FolderPath = FolderPath & "\"
         Set fso = CreateObject("Scripting.FileSystemObject")
         If Not fso.FolderExists(FolderPath) Then
             FolderIsWritable = False
         Else
             Do
                 FName = FolderPath & "TempFile" & Counter & ".tmp"
                 Counter = Counter + 1
            Loop Until Not FileExists(FName)
            On Error Resume Next
            Set T = fso.OpenTextFile(FName, 2, True)
            EN = Err.Number
            On Error GoTo 0
            If EN = 0 Then
                T.Close
                fso.GetFile(FName).Delete
                FolderIsWritable = True
            Else
                FolderIsWritable = False
            End If
        End If
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
Function CopyNamedFiles(ByVal TheSourceFolder, ByVal TheDestinationFolder, _
                        ByVal FileNames, ThrowErrorIfNoSourceFile)
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
            If FileExists(TheDestinationFolder & FileNamesArray(i)) Then
                On Error Resume Next
                MakeFileWritable TheDestinationFolder & FileNamesArray(i)
            End If
            On Error Resume Next
            fso.CopyFile TheSourceFolder & FileNamesArray(i), _
                         TheDestinationFolder & FileNamesArray(i), True
            If Err.Number <> 0 Then
                gErrorsEncountered = True
                ErrorMessage = "Failed to copy from: " & _
                    TheSourceFolder & FileNamesArray(i) & vbLf & _
                    "to: " & TheDestinationFolder & FileNamesArray(i) & vbLf & _
                    "Error: " & Err.Description
                    If FileExists(TheSourceFolder & FileNamesArray(i)) Then
                        If FileExists(TheDestinationFolder & FileNamesArray(i)) Then
                            ErrorMessage = ErrorMessage & vblf & vbLf & _
                                "Does another user of this PC have the file open in Excel? Check that no other users of the PC are logged in"
                        End If
                    End If
                MsgBox ErrorMessage, vbOKOnly + vbExclamation, MsgBoxTitleBad
            End If
        End If
    Next
End Function

Function MakeFileWritable(FileName)
    Const ReadOnly = 1
    Dim fso
    Dim f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(FileName)
    If f.Attributes And ReadOnly Then
       f.Attributes = f.Attributes XOR ReadOnly 
    End If
End Function

Function MakeFileReadOnly(FileName)
    Const ReadOnly = 1
    Dim fso
    Dim f
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.GetFile(FileName)
    If Not (f.Attributes And ReadOnly) Then
       f.Attributes = f.Attributes XOR ReadOnly 
    End If
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
' Purpose   : Gets the AltStartupPath, by looking in the Registry
'             There is some chance that this returns the wrong result - e.g. on a PC
'             where Office 16.0 was previously installed (leaving data in the Registry)
'             but the version of Office used is Office 15.0 
'---------------------------------------------------------------------------------------
Function GetAltStartupPath() 'App)
    GetAltStartupPath = RegistryRead("HKEY_CURRENT_USER\Software\Microsoft\Office\" & _
    gOfficeVersion & "\Excel\Options\AltStartup", "Not found")
End Function

'---------------------------------------------------------------------------------------
' Procedure : SetAltStartupPath
' Purpose   : Sets the AltStartupPath, by looking in the Registry. See caution for 
'             GetAltStartupPath
'---------------------------------------------------------------------------------------
Function SetAltStartupPath(Path) '(App,Path)
    RegistryWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\" & gOfficeVersion & _
    "\Excel\Options\AltStartup", Path
End Function

'---------------------------------------------------------------------------------------
' Procedure : RegistryRead
' Purpose   : Read a value from the Registry
' https://msdn.microsoft.com/en-us/library/x05fawxd(v=vs.84).aspx
'---------------------------------------------------------------------------------------
Function RegistryRead(RegKey, DefaultValue)
    Dim myWS
    RegistryRead = DefaultValue
    Set myWS = CreateObject("WScript.Shell")
    On Error Resume Next
    RegistryRead = myWS.RegRead(RegKey) 
    Exit Function
End Function

'---------------------------------------------------------------------------------------
' Procedure : RegistryWrite
' Purpose   : Write to the Registry
' https://msdn.microsoft.com/en-us/library/yfdfhz1b(v=vs.84).aspx
'---------------------------------------------------------------------------------------
Function RegistryWrite(RegKey, NewValue)
    Dim myWS
    Set myWS = CreateObject("WScript.Shell")
    myWS.RegWrite RegKey, NewValue, "REG_SZ"
End Function

'---------------------------------------------------------------------------------------
' Procedure : sRegistryKeyExists
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
    Dim myWS
    Set myWS = CreateObject("WScript.Shell")
    myWS.regDelete RegKey
End Function

'Apparently VBScript has no in-line if. So create one, but note that unlike
'VB6/VBA's Iif this one does not evaluate both truepart and falsepart.
Function IIf( expr, truepart, falsepart )
    If expr Then
        IIf = truepart
    Else
        IIf = falsepart
    End If
End Function

Function Environ(Expression)
	Dim WshShell
	Set WshShell = CreateObject("WScript.Shell")
	Environ = WshShell.ExpandEnvironmentStrings("%" & Expression & "%")
End Function

Sub InstallExcelAddin(AddinFullName, WithSlashR)
    Dim RegKeyBranch
    Dim RegKeyLeaf
    Dim i
    Dim Found
    Dim NumAddins
    Dim RegValue

    RegKeyBranch = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & _
                    gOfficeVersion & "\Excel\Options\"
    i = 0
    Do
        i = i + 1
        RegKeyLeaf = "OPEN" & IIf(i > 1, CStr(i - 1), "")
        If RegistryKeyExists(RegKeyBranch & RegKeyLeaf) Then
            NumAddins = NumAddins + 1
            RegValue = RegistryRead(RegKeyBranch & RegKeyLeaf, "")
            Found = InStr(LCase(RegValue), LCase(AddinFullName)) > 0
            If Found Then Exit Sub
        Else
            Exit Do
        End If
    Loop

    RegKeyLeaf = "OPEN" & IIf(NumAddins > 0, CStr(NumAddins), "")
    'I can't discover what is the significance of the /R that appears in the Registry for
    'some addins but not for others...
    If WithSlashR Then
        RegValue = "/R """ & AddinFullName & """"
    Else
        RegValue = AddinFullName
    End If
    RegistryWrite RegKeyBranch + RegKeyLeaf, RegValue
End Sub

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : GetOfficeVersionAndBitness
' Author     : Philip Swannell
' Date       : 14-Dec-2021
' Notes      : Previously was trying to determine office version and bitness by reading the registry, which turns out to
'              be hard to do, for example when a PC has had various versions of Office installed. So reverted to 
'              launching Excel via CreateObject.
'              I posted something along these lines at
'              https://stackoverflow.com/questions/2203980/detect-whether-office-is-32bit-or-64bit-via-the-registry
' -----------------------------------------------------------------------------------------------------------------------
Function GetOfficeVersionAndBitness(OfficeVersion,OfficeBitness)
    Dim Excel, EN

    On Error Resume Next
    Set Excel = CreateObject("Excel.Application")
    EN = Err.Number
    Excel.Visible = False
    On Error GoTo 0

    If EN = 0 Then
        If InStr(Excel.OperatingSystem,"64") > 0 Then
            OfficeBitness = 64
            OfficeVersion = Excel.Version
        Else
            OfficeBitness = 32
            OfficeVersion = Excel.Version
        End if
        Excel.Quit
    Else
        OfficeBitness = 0
        OfficeVersion = "Office Not found"
    End If

    Set Excel = Nothing
End Function

' -----------------------------------------------------------------------------------------------------------------------
' Procedure  : DeleteExcelAddinFromRegistry
' Purpose    : Edits the Windows Registry to ensure that excel does not load a particular addin. Will not work if the addin
'              is located in the AltStartUp path
' Parameters :
'  AddinName:  The file name of the addin e.g. "ExcelDna.IntelliSense64.xll" can include the path if we want to remove an
'              addin only if it's currently being loaded from the "wrong" location.
' -----------------------------------------------------------------------------------------------------------------------
Sub DeleteExcelAddinFromRegistry(AddinName)
    Dim RegKey
    Dim AllKeys()
    Dim i, j
    Dim RegKeyLeaf
    Dim NumAddins
    Dim Found

    RegKey = "HKEY_CURRENT_USER\Software\Microsoft\Office\" & gOfficeVersion & "\Excel\Options\"
    i = 0
    Do
        i = i + 1
        RegKeyLeaf = "OPEN" & IIf(i > 1, CStr(i - 1), "")
        If RegistryKeyExists(RegKey & RegKeyLeaf) Then
            NumAddins = NumAddins + 1
        Else
            Exit Do
        End If
    Loop

    Found = False
    
    ReDim AllKeys(NumAddins - 1, 1) 'VBScript has base 0 so that's two columns
    For i = 0 To NumAddins - 1
        RegKeyLeaf = "OPEN" & IIf(i > 0, CStr(i), "")
        AllKeys(i, 0) = RegKeyLeaf
        AllKeys(i, 1) = RegistryRead(RegKey & RegKeyLeaf, "")
        If InStr(LCase(AllKeys(i, 1)), LCase(AddinName)) > 0 Then
            Found = True
        End If
    Next

    If Not Found Then Exit Sub

    For i = 0 To NumAddins - 1
        RegistryDelete RegKey & AllKeys(i, 0)
    Next

    j = 0
    For i = 0 To NumAddins - 1
        If InStr(LCase(AllKeys(i, 1)), LCase(AddinName)) = 0 Then
            j = j + 1
            RegKeyLeaf = "OPEN" & IIf(j > 1, CStr(j - 1), "")
            RegistryWrite RegKey & RegKeyLeaf, AllKeys(i, 1)
            'Debug.Print "Writing " + RegKey & RegKeyLeaf + " with value " + AllKeys(i, 1)
        End If
    Next
End Sub

'*******************************************************************************************
'Effective start of this VBScript. Note elevating to admin as per 
'http://www.winhelponline.com/blog/vbscripts-and-uac-elevation/
'We install to C:\ProgramData see
'https://stackoverflow.com/questions/22107812/privileges-owner-issue-when-writing-in-c-programdata
'*******************************************************************************************
If (WScript.Arguments.length = 0) And ElevateToAdmin Then
   Dim objShell, ThisFileName
   Set objShell = CreateObject("Shell.Application")
   'Pass a bogus argument, say [ uac]
   ThisFileName = WScript.ScriptFullName
   objShell.ShellExecute "wscript.exe", Chr(34) & _
      ThisFileName & Chr(34) & " uac", "", "runas", 1
Else
    Dim gOfficeVersion, gOfficeBitness
    Set myWS = CreateObject("WScript.Shell")

    gErrorsEncountered = False
    If Not GIFRecordingMode Then
        'CheckExcel must be called BEFORE GetOfficeVersionAndBitness
        CheckExcel
    End If

    GetOfficeVersionAndBitness gOfficeVersion, gOfficeBitness

    GIFRecordingMode = FileExists(GIFRecordingFlagFile)

    If gOfficeVersion = "Office Not found" Then
    Prompt = "Installation cannot proceed because no version of Microsoft Office has " & _
               "been detected on this PC." & vblf  & vblf & _
               "The script attempts to detect the installed versions of Office by " & _
               "executing the code `CreateObject(""Excel.Application"")` which should " & _
               "launch Excel so that its version can be determined." & _ 
               vblf & vblf & "However, that didn't work. So it seems you need to " & _
               "install Microsoft Office before installing JuliaExcel."

        MsgBox Prompt,vbCritical,MsgBoxTitleBad
        WScript.Quit
    End If    

    AddinsSource = WScript.ScriptFullName
    AddinsSource = Left(AddinsSource, InStrRev(AddinsSource, "\") - 1)
    AddinsSource = Left(AddinsSource, InStrRev(AddinsSource, "\"))
    AddinsSource = AddinsSource & "workbooks\"
    IntellisenseSource = WScript.ScriptFullName
    IntellisenseSource = Left(IntellisenseSource, InStrRev(IntellisenseSource, "\") - 1)
    IntellisenseSource = Left(IntellisenseSource, InStrRev(IntellisenseSource, "\"))
    IntellisenseSource = IntellisenseSource & "ExcelDNA\"

        Select Case gOfficeBitness
        Case 32
            IntellisenseName = "ExcelDna.IntelliSense.xll"
            InstallIntellisense = True
        Case 64
            IntellisenseName = "ExcelDna.IntelliSense64.xll"
            InstallIntellisense = True
        Case Else
            InstallIntellisense = False
        End Select

    Prompt = "This will install JuliaExcel by copying two files from: " & vbLf & vblf & _
        AddinsSource & AddinName  & vbLf & _
        IntellisenseSource & IntellisenseName & vbLf & vbLf & _
        "to:" & vblf & vblf & _
        AddinsDest & AddinName & vblf & _ 
        AddinsDest & IntellisenseName & vbLf & vbLf & _
        "and making them both be Excel add-ins," & vblf & _
        "via Excel > File > Options > Add-ins > Excel Add-ins." & vblf & vblf & _
        "Do you wish to continue?" & vblf  & vblf & _
        "More information at:" & vblf & Website

    If MsgBox(Prompt, vbYesNo + vbQuestion, MsgBoxTitle) <> vbYes Then WScript.Quit

    ForceFolderToExist AddinsDest

    If not GIFRecordingMode Then

        CopyNamedFiles AddinsSource, AddinsDest, AddinName, True
        MakeFileReadOnly AddinsDest & AddinName
        DeleteExcelAddinFromRegistry AddinName
        InstallExcelAddin AddinsDest & AddinName, True

        If InstallIntellisense Then
            CopyNamedFiles IntellisenseSource, AddinsDest, IntellisenseName, True
            DeleteExcelAddinFromRegistry IntellisenseName
            InstallExcelAddin AddinsDest & IntellisenseName, True
        End If
    End If

    If gErrorsEncountered Then
        Prompt = "The install script has finished, but errors were encountered, " & _
                 "which may mean the software will not work correctly." & vblf & vblf & _
                 Website
        MsgBox Prompt, vbOKOnly + vbCritical, MsgBoxTitleBad
    Else
        Prompt = "JuliaExcel is installed, and its functions such as JuliaEval and " & _
                 "JuliaCall will be available the next time you start Excel." & vblf & _
                 vblf & Website
        MsgBox Prompt, vbOKOnly + vbInformation, MsgBoxTitle
    End If

    'Don't record this bit. Have this warning after forgetting about the flag file!
    If GIFRecordingMode Then 
        MsgBox "That previous message was false. The installation was blocked by the " & _
                "existence of file '" & GIFRecordingFlagFile & "'",vbOKOnly + vbCritical, _
                MsgBoxTitleBad
    End If

    WScript.Quit
End If
