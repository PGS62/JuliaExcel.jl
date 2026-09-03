# Installer for JuliaExcel.xlam
# PowerShell port of the now-deleted Install.vbs (see git history) - VBScript is being phased
# out of Windows, and its MsgBox dialogs render blurrily on high-DPI displays; this version uses
# WinForms message boxes, which don't have that problem.
#
# Targets Windows PowerShell 5.1 (ships on every Windows install) deliberately - no PS7-only
# syntax (ternary operator, null-coalescing, etc.) so nothing extra needs to be installed.
#
# To debug: open in VS Code with the PowerShell extension and set breakpoints, or run
# `powershell.exe -File Install.ps1` directly from a terminal.

Set-StrictMode -Version Latest

$AddinName = "JuliaExcel.xlam"
$Website = "https://github.com/PGS62/JuliaExcel.jl"
$WebsiteIntellisense = "https://github.com/Excel-DNA/IntelliSense"
$GIFRecordingFlagFile = "C:\Temp\RecordingGIF.tmp"
$MsgBoxTitle = "Install JuliaExcel"
$MsgBoxTitleBad = "Install JuliaExcel - Error Encountered"
# $AddinsDest = "C:\ProgramData\JuliaExcel\"  # Would need Admin rights to write to
$AddinsDest = "C:\Users\Public\JuliaExcel\"   # Does not need admin rights
$ElevateToAdmin = $false                      # Since writing to c:\Users\Public does not need admin rights

# Putting the add-in in the same folder for all users has both advantages and
# disadvantages:
# Advantage: Avoid "Excel Link Hell" caused by the fact that workbooks store the
#     absolute address of files to which they link (unless the file is in the same
#     folder). Causes endless problems when two users share a workbook.
# Disadvantage: Two different users of the same PC would share copy of the add-in and
#     thus be forced to use the same version of the add-in, though they don't both have
#     to have the addin installed since that's controlled via the registry, which _is_
#     user specific.

Add-Type -AssemblyName System.Windows.Forms

# powershell.exe has no DPI-aware manifest, so without this, Windows applies its legacy
# compatibility behaviour to any window the process creates: render at "normal" size, then
# bitmap-stretch the whole thing to match the display's scale factor - which is exactly what
# produces blurry, undersized-looking text on a high-DPI screen (the same root cause as the old
# VBScript dialogs, since wscript.exe isn't DPI-aware either). Declaring Per-Monitor-V2 awareness
# (Windows 10 1703+) makes WinForms render the dialog natively at the correct size instead;
# SetProcessDPIAware (Vista+) is a coarser but more broadly compatible fallback if the newer API
# call fails for any reason.
Add-Type -Name NativeDpi -Namespace JuliaExcelInstaller -MemberDefinition @'
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool SetProcessDpiAwarenessContext(IntPtr value);
    [System.Runtime.InteropServices.DllImport("user32.dll")]
    public static extern bool SetProcessDPIAware();
'@
try {
    # DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4
    $dpiOk = [JuliaExcelInstaller.NativeDpi]::SetProcessDpiAwarenessContext([IntPtr]::new(-4))
    if (-not $dpiOk) { [JuliaExcelInstaller.NativeDpi]::SetProcessDPIAware() | Out-Null }
} catch {
    try { [JuliaExcelInstaller.NativeDpi]::SetProcessDPIAware() | Out-Null } catch {}
}

$script:gErrorsEncountered = $false

# -----------------------------------------------------------------------------------------------
# Function  : Show-MsgBox
# Purpose   : Thin wrapper around a WinForms MessageBox, standing in for VBScript's MsgBox.
#             Unlike the classic Win32 MsgBox VBScript hosts, this renders crisply on high-DPI
#             displays. Buttons: 'OK', 'OKCancel', 'YesNo'. Icon: 'Info', 'Warning', 'Error',
#             'Question'. Returns a [System.Windows.Forms.DialogResult].
# -----------------------------------------------------------------------------------------------
function Show-MsgBox {
    param(
        [string]$Prompt,
        [string]$Title,
        [ValidateSet('OK', 'OKCancel', 'YesNo')][string]$Buttons = 'OK',
        [ValidateSet('Info', 'Warning', 'Error', 'Question')][string]$Icon = 'Info'
    )
    $buttonsEnum = switch ($Buttons) {
        'OK' { [System.Windows.Forms.MessageBoxButtons]::OK }
        'OKCancel' { [System.Windows.Forms.MessageBoxButtons]::OKCancel }
        'YesNo' { [System.Windows.Forms.MessageBoxButtons]::YesNo }
    }
    $iconEnum = switch ($Icon) {
        'Info' { [System.Windows.Forms.MessageBoxIcon]::Information }
        'Warning' { [System.Windows.Forms.MessageBoxIcon]::Warning }
        'Error' { [System.Windows.Forms.MessageBoxIcon]::Error }
        'Question' { [System.Windows.Forms.MessageBoxIcon]::Question }
    }
    return [System.Windows.Forms.MessageBox]::Show($Prompt, $Title, $buttonsEnum, $iconEnum)
}

function Test-ExcelRunning {
    return [bool](Get-Process -Name "EXCEL" -ErrorAction SilentlyContinue)
}

# -----------------------------------------------------------------------------------------------
# Function  : Confirm-ExcelClosed
# Purpose   : Invite user to shut down Excel, returns once the user does so or exits the
#             script if they decline.
# -----------------------------------------------------------------------------------------------
function Confirm-ExcelClosed {
    $FriendlyName = "Microsoft Excel"
    if (-not (Test-ExcelRunning)) { return }
    Show-MsgBox -Prompt "$FriendlyName is running. Please close it and then click OK to continue." `
        -Title $MsgBoxTitle -Buttons OK -Icon Warning | Out-Null
    while ($true) {
        if (Test-ExcelRunning) {
            $result = Show-MsgBox -Prompt ("$FriendlyName is still running. Please close it and then click OK to continue, or click Cancel to quit.`n`n" + `
                    "Can't see $FriendlyName`? Use Windows Task Manager to check if $FriendlyName is running as a ""background process"", and if so use the right-click menu to ""End task"" the process.") `
                -Title $MsgBoxTitle -Buttons OKCancel -Icon Warning
            if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
                exit 1
            }
        } else {
            return
        }
    }
}

function Copy-AddinFile {
    param(
        [string]$SourceFolder,
        [string]$DestFolder,
        [string]$FileName,
        [bool]$ThrowErrorIfNoSourceFile
    )
    $sourcePath = Join-Path $SourceFolder $FileName
    $destPath = Join-Path $DestFolder $FileName

    if (-not (Test-Path -Path $sourcePath -PathType Leaf)) {
        if ($ThrowErrorIfNoSourceFile) {
            $script:gErrorsEncountered = $true
            Show-MsgBox -Prompt "Cannot find file: $sourcePath" -Title $MsgBoxTitleBad -Buttons OK -Icon Warning | Out-Null
        }
        return
    }

    if (Test-Path -Path $destPath -PathType Leaf) {
        try { Set-ItemProperty -Path $destPath -Name IsReadOnly -Value $false -ErrorAction Stop } catch {}
    }

    try {
        Copy-Item -Path $sourcePath -Destination $destPath -Force -ErrorAction Stop
    } catch {
        $script:gErrorsEncountered = $true
        $errorMessage = "Failed to copy from: $sourcePath`nto: $destPath`nError: $($_.Exception.Message)"
        if ((Test-Path -Path $sourcePath -PathType Leaf) -and (Test-Path -Path $destPath -PathType Leaf)) {
            $errorMessage += "`n`nDoes another user of this PC have the file open in Excel? Check that no other users of the PC are logged in"
        }
        Show-MsgBox -Prompt $errorMessage -Title $MsgBoxTitleBad -Buttons OK -Icon Warning | Out-Null
    }
}

function Get-ExcelOptionsRegPath {
    return "HKCU:\Software\Microsoft\Office\$script:gOfficeVersion\Excel\Options"
}

function Test-RegValueExists {
    param([string]$Path, [string]$Name)
    if (-not (Test-Path -Path $Path)) { return $false }
    return $null -ne (Get-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue)
}

function Get-RegValue {
    param([string]$Path, [string]$Name, [string]$Default = "")
    if (-not (Test-RegValueExists -Path $Path -Name $Name)) { return $Default }
    return (Get-ItemProperty -Path $Path -Name $Name).$Name
}

function Set-RegValue {
    param([string]$Path, [string]$Name, [string]$Value)
    if (-not (Test-Path -Path $Path)) { New-Item -Path $Path -Force | Out-Null }
    Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type String
}

function Remove-RegValue {
    param([string]$Path, [string]$Name)
    Remove-ItemProperty -Path $Path -Name $Name -ErrorAction SilentlyContinue
}

# -----------------------------------------------------------------------------------------------
# Function  : Get-OfficeVersionAndBitness
# Purpose   : Determines Office's version and bitness by launching Excel via COM, rather than by
#             reading the registry directly, which turns out to be hard to do reliably - e.g.
#             when a PC has had various versions of Office installed over time.
#             https://stackoverflow.com/questions/2203980/detect-whether-office-is-32bit-or-64bit-via-the-registry
# -----------------------------------------------------------------------------------------------
function Get-OfficeVersionAndBitness {
    try {
        $excelApp = New-Object -ComObject Excel.Application
        try {
            $excelApp.Visible = $false
            if ($excelApp.OperatingSystem -like "*64*") { $bitness = 64 } else { $bitness = 32 }
            $version = $excelApp.Version
        } finally {
            $excelApp.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excelApp) | Out-Null
        }
        return @{ Version = $version; Bitness = $bitness }
    } catch {
        return @{ Version = "Office Not found"; Bitness = 0 }
    }
}

# -----------------------------------------------------------------------------------------------
# Function  : Install-ExcelAddin
# Purpose   : Registers an add-in to auto-load with Excel by writing to the next free
#             OPEN/OPENn value under HKCU:\...\Excel\Options - does nothing if the add-in
#             (matched by filename, case-insensitively, anywhere in an existing value) is
#             already registered.
# -----------------------------------------------------------------------------------------------
function Install-ExcelAddin {
    param([string]$AddinFullName, [bool]$WithSlashR)

    $regPath = Get-ExcelOptionsRegPath
    $i = 0
    $numAddins = 0
    while ($true) {
        $i++
        $valueName = "OPEN" + $(if ($i -gt 1) { [string]($i - 1) } else { "" })
        if (Test-RegValueExists -Path $regPath -Name $valueName) {
            $numAddins++
            $existing = Get-RegValue -Path $regPath -Name $valueName -Default ""
            if ($existing.ToLower().Contains($AddinFullName.ToLower())) { return }
        } else {
            break
        }
    }

    $valueName = "OPEN" + $(if ($numAddins -gt 0) { [string]$numAddins } else { "" })
    # I can't discover what is the significance of the /R that appears in the Registry for
    # some addins but not for others...
    if ($WithSlashR) {
        $value = "/R `"$AddinFullName`""
    } else {
        $value = $AddinFullName
    }
    Set-RegValue -Path $regPath -Name $valueName -Value $value
}

# -----------------------------------------------------------------------------------------------
# Function  : Remove-ExcelAddinFromRegistry
# Purpose   : Edits the Windows Registry to ensure that Excel does not load a particular addin.
#             Will not work if the addin is located in the AltStartUp path.
# Parameters:
#  AddinName: The file name of the addin, e.g. "ExcelDna.IntelliSense64.xll" - can include the
#             path if we want to remove an addin only if it's currently being loaded from the
#             "wrong" location.
# -----------------------------------------------------------------------------------------------
function Remove-ExcelAddinFromRegistry {
    param([string]$AddinName)

    $regPath = Get-ExcelOptionsRegPath
    $i = 0
    $numAddins = 0
    while ($true) {
        $i++
        $valueName = "OPEN" + $(if ($i -gt 1) { [string]($i - 1) } else { "" })
        if (Test-RegValueExists -Path $regPath -Name $valueName) {
            $numAddins++
        } else {
            break
        }
    }

    $allValues = @()
    $found = $false
    for ($j = 0; $j -lt $numAddins; $j++) {
        $valueName = "OPEN" + $(if ($j -gt 0) { [string]$j } else { "" })
        $value = Get-RegValue -Path $regPath -Name $valueName -Default ""
        $allValues += [PSCustomObject]@{ Name = $valueName; Value = $value }
        if ($value.ToLower().Contains($AddinName.ToLower())) { $found = $true }
    }

    if (-not $found) { return }

    foreach ($entry in $allValues) {
        Remove-RegValue -Path $regPath -Name $entry.Name
    }

    $k = 0
    foreach ($entry in $allValues) {
        if (-not $entry.Value.ToLower().Contains($AddinName.ToLower())) {
            $k++
            $valueName = "OPEN" + $(if ($k -gt 1) { [string]($k - 1) } else { "" })
            Set-RegValue -Path $regPath -Name $valueName -Value $entry.Value
        }
    }
}

#***************************************************************************************
# Effective start of this script. Note elevating to admin as per
# http://www.winhelponline.com/blog/vbscripts-and-uac-elevation/ (PowerShell equivalent:
# Start-Process -Verb RunAs). We install to C:\Users\Public - see
# https://stackoverflow.com/questions/22107812/privileges-owner-issue-when-writing-in-c-programdata
#***************************************************************************************
if ($args.Count -eq 0 -and $ElevateToAdmin) {
    Start-Process -FilePath "powershell.exe" `
        -ArgumentList @("-ExecutionPolicy", "Bypass", "-NoProfile", "-File", "`"$PSCommandPath`"", "uac") `
        -Verb RunAs
    exit 0
}

$GIFRecordingMode = Test-Path -Path $GIFRecordingFlagFile -PathType Leaf

# CheckExcel/Confirm-ExcelClosed must be called BEFORE Get-OfficeVersionAndBitness (which also
# launches Excel, via COM). Skipped entirely in GIF-recording mode, so that the installation can
# be recorded with Excel already open, rather than blocking on it - the actual file copy is
# skipped for the same reason, below.
if (-not $GIFRecordingMode) {
    Confirm-ExcelClosed
}

$officeInfo = Get-OfficeVersionAndBitness
$gOfficeVersion = $officeInfo.Version
$gOfficeBitness = $officeInfo.Bitness

if ($gOfficeVersion -eq "Office Not found") {
    $prompt = "Installation cannot proceed because no version of Microsoft Office has been " + `
        "detected on this PC.`n`n" + `
        "The script attempts to detect the installed versions of Office by launching Excel via " + `
        "COM automation (New-Object -ComObject Excel.Application), so its version can be " + `
        "determined.`n`n" + `
        "However, that didn't work. So it seems you need to install Microsoft Office before " + `
        "installing JuliaExcel."
    Show-MsgBox -Prompt $prompt -Title $MsgBoxTitleBad -Buttons OK -Icon Error | Out-Null
    exit 1
}

$repoRoot = Split-Path -Path $PSScriptRoot -Parent
$AddinsSource = Join-Path $repoRoot "workbooks"
$IntellisenseSource = Join-Path $repoRoot "ExcelDNA"

switch ($gOfficeBitness) {
    32 { $IntellisenseName = "ExcelDna.IntelliSense.xll"; $InstallIntellisense = $true }
    64 { $IntellisenseName = "ExcelDna.IntelliSense64.xll"; $InstallIntellisense = $true }
    default { $InstallIntellisense = $false }
}

$prompt = "This will install JuliaExcel by copying two files: `n`n" + `
    (Join-Path $AddinsSource $AddinName) + "`n" + `
    (Join-Path $IntellisenseSource $IntellisenseName) + "`n`n" + `
    "to:`n`n" + `
    (Join-Path $AddinsDest $AddinName) + "`n" + `
    (Join-Path $AddinsDest $IntellisenseName) + "`n`n" + `
    "and make them both Excel add-ins,`n" + `
    "via Excel > File > Options > Add-ins > Excel Add-ins.`n`n" + `
    "Do you wish to continue?`n`n`n" + `
    "$Website`n$WebsiteIntellisense"

if ((Show-MsgBox -Prompt $prompt -Title $MsgBoxTitle -Buttons YesNo -Icon Question) -ne [System.Windows.Forms.DialogResult]::Yes) {
    exit 1
}

if (-not (Test-Path -Path $AddinsDest)) { New-Item -ItemType Directory -Path $AddinsDest -Force | Out-Null }

if (-not $GIFRecordingMode) {
    Copy-AddinFile -SourceFolder $AddinsSource -DestFolder $AddinsDest -FileName $AddinName -ThrowErrorIfNoSourceFile $true
    try { Set-ItemProperty -Path (Join-Path $AddinsDest $AddinName) -Name IsReadOnly -Value $true -ErrorAction Stop } catch {}
    Remove-ExcelAddinFromRegistry -AddinName $AddinName
    Install-ExcelAddin -AddinFullName (Join-Path $AddinsDest $AddinName) -WithSlashR $true

    if ($InstallIntellisense) {
        Copy-AddinFile -SourceFolder $IntellisenseSource -DestFolder $AddinsDest -FileName $IntellisenseName -ThrowErrorIfNoSourceFile $true
        Remove-ExcelAddinFromRegistry -AddinName $IntellisenseName
        Install-ExcelAddin -AddinFullName (Join-Path $AddinsDest $IntellisenseName) -WithSlashR $true
    }
}

if ($script:gErrorsEncountered) {
    $prompt = "The install script has finished, but errors were encountered, which may mean " + `
        "the software will not work correctly.`n`n$Website"
    Show-MsgBox -Prompt $prompt -Title $MsgBoxTitleBad -Buttons OK -Icon Error | Out-Null
} else {
    $prompt = "JuliaExcel is installed, and its functions such as JuliaEval and JuliaCall will " + `
        "be available the next time you start Excel.`n`n$Website"
    Show-MsgBox -Prompt $prompt -Title $MsgBoxTitle -Buttons OK -Icon Info | Out-Null
}

if ($script:gErrorsEncountered) {
    exit 1
} else {
    exit 0
}
