# Push-VbaModules.ps1
# Pushes .bas/.cls/.frm files from vba\JuliaExcel.xlam\VBA\ on disk into the running Excel
# workbook, replacing any existing modules of the same name.
# Before overwriting, backs up the workbook's current VBA to .vba-backups\pre-push-<timestamp>\.
#
# Prerequisites:
#   Excel Trust Center -> Trust Center Settings -> Macro Settings ->
#   [x] Trust access to the VBA project object model
#
# Run from VSCode with Ctrl+Shift+B (wired as the default build task in .vscode/tasks.json).

$ErrorActionPreference = "Stop"

# Locate Excel
try {
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    Write-Error "Excel is not running. Open workbooks\JuliaExcel.xlam in Excel first."
    exit 1
}

# Locate workbook by name, then verify full path.
# Note: when the workbook is loaded as an installed add-in (IsAddin = True), it is excluded
# from Workbooks.Count and foreach enumeration, but Workbooks.Item(name) still finds it.
$xlPath = (Resolve-Path (Join-Path $PSScriptRoot "..\workbooks\JuliaExcel.xlam")).Path
$xlName = [IO.Path]::GetFileName($xlPath)
$book = $null
try {
    $wb = $excel.Workbooks.Item($xlName)
    if ($wb.FullName -ieq $xlPath) { $book = $wb }
} catch { }
if ($null -eq $book) {
    Write-Error "workbooks\JuliaExcel.xlam is not open in Excel. Open it from:`n  $xlPath"
    exit 1
}
Write-Host "Found workbook: $($book.FullName)"

$proj = $book.VBProject
if ($null -eq $proj -or $null -eq $proj.VBComponents) {
    Write-Error @"
Cannot access the VBA project object model.
In Excel: File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings
Check: Trust access to the VBA project object model
"@
    exit 1
}

# Find source files on disk
$basDir = (Resolve-Path (Join-Path $PSScriptRoot "..\vba\JuliaExcel.xlam\VBA")).Path
$files = Get-ChildItem "$basDir\*.bas", "$basDir\*.cls", "$basDir\*.frm" -ErrorAction SilentlyContinue
if ($files.Count -eq 0) {
    Write-Warning "No .bas/.cls/.frm files found in $basDir"
    exit 0
}

# Confirm
$confirm = Read-Host "Push $($files.Count) module(s) from disk to Excel workbook? [Y/N]"
if ($confirm -notmatch '^[Yy]') { Write-Host "Cancelled."; exit 0 }

# Backup workbook's current VBA to .vba-backups\pre-push-<timestamp>\
$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$backupDir = Join-Path $projectRoot ".vba-backups\pre-push-$timestamp"
New-Item -ItemType Directory -Path $backupDir -Force | Out-Null

foreach ($comp in $proj.VBComponents) {
    if ($comp.Type -notin @(1, 2, 3)) { continue }   # skip document modules (Sheet1, ThisWorkbook etc.)
    $ext = switch ($comp.Type) {
        1 { ".bas" }
        2 { ".cls" }
        3 { ".frm" }
    }
    $comp.Export($(Join-Path $backupDir "$($comp.Name)$ext"))
}
Write-Host "Backup saved to $backupDir"
Write-Host ""

# Push: import each file into the workbook
foreach ($file in $files) {
    $name = [IO.Path]::GetFileNameWithoutExtension($file.Name)
    Write-Host "  Pushing $name ..."

    # Refresh project reference each iteration in case the COM object goes stale
    $proj = $book.VBProject
    if ($null -eq $proj -or $null -eq $proj.VBComponents) {
        Write-Error "VBProject/VBComponents became null while processing $name"
        exit 1
    }

    try {
        $proj.VBComponents.Remove($proj.VBComponents.Item($name))
    } catch { }
    $proj.VBComponents.Import($file.FullName) | Out-Null
    Write-Host "  Pushed $name"
}

Write-Host ""

# Save the workbook so the pushed modules persist to disk.
# Workbook.Save throws (or silently no-ops, depending on Excel version) unless IsAddin is True,
# so toggle it on if necessary and restore the original value afterward - leaving IsAddin as we
# found it, e.g. False while the workbook window is shown for development.
$wasAddin = $book.IsAddin
try {
    if (-not $wasAddin) { $book.IsAddin = $true }
    $book.Save()
    Write-Host "Workbook saved: $($book.FullName)"
} finally {
    if ($book.IsAddin -ne $wasAddin) { $book.IsAddin = $wasAddin }
}

Write-Host ""
Write-Host "Done - $($files.Count) module(s) pushed from disk to Excel and saved."
