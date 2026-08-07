# Pull-VbaModules.ps1
# Pulls VBA modules from the running Excel workbook to vba\JuliaExcel.xlam\VBA\ on disk,
# replacing existing files.
# Before overwriting, backs up the current disk files to .vba-backups\pre-pull-<timestamp>\.
#
# Prerequisites:
#   Excel Trust Center -> Trust Center Settings -> Macro Settings ->
#   [x] Trust access to the VBA project object model
#
# Run from VSCode via Terminal -> Run Task -> "VBA: Pull from Excel".

$ErrorActionPreference = "Stop"

# Locate Excel
try {
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    Write-Error "Excel is not running. Open workbooks\JuliaExcel.xlam in Excel first."
    exit 1
}

# Locate workbook by full path (avoids matching an installed add-in of the same name)
$xlPath = (Resolve-Path (Join-Path $PSScriptRoot "..\workbooks\JuliaExcel.xlam")).Path
$book = $null
foreach ($wb in $excel.Workbooks) {
    if ($wb.FullName -ieq $xlPath) { $book = $wb; break }
}
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

# Collect exportable modules (skip document modules: Sheet1, ThisWorkbook etc.)
$components = @($proj.VBComponents | Where-Object { $_.Type -in @(1, 2, 3) })
if ($components.Count -eq 0) {
    Write-Warning "No exportable modules found in the workbook."
    exit 0
}

# Confirm
$confirm = Read-Host "Pull $($components.Count) module(s) from Excel to disk? [Y/N]"
if ($confirm -notmatch '^[Yy]') { Write-Host "Cancelled."; exit 0 }

# Backup current disk files to .vba-backups\pre-pull-<timestamp>\
$basDir = (Resolve-Path (Join-Path $PSScriptRoot "..\vba\JuliaExcel.xlam\VBA")).Path
$projectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..")).Path
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$backupDir = Join-Path $projectRoot ".vba-backups\pre-pull-$timestamp"
New-Item -ItemType Directory -Path $backupDir -Force | Out-Null

$existingFiles = Get-ChildItem "$basDir\*.bas", "$basDir\*.cls", "$basDir\*.frm", "$basDir\*.frx" -ErrorAction SilentlyContinue
foreach ($f in $existingFiles) {
    Copy-Item $f.FullName -Destination $backupDir
}
Write-Host "Backup saved to $backupDir"
Write-Host ""

# Pull: export each module from the workbook to disk
foreach ($comp in $components) {
    $ext = switch ($comp.Type) {
        1 { ".bas" }
        2 { ".cls" }
        3 { ".frm" }
    }
    $outPath = Join-Path $basDir "$($comp.Name)$ext"
    Write-Host "  Pulling $($comp.Name) ..."
    $comp.Export($outPath)
    Write-Host "  Pulled $($comp.Name)"
}

Write-Host ""
Write-Host "Done - $($components.Count) module(s) pulled from Excel to disk."
