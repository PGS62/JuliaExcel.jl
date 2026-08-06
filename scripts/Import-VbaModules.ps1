# Import-VbaModules.ps1
# Imports all .bas files from vba\JuliaExcel.xlam\VBA\ into the running JuliaExcel
# workbook/addin via COM automation.
#
# Prerequisites:
#   Excel Trust Center -> Trust Center Settings -> Macro Settings ->
#   [x] Trust access to the VBA project object model
#
# Run from VSCode with Ctrl+Shift+B (wired as the default build task in .vscode/tasks.json).

$ErrorActionPreference = "Stop"

try {
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    Write-Error "Excel is not running. Open JuliaExcel.xlam in Excel first."
    exit 1
}

# Locate workbooks\JuliaExcel.xlam by its full path so we always get the editable
# copy, not an installed add-in whose VBProject is inaccessible.
$xlPath = (Resolve-Path (Join-Path $PSScriptRoot "..\workbooks\JuliaExcel.xlam")).Path

$book = $null
foreach ($wb in $excel.Workbooks) {
    if ($wb.FullName -ieq $xlPath) {
        $book = $wb
        break
    }
}
if ($null -eq $book) {
    Write-Error "workbooks\JuliaExcel.xlam is not open in Excel. Open it from:`n  $xlPath"
    exit 1
}
Write-Host "Found workbook: $($book.FullName)"

$proj   = $book.VBProject
if ($null -eq $proj -or $null -eq $proj.VBComponents) {
    Write-Error @"
Cannot access the VBA project object model.
In Excel: File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings
Check: Trust access to the VBA project object model
"@
    exit 1
}

$basDir = Join-Path $PSScriptRoot "..\vba\JuliaExcel.xlam\VBA"
$basDir = (Resolve-Path $basDir).Path

$files = Get-ChildItem "$basDir\*.bas"
if ($files.Count -eq 0) {
    Write-Warning "No .bas files found in $basDir"
    exit 0
}

foreach ($file in $files) {
    $name = [IO.Path]::GetFileNameWithoutExtension($file.Name)
    Write-Host "  Importing $name ..."

    # Refresh the project reference each iteration in case the COM object goes stale.
    $proj = $book.VBProject
    if ($null -eq $proj -or $null -eq $proj.VBComponents) {
        Write-Error "VBProject/VBComponents became null while processing $name"
        exit 1
    }

    try {
        $proj.VBComponents.Remove($proj.VBComponents.Item($name))
    } catch {
        # Module not present yet - that's fine
    }
    $proj.VBComponents.Import($file.FullName) | Out-Null
    Write-Host "  Imported $name"
}

Write-Host ""
Write-Host "Done - $($files.Count) module(s) imported from $basDir"
