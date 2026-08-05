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

try {
    $book = $excel.Workbooks.Item("JuliaExcel.xlam")
} catch {
    $book = $null
}
if ($null -eq $book) {
    Write-Error "JuliaExcel.xlam not found in Excel. Open or install it first."
    exit 1
}

$proj   = $book.VBProject
$basDir = Join-Path $PSScriptRoot "..\vba\JuliaExcel.xlam\VBA"
$basDir = (Resolve-Path $basDir).Path

$files = Get-ChildItem "$basDir\*.bas"
if ($files.Count -eq 0) {
    Write-Warning "No .bas files found in $basDir"
    exit 0
}

foreach ($file in $files) {
    $name = [IO.Path]::GetFileNameWithoutExtension($file.Name)
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
