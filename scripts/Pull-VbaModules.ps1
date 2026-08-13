# Pull-VbaModules.ps1
# Saves the JuliaExcel.xlam workbook and exports its VBA to disk, by delegating to SolumAddin.xlam's
# SaveAddInAndExportVBA method.
#
# Prerequisites:
#   - Excel Trust Center -> Trust Center Settings -> Macro Settings ->
#     [x] Trust access to the VBA project object model
#   - SolumAddin.xlam and JuliaExcel.xlam both open in Excel (either as ordinary workbooks or as
#     installed add-ins - IsAddin = True is fine, see note below).
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

# Locate a workbook by name, then optionally verify its full path.
# Note: when a workbook is loaded as an installed add-in (IsAddin = True), it is excluded from
# Workbooks.Count and foreach enumeration, but Workbooks.Item(name) still finds it.
function Find-Workbook {
    param(
        [Parameter(Mandatory)][string]$Name,
        [string]$ExpectedFullName
    )
    try {
        $wb = $excel.Workbooks.Item($Name)
    } catch {
        return $null
    }
    if ($ExpectedFullName -and $wb.FullName -ine $ExpectedFullName) {
        return $null
    }
    return $wb
}

$xlPath = (Resolve-Path (Join-Path $PSScriptRoot "..\workbooks\JuliaExcel.xlam")).Path
$juliaBook = Find-Workbook -Name "JuliaExcel.xlam" -ExpectedFullName $xlPath
if ($null -eq $juliaBook) {
    Write-Error "workbooks\JuliaExcel.xlam is not open in Excel. Open it from:`n  $xlPath"
    exit 1
}
Write-Host "Found workbook: $($juliaBook.FullName)"

# SolumAddin.xlam lives outside this repo, so there's no known path to verify against - just
# confirm Excel has a workbook of that name open.
$solumBook = Find-Workbook -Name "SolumAddin.xlam"
if ($null -eq $solumBook) {
    Write-Error "SolumAddin.xlam is not open in Excel. Open (or install) it first."
    exit 1
}
Write-Host "Found workbook: $($solumBook.FullName)"

# Delegate to SolumAddin.xlam via a wrapper function in JuliaExcel.xlam (modUtils.CallSaveAddInAndExportVBA),
# which passes ThisWorkbook (i.e. JuliaExcel.xlam itself) through to SaveAddInAndExportVBA.
# The wrapper exists because when a macro invoked via Application.Run raises an unhandled error,
# Excel's automation interface does not propagate the error's Description back to a COM caller like
# this script - only an opaque HRESULT. The wrapper catches the error in VBA and returns it as a
# string instead (in this project's usual "#FunctionName (line N): message!" error-string format),
# so we detect failure by inspecting the returned string rather than by catching an exception here.
Write-Host ""
Write-Host "Calling SaveAddInAndExportVBA for $($juliaBook.Name) ..."
$result = $excel.Run("'$xlPath'!modUtils.CallSaveAddInAndExportVBA")

if ($result -like "#*!") {
    Write-Error "SaveAddInAndExportVBA failed: $result"
    exit 1
}

Write-Host "Done - $result"
