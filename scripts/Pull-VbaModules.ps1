# Pull-VbaModules.ps1
# Saves the target workbook and exports its VBA to disk, by delegating to SolumAddin.xlam's
# SaveWorkbookAndExportForGit method via a CallSaveAddInAndExportVBA wrapper function that must
# exist somewhere in the target workbook's own VBA project (see modUtils.bas in JuliaExcel.xlam for
# the canonical example - any workbook wanting Pull support needs its own copy of that thin
# wrapper, since "ThisWorkbook" inside it only resolves correctly when the code lives in that
# workbook). SaveWorkbookAndExportForGit works for both .xlam add-ins and ordinary .xlsm workbooks
# (its predecessor, SaveAddInAndExportVBA, was .xlam-only, which is why this script and a separate
# Pull-VbaModules-Simple.ps1 used to be needed side by side - now this one script covers both).
#
# Prerequisites:
#   - Excel Trust Center -> Trust Center Settings -> Macro Settings ->
#     [x] Trust access to the VBA project object model
#   - SolumAddin.xlam and the target workbook both open in Excel (either as ordinary workbooks or
#     as installed add-ins - IsAddin = True is fine, see note below).
#
# Run from VSCode via Terminal -> Run Task -> "VBA: Pull from Excel", which invokes this with no
# arguments, i.e. against workbooks\JuliaExcel.xlam.

param(
    [string]$WorkbookName = "JuliaExcel.xlam"
)

$ErrorActionPreference = "Stop"

# Locate Excel
try {
    $excel = [Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
} catch {
    Write-Error "Excel is not running. Open workbooks\$WorkbookName in Excel first."
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

$xlPath = (Resolve-Path (Join-Path $PSScriptRoot "..\workbooks\$WorkbookName")).Path
$targetBook = Find-Workbook -Name $WorkbookName -ExpectedFullName $xlPath
if ($null -eq $targetBook) {
    Write-Error "workbooks\$WorkbookName is not open in Excel. Open it from:`n  $xlPath"
    exit 1
}
Write-Host "Found workbook: $($targetBook.FullName)"

# SolumAddin.xlam lives outside this repo, so there's no known path to verify against - just
# confirm Excel has a workbook of that name open.
$solumBook = Find-Workbook -Name "SolumAddin.xlam"
if ($null -eq $solumBook) {
    Write-Error "SolumAddin.xlam is not open in Excel. Open (or install) it first."
    exit 1
}
Write-Host "Found workbook: $($solumBook.FullName)"

# Delegate to SolumAddin.xlam via a wrapper function (CallSaveAddInAndExportVBA) that must exist
# somewhere in the target workbook's own VBA project, which passes ThisWorkbook (i.e. the target
# workbook itself) through to SaveWorkbookAndExportForGit. The wrapper exists because when a macro
# invoked via Application.Run raises an unhandled error, Excel's automation interface does not
# propagate the error's Description back to a COM caller like this script - only an opaque HRESULT.
# The wrapper catches the error in VBA and returns it as a string instead (in this project's usual
# "#FunctionName (line N): message!" error-string format), so we detect failure by inspecting the
# returned string rather than by catching an exception here. Called unqualified (no module prefix)
# since each workbook's copy of this wrapper is the only procedure of that name in its project.
Write-Host ""
Write-Host "Calling SaveWorkbookAndExportForGit for $($targetBook.Name) ..."
$result = $excel.Run("'$xlPath'!CallSaveAddInAndExportVBA")

if ($result -like "#*!") {
    Write-Error "SaveWorkbookAndExportForGit failed: $result"
    exit 1
}

Write-Host "Done - $result"
