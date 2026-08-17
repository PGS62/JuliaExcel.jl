# Run-VbaTests.ps1
# Runs the VBA test suite (modTest.RunTests) in the currently open JuliaExcel.xlam workbook and
# reports pass/fail. Exits 0 if all tests passed, 1 otherwise (including if Excel/the workbook
# can't be found), so this can be used as a VS Code task or from a script.
#
# Prerequisites:
#   - workbooks\JuliaExcel.xlam already open in Excel (either as an ordinary workbook or as an
#     installed add-in - IsAddin = True is fine, see note below).
#   - Excel Trust Center -> Trust Center Settings -> Macro Settings ->
#     [x] Trust access to the VBA project object model
#
# Run from VSCode via Terminal -> Run Task -> "VBA: Run Tests".

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

Write-Host ""
Write-Host "Running RunTests (SilentMode) ..."
# SilentMode:=True suppresses the MsgBox summary and Function Wizard checks that would otherwise
# block this script waiting for user input; results still go to the VBA Immediate window as usual.
$result = $excel.Run("'$xlPath'!modTest.RunTests", $true)

Write-Host ""
if ($result) {
    Write-Host "RunTests: ALL TESTS PASSED"
    exit 0
} else {
    Write-Host "RunTests: ONE OR MORE TESTS FAILED - see the VBA Immediate window for details"
    exit 1
}
