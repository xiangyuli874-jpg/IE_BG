[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$toolRoot = $PSScriptRoot
$repoRoot = Split-Path (Split-Path $toolRoot -Parent) -Parent
$outputRoot = Join-Path $repoRoot 'outputs\man_machine_analysis'
$workbookPath = (Get-ChildItem -LiteralPath $outputRoot -Filter '*_v1.xlsm' |
    Select-Object -First 1 -ExpandProperty FullName)
$requiredTables = @{
    1 = @('tblParameters', 'tblPeople', 'tblDevices', 'tblMoveTime')
    2 = @('tblSteps')
    3 = @('tblSchedule')
}
$formulaErrors = @('#REF!', '#DIV/0!', '#VALUE!', '#NAME?', '#N/A')

function Release-ComObject {
    param([object]$ComObject)
    if ($null -ne $ComObject -and
            [Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
    }
}

if (-not (Test-Path -LiteralPath $workbookPath)) {
    throw "Workbook not found: $workbookPath"
}

$excel = $null
$books = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $books = $excel.Workbooks
    $book = $books.Open($workbookPath)

    if ($book.FileFormat -ne 52) { throw "Expected XLSM format 52, got $($book.FileFormat)" }
    if ($book.Worksheets.Count -ne 5) { throw "Expected 5 worksheets, got $($book.Worksheets.Count)" }
    foreach ($sheetIndex in $requiredTables.Keys) {
        $sheet = $null
        try {
            $sheet = $book.Worksheets.Item([int]$sheetIndex)
            foreach ($tableName in $requiredTables[$sheetIndex]) {
                $table = $null
                try {
                    $table = $sheet.ListObjects.Item($tableName)
                    if ($null -eq $table) { throw "Missing table at sheet $sheetIndex`: $tableName" }
                }
                finally { Release-ComObject $table }
            }
        }
        finally { Release-ComObject $sheet }
    }

    $selfTests = $excel.Run("'$($book.Name)'!RunAllSelfTests",
        (Join-Path $toolRoot 'fixtures\baseline_1p3m.csv'))
    if ([string]$selfTests -like 'SELF_TESTS FAIL*') { throw [string]$selfTests }

    $excel.Run("'$($book.Name)'!CmdBuildInitial")
    $excel.Run("'$($book.Name)'!CmdOptimize")

    $scheduleSheet = $null
    $ganttSheet = $null
    $vbProject = $null
    try {
        $scheduleSheet = $book.Worksheets.Item(3)
        $ganttSheet = $book.Worksheets.Item(4)
        $cycle = [double]$scheduleSheet.Range('A4').Value2
        $takt = [double]$scheduleSheet.Range('B4').Value2
        $rowCount = $scheduleSheet.ListObjects.Item('tblSchedule').ListRows.Count
        $chartCount = $ganttSheet.ChartObjects().Count
        $buttonCount = $scheduleSheet.Buttons().Count
        if ([Math]::Abs($cycle - 56.0) -gt 0.01) { throw "Baseline cycle expected 56.0, got $cycle" }
        if ([Math]::Abs($takt - (56.0 / 3.0)) -gt 0.01) { throw "Baseline takt expected 18.67, got $takt" }
        if ($rowCount -ne 45) { throw "Expected 45 scheduled tasks, got $rowCount" }
        if ($chartCount -lt 1) { throw 'Gantt chart is missing' }
        if ($buttonCount -lt 8) { throw "Expected at least 8 schedule buttons, got $buttonCount" }
        $vbProject = $book.VBProject
        if ($vbProject.VBComponents.Count -lt 10) { throw 'VBA project is incomplete' }
    }
    finally {
        Release-ComObject $vbProject
        Release-ComObject $ganttSheet
        Release-ComObject $scheduleSheet
    }

    $errorCount = 0
    foreach ($sheet in $book.Worksheets) {
        $usedRange = $null
        try {
            $usedRange = $sheet.UsedRange
            foreach ($errorText in $formulaErrors) {
                $found = $null
                try {
                    $found = $usedRange.Find($errorText)
                    if ($null -ne $found) { $errorCount++ }
                }
                finally { Release-ComObject $found }
            }
        }
        finally { Release-ComObject $usedRange }
    }
    if ($errorCount -ne 0) { throw "Formula errors found: $errorCount" }

    Write-Output 'SELF_TESTS: PASS'
    Write-Output ('BASELINE_CYCLE_SEC: {0:N1}' -f $cycle)
    Write-Output ('BASELINE_TAKT_SEC: {0:N2}' -f $takt)
    Write-Output "FORMULA_ERRORS: $errorCount"
    Write-Output 'VERIFY: PASS'
}
finally {
    if ($null -ne $book) {
        try { $book.Close($false) } catch { Write-Warning $_.Exception.Message }
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { Write-Warning $_.Exception.Message }
    }
    Release-ComObject $book
    Release-ComObject $books
    Release-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
