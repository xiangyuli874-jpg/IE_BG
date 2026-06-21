[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$toolRoot = $PSScriptRoot
$repoRoot = Split-Path (Split-Path $toolRoot -Parent) -Parent
$outputRoot = Join-Path $repoRoot 'outputs\人机作业分析表'
$workbookPath = Join-Path $outputRoot '人机作业分析自动排程模板_v2.xlsm'
$requiredTables = @{
    '基础设置' = @('tblParameters', 'tblPeople', 'tblDevices', 'tblMoveTime')
    '作业步骤' = @('tblSteps')
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
    $expectedSheets = @('使用说明', '基础设置', '作业步骤', '自动排程', '人机作业图', '改善对比报告')
    if ($book.Worksheets.Count -ne 6) { throw "Expected 6 worksheets, got $($book.Worksheets.Count)" }
    for ($index = 0; $index -lt $expectedSheets.Count; $index++) {
        if ($book.Worksheets.Item($index + 1).Name -ne $expectedSheets[$index]) {
            throw "Unexpected worksheet at position $($index + 1)"
        }
    }
    foreach ($sheetName in $requiredTables.Keys) {
        $sheet = $null
        try {
            $sheet = $book.Worksheets.Item([string]$sheetName)
            foreach ($tableName in $requiredTables[$sheetName]) {
                $table = $null
                try {
                    $table = $sheet.ListObjects.Item($tableName)
                    if ($null -eq $table) { throw "Missing table at sheet $sheetName`: $tableName" }
                }
                finally { Release-ComObject $table }
            }
        }
        finally { Release-ComObject $sheet }
    }

    foreach ($sheetName in $expectedSheets) {
        $viewSheet = $book.Worksheets.Item($sheetName)
        try {
            $viewSheet.Activate()
            if ($excel.ActiveWindow.DisplayGridlines) {
                throw "Gridlines are visible on worksheet: $sheetName"
            }
        }
        finally { Release-ComObject $viewSheet }
    }

    $selfTests = $excel.Run("'$($book.Name)'!RunAllSelfTests",
        (Join-Path $toolRoot 'fixtures\baseline_1p3m.csv'))
    if ([string]$selfTests -like 'SELF_TESTS FAIL*') { throw [string]$selfTests }

    $settingsSheet = $book.Worksheets.Item('基础设置')
    $stepsSheet = $book.Worksheets.Item('作业步骤')
    $scheduleSheet = $book.Worksheets.Item('自动排程')
    $ganttSheet = $book.Worksheets.Item('人机作业图')
    $instructionsSheet = $book.Worksheets.Item('使用说明')
    $parameterTable = $settingsSheet.ListObjects.Item('tblParameters')
        $peopleCount = $parameterTable.ListColumns.Item('人员数量').DataBodyRange.Cells(1, 1).Value2
        $deviceCount = $parameterTable.ListColumns.Item('设备数量').DataBodyRange.Cells(1, 1).Value2
        if ($null -ne $peopleCount -and [string]$peopleCount -ne '') {
            throw 'Default people count is not blank'
        }
        if ($null -ne $deviceCount -and [string]$deviceCount -ne '') {
            throw 'Default device count is not blank'
        }
        if ($instructionsSheet.Buttons().Count -lt 5) { throw 'Instruction buttons are missing' }
        if ($instructionsSheet.Hyperlinks.Count -ne 3) { throw 'Instruction video links are missing' }
        $expectedInstructionButtons = @{
            '新建空白分析' = 'CmdNewBlankAnalysis'
            '载入1人3机示例并试算' = 'CmdLoadExampleAndRun'
            '前往基础设置' = 'CmdGoSettings'
            '前往作业步骤' = 'CmdGoSteps'
            '前往自动排程' = 'CmdGoSchedule'
        }
        foreach ($button in $instructionsSheet.Buttons()) {
            if (-not $expectedInstructionButtons.ContainsKey([string]$button.Caption)) {
                throw "Unexpected instruction button: $($button.Caption)"
            }
            if ([string]$button.OnAction -ne $expectedInstructionButtons[[string]$button.Caption]) {
                throw "Incorrect macro binding for instruction button: $($button.Caption)"
            }
        }
        $validationChecks = @(
            @($settingsSheet, 'tblParameters', '优化目标'),
            @($settingsSheet, 'tblPeople', '启用'),
            @($settingsSheet, 'tblDevices', '关系类型'),
            @($settingsSheet, 'tblDevices', '启用'),
            @($stepsSheet, 'tblSteps', '设备'),
            @($stepsSheet, 'tblSteps', '类型'),
            @($stepsSheet, 'tblSteps', '允许等待')
        )
        foreach ($check in $validationChecks) {
            $validationRange = $check[0].ListObjects.Item($check[1]).ListColumns.Item($check[2]).DataBodyRange
            if ($validationRange.Validation.Type -ne 3) {
                throw "Missing list validation: $($check[1]).$($check[2])"
            }
        }
        if ($excel.WorksheetFunction.CountA(
                $settingsSheet.ListObjects.Item('tblPeople').DataBodyRange.Columns.Item(2)) -gt 0) {
            throw 'Default people data is not blank'
        }
        if ($excel.WorksheetFunction.CountA(
                $settingsSheet.ListObjects.Item('tblDevices').DataBodyRange.Columns.Item(2)) -gt 0) {
            throw 'Default device data is not blank'
        }
        if ($excel.WorksheetFunction.CountA(
                $stepsSheet.ListObjects.Item('tblSteps').DataBodyRange.Columns.Item(4)) -gt 0) {
            throw 'Default step data is not blank'
        }
        if ($scheduleSheet.ListObjects.Count -ne 0) { throw 'Default schedule is not blank' }
    if ($ganttSheet.ChartObjects().Count -ne 0) { throw 'Default chart is not blank' }

    $excel.Run("'$($book.Name)'!LoadExampleDataAndRun")

    $vbProject = $null
    try {
        $cycle = [double]$scheduleSheet.Range('A4').Value2
        $takt = [double]$scheduleSheet.Range('B4').Value2
        $rowCount = $scheduleSheet.ListObjects.Item('tblSchedule').ListRows.Count
        $chartCount = $ganttSheet.ChartObjects().Count
        $buttonCount = $scheduleSheet.Buttons().Count
        if ([Math]::Abs($cycle - 56.0) -gt 0.01) { throw "Baseline cycle expected 56.0, got $cycle" }
        if ([Math]::Abs($takt - (56.0 / 3.0)) -gt 0.01) { throw "Baseline takt expected 18.67, got $takt" }
        if ($rowCount -ne 45) { throw "Expected 45 scheduled tasks, got $rowCount" }
        if ($chartCount -lt 1) { throw 'Gantt chart is missing' }
        $assignedTaskCount = $excel.WorksheetFunction.CountA(
            $scheduleSheet.ListObjects.Item('tblSchedule').ListColumns.Item('执行人员').DataBodyRange)
        $ganttChart = $ganttSheet.ChartObjects('chtManMachine').Chart
        $durationSeries = $ganttChart.SeriesCollection(2)
        $ganttPointCount = $durationSeries.Points().Count
        $expectedGanttPointCount = $rowCount + $assignedTaskCount
        if ($ganttPointCount -ne $expectedGanttPointCount) {
            throw "Expected $expectedGanttPointCount Gantt points, got $ganttPointCount"
        }
        if ($buttonCount -lt 8) { throw "Expected at least 8 schedule buttons, got $buttonCount" }
        $vbProject = $book.VBProject
        if ($vbProject.VBComponents.Count -lt 10) { throw 'VBA project is incomplete' }
    }
    finally {
        Release-ComObject $vbProject
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

    $excel.Run("'$($book.Name)'!ResetToBlankState")
    if ($scheduleSheet.ListObjects.Count -ne 0) { throw 'Schedule was not cleared by blank reset' }
    if ($ganttSheet.ChartObjects().Count -ne 0) { throw 'Chart was not cleared by blank reset' }

    Write-Output 'SELF_TESTS: PASS'
    Write-Output ('BASELINE_CYCLE_SEC: {0:N1}' -f $cycle)
    Write-Output ('BASELINE_TAKT_SEC: {0:N2}' -f $takt)
    Write-Output "FORMULA_ERRORS: $errorCount"
    Write-Output 'VERIFY: PASS'
}
finally {
    Release-ComObject $instructionsSheet
    Release-ComObject $ganttSheet
    Release-ComObject $scheduleSheet
    Release-ComObject $stepsSheet
    Release-ComObject $settingsSheet
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
