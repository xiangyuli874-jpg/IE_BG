[CmdletBinding()]
param(
    [switch]$TestOnly
)

$ErrorActionPreference = 'Stop'

$toolRoot = $PSScriptRoot
$repoRoot = Split-Path (Split-Path $toolRoot -Parent) -Parent
$vbaRoot = Join-Path $toolRoot 'vba'
$tempRoot = Join-Path $repoRoot '.codex_tmp\man_machine_analysis'
$outputRoot = Join-Path $repoRoot 'outputs\man_machine_analysis'
$finalWorkbook = Join-Path $outputRoot '人机作业分析自动排程模板_v1.xlsm'
$buildingWorkbook = Join-Path $tempRoot '人机作业分析自动排程模板_v1.building.xlsm'
$expectedResult = 'Test_ChineseSourceRoundTrip PASS; Test_DomainConstants PASS; Test_DomainTypes PASS; Test_BaselineFixtureExpansion PASS; Test_WorkbookStructure PASS; Test_WorkbookReaders PASS; Test_ValidationRejectsBadInputs PASS'
$MAX_PEOPLE = 5
$MAX_DEVICES = 10
$MAX_STEPS_PER_DEVICE = 20

$xlSrcRange = 1
$xlYes = 1
$xlSheetVisible = -1
$xlValidateList = 3
$xlValidAlertStop = 1
$xlBetween = 1
$xlOpenXMLWorkbookMacroEnabled = 52
$fontName = '微软雅黑'
$requiredColor = 13431551 # RGB(255,242,204)
$optionalColor = 16247773 # RGB(221,235,247)
$outputColor = 15921906   # RGB(242,242,242)
$headerColor = 7949855    # RGB(31,78,121)

function Release-ComObject {
    param([object]$ComObject)
    if ($null -ne $ComObject -and
            [Runtime.InteropServices.Marshal]::IsComObject($ComObject)) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($ComObject)
    }
}

function Set-RangeValues {
    param(
        [object]$Worksheet,
        [int]$StartRow,
        [int]$StartColumn,
        [object[]]$Rows
    )

    $rowCount = $Rows.Count
    $columnCount = $Rows[0].Count
    $values = New-Object 'object[,]' $rowCount, $columnCount
    for ($rowIndex = 0; $rowIndex -lt $rowCount; $rowIndex++) {
        for ($columnIndex = 0; $columnIndex -lt $columnCount; $columnIndex++) {
            $values[$rowIndex, $columnIndex] = $Rows[$rowIndex][$columnIndex]
        }
    }

    $startCell = $null
    $endCell = $null
    $targetRange = $null
    try {
        $startCell = $Worksheet.Cells.Item($StartRow, $StartColumn)
        $endCell = $Worksheet.Cells.Item(
            $StartRow + $rowCount - 1,
            $StartColumn + $columnCount - 1
        )
        $targetRange = $Worksheet.Range($startCell, $endCell)
        $targetRange.Value2 = $values
    }
    finally {
        Release-ComObject $targetRange
        Release-ComObject $endCell
        Release-ComObject $startCell
    }
}

function Add-InputTable {
    param(
        [object]$Worksheet,
        [string]$TableName,
        [int]$HeaderRow,
        [int]$StartColumn,
        [string[]]$Headers,
        [int]$DataRowCount
    )

    $blankRows = @()
    $blankRows += ,([object[]]$Headers)
    for ($rowIndex = 0; $rowIndex -lt $DataRowCount; $rowIndex++) {
        $blankRows += ,([object[]](1..$Headers.Count | ForEach-Object { '' }))
    }
    Set-RangeValues -Worksheet $Worksheet -StartRow $HeaderRow `
        -StartColumn $StartColumn -Rows $blankRows

    $startCell = $null
    $endCell = $null
    $tableRange = $null
    $listObjects = $null
    $table = $null
    try {
        $startCell = $Worksheet.Cells.Item($HeaderRow, $StartColumn)
        $endCell = $Worksheet.Cells.Item(
            $HeaderRow + $DataRowCount,
            $StartColumn + $Headers.Count - 1
        )
        $tableRange = $Worksheet.Range($startCell, $endCell)
        $listObjects = $Worksheet.ListObjects
        $table = $listObjects.Add($xlSrcRange, $tableRange, $null, $xlYes)
        $table.Name = $TableName
        $table.TableStyle = 'TableStyleMedium2'
        $table.ShowTableStyleRowStripes = $true
        return $table
    }
    finally {
        Release-ComObject $listObjects
        Release-ComObject $tableRange
        Release-ComObject $endCell
        Release-ComObject $startCell
    }
}

function Set-ColumnFill {
    param(
        [object]$Table,
        [string[]]$ColumnNames,
        [int]$Color
    )

    foreach ($columnName in $ColumnNames) {
        $listColumn = $null
        $dataRange = $null
        try {
            $listColumn = $Table.ListColumns.Item($columnName)
            $dataRange = $listColumn.DataBodyRange
            $dataRange.Interior.Color = $Color
        }
        finally {
            Release-ComObject $dataRange
            Release-ComObject $listColumn
        }
    }
}

function Add-ListValidation {
    param(
        [object]$Table,
        [string]$ColumnName,
        [string]$Formula
    )

    $listColumn = $null
    $dataRange = $null
    $validation = $null
    try {
        $listColumn = $Table.ListColumns.Item($ColumnName)
        $dataRange = $listColumn.DataBodyRange
        $validation = $dataRange.Validation
        $validation.Delete()
        $validation.Add($xlValidateList, $xlValidAlertStop, $xlBetween, $Formula)
        $validation.IgnoreBlank = $true
        $validation.InCellDropdown = $true
        $validation.ShowError = $true
    }
    finally {
        Release-ComObject $validation
        Release-ComObject $dataRange
        Release-ComObject $listColumn
    }
}

function Set-SectionTitle {
    param(
        [object]$Worksheet,
        [string]$Address,
        [string]$Text
    )

    $range = $null
    try {
        $range = $Worksheet.Range($Address)
        [void]$range.Merge()
        $range.Value2 = $Text
        $range.Font.Bold = $true
        $range.Font.Size = 12
        $range.Font.Color = 16777215
        $range.Interior.Color = $headerColor
        $range.HorizontalAlignment = -4131
        $range.VerticalAlignment = -4108
    }
    finally {
        Release-ComObject $range
    }
}

function Set-SheetTitle {
    param(
        [object]$Worksheet,
        [string]$Address,
        [string]$Text
    )

    $range = $null
    try {
        $range = $Worksheet.Range($Address)
        [void]$range.Merge()
        $range.Value2 = $Text
        $range.Font.Bold = $true
        $range.Font.Size = 18
        $range.Font.Color = 16777215
        $range.Interior.Color = $headerColor
        $range.HorizontalAlignment = -4108
        $range.VerticalAlignment = -4108
        $range.RowHeight = 30
    }
    finally {
        Release-ComObject $range
    }
}

function Set-FreezePane {
    param(
        [object]$Excel,
        [object]$Worksheet,
        [string]$CellAddress
    )

    $cell = $null
    try {
        [void]$Worksheet.Activate()
        $cell = $Worksheet.Range($CellAddress)
        [void]$cell.Select()
        $Excel.ActiveWindow.FreezePanes = $false
        $Excel.ActiveWindow.FreezePanes = $true
    }
    finally {
        Release-ComObject $cell
    }
}

function Add-WorkbookName {
    param(
        [object]$Book,
        [string]$Name,
        [object]$TargetRange
    )

    $names = $null
    $definedName = $null
    try {
        $names = $Book.Names
        $definedName = $names.Add($Name, "=$($TargetRange.Address($true, $true, 1, $true))")
    }
    finally {
        Release-ComObject $definedName
        Release-ComObject $names
    }
}

function Initialize-WorkbookStructure {
    param(
        [object]$Excel,
        [object]$Book
    )

    $sheetNames = @('基础设置', '作业步骤', '自动排程', '人机作业图', '改善对比报告')
    $worksheets = $null
    try {
        $worksheets = $Book.Worksheets
        while ($worksheets.Count -lt 5) {
            $newSheet = $null
            try {
                $newSheet = $worksheets.Add()
            }
            finally {
                Release-ComObject $newSheet
            }
        }
        while ($worksheets.Count -gt 5) {
            $extraSheet = $null
            try {
                $extraSheet = $worksheets.Item($worksheets.Count)
                $extraSheet.Delete()
            }
            finally {
                Release-ComObject $extraSheet
            }
        }

        for ($sheetIndex = 1; $sheetIndex -le 5; $sheetIndex++) {
            $worksheet = $null
            $allCells = $null
            try {
                $worksheet = $worksheets.Item($sheetIndex)
                $worksheet.Name = $sheetNames[$sheetIndex - 1]
                $worksheet.Visible = $xlSheetVisible
                $allCells = $worksheet.Cells
                $allCells.Font.Name = $fontName
                $allCells.Font.Size = 10
                $worksheet.Tab.Color = $headerColor
            }
            finally {
                Release-ComObject $allCells
                Release-ComObject $worksheet
            }
        }

        $settings = $null
        $steps = $null
        $schedule = $null
        $gantt = $null
        $report = $null
        try {
            $settings = $worksheets.Item('基础设置')
            $steps = $worksheets.Item('作业步骤')
            $schedule = $worksheets.Item('自动排程')
            $gantt = $worksheets.Item('人机作业图')
            $report = $worksheets.Item('改善对比报告')

            Build-SettingsSheet -Book $Book -Worksheet $settings
            Build-StepsSheet -Worksheet $steps
            Build-OutputSheet -Worksheet $schedule -Title '自动排程'
            Build-OutputSheet -Worksheet $gantt -Title '人机作业图'
            Build-OutputSheet -Worksheet $report -Title '改善对比报告'

            Set-FreezePane -Excel $Excel -Worksheet $settings -CellAddress 'A4'
            Set-FreezePane -Excel $Excel -Worksheet $steps -CellAddress 'A5'
            Set-FreezePane -Excel $Excel -Worksheet $schedule -CellAddress 'A4'
            Set-FreezePane -Excel $Excel -Worksheet $gantt -CellAddress 'A4'
            Set-FreezePane -Excel $Excel -Worksheet $report -CellAddress 'A4'
        }
        finally {
            Release-ComObject $report
            Release-ComObject $gantt
            Release-ComObject $schedule
            Release-ComObject $steps
            Release-ComObject $settings
        }
    }
    finally {
        Release-ComObject $worksheets
    }
}

function Build-SettingsSheet {
    param(
        [object]$Book,
        [object]$Worksheet
    )

    $resourceHeaderRow = 8
    $resourceDataRow = $resourceHeaderRow + 1
    $moveSectionRow = $resourceHeaderRow + [Math]::Max($MAX_PEOPLE, $MAX_DEVICES) + 3
    $moveHeaderRow = $moveSectionRow + 1
    $moveDataRow = $moveHeaderRow + 1

    Set-SheetTitle -Worksheet $Worksheet -Address 'A1:L1' -Text '人机作业分析｜基础设置'
    Set-SectionTitle -Worksheet $Worksheet -Address 'A3:J3' -Text '分析参数'
    Set-SectionTitle -Worksheet $Worksheet -Address 'A7:E7' -Text '人员与技能'
    Set-SectionTitle -Worksheet $Worksheet -Address 'G7:L7' -Text '设备设置'
    Set-SectionTitle -Worksheet $Worksheet `
        -Address "A${moveSectionRow}:K${moveSectionRow}" -Text '设备间移动时间矩阵（秒）'

    $parameterHeaders = @(
        '方案名称', '人员数量', '设备数量', '计划分析循环数', '优化目标', '搜索迭代次数',
        '周期权重', '均衡权重', '等待权重', '移动权重'
    )
    $peopleHeaders = @('人员编号', '人员名称', '技能', '可操作设备', '启用')
    $deviceHeaders = @('设备编号', '设备名称', '产品', '关系类型', '前置设备', '启用')
    $moveHeaders = @('起点') + (1..$MAX_DEVICES | ForEach-Object { "M$_" })

    $parameters = $null
    $people = $null
    $devices = $null
    $moveTime = $null
    try {
        $parameters = Add-InputTable -Worksheet $Worksheet -TableName 'tblParameters' `
            -HeaderRow 4 -StartColumn 1 -Headers $parameterHeaders -DataRowCount 1
        $people = Add-InputTable -Worksheet $Worksheet -TableName 'tblPeople' `
            -HeaderRow $resourceHeaderRow -StartColumn 1 -Headers $peopleHeaders `
            -DataRowCount $MAX_PEOPLE
        $devices = Add-InputTable -Worksheet $Worksheet -TableName 'tblDevices' `
            -HeaderRow $resourceHeaderRow -StartColumn 7 -Headers $deviceHeaders `
            -DataRowCount $MAX_DEVICES
        $moveTime = Add-InputTable -Worksheet $Worksheet -TableName 'tblMoveTime' `
            -HeaderRow $moveHeaderRow -StartColumn 1 -Headers $moveHeaders `
            -DataRowCount $MAX_DEVICES

        Set-RangeValues -Worksheet $Worksheet -StartRow 5 -StartColumn 1 -Rows @(
            ,([object[]]@('当前方案', 1, 3, 3, '综合优化', 500, 0.4, 0.2, 0.3, 0.1))
        )
        Set-RangeValues -Worksheet $Worksheet -StartRow $resourceDataRow -StartColumn 1 -Rows @(
            ,([object[]]@('P1', '操作员1', '通用', 'M1,M2,M3', '是'))
        )
        $deviceDefaults = @()
        for ($deviceIndex = 1; $deviceIndex -le $MAX_DEVICES; $deviceIndex++) {
            $deviceDefaults += ,([object[]]@(
                "M$deviceIndex", "设备$deviceIndex", '', '独立循环', '', $(if ($deviceIndex -le 3) { '是' } else { '否' })
            ))
        }
        Set-RangeValues -Worksheet $Worksheet -StartRow $resourceDataRow `
            -StartColumn 7 -Rows $deviceDefaults
        $moveDefaults = @()
        for ($fromIndex = 1; $fromIndex -le $MAX_DEVICES; $fromIndex++) {
            $row = [object[]]::new($MAX_DEVICES + 1)
            $row[0] = "M$fromIndex"
            for ($toIndex = 1; $toIndex -le $MAX_DEVICES; $toIndex++) {
                $row[$toIndex] = 0
            }
            $moveDefaults += ,$row
        }
        Set-RangeValues -Worksheet $Worksheet -StartRow $moveDataRow `
            -StartColumn 1 -Rows $moveDefaults

        Set-ColumnFill -Table $parameters -ColumnNames @(
            '人员数量', '设备数量', '计划分析循环数', '优化目标', '搜索迭代次数',
            '周期权重', '均衡权重', '等待权重', '移动权重'
        ) -Color $requiredColor
        Set-ColumnFill -Table $parameters -ColumnNames @('方案名称') -Color $optionalColor
        Set-ColumnFill -Table $people -ColumnNames @('人员编号', '技能', '可操作设备', '启用') `
            -Color $requiredColor
        Set-ColumnFill -Table $people -ColumnNames @('人员名称') -Color $optionalColor
        Set-ColumnFill -Table $devices -ColumnNames @(
            '设备编号', '设备名称', '关系类型', '启用'
        ) -Color $requiredColor
        Set-ColumnFill -Table $devices -ColumnNames @('产品', '前置设备') -Color $optionalColor
        Set-ColumnFill -Table $moveTime -ColumnNames $moveHeaders -Color $requiredColor

        Add-ListValidation -Table $parameters -ColumnName '优化目标' `
            -Formula '"最高产能,人员均衡,设备少等待,综合优化"'
        Add-ListValidation -Table $people -ColumnName '启用' -Formula '"是,否"'
        Add-ListValidation -Table $devices -ColumnName '关系类型' `
            -Formula '"独立循环,有先后顺序"'
        Add-ListValidation -Table $devices -ColumnName '启用' -Formula '"是,否"'

        $nameTargets = @{
            nmCycleCount = 'D5'
            nmOptimizationTarget = 'E5'
            nmWeightCycle = 'G5'
            nmWeightBalance = 'H5'
            nmWeightWait = 'I5'
            nmWeightMove = 'J5'
        }
        foreach ($entry in $nameTargets.GetEnumerator()) {
            $targetRange = $null
            try {
                $targetRange = $Worksheet.Range($entry.Value)
                Add-WorkbookName -Book $Book -Name $entry.Key -TargetRange $targetRange
            }
            finally {
                Release-ComObject $targetRange
            }
        }
        $deviceIds = $null
        $personIds = $null
        try {
            $deviceIds = $devices.ListColumns.Item('设备编号').DataBodyRange
            Add-WorkbookName -Book $Book -Name 'nmDeviceIds' -TargetRange $deviceIds
            $personIds = $people.ListColumns.Item('人员编号').DataBodyRange
            Add-WorkbookName -Book $Book -Name 'nmPersonIds' -TargetRange $personIds
        }
        finally {
            Release-ComObject $personIds
            Release-ComObject $deviceIds
        }

        $Worksheet.Columns.Item('A').ColumnWidth = 13
        $Worksheet.Columns.Item('B').ColumnWidth = 16
        $Worksheet.Columns.Item('C').ColumnWidth = 15
        $Worksheet.Columns.Item('D').ColumnWidth = 18
        $Worksheet.Columns.Item('E').ColumnWidth = 12
        $Worksheet.Columns.Item('F').ColumnWidth = 12
        $Worksheet.Columns.Item('G').ColumnWidth = 13
        $Worksheet.Columns.Item('H').ColumnWidth = 13
        $Worksheet.Columns.Item('I:J').ColumnWidth = 13
        $Worksheet.Columns.Item('K:L').ColumnWidth = 11
    }
    finally {
        Release-ComObject $moveTime
        Release-ComObject $devices
        Release-ComObject $people
        Release-ComObject $parameters
    }
}

function Build-StepsSheet {
    param([object]$Worksheet)

    Set-SheetTitle -Worksheet $Worksheet -Address 'A1:K1' -Text '人机作业分析｜作业步骤'
    $instruction = $null
    try {
        $instruction = $Worksheet.Range('A2:K2')
        [void]$instruction.Merge()
        $instruction.Value2 = '每台设备最多 20 步；黄色为必填，浅蓝为选填。排程算法将在后续模块生成结果。'
        $instruction.Interior.Color = $optionalColor
        $instruction.Font.Color = 8210719
    }
    finally {
        Release-ComObject $instruction
    }
    Set-SectionTitle -Worksheet $Worksheet -Address 'A3:K3' -Text '集中作业步骤输入'

    $headers = @(
        '设备', '步骤号', '名称', '类型', '工时(s)', '前置步骤',
        '所需技能', '锁定人员', '锁定开始(s)', '允许等待', '备注'
    )
    $steps = $null
    try {
        $steps = Add-InputTable -Worksheet $Worksheet -TableName 'tblSteps' `
            -HeaderRow 4 -StartColumn 1 -Headers $headers `
            -DataRowCount ($MAX_DEVICES * $MAX_STEPS_PER_DEVICE)
        Set-ColumnFill -Table $steps -ColumnNames @(
            '设备', '步骤号', '名称', '类型', '工时(s)'
        ) -Color $requiredColor
        Set-ColumnFill -Table $steps -ColumnNames @(
            '前置步骤', '所需技能', '锁定人员', '锁定开始(s)', '允许等待', '备注'
        ) -Color $optionalColor
        Add-ListValidation -Table $steps -ColumnName '类型' `
            -Formula '"人工,自动运行,人机协同,等待"'
        Add-ListValidation -Table $steps -ColumnName '允许等待' -Formula '"是,否"'
        Add-ListValidation -Table $steps -ColumnName '设备' `
            -Formula '=nmDeviceIds'
        Add-ListValidation -Table $steps -ColumnName '锁定人员' `
            -Formula '=nmPersonIds'

        $Worksheet.Columns.Item('A').ColumnWidth = 12
        $Worksheet.Columns.Item('B').ColumnWidth = 10
        $Worksheet.Columns.Item('C').ColumnWidth = 22
        $Worksheet.Columns.Item('D').ColumnWidth = 14
        $Worksheet.Columns.Item('E').ColumnWidth = 12
        $Worksheet.Columns.Item('F').ColumnWidth = 14
        $Worksheet.Columns.Item('G').ColumnWidth = 15
        $Worksheet.Columns.Item('H').ColumnWidth = 14
        $Worksheet.Columns.Item('I').ColumnWidth = 15
        $Worksheet.Columns.Item('J').ColumnWidth = 12
        $Worksheet.Columns.Item('K').ColumnWidth = 24
    }
    finally {
        Release-ComObject $steps
    }
}

function Build-OutputSheet {
    param(
        [object]$Worksheet,
        [string]$Title
    )

    Set-SheetTitle -Worksheet $Worksheet -Address 'A1:L1' -Text "人机作业分析｜$Title"
    $outputArea = $null
    try {
        $outputArea = $Worksheet.Range('A2:L40')
        $outputArea.Interior.Color = $outputColor
        $outputArea.Borders.LineStyle = 1
        $outputArea.Borders.Color = 14277081
        $outputArea.WrapText = $true
        $Worksheet.Range('A2').Value2 = '此页由 VBA 自动生成，请勿手工覆盖。'
        $Worksheet.Range('A2').Font.Bold = $true
        $Worksheet.Columns.Item('A:L').ColumnWidth = 14
    }
    finally {
        Release-ComObject $outputArea
    }
}

function Import-VbaModules {
    param(
        [object]$Book,
        [string]$VbaDirectory,
        [string]$ImportDirectory
    )

    $vbProject = $null
    $vbComponents = $null
    try {
        try {
            $vbProject = $Book.VBProject
            $vbComponents = $vbProject.VBComponents
        }
        catch {
            throw 'Excel blocked VBA project access. Enable "Trust access to the VBA project object model".'
        }

        Get-ChildItem -LiteralPath $VbaDirectory -Filter '*.bas' |
            Sort-Object Name |
            ForEach-Object {
                $sourceText = [IO.File]::ReadAllText(
                    $_.FullName,
                    (New-Object Text.UTF8Encoding($false, $true))
                )
                $importPath = Join-Path $ImportDirectory ("import_" + $_.Name)
                [IO.File]::WriteAllText($importPath, $sourceText, [Text.Encoding]::Default)

                $importedComponent = $null
                try {
                    $importedComponent = $vbComponents.Import($importPath)
                }
                finally {
                    Release-ComObject $importedComponent
                }
            }
    }
    finally {
        Release-ComObject $vbComponents
        Release-ComObject $vbProject
    }
}

New-Item -ItemType Directory -Force -Path $tempRoot | Out-Null
if (-not $TestOnly) {
    New-Item -ItemType Directory -Force -Path $outputRoot | Out-Null
}

$stamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
$workbookPath = if ($TestOnly) {
    Join-Path $tempRoot "workbook_structure_test_$stamp.xlsm"
}
else {
    $buildingWorkbook
}

if (-not $TestOnly -and (Test-Path -LiteralPath $buildingWorkbook)) {
    Remove-Item -LiteralPath $buildingWorkbook
}

$excel = $null
$books = $null
$book = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.ScreenUpdating = $false

    $books = $excel.Workbooks
    $book = $books.Add()
    Initialize-WorkbookStructure -Excel $excel -Book $book
    $book.SaveAs($workbookPath, $xlOpenXMLWorkbookMacroEnabled)

    Import-VbaModules -Book $book -VbaDirectory $vbaRoot -ImportDirectory $tempRoot
    $book.Save()

    $macroName = "'$($book.Name)'!RunAllSelfTests"
    $fixturePath = Join-Path $toolRoot 'fixtures\baseline_1p3m.csv'
    $result = $excel.Run($macroName, $fixturePath)
    Write-Output $result
    if ($null -eq $result -or [string]$result -cne $expectedResult) {
        throw "Unexpected self-test result. Expected [$expectedResult], actual [$result]"
    }

}
finally {
    if ($null -ne $book) {
        try {
            $book.Close($false)
        }
        catch {
            Write-Warning "Failed to close workbook: $($_.Exception.Message)"
        }
    }
    if ($null -ne $excel) {
        try {
            $excel.Quit()
        }
        catch {
            Write-Warning "Failed to quit Excel instance: $($_.Exception.Message)"
        }
    }

    Release-ComObject $book
    Release-ComObject $books
    Release-ComObject $excel
    $book = $null
    $books = $null
    $excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()

    if ($TestOnly -and (Test-Path -LiteralPath $workbookPath)) {
        Remove-Item -LiteralPath $workbookPath
    }
}

if (-not $TestOnly) {
    Move-Item -LiteralPath $buildingWorkbook -Destination $finalWorkbook -Force
    Write-Output "OUTPUT: $finalWorkbook"
}
