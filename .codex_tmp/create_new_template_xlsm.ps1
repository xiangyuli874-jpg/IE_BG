$ErrorActionPreference = 'Stop'

$outputDir = 'E:\AI\gongshibiao\outputs\worktime_new_template'
$outputPath = Join-Path $outputDir '工时测量与负荷山积自动扩展模板_v5.xlsx'
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

$excel = $null
$wb = $null

function Release-ComObject {
    param($Object)
    if ($null -ne $Object) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object) | Out-Null
    }
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Add()

    while ($wb.Worksheets.Count -lt 4) {
        $null = $wb.Worksheets.Add()
    }

    $wsIntro = $wb.Worksheets.Item(1)
    $wsParam = $wb.Worksheets.Item(2)
    $wsMeasure = $wb.Worksheets.Item(3)
    $wsChart = $wb.Worksheets.Item(4)

    $wsIntro.Name = '使用说明'
    $wsParam.Name = '参数设置'
    $wsMeasure.Name = '工序测量'
    $wsChart.Name = '山积分析'
    Write-Output 'STEP: worksheets ready'

    # 清理多余工作表
    while ($wb.Worksheets.Count -gt 4) {
        $wb.Worksheets.Item($wb.Worksheets.Count).Delete()
    }

    # ---------- 使用说明 ----------
    $wsIntro.Range('A1:H1').Merge()
    $wsIntro.Range('A1').Value2 = '工时测量与负荷山积自动化模板'
    $wsIntro.Range('A1').Font.Bold = $true
    $wsIntro.Range('A1').Font.Size = 18
    $wsIntro.Range('A1').HorizontalAlignment = -4108

    $intro = @(
        @('用途', '用于录入工序、采样时间、自动计算标准工时和工序负荷，并自动维护山积分析。'),
        @('怎么加工序', '直接在“工序测量”表格最后一行的下一行继续录入。Excel Table 会自动扩表，公式会自动续接。'),
        @('为什么别再手插普通行', '旧表的公式和图表基于固定区域，手插普通行容易打断引用。这个新模板改成“表格驱动”，建议只在表尾续写。'),
        @('图表更新', '山积图预设覆盖 140 个工序行。继续在工序表中录入时，图表会显示新增工序。'),
        @('录入建议', '先填参数，再填工序名称、动作说明、人员和测量值。')
    )
    $startRow = 3
    for ($i = 0; $i -lt $intro.Count; $i++) {
        $wsIntro.Cells.Item($startRow + $i, 1).Value2 = $intro[$i][0]
        $wsIntro.Cells.Item($startRow + $i, 2).Value2 = $intro[$i][1]
    }
    $wsIntro.Columns.Item('A').ColumnWidth = 18
    $wsIntro.Columns.Item('B').ColumnWidth = 90
    $wsIntro.Range('A3:A7').Font.Bold = $true
    $wsIntro.Range('B3:B7').WrapText = $true
    Write-Output 'STEP: intro ready'

    # ---------- 参数设置 ----------
    $wsParam.Range('A1:D1').Merge()
    Write-Output 'STEP: params title merged'
    $wsParam.Range('A1').Value2 = '参数设置'
    $wsParam.Range('A1').Font.Bold = $true
    $wsParam.Range('A1').Font.Size = 16
    $wsParam.Range('A1').HorizontalAlignment = -4108
    Write-Output 'STEP: params title styled'

    $wsParam.Range('A2').Value2 = '参数'
    $wsParam.Range('B2').Value2 = '数值'
    $wsParam.Range('C2').Value2 = '说明'
    $wsParam.Range('A3').Value2 = '计划单班产量(pcs/班)'
    $wsParam.Range('B3').Value2 = ''
    $wsParam.Range('C3').Value2 = '由用户自行填写，用于计算目标节拍 TT'
    $wsParam.Range('A4').Value2 = '班次工时(H/班)'
    $wsParam.Range('B4').Value2 = 11
    $wsParam.Range('C4').Value2 = '单班可用工时'
    $wsParam.Range('A5').Value2 = '稼动率'
    $wsParam.Range('B5').Value2 = 1
    $wsParam.Range('C5').Value2 = '有效生产时间比例'
    $wsParam.Range('A6').Value2 = '宽放率'
    $wsParam.Range('B6').Value2 = 0.05
    $wsParam.Range('C6').Value2 = '疲劳、生理、管理等综合宽放'
    $wsParam.Range('A7').Value2 = '熟练系数'
    $wsParam.Range('B7').Value2 = 0
    $wsParam.Range('C7').Value2 = '评比系数组成项，默认 0'
    $wsParam.Range('A8').Value2 = '努力系数'
    $wsParam.Range('B8').Value2 = 0
    $wsParam.Range('C8').Value2 = '评比系数组成项，默认 0'
    $wsParam.Range('A9').Value2 = '工作环境系数'
    $wsParam.Range('B9').Value2 = 0
    $wsParam.Range('C9').Value2 = '评比系数组成项，默认 0'
    $wsParam.Range('A10').Value2 = '一致性系数'
    $wsParam.Range('B10').Value2 = 0
    $wsParam.Range('C10').Value2 = '评比系数组成项，默认 0'
    $wsParam.Range('A11').Value2 = '综合评比系数'
    $wsParam.Range('C11').Value2 = '熟练系数、努力系数、工作环境系数、一致性系数四项合计'
    $wsParam.Range('A12').Value2 = '目标节拍TT(s)'
    $wsParam.Range('C12').Value2 = '单班可用秒数除以计划单班产量'
    Write-Output 'STEP: params values written'
    $wsParam.Range('B11').Formula = '=SUM(B7:B10)'
    $wsParam.Range('B12').Formula = '=IFERROR(B4*3600*B5/B3,"")'
    Write-Output 'STEP: params formula written'
    $wsParam.Range('B5:B6').NumberFormat = '0.0%'
    $wsParam.Range('B7:B11').NumberFormat = '0.00'
    $wsParam.Range('B12').NumberFormat = '0.00'
    $wsParam.Columns.Item('A').ColumnWidth = 24
    $wsParam.Columns.Item('B').ColumnWidth = 16
    $wsParam.Columns.Item('C').ColumnWidth = 48
    $wsParam.Range('A2:C12').Borders.LineStyle = 1
    $wsParam.Range('A2:C2').Font.Bold = $true
    Write-Output 'STEP: params ready'

    # ---------- 工序测量 ----------
    $wsMeasure.Range('A1:O1').Merge()
    $wsMeasure.Range('A1').Value2 = '工序测量'
    $wsMeasure.Range('A1').Font.Bold = $true
    $wsMeasure.Range('A1').Font.Size = 16
    $wsMeasure.Range('A1').HorizontalAlignment = -4108

    $headers = @(
        '工序号','工序名称','作业要素','人员',
        '测量1','测量2','测量3','测量4','测量5',
        '平均实测(s)','标准工时(s)','负荷ST(s)','TT基准(s)','状态','备注'
    )
    for ($c = 0; $c -lt $headers.Count; $c++) {
        $wsMeasure.Cells.Item(2, $c + 1).Value2 = $headers[$c]
    }

    $seedRows = 10
    for ($i = 0; $i -lt $seedRows; $i++) {
        $row = 3 + $i
        $wsMeasure.Cells.Item($row, 1).Value2 = [double]($i + 1)
        $wsMeasure.Cells.Item($row, 4).Value2 = [double]1
    }

    $lastSeedRow = 2 + $seedRows
    $range = $wsMeasure.Range("A2:O$lastSeedRow")
    $table = $wsMeasure.ListObjects.Add(1, $range, $null, 1)
    $table.Name = 'tblProcess'
    $table.TableStyle = 'TableStyleMedium2'

    $dataStart = 3
    $dataEnd = $lastSeedRow
    $wsMeasure.Range("J$dataStart:J$dataEnd").FormulaR1C1 = '=IF(COUNT(RC[-5]:RC[-1])=0,"",AVERAGE(RC[-5]:RC[-1]))'
    $wsMeasure.Range("K$dataStart:K$dataEnd").FormulaR1C1 = '=IF(RC[-1]="","",RC[-1]*(1+参数设置!R11C2)*(1+参数设置!R6C2))'
    $wsMeasure.Range("L$dataStart:L$dataEnd").FormulaR1C1 = '=IF(OR(RC[-1]="",RC[-8]=""),"",RC[-1]/RC[-8])'
    $wsMeasure.Range("M$dataStart:M$dataEnd").FormulaR1C1 = '=IF(RC[-12]="","",参数设置!R12C2)'
    $wsMeasure.Range("N$dataStart:N$dataEnd").FormulaR1C1 = '=IF(RC[-2]="","",IF(RC[-2]>RC[-1],"超节拍",IF(RC[-2]>=RC[-1]*0.9,"接近上限","正常")))'

    $wsMeasure.Range("E$dataStart:M$dataEnd").NumberFormat = '0.00'
    $wsMeasure.Range("D$dataStart:D$dataEnd").NumberFormat = '0'
    $wsMeasure.Columns.Item('A').ColumnWidth = 10
    $wsMeasure.Columns.Item('B').ColumnWidth = 18
    $wsMeasure.Columns.Item('C').ColumnWidth = 28
    $wsMeasure.Columns.Item('D').ColumnWidth = 8
    $wsMeasure.Columns.Item('E:I').ColumnWidth = 10
    $wsMeasure.Columns.Item('J:M').ColumnWidth = 14
    $wsMeasure.Columns.Item('N').ColumnWidth = 14
    $wsMeasure.Columns.Item('O').ColumnWidth = 20
    $wsMeasure.Activate() | Out-Null
    $wsMeasure.Range('A3').Select() | Out-Null
    $excel.ActiveWindow.FreezePanes = $true
    Write-Output 'STEP: measure table ready'

    Write-Output 'STEP: conditional formatting skipped'

    # ---------- 山积分析 ----------
    $wsChart.Range('A1:H1').Merge()
    $wsChart.Range('A1').Value2 = '负荷山积图'
    $wsChart.Range('A1').Font.Bold = $true
    $wsChart.Range('A1').Font.Size = 16
    $wsChart.Range('A1').HorizontalAlignment = -4108
    $wsChart.Range('A3').Value2 = '说明'
    $wsChart.Range('B3').Value2 = '柱形 = 负荷ST；折线 = TT基准。图表预设覆盖工序测量表前 140 个工序行。'
    $wsChart.Range('D3').Value2 = '线平衡率'
    $wsChart.Range('E3').Formula = '=IFERROR(SUM(工序测量!L3:L142)/(MAX(工序测量!L3:L142)*COUNT(工序测量!L3:L142)),"")'
    $wsChart.Range('E3').NumberFormat = '0.0%'
    $wsChart.Columns.Item('A').ColumnWidth = 14
    $wsChart.Columns.Item('B').ColumnWidth = 80

    $chartObj = $wsChart.ChartObjects().Add(40, 90, 980, 460)
    Write-Output 'STEP: chart object added'
    $chart = $chartObj.Chart
    Write-Output 'STEP: chart object resolved'
    $chart.ChartType = 51
    $chart.HasTitle = $true
    $chart.ChartTitle.Text = '工序负荷山积图'
    $chart.HasLegend = $true
    $chart.Legend.Position = -4160
    Write-Output 'STEP: chart basic settings ready'

    $series1 = $chart.SeriesCollection().NewSeries()
    $series1.Formula = '=SERIES("负荷ST",工序测量!$B$3:$B$142,工序测量!$L$3:$L$142,1)'
    Write-Output 'STEP: series1 140-row formula set'
    $series2 = $chart.SeriesCollection().NewSeries()
    $series2.Formula = '=SERIES("TT基准",工序测量!$B$3:$B$142,工序测量!$M$3:$M$142,2)'
    $series2.ChartType = 4
    Write-Output 'STEP: series2 140-row line formula set'
    $chart.ChartGroups(1).GapWidth = 35

    $chart.Axes(2).HasTitle = $true
    $chart.Axes(2).AxisTitle.Text = '秒'
    Write-Output 'STEP: chart ready'

    # 保存 xlsx
    Write-Output 'STEP: saving'
    $wb.SaveAs($outputPath, 51)
    Write-Output $outputPath
}
finally {
    if ($wb) {
        $wb.Close($true) | Out-Null
        Release-ComObject $wb
    }
    if ($excel) {
        $excel.Quit() | Out-Null
        Release-ComObject $excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

















