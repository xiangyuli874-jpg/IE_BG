$ErrorActionPreference = 'Stop'

$outputDir = 'E:\AI\gongshibiao\outputs\worktime_action_element_template'
$outputPath = Join-Path $outputDir '动作要素工时与堆叠山积图模板_v2.xlsx'
New-Item -ItemType Directory -Force -Path $outputDir | Out-Null

$excel = $null
$wb = $null

function Release-ComObject {
    param($Object)
    if ($null -ne $Object) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Object) | Out-Null
    }
}

function RgbValue {
    param([int]$R, [int]$G, [int]$B)
    return ($R + ($G * 256) + ($B * 65536))
}

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Add()

    while ($wb.Worksheets.Count -lt 5) {
        $null = $wb.Worksheets.Add()
    }

    $wsIntro = $wb.Worksheets.Item(1)
    $wsParam = $wb.Worksheets.Item(2)
    $wsMeasure = $wb.Worksheets.Item(3)
    $wsLoad = $wb.Worksheets.Item(4)
    $wsChart = $wb.Worksheets.Item(5)

    $wsIntro.Name = '使用说明'
    $wsParam.Name = '参数设置'
    $wsMeasure.Name = '动作要素测量'
    $wsLoad.Name = '员工负荷表'
    $wsChart.Name = '山积分析'

    while ($wb.Worksheets.Count -gt 5) {
        $wb.Worksheets.Item($wb.Worksheets.Count).Delete()
    }

    $processCount = 140
    $actionsPerProcess = 5
    $dataStart = 3
    $dataRows = $processCount * $actionsPerProcess
    $dataEnd = $dataStart + $dataRows - 1
    $actionNames = @('ST1','ST2','ST3','ST4','ST5')
    $seriesColors = @(
        (RgbValue 31 78 121),
        (RgbValue 192 0 0),
        (RgbValue 80 136 69),
        (RgbValue 112 48 160),
        (RgbValue 237 125 49)
    )
    $ttColor = RgbValue 255 102 0

    # ---------- 使用说明 ----------
    $wsIntro.Range('A1:H1').Merge()
    $wsIntro.Range('A1').Value2 = '动作要素工时与员工负荷表模板'
    $wsIntro.Range('A1').Font.Bold = $true
    $wsIntro.Range('A1').Font.Size = 18
    $wsIntro.Range('A1').HorizontalAlignment = -4108

    $intro = @(
        @('用途', '用于录入单个工序下的多个动作要素测量时间，自动计算动作平均ST、工序总ST、宽放ST，并生成动作堆叠山积图和员工负荷表。'),
        @('容量', '无宏稳定模板，预置 140 个工序，每个工序 5 行动作要素，共 700 行。'),
        @('动作序号', '动作序号按 ST1、ST2、ST3、ST4、ST5 填充，表示一个工序内的五个动作段。'),
        @('山积分析', '山积图柱形展示真实动作平均ST堆叠，并显示每段时间标签；TT基准用折线显示。'),
        @('员工负荷表', '横向自动抓取工序号、工序名称、动作名称、动作时间、总工时、目标节拍和实际节拍。')
    )
    for ($i = 0; $i -lt $intro.Count; $i++) {
        $row = 3 + $i
        $wsIntro.Cells.Item($row, 1).Value2 = $intro[$i][0]
        $wsIntro.Cells.Item($row, 2).Value2 = $intro[$i][1]
    }
    $wsIntro.Columns.Item('A').ColumnWidth = 18
    $wsIntro.Columns.Item('B').ColumnWidth = 95
    $wsIntro.Range('A3:A7').Font.Bold = $true
    $wsIntro.Range('B3:B7').WrapText = $true

    # ---------- 参数设置 ----------
    $wsParam.Range('A1:D1').Merge()
    $wsParam.Range('A1').Value2 = '参数设置'
    $wsParam.Range('A1').Font.Bold = $true
    $wsParam.Range('A1').Font.Size = 16
    $wsParam.Range('A1').HorizontalAlignment = -4108

    $wsParam.Range('A2').Value2 = '参数'
    $wsParam.Range('B2').Value2 = '数值'
    $wsParam.Range('C2').Value2 = '说明'
    $wsParam.Range('A3').Value2 = '计划单班产量(pcs/班)'
    $wsParam.Range('B3').Value2 = ''
    $wsParam.Range('C3').Value2 = '由用户自行填写，用于计算目标节拍 TT'
    $wsParam.Range('A4').Value2 = '班次工时(h/班)'
    $wsParam.Range('B4').Value2 = 11
    $wsParam.Range('C4').Value2 = '单班可用工时'
    $wsParam.Range('A5').Value2 = '稼动率'
    $wsParam.Range('B5').Value2 = 1
    $wsParam.Range('C5').Value2 = '有效生产时间比例'
    $wsParam.Range('A6').Value2 = '宽放率'
    $wsParam.Range('B6').Value2 = 0.05
    $wsParam.Range('C6').Value2 = '用于补充计算宽放ST，不参与山积图柱形'
    $wsParam.Range('A7').Value2 = '目标节拍TT(s)'
    $wsParam.Range('B7').Formula = '=IFERROR(B4*3600*B5/B3,"")'
    $wsParam.Range('C7').Value2 = '单班可用秒数除以计划单班产量'

    $wsParam.Range('B5:B6').NumberFormat = '0.0%'
    $wsParam.Range('B7').NumberFormat = '0.00'
    $wsParam.Columns.Item('A').ColumnWidth = 24
    $wsParam.Columns.Item('B').ColumnWidth = 16
    $wsParam.Columns.Item('C').ColumnWidth = 58
    $wsParam.Range('A2:C7').Borders.LineStyle = 1
    $wsParam.Range('A2:C2').Font.Bold = $true
    $wsParam.Range('A2:C2').Interior.Color = 15773696

    # ---------- 动作要素测量 ----------
    $wsMeasure.Range('A1:O1').Merge()
    $wsMeasure.Range('A1').Value2 = '动作要素测量'
    $wsMeasure.Range('A1').Font.Bold = $true
    $wsMeasure.Range('A1').Font.Size = 16
    $wsMeasure.Range('A1').HorizontalAlignment = -4108

    $headers = @(
        '工序号','工序名称','动作序号','动作单元描述',
        'ST1','ST2','ST3','ST4','ST5',
        '平均ST','工序总ST','宽放ST','TT基准','状态','备注'
    )
    for ($c = 0; $c -lt $headers.Count; $c++) {
        $wsMeasure.Cells.Item(2, $c + 1).Value2 = $headers[$c]
    }

    for ($p = 1; $p -le $processCount; $p++) {
        for ($a = 1; $a -le $actionsPerProcess; $a++) {
            $row = $dataStart + (($p - 1) * $actionsPerProcess) + ($a - 1)
            $wsMeasure.Cells.Item($row, 1).Value2 = [double]$p
            $wsMeasure.Cells.Item($row, 3).Value2 = $actionNames[$a - 1]
        }
    }

    $tableRange = $wsMeasure.Range("A2:O$dataEnd")
    $table = $wsMeasure.ListObjects.Add(1, $tableRange, $null, 1)
    $table.Name = 'tblActionElements'
    $table.TableStyle = 'TableStyleMedium2'

    $wsMeasure.Range("J$dataStart:J$dataEnd").FormulaR1C1 = '=IF(COUNT(RC[-5]:RC[-1])=0,"",AVERAGE(RC[-5]:RC[-1]))'
    $wsMeasure.Range("K$dataStart:K$dataEnd").FormulaR1C1 = '=IF(RC[-8]<>"ST1","",IF(COUNTIFS(C1,RC[-10],C10,"<>")=0,"",SUMIFS(C10,C1,RC[-10])))'
    $wsMeasure.Range("L$dataStart:L$dataEnd").FormulaR1C1 = '=IF(RC[-9]<>"ST1","",IF(RC[-1]="","",RC[-1]*(1+参数设置!R6C2)))'
    $wsMeasure.Range("M$dataStart:M$dataEnd").FormulaR1C1 = '=IF(RC[-12]="","",参数设置!R7C2)'
    $wsMeasure.Range("N$dataStart:N$dataEnd").FormulaR1C1 = '=IF(RC[-3]="","",IF(RC[-1]="","",IF(RC[-3]>RC[-1],"超节拍",IF(RC[-3]>=RC[-1]*0.9,"接近上限","正常"))))'

    $wsMeasure.Range("E$dataStart:M$dataEnd").NumberFormat = '0.00'
    $wsMeasure.Range("A$dataStart:A$dataEnd").NumberFormat = '0'
    $wsMeasure.Columns.Item('A').ColumnWidth = 10
    $wsMeasure.Columns.Item('B').ColumnWidth = 22
    $wsMeasure.Columns.Item('C').ColumnWidth = 10
    $wsMeasure.Columns.Item('D').ColumnWidth = 34
    $wsMeasure.Columns.Item('E:I').ColumnWidth = 9
    $wsMeasure.Columns.Item('J:M').ColumnWidth = 13
    $wsMeasure.Columns.Item('N').ColumnWidth = 13
    $wsMeasure.Columns.Item('O').ColumnWidth = 22
    $wsMeasure.Range("A2:O2").Font.Bold = $true
    $wsMeasure.Range("D$dataStart:D$dataEnd").WrapText = $true

    $statusRange = $wsMeasure.Range("N$dataStart:N$dataEnd")
    $statusRange.FormatConditions.Delete()
    $fcOver = $statusRange.FormatConditions.Add(1, 3, '="超节拍"')
    $fcOver.Interior.Color = 13421823
    $fcOver.Font.Color = 192
    $fcNear = $statusRange.FormatConditions.Add(1, 3, '="接近上限"')
    $fcNear.Interior.Color = 10092543
    $fcNear.Font.Color = 49407

    $wsMeasure.Activate() | Out-Null
    $wsMeasure.Range('A3').Select() | Out-Null
    $excel.ActiveWindow.FreezePanes = $true

    # ---------- 员工负荷表 ----------
    $rowLabels = @('工序序号','工序名称','ST1动作','ST2动作','ST3动作','ST4动作','ST5动作','ST1时间','ST2时间','ST3时间','ST4时间','ST5时间','总工时','目标节拍','实际节拍')
    for ($r = 0; $r -lt $rowLabels.Count; $r++) {
        $wsLoad.Cells.Item($r + 1, 1).Value2 = $rowLabels[$r]
    }
    $wsLoad.Range('A1:A15').Font.Bold = $true
    $wsLoad.Range('A1:A15').Interior.Color = 15773696

    for ($p = 1; $p -le $processCount; $p++) {
        $col = 1 + $p
        $srcStart = $dataStart + (($p - 1) * $actionsPerProcess)
        $srcEnd = $srcStart + $actionsPerProcess - 1
        $wsLoad.Cells.Item(1, $col).Formula = "=动作要素测量!A$srcStart"
        $wsLoad.Cells.Item(2, $col).Formula = "=IF(动作要素测量!B$srcStart="""","""",动作要素测量!B$srcStart)"
        for ($a = 1; $a -le $actionsPerProcess; $a++) {
            $srcRow = $srcStart + $a - 1
            $wsLoad.Cells.Item(2 + $a, $col).Formula = "=IF(动作要素测量!D$srcRow="""","""",动作要素测量!D$srcRow)"
            $wsLoad.Cells.Item(7 + $a, $col).Formula = "=IF(动作要素测量!J$srcRow="""","""",动作要素测量!J$srcRow)"
        }
        $wsLoad.Cells.Item(13, $col).Formula = "=IF(COUNT(R8C:R12C)=0,"""",SUM(R8C:R12C))"
        $wsLoad.Cells.Item(14, $col).Formula = "=IF(R13C="""","""",参数设置!R7C2)"
        $wsLoad.Cells.Item(15, $col).Formula = "=R13C"
    }

    $wsLoad.Columns.Item('A').ColumnWidth = 12
    $wsLoad.Range('B:EK').ColumnWidth = 7
    $wsLoad.Rows.Item('2:7').RowHeight = 62
    $wsLoad.Range('B2:EK7').Orientation = -4166
    $wsLoad.Range('B2:EK7').WrapText = $true
    $wsLoad.Range('B8:EK15').NumberFormat = '0.0'
    $wsLoad.Range('A1:EK15').Borders.LineStyle = 1
    $wsLoad.Range('A13:EK13').Interior.Color = 10079487
    $wsLoad.Range('A14:EK14').Interior.Color = 5287936
    $wsLoad.Range('A14:EK14').Font.Color = 16777215
    $wsLoad.Range('A15:EK15').Interior.Color = 49407
    $wsLoad.Range('A15:EK15').Font.Color = 16777215
    $wsLoad.Activate() | Out-Null
    $wsLoad.Range('B16').Select() | Out-Null
    $excel.ActiveWindow.FreezePanes = $true

    $loadChartObj = $wsLoad.ChartObjects().Add(25, 330, 1250, 420)
    $loadChart = $loadChartObj.Chart
    $loadChart.ChartType = 52
    $loadChart.HasTitle = $true
    $loadChart.ChartTitle.Text = '员工负荷堆叠图'
    $loadChart.HasLegend = $true
    $loadChart.Legend.Position = -4160
    for ($a = 1; $a -le $actionsPerProcess; $a++) {
        $row = 7 + $a
        $series = $loadChart.SeriesCollection().NewSeries()
        $series.Name = $actionNames[$a - 1]
        $series.XValues = $wsLoad.Range('$B$1:$EK$1')
        $series.Values = $wsLoad.Range("B${row}:EK${row}")
        $series.ChartType = 52
        $series.Format.Fill.ForeColor.RGB = $seriesColors[$a - 1]
        $series.Format.Line.Visible = 0
    }
    $loadTtSeries = $loadChart.SeriesCollection().NewSeries()
    $loadTtSeries.Name = '目标节拍'
    $loadTtSeries.XValues = $wsLoad.Range('$B$1:$EK$1')
    $loadTtSeries.Values = $wsLoad.Range('$B$14:$EK$14')
    $loadTtSeries.ChartType = 4
    $loadTtSeries.Format.Line.ForeColor.RGB = $ttColor
    $loadTtSeries.Format.Line.Weight = 2.25
    $loadChart.ChartGroups(1).GapWidth = 35
    $loadChart.Axes(2).HasTitle = $true
    $loadChart.Axes(2).AxisTitle.Text = 'ST(s)'

    # ---------- 山积分析 ----------
    $wsChart.Range('A1:H1').Merge()
    $wsChart.Range('A1').Value2 = '动作要素堆叠山积图'
    $wsChart.Range('A1').Font.Bold = $true
    $wsChart.Range('A1').Font.Size = 16
    $wsChart.Range('A1').HorizontalAlignment = -4108

    $wsChart.Range('A3').Value2 = '线平衡率'
    $wsChart.Range('B3').Formula = '=IFERROR(SUM(O3:O142)/(MAX(O3:O142)*COUNT(O3:O142)),"")'
    $wsChart.Range('D3').Value2 = '最大工序ST'
    $wsChart.Range('E3').Formula = '=IFERROR(MAX(O3:O142),"")'
    $wsChart.Range('G3').Value2 = '超过TT工序数'
    $wsChart.Range('H3').Formula = '=IFERROR(COUNTIF(Q3:Q142,"超节拍"),"")'
    $wsChart.Range('A4').Value2 = '平均工序ST'
    $wsChart.Range('B4').Formula = '=IFERROR(AVERAGE(O3:O142),"")'
    $wsChart.Range('D4').Value2 = '说明'
    $wsChart.Range('E4').Value2 = '柱形=ST1到ST5平均ST堆叠；折线=TT基准。'
    $wsChart.Range('B3').NumberFormat = '0.0%'
    $wsChart.Range('E3,B4').NumberFormat = '0.00'

    $helperHeaders = @('工序','ST1','ST2','ST3','ST4','ST5','工序总ST','TT基准','状态')
    for ($c = 0; $c -lt $helperHeaders.Count; $c++) {
        $wsChart.Cells.Item(2, 9 + $c).Value2 = $helperHeaders[$c]
    }

    for ($p = 1; $p -le $processCount; $p++) {
        $row = 2 + $p
        $srcStart = $dataStart + (($p - 1) * $actionsPerProcess)
        $wsChart.Cells.Item($row, 9).Formula = "=IF(动作要素测量!B$srcStart<>"""",动作要素测量!A$srcStart&""-""&动作要素测量!B$srcStart,动作要素测量!A$srcStart)"
        for ($a = 1; $a -le $actionsPerProcess; $a++) {
            $srcRow = $srcStart + $a - 1
            $wsChart.Cells.Item($row, 9 + $a).Formula = "=IF(动作要素测量!J$srcRow="""","""",动作要素测量!J$srcRow)"
        }
        $wsChart.Cells.Item($row, 15).Formula = "=IF(COUNT(J${row}:N${row})=0,"""",SUM(J${row}:N${row}))"
        $wsChart.Cells.Item($row, 16).Formula = "=IF(O${row}="""","""",参数设置!B7)"
        $wsChart.Cells.Item($row, 17).Formula = "=IF(O${row}="""","""",IF(P${row}="""","""",IF(O${row}>P${row},""超节拍"",IF(O${row}>=P${row}*0.9,""接近上限"",""正常""))))"
    }

    $wsChart.Columns.Item('A').ColumnWidth = 13
    $wsChart.Columns.Item('B').ColumnWidth = 12
    $wsChart.Columns.Item('D').ColumnWidth = 13
    $wsChart.Columns.Item('E').ColumnWidth = 16
    $wsChart.Columns.Item('G').ColumnWidth = 14
    $wsChart.Columns.Item('H').ColumnWidth = 12
    $wsChart.Columns.Item('I').ColumnWidth = 20
    $wsChart.Columns.Item('J:Q').ColumnWidth = 11
    $wsChart.Range('A3:H4').Borders.LineStyle = 1
    $wsChart.Range('I2:Q142').Borders.LineStyle = 1
    $wsChart.Range('I2:Q2').Font.Bold = $true
    $wsChart.Range('I2:Q2').Interior.Color = 15773696
    $wsChart.Range('J3:P142').NumberFormat = '0.00'

    $chartObj = $wsChart.ChartObjects().Add(25, 120, 1150, 500)
    $chart = $chartObj.Chart
    $chart.ChartType = 52
    $chart.HasTitle = $true
    $chart.ChartTitle.Text = '工序动作要素堆叠山积图'
    $chart.HasLegend = $true
    $chart.Legend.Position = -4160

    for ($a = 1; $a -le $actionsPerProcess; $a++) {
        $colLetter = [char]([int][char]'J' + $a - 1)
        $series = $chart.SeriesCollection().NewSeries()
        $series.Formula = "=SERIES(""$($actionNames[$a - 1])"",山积分析!`$I`$3:`$I`$142,山积分析!`$$colLetter`$3:`$$colLetter`$142,$a)"
        $series.XValues = $wsChart.Range('$I$3:$I$142')
        $series.ChartType = 52
        $series.Format.Fill.ForeColor.RGB = $seriesColors[$a - 1]
        $series.Format.Line.Visible = 0
        $series.HasDataLabels = $true
        $series.DataLabels().NumberFormat = '0.0'
        $series.DataLabels().Position = -4108
        $series.DataLabels().Font.Size = 8
        $series.DataLabels().Font.Color = 16777215
    }
    $ttSeries = $chart.SeriesCollection().NewSeries()
    $ttSeries.Formula = '=SERIES("TT基准",山积分析!$I$3:$I$142,山积分析!$P$3:$P$142,6)'
    $ttSeries.XValues = $wsChart.Range('$I$3:$I$142')
    $ttSeries.ChartType = 4
    $ttSeries.Format.Line.ForeColor.RGB = $ttColor
    $ttSeries.Format.Line.Weight = 2.25
    $ttSeries.HasDataLabels = $false

    $chart.ChartGroups(1).GapWidth = 35
    $chart.Axes(2).HasTitle = $true
    $chart.Axes(2).AxisTitle.Text = 'ST(s)'

    $wsChart.Activate() | Out-Null
    $wsChart.Range('A1').Select() | Out-Null

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




