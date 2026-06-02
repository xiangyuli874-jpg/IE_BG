$ErrorActionPreference = 'Stop'

$outputDir = 'E:\AI\gongshibiao\outputs\worktime_action_element_template'
$outputPath = Join-Path $outputDir '动作要素工时与堆叠山积图模板_v1.xlsx'
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
    $wsMeasure.Name = '动作要素测量'
    $wsChart.Name = '山积分析'

    while ($wb.Worksheets.Count -gt 4) {
        $wb.Worksheets.Item($wb.Worksheets.Count).Delete()
    }

    # ---------- 使用说明 ----------
    $wsIntro.Range('A1:H1').Merge()
    $wsIntro.Range('A1').Value2 = '动作要素工时与堆叠山积图模板'
    $wsIntro.Range('A1').Font.Bold = $true
    $wsIntro.Range('A1').Font.Size = 18
    $wsIntro.Range('A1').HorizontalAlignment = -4108

    $intro = @(
        @('用途', '用于录入单个工序下的多个动作要素测量时间，自动计算动作平均ST、工序总ST、宽放ST，并生成动作要素堆叠山积图。'),
        @('容量', '第一版无宏稳定模板，预置 140 个工序，每个工序 5 行动作要素，共 700 行。'),
        @('填写方式', '在“动作要素测量”中按工序号填写工序名称、动作单元描述和 ST1-ST5。每个工序最多填写 5 个动作。'),
        @('图表口径', '山积图柱形只展示真实动作平均ST的堆叠；宽放ST仅作为工序级补充列，不参与柱形图。'),
        @('超过140工序', '当前版本不使用宏自动插入第141个工序。如超过容量，建议后续生成 200 或 300 工序扩容版。')
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

    $processCount = 140
    $actionsPerProcess = 5
    $dataStart = 3
    $dataRows = $processCount * $actionsPerProcess
    $dataEnd = $dataStart + $dataRows - 1

    for ($p = 1; $p -le $processCount; $p++) {
        for ($a = 1; $a -le $actionsPerProcess; $a++) {
            $row = $dataStart + (($p - 1) * $actionsPerProcess) + ($a - 1)
            $wsMeasure.Cells.Item($row, 1).Value2 = [double]$p
            $wsMeasure.Cells.Item($row, 3).Value2 = [double]$a
        }
    }

    $tableRange = $wsMeasure.Range("A2:O$dataEnd")
    $table = $wsMeasure.ListObjects.Add(1, $tableRange, $null, 1)
    $table.Name = 'tblActionElements'
    $table.TableStyle = 'TableStyleMedium2'

    $wsMeasure.Range("J$dataStart:J$dataEnd").FormulaR1C1 = '=IF(COUNT(RC[-5]:RC[-1])=0,"",AVERAGE(RC[-5]:RC[-1]))'
    $wsMeasure.Range("K$dataStart:K$dataEnd").FormulaR1C1 = '=IF(RC[-8]<>1,"",IF(COUNTIFS(C1,RC[-10],C10,"<>")=0,"",SUMIFS(C10,C1,RC[-10])))'
    $wsMeasure.Range("L$dataStart:L$dataEnd").FormulaR1C1 = '=IF(RC[-9]<>1,"",IF(RC[-1]="","",RC[-1]*(1+参数设置!R6C2)))'
    $wsMeasure.Range("M$dataStart:M$dataEnd").FormulaR1C1 = '=IF(RC[-12]="","",参数设置!R7C2)'
    $wsMeasure.Range("N$dataStart:N$dataEnd").FormulaR1C1 = '=IF(RC[-3]="","",IF(RC[-1]="","",IF(RC[-3]>RC[-1],"超节拍",IF(RC[-3]>=RC[-1]*0.9,"接近上限","正常"))))'

    $wsMeasure.Range("E$dataStart:M$dataEnd").NumberFormat = '0.00'
    $wsMeasure.Range("A$dataStart:A$dataEnd").NumberFormat = '0'
    $wsMeasure.Range("C$dataStart:C$dataEnd").NumberFormat = '0'
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
    $wsChart.Range('E4').Value2 = '柱形=动作1到动作5平均ST堆叠；折线=TT基准。'
    $wsChart.Range('B3').NumberFormat = '0.0%'
    $wsChart.Range('E3,B4').NumberFormat = '0.00'

    $helperHeaders = @('工序','动作1','动作2','动作3','动作4','动作5','工序总ST','TT基准','状态')
    for ($c = 0; $c -lt $helperHeaders.Count; $c++) {
        $wsChart.Cells.Item(2, 9 + $c).Value2 = $helperHeaders[$c]
    }

    for ($p = 1; $p -le $processCount; $p++) {
        $row = 2 + $p
        $srcStart = $dataStart + (($p - 1) * $actionsPerProcess)
        $srcEnd = $srcStart + $actionsPerProcess - 1
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
        $series.Formula = "=SERIES(""动作$a"",山积分析!`$I`$3:`$I`$142,山积分析!`$$colLetter`$3:`$$colLetter`$142,$a)"
        $series.XValues = $wsChart.Range('$I$3:$I$142')
        $series.ChartType = 52
    }
    $ttSeries = $chart.SeriesCollection().NewSeries()
    $ttSeries.Formula = '=SERIES("TT基准",山积分析!$I$3:$I$142,山积分析!$P$3:$P$142,6)'
    $ttSeries.XValues = $wsChart.Range('$I$3:$I$142')
    $ttSeries.ChartType = 4
    $ttSeries.Format.Line.ForeColor.RGB = 49407
    $ttSeries.Format.Line.Weight = 2.25

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



