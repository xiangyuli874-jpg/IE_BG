$ErrorActionPreference = 'Stop'

$outputDir = 'E:\AI\gongshibiao\outputs\worktime_action_element_template'
$outputPath = Join-Path $outputDir '动作要素工时与堆叠山积图模板_v3.xlsm'
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

$moduleCode = @'
Option Explicit

Public Const SHEET_INTRO As String = "使用说明"
Public Const SHEET_PARAM As String = "参数设置"
Public Const SHEET_MEASURE As String = "动作要素测量"
Public Const SHEET_LOAD As String = "员工负荷表"
Public Const SHEET_CHART As String = "山积分析"
Public Const DATA_START As Long = 3
Public Const ACTIONS_PER_PROCESS As Long = 5

Public Sub AddProcessAfterSelection()
    If ActiveSheet.Name <> SHEET_MEASURE Then
        MsgBox "请先在“动作要素测量”表中选中某个工序的任意一行。", vbInformation
        Exit Sub
    End If
    If ActiveCell.Row < DATA_START Or ActiveCell.Row > LastDataRow() Then
        MsgBox "请选择某个工序的数据行。", vbInformation
        Exit Sub
    End If

    Dim ws As Worksheet
    Dim insertAt As Long
    Dim oldLastRow As Long
    Set ws = ThisWorkbook.Worksheets(SHEET_MEASURE)
    insertAt = GroupStartRow(ActiveCell.Row) + ACTIONS_PER_PROCESS
    oldLastRow = LastDataRow()

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ws.Rows(insertAt & ":" & insertAt + ACTIONS_PER_PROCESS - 1).Insert Shift:=xlDown
    ResizeMeasureTable oldLastRow + ACTIONS_PER_PROCESS
    RefreshAll
    ws.Activate
    ws.Cells(insertAt, 2).Select
CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Sub DeleteSelectedProcess(Optional ByVal skipConfirm As Boolean = False)
    If ActiveSheet.Name <> SHEET_MEASURE Then
        MsgBox "请先在“动作要素测量”表中选中某个工序的任意一行。", vbInformation
        Exit Sub
    End If
    If ActiveCell.Row < DATA_START Or ActiveCell.Row > LastDataRow() Then
        MsgBox "请选择某个工序的数据行。", vbInformation
        Exit Sub
    End If
    If ProcessCount() <= 1 Then
        MsgBox "至少保留 1 个工序。", vbInformation
        Exit Sub
    End If
    If Not skipConfirm Then
        If MsgBox("确认删除当前工序的 5 行动作吗？", vbQuestion + vbYesNo) <> vbYes Then Exit Sub
    End If

    Dim ws As Worksheet
    Dim startRow As Long
    Dim oldLastRow As Long
    Set ws = ThisWorkbook.Worksheets(SHEET_MEASURE)
    startRow = GroupStartRow(ActiveCell.Row)
    oldLastRow = LastDataRow()

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ws.Rows(startRow & ":" & startRow + ACTIONS_PER_PROCESS - 1).Delete
    ResizeMeasureTable oldLastRow - ACTIONS_PER_PROCESS
    RefreshAll
    ws.Activate
    ws.Cells(Application.Min(startRow, LastDataRow()), 2).Select
CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Sub RefreshAll()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    EnsureValidMeasureShape
    RefreshActionRows
    BuildLoadReport
    BuildChartAnalysis
CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Sub AutoExtendFromBottom(ByVal changedRow As Long, ByVal changedCol As Long, ByVal changedValue As Variant)
    Dim currentLast As Long
    Dim totalRows As Long
    Dim remainder As Long
    currentLast = LastDataRow()
    If changedCol < 2 Or changedCol > 9 Then Exit Sub
    If changedRow < DATA_START Then Exit Sub

    totalRows = currentLast - DATA_START + 1
    remainder = totalRows Mod ACTIONS_PER_PROCESS
    If changedRow <= currentLast And remainder <> 0 Then
        currentLast = currentLast + (ACTIONS_PER_PROCESS - remainder)
    ElseIf changedRow = currentLast + 1 Then
        currentLast = currentLast + ACTIONS_PER_PROCESS
    Else
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ResizeMeasureTable currentLast
    ThisWorkbook.Worksheets(SHEET_MEASURE).Cells(changedRow, changedCol).Value = changedValue
    RefreshAll
CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Public Function ProcessCount() As Long
    ProcessCount = (LastDataRow() - DATA_START + 1) \ ACTIONS_PER_PROCESS
End Function

Public Function ActiveReportProcessCount() As Long
    Dim ws As Worksheet
    Dim p As Long, a As Long, c As Long
    Dim r As Long
    Dim hasUserInput As Boolean
    Dim lastActive As Long
    Set ws = ThisWorkbook.Worksheets(SHEET_MEASURE)

    For p = 1 To ProcessCount()
        hasUserInput = False
        For a = 0 To ACTIONS_PER_PROCESS - 1
            r = DATA_START + ((p - 1) * ACTIONS_PER_PROCESS) + a
            If Len(Trim(CStr(ws.Cells(r, 2).Value))) > 0 Then hasUserInput = True
            If Len(Trim(CStr(ws.Cells(r, 4).Value))) > 0 Then hasUserInput = True
            If Len(Trim(CStr(ws.Cells(r, 15).Value))) > 0 Then hasUserInput = True
            For c = 5 To 9
                If Len(Trim(CStr(ws.Cells(r, c).Value))) > 0 Then hasUserInput = True
            Next c
        Next a
        If hasUserInput Then lastActive = p
    Next p

    ActiveReportProcessCount = lastActive
End Function

Public Function LastDataRow() As Long
    Dim ws As Worksheet
    Dim lo As ListObject
    Set ws = ThisWorkbook.Worksheets(SHEET_MEASURE)
    Set lo = ws.ListObjects("tblActionElements")
    LastDataRow = lo.Range.Row + lo.Range.Rows.Count - 1
End Function

Private Function GroupStartRow(ByVal rowNumber As Long) As Long
    GroupStartRow = DATA_START + ((rowNumber - DATA_START) \ ACTIONS_PER_PROCESS) * ACTIONS_PER_PROCESS
End Function

Private Sub EnsureValidMeasureShape()
    Dim totalRows As Long
    Dim remainder As Long
    totalRows = LastDataRow() - DATA_START + 1
    remainder = totalRows Mod ACTIONS_PER_PROCESS
    If remainder <> 0 Then
        ResizeMeasureTable LastDataRow() + (ACTIONS_PER_PROCESS - remainder)
    End If
End Sub

Private Sub ResizeMeasureTable(ByVal newLastRow As Long)
    Dim ws As Worksheet
    Dim lo As ListObject
    Set ws = ThisWorkbook.Worksheets(SHEET_MEASURE)
    Set lo = ws.ListObjects("tblActionElements")
    If newLastRow < DATA_START + ACTIONS_PER_PROCESS - 1 Then newLastRow = DATA_START + ACTIONS_PER_PROCESS - 1
    lo.Resize ws.Range("A2:O" & newLastRow)
End Sub

Private Sub RefreshActionRows()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim processNo As Long
    Dim actionNo As Long
    Set ws = ThisWorkbook.Worksheets(SHEET_MEASURE)
    lastRow = LastDataRow()

    For r = DATA_START To lastRow
        processNo = ((r - DATA_START) \ ACTIONS_PER_PROCESS) + 1
        actionNo = ((r - DATA_START) Mod ACTIONS_PER_PROCESS) + 1
        ws.Cells(r, 1).Value = processNo
        ws.Cells(r, 3).Value = "ST" & actionNo
    Next r

    ws.Range("J" & DATA_START & ":J" & lastRow).FormulaR1C1 = "=IF(COUNT(RC[-5]:RC[-1])=0,"""",AVERAGE(RC[-5]:RC[-1]))"
    ws.Range("K" & DATA_START & ":K" & lastRow).FormulaR1C1 = "=IF(RC[-8]<>""ST1"","""",IF(COUNTIFS(C1,RC[-10],C10,""<>"")=0,"""",SUMIFS(C10,C1,RC[-10])))"
    ws.Range("L" & DATA_START & ":L" & lastRow).FormulaR1C1 = "=IF(RC[-9]<>""ST1"","""",IF(RC[-1]="""","""",RC[-1]*(1+" & SHEET_PARAM & "!R6C2)))"
    ws.Range("M" & DATA_START & ":M" & lastRow).FormulaR1C1 = "=IF(RC[-12]="""",""""," & SHEET_PARAM & "!R7C2)"
    ws.Range("N" & DATA_START & ":N" & lastRow).FormulaR1C1 = "=IF(RC[-3]="""","""",IF(RC[-1]="""","""",IF(RC[-3]>RC[-1],""超节拍"",IF(RC[-3]>=RC[-1]*0.9,""接近上限"",""正常""))))"
    ws.Range("E" & DATA_START & ":M" & lastRow).NumberFormat = "0.00"
End Sub

Private Sub BuildLoadReport()
    Dim ws As Worksheet
    Dim p As Long, a As Long
    Dim n As Long, col As Long
    Dim srcStart As Long, srcRow As Long
    Dim endCol As Long
    Dim labels As Variant
    Dim chartObj As ChartObject
    Dim s As Series
    Dim colors As Variant

    Set ws = ThisWorkbook.Worksheets(SHEET_LOAD)
    n = ActiveReportProcessCount()
    endCol = n + 1
    labels = Array("工序序号", "工序名称", "ST1动作", "ST2动作", "ST3动作", "ST4动作", "ST5动作", "ST1时间", "ST2时间", "ST3时间", "ST4时间", "ST5时间", "总工时", "目标节拍", "实际节拍")
    colors = Array(RGB(31, 78, 121), RGB(192, 0, 0), RGB(80, 136, 69), RGB(112, 48, 160), RGB(237, 125, 49))

    ws.Cells.Clear
    For Each chartObj In ws.ChartObjects
        chartObj.Delete
    Next chartObj

    Dim i As Long
    For i = 0 To UBound(labels)
        ws.Cells(i + 1, 1).Value = labels(i)
    Next i

    For p = 1 To n
        col = p + 1
        srcStart = DATA_START + ((p - 1) * ACTIONS_PER_PROCESS)
        ws.Cells(1, col).Formula = "='" & SHEET_MEASURE & "'!A" & srcStart
        ws.Cells(2, col).Formula = "=IF('" & SHEET_MEASURE & "'!B" & srcStart & "="""","""",'" & SHEET_MEASURE & "'!B" & srcStart & ")"
        For a = 1 To ACTIONS_PER_PROCESS
            srcRow = srcStart + a - 1
            ws.Cells(2 + a, col).Formula = "=IF('" & SHEET_MEASURE & "'!D" & srcRow & "="""","""",'" & SHEET_MEASURE & "'!D" & srcRow & ")"
            ws.Cells(7 + a, col).Formula = "=IF('" & SHEET_MEASURE & "'!J" & srcRow & "="""","""",'" & SHEET_MEASURE & "'!J" & srcRow & ")"
        Next a
        ws.Cells(13, col).FormulaR1C1 = "=IF(COUNT(R8C:R12C)=0,"""",SUM(R8C:R12C))"
        ws.Cells(14, col).Formula = "=IF(" & ws.Cells(13, col).Address(False, False) & "="""","""",'" & SHEET_PARAM & "'!B7)"
        ws.Cells(15, col).Formula = "=" & ws.Cells(13, col).Address(False, False)
    Next p

    ws.Columns(1).ColumnWidth = 12
    ws.Rows("2:2").RowHeight = 62
    ws.Rows("3:7").RowHeight = 42
    ws.Range("A1:A15").Font.Bold = True
    ws.Range("A1:A15").Interior.Color = RGB(0, 176, 240)
    ws.Range("A1:A15").Borders.LineStyle = 1
    If n = 0 Then Exit Sub

    ws.Range(ws.Cells(1, 2), ws.Cells(15, endCol)).ColumnWidth = 8
    ws.Range(ws.Cells(1, 1), ws.Cells(15, endCol)).Borders.LineStyle = 1
    ws.Range(ws.Cells(2, 2), ws.Cells(2, endCol)).Orientation = xlVertical
    ws.Range(ws.Cells(3, 2), ws.Cells(7, endCol)).Orientation = xlHorizontal
    ws.Range(ws.Cells(3, 2), ws.Cells(7, endCol)).WrapText = True
    ws.Range(ws.Cells(8, 2), ws.Cells(15, endCol)).NumberFormat = "0.0"
    ws.Range(ws.Cells(13, 1), ws.Cells(13, endCol)).Interior.Color = RGB(255, 230, 153)
    ws.Range(ws.Cells(14, 1), ws.Cells(14, endCol)).Interior.Color = RGB(0, 176, 80)
    ws.Range(ws.Cells(14, 1), ws.Cells(14, endCol)).Font.Color = RGB(255, 255, 255)
    ws.Range(ws.Cells(15, 1), ws.Cells(15, endCol)).Interior.Color = RGB(255, 192, 0)

    Set chartObj = ws.ChartObjects.Add(25, 330, 1250, 420)
    With chartObj.Chart
        .ChartType = xlColumnStacked
        .HasTitle = True
        .ChartTitle.Text = "员工负荷堆叠图"
        .HasLegend = True
        .Legend.Position = xlLegendPositionTop
        For a = 1 To ACTIONS_PER_PROCESS
            Set s = .SeriesCollection.NewSeries
            s.Name = "ST" & a
            s.XValues = ws.Range(ws.Cells(1, 2), ws.Cells(1, endCol))
            s.Values = ws.Range(ws.Cells(7 + a, 2), ws.Cells(7 + a, endCol))
            s.ChartType = xlColumnStacked
            s.Format.Fill.ForeColor.RGB = colors(a - 1)
            s.Format.Line.Visible = msoFalse
        Next a
        Set s = .SeriesCollection.NewSeries
        s.Name = "目标节拍"
        s.XValues = ws.Range(ws.Cells(1, 2), ws.Cells(1, endCol))
        s.Values = ws.Range(ws.Cells(14, 2), ws.Cells(14, endCol))
        s.ChartType = xlLine
        s.Format.Line.ForeColor.RGB = RGB(255, 102, 0)
        s.Format.Line.Weight = 2.25
        .ChartGroups(1).GapWidth = 35
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "ST(s)"
    End With
End Sub

Private Sub BuildChartAnalysis()
    Dim ws As Worksheet
    Dim p As Long, a As Long
    Dim n As Long, row As Long
    Dim srcStart As Long, srcRow As Long
    Dim chartObj As ChartObject
    Dim s As Series
    Dim colors As Variant

    Set ws = ThisWorkbook.Worksheets(SHEET_CHART)
    n = ActiveReportProcessCount()
    colors = Array(RGB(31, 78, 121), RGB(192, 0, 0), RGB(80, 136, 69), RGB(112, 48, 160), RGB(237, 125, 49))

    ws.Cells.Clear
    For Each chartObj In ws.ChartObjects
        chartObj.Delete
    Next chartObj

    ws.Range("A1:H1").Merge
    ws.Range("A1").Value = "动作要素堆叠山积图"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
    ws.Range("A1").HorizontalAlignment = xlCenter

    ws.Range("A3").Value = "线平衡率"
    If n = 0 Then
        ws.Range("B3").Value = ""
    Else
        ws.Range("B3").Formula = "=IFERROR(SUM(O3:O" & (n + 2) & ")/(MAX(O3:O" & (n + 2) & ")*COUNT(O3:O" & (n + 2) & ")),"""")"
    End If
    ws.Range("D3").Value = "最大工序ST"
    If n = 0 Then
        ws.Range("E3").Value = ""
    Else
        ws.Range("E3").Formula = "=IFERROR(MAX(O3:O" & (n + 2) & "),"""")"
    End If
    ws.Range("G3").Value = "超过TT工序数"
    If n = 0 Then
        ws.Range("H3").Value = ""
    Else
        ws.Range("H3").Formula = "=IFERROR(COUNTIF(Q3:Q" & (n + 2) & ",""超节拍""),"""")"
    End If
    ws.Range("A4").Value = "平均工序ST"
    If n = 0 Then
        ws.Range("B4").Value = ""
    Else
        ws.Range("B4").Formula = "=IFERROR(AVERAGE(O3:O" & (n + 2) & "),"""")"
    End If
    ws.Range("D4").Value = "说明"
    ws.Range("E4").Value = "柱形=ST1到ST5平均ST堆叠；折线=TT基准。"

    ws.Range("I2:Q2").Value = Array("工序", "ST1", "ST2", "ST3", "ST4", "ST5", "工序总ST", "TT基准", "状态")
    If n = 0 Then
        ws.Range("A3:H4").Borders.LineStyle = 1
        ws.Range("I2:Q2").Font.Bold = True
        ws.Range("I2:Q2").Interior.Color = RGB(0, 176, 240)
        Exit Sub
    End If

    For p = 1 To n
        row = p + 2
        srcStart = DATA_START + ((p - 1) * ACTIONS_PER_PROCESS)
        ws.Cells(row, 9).Formula = "=IF('" & SHEET_MEASURE & "'!B" & srcStart & "="""",'" & SHEET_MEASURE & "'!A" & srcStart & ",'" & SHEET_MEASURE & "'!A" & srcStart & "&""-""&'" & SHEET_MEASURE & "'!B" & srcStart & ")"
        For a = 1 To ACTIONS_PER_PROCESS
            srcRow = srcStart + a - 1
            ws.Cells(row, 9 + a).Formula = "=IF('" & SHEET_MEASURE & "'!J" & srcRow & "="""","""",'" & SHEET_MEASURE & "'!J" & srcRow & ")"
        Next a
        ws.Cells(row, 15).Formula = "=IF(COUNT(J" & row & ":N" & row & ")=0,"""",SUM(J" & row & ":N" & row & "))"
        ws.Cells(row, 16).Formula = "=IF(O" & row & "="""","""",'" & SHEET_PARAM & "'!B7)"
        ws.Cells(row, 17).Formula = "=IF(O" & row & "="""","""",IF(P" & row & "="""","""",IF(O" & row & ">P" & row & ",""超节拍"",IF(O" & row & ">=P" & row & "*0.9,""接近上限"",""正常""))))"
    Next p

    ws.Range("B3").NumberFormat = "0.0%"
    ws.Range("E3,B4").NumberFormat = "0.00"
    ws.Range(ws.Cells(3, 10), ws.Cells(n + 2, 16)).NumberFormat = "0.00"
    ws.Range("A3:H4").Borders.LineStyle = 1
    ws.Range(ws.Cells(2, 9), ws.Cells(n + 2, 17)).Borders.LineStyle = 1
    ws.Range("I2:Q2").Font.Bold = True
    ws.Range("I2:Q2").Interior.Color = RGB(0, 176, 240)

    Set chartObj = ws.ChartObjects.Add(25, 120, 1150, 500)
    With chartObj.Chart
        .ChartType = xlColumnStacked
        .HasTitle = True
        .ChartTitle.Text = "工序动作要素堆叠山积图"
        .HasLegend = True
        .Legend.Position = xlLegendPositionTop
        For a = 1 To ACTIONS_PER_PROCESS
            Set s = .SeriesCollection.NewSeries
            s.Name = "ST" & a
            s.XValues = ws.Range(ws.Cells(3, 9), ws.Cells(n + 2, 9))
            s.Values = ws.Range(ws.Cells(3, 9 + a), ws.Cells(n + 2, 9 + a))
            s.ChartType = xlColumnStacked
            s.Format.Fill.ForeColor.RGB = colors(a - 1)
            s.Format.Line.Visible = msoFalse
            FormatNonZeroDataLabels s
        Next a
        Set s = .SeriesCollection.NewSeries
        s.Name = "TT基准"
        s.XValues = ws.Range(ws.Cells(3, 9), ws.Cells(n + 2, 9))
        s.Values = ws.Range(ws.Cells(3, 16), ws.Cells(n + 2, 16))
        s.ChartType = xlLine
        s.Format.Line.ForeColor.RGB = RGB(255, 102, 0)
        s.Format.Line.Weight = 2.25
        s.HasDataLabels = False
        .ChartGroups(1).GapWidth = 35
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "ST(s)"
    End With
End Sub

Private Sub FormatNonZeroDataLabels(ByVal s As Series)
    s.HasDataLabels = True
    With s.DataLabels
        .NumberFormat = "0.0;-0.0;;@"
        .Position = xlLabelPositionCenter
        .Font.Name = "微软雅黑"
        .Font.Size = 10
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
    End With
End Sub
'@

$sheetCode = @'
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit
    If Target.CountLarge > 1 Then Exit Sub
    If Target.Row < modActionElementTemplate.DATA_START Then Exit Sub
    If Target.Column < 2 Or Target.Column > 9 Then Exit Sub

    Dim v As Variant
    Dim r As Long
    Dim c As Long
    v = Target.Value
    r = Target.Row
    c = Target.Column
    modActionElementTemplate.AutoExtendFromBottom r, c, v
SafeExit:
End Sub
'@

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $wb = $excel.Workbooks.Add()

    try {
        $null = $wb.VBProject
    }
    catch {
        throw 'Excel 当前仍不能访问 VBProject，请确认已勾选“信任对 VBA 工程对象模型的访问”，并重新启动 Excel 后再试。'
    }

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

    $processCount = 25
    $actionsPerProcess = 5
    $dataStart = 3
    $dataRows = $processCount * $actionsPerProcess
    $dataEnd = $dataStart + $dataRows - 1
    $actionNames = @('ST1','ST2','ST3','ST4','ST5')

    # ---------- 使用说明 ----------
    $wsIntro.Range('A1:H1').Merge()
    $wsIntro.Range('A1').Value2 = '动作要素工时与员工负荷表宏版模板'
    $wsIntro.Range('A1').Font.Bold = $true
    $wsIntro.Range('A1').Font.Size = 18
    $wsIntro.Range('A1').HorizontalAlignment = -4108
    $intro = @(
        @('用途', '宏版模板：录入动作要素测量时间，自动计算平均ST、工序总ST、宽放ST，并维护员工负荷表和山积图。'),
        @('默认容量', '默认 25 个工序，每个工序 5 行动作。需要更多工序时，可在表格底部继续填写，也可用按钮在中间新增。'),
        @('新增工序', '选中某工序任意一行，点击“新增工序”，会在当前工序后插入一组 5 行。'),
        @('删除工序', '选中某工序任意一行，点击“删除工序”，会删除当前工序的 5 行。'),
        @('底部扩展', '在最后一个工序下方第一行填写工序名称、动作描述或测量值时，会自动新增下一组 5 行。')
    )
    for ($i = 0; $i -lt $intro.Count; $i++) {
        $row = 3 + $i
        $wsIntro.Cells.Item($row, 1).Value2 = $intro[$i][0]
        $wsIntro.Cells.Item($row, 2).Value2 = $intro[$i][1]
    }
    $wsIntro.Columns.Item('A').ColumnWidth = 18
    $wsIntro.Columns.Item('B').ColumnWidth = 100
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
    $wsParam.Range('A2:C2').Interior.Color = RgbValue 0 176 240

    # ---------- 动作要素测量 ----------
    $wsMeasure.Range('A1:R1').Merge()
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
    $wsMeasure.Range("E$dataStart:M$dataEnd").NumberFormat = '0.00'
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

    $btnTop = $wsMeasure.Range('Q2').Top
    $btnLeft = $wsMeasure.Range('Q2').Left
    $btn = $wsMeasure.Buttons().Add($btnLeft, $btnTop, 80, 24)
    $btn.Caption = '新增工序'
    $btn.OnAction = 'AddProcessAfterSelection'
    $btn = $wsMeasure.Buttons().Add($btnLeft + 90, $btnTop, 80, 24)
    $btn.Caption = '删除工序'
    $btn.OnAction = 'DeleteSelectedProcess'
    $btn = $wsMeasure.Buttons().Add($btnLeft + 180, $btnTop, 80, 24)
    $btn.Caption = '刷新报表'
    $btn.OnAction = 'RefreshAll'
    $wsMeasure.Columns.Item('Q:S').ColumnWidth = 12

    # ---------- VBA ----------
    $vbProject = $wb.VBProject
    $module = $vbProject.VBComponents.Add(1)
    $module.Name = 'modActionElementTemplate'
    $module.CodeModule.AddFromString($moduleCode)
    $sheetModule = $vbProject.VBComponents.Item($wsMeasure.CodeName)
    $sheetModule.CodeModule.AddFromString($sheetCode)

    $excel.Run('RefreshAll')
    $wsMeasure.Activate() | Out-Null
    $wsMeasure.Range('A3').Select() | Out-Null
    $excel.ActiveWindow.FreezePanes = $true

    $wb.SaveAs($outputPath, 52)
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







