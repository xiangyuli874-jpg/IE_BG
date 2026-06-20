Attribute VB_Name = "modVisualization"
Option Explicit

Public Sub WriteScheduleOutput(ByVal schedule As Collection)
    Dim sheet As Worksheet, table As ListObject, target As Range
    Dim headers As Variant, task As Object, rowIndex As Long
    Set sheet = ThisWorkbook.Worksheets("自动排程")
    headers = Array("任务编号", "循环", "设备", "步骤", "作业内容", "执行人员", _
        "开始(s)", "结束(s)", "持续(s)", "类型", "前置任务", "是否锁定", "等待原因")

    On Error Resume Next
    Set table = sheet.ListObjects("tblSchedule")
    On Error GoTo 0
    If Not table Is Nothing Then table.Delete
    sheet.Range("A8:M1000").ClearContents
    For rowIndex = 0 To UBound(headers)
        sheet.Cells(8, rowIndex + 1).Value2 = headers(rowIndex)
    Next rowIndex
    rowIndex = 9
    For Each task In schedule
        sheet.Cells(rowIndex, 1).Value2 = task("TaskId")
        sheet.Cells(rowIndex, 2).Value2 = task("CycleNo")
        sheet.Cells(rowIndex, 3).Value2 = task("DeviceId")
        sheet.Cells(rowIndex, 4).Value2 = task("StepNo")
        sheet.Cells(rowIndex, 5).Value2 = task("StepName")
        sheet.Cells(rowIndex, 6).Value2 = task("PersonId")
        sheet.Cells(rowIndex, 7).Value2 = task("StartSec")
        sheet.Cells(rowIndex, 8).Value2 = task("EndSec")
        sheet.Cells(rowIndex, 9).Value2 = task("DurationSec")
        sheet.Cells(rowIndex, 10).Value2 = DisplayStepType(CStr(task("StepType")))
        sheet.Cells(rowIndex, 11).Value2 = task("PredecessorIds")
        sheet.Cells(rowIndex, 12).Value2 = IIf(CBool(task("IsLocked")), "是", "否")
        sheet.Cells(rowIndex, 13).Value2 = IIf(CDbl(task("WaitSec")) > 0, _
            "等待资源 " & Format$(CDbl(task("WaitSec")), "0.0") & "s", "")
        rowIndex = rowIndex + 1
    Next task
    If rowIndex = 9 Then rowIndex = 10
    Set target = sheet.Range("A8:M" & rowIndex - 1)
    Set table = sheet.ListObjects.Add(xlSrcRange, target, , xlYes)
    table.Name = "tblSchedule"
    table.TableStyle = "TableStyleMedium2"
    sheet.Columns("A:M").AutoFit
    sheet.Columns("E").ColumnWidth = 20
    sheet.Columns("K:M").ColumnWidth = 18
End Sub

Public Sub RefreshGanttView(Optional ByVal viewMode As String = "全部")
    Dim sheet As Worksheet, schedule As Collection, task As Object
    Dim rowIndex As Long, chartObject As ChartObject, chart As Chart
    Dim pointIndex As Long
    Set sheet = ThisWorkbook.Worksheets("人机作业图")
    If gCurrentSchedule Is Nothing Then
        If SnapshotExists("UNDO") Then Set gCurrentSchedule = RestoreScheduleSnapshot("UNDO")
    End If
    If gCurrentSchedule Is Nothing Then Exit Sub
    Set schedule = gCurrentSchedule

    sheet.Range("A4:D1000").ClearContents
    sheet.Range("A4:D4").Value = Array("资源与任务", "开始偏移", "持续时间", "类型")
    rowIndex = 5
    If viewMode = "全部" Or viewMode = "仅设备" Then
        For Each task In schedule
            WriteGanttRow sheet, rowIndex, "设备 " & task("DeviceId") & "｜" & _
                task("StepName"), task
        Next task
    End If
    If viewMode = "全部" Or viewMode = "仅人员" Then
        For Each task In schedule
            If Len(CStr(task("PersonId"))) > 0 Then
                WriteGanttRow sheet, rowIndex, "人员 " & task("PersonId") & "｜" & _
                    task("StepName"), task
            End If
        Next task
    End If

    On Error Resume Next
    sheet.ChartObjects("chtManMachine").Delete
    On Error GoTo 0
    Set chartObject = sheet.ChartObjects.Add( _
        sheet.Range("F4").Left, sheet.Range("F4").Top, 820, 480)
    chartObject.Name = "chtManMachine"
    Set chart = chartObject.Chart
    chart.ChartType = xlBarStacked
    chart.SetSourceData sheet.Range("A4:C" & rowIndex - 1)
    chart.HasTitle = True
    chart.ChartTitle.Text = "人机作业时间轴（" & viewMode & "）"
    chart.HasLegend = False
    chart.SeriesCollection(1).Format.Fill.Visible = msoFalse
    chart.SeriesCollection(1).Format.Line.Visible = msoFalse
    For pointIndex = 1 To chart.SeriesCollection(2).Points.Count
        chart.SeriesCollection(2).Points(pointIndex).Format.Fill.ForeColor.RGB = _
            StepTypeColor(CStr(sheet.Cells(pointIndex + 4, 4).Value2))
    Next pointIndex
    chart.Axes(xlCategory).ReversePlotOrder = True
    sheet.Columns("A").ColumnWidth = 36
    sheet.Columns("B:D").ColumnWidth = 12
End Sub

Private Sub WriteGanttRow(ByVal sheet As Worksheet, ByRef rowIndex As Long, _
        ByVal labelText As String, ByVal task As Object)
    sheet.Cells(rowIndex, 1).Value2 = labelText
    sheet.Cells(rowIndex, 2).Value2 = task("StartSec")
    sheet.Cells(rowIndex, 3).Value2 = task("DurationSec")
    sheet.Cells(rowIndex, 4).Value2 = task("StepType")
    rowIndex = rowIndex + 1
End Sub

Public Sub RefreshComparisonReport()
    Dim sheet As Worksheet, beforeSchedule As Collection, afterSchedule As Collection
    Dim beforeMetrics As Object, afterMetrics As Object
    Dim labels As Variant, keys As Variant, rowIndex As Long
    Dim beforeValue As Double, afterValue As Double
    Set sheet = ThisWorkbook.Worksheets("改善对比报告")
    sheet.Range("A3:F30").ClearContents
    sheet.Range("A3:E3").Value = Array("指标", "改善前", "改善后", "差值", "改善率")
    labels = Array("循环周期(s)", "平均节拍(s/件)", "小时产能(pcs/h)", _
        "平均人员负荷率", "最大人员负荷率", "设备等待人员(s)", "总移动时间(s)")
    keys = Array("CycleTimeSec", "TaktSec", "HourlyCapacity", "AveragePersonLoad", _
        "MaxPersonLoad", "TotalDeviceWaitSec", "TotalMoveSec")
    If SnapshotExists("BEFORE") Then Set beforeSchedule = RestoreScheduleSnapshot("BEFORE")
    If SnapshotExists("AFTER") Then Set afterSchedule = RestoreScheduleSnapshot("AFTER")
    If beforeSchedule Is Nothing Then
        If Not gCurrentSchedule Is Nothing Then Set beforeSchedule = gCurrentSchedule
    End If
    If afterSchedule Is Nothing Then
        If Not gCurrentSchedule Is Nothing Then Set afterSchedule = gCurrentSchedule
    End If
    If beforeSchedule Is Nothing Or afterSchedule Is Nothing Then Exit Sub
    Set beforeMetrics = CalculateMetrics(beforeSchedule)
    Set afterMetrics = CalculateMetrics(afterSchedule)
    For rowIndex = 0 To UBound(labels)
        beforeValue = CDbl(beforeMetrics(keys(rowIndex)))
        afterValue = CDbl(afterMetrics(keys(rowIndex)))
        sheet.Cells(rowIndex + 4, 1).Value2 = labels(rowIndex)
        sheet.Cells(rowIndex + 4, 2).Value2 = beforeValue
        sheet.Cells(rowIndex + 4, 3).Value2 = afterValue
        sheet.Cells(rowIndex + 4, 4).Value2 = afterValue - beforeValue
        If Abs(beforeValue) > 0.000001 Then
            sheet.Cells(rowIndex + 4, 5).Value2 = (afterValue - beforeValue) / beforeValue
        End If
    Next rowIndex
    sheet.Range("E4:E10").NumberFormat = "0.0%"
    sheet.Range("B4:D10").NumberFormat = "0.00"
    sheet.Cells(12, 1).Value2 = "改善建议"
    sheet.Cells(12, 1).Font.Bold = True
    sheet.Cells(13, 1).Value2 = ImprovementAdvice(afterMetrics)
    sheet.Range("A13:E16").Merge
    sheet.Range("A13:E16").WrapText = True
    sheet.Columns("A").ColumnWidth = 24
    sheet.Columns("B:E").ColumnWidth = 16
End Sub

Public Sub WriteMetricCards(ByVal metrics As Object)
    Dim sheet As Worksheet
    Set sheet = ThisWorkbook.Worksheets("自动排程")
    sheet.Range("A3:H5").ClearContents
    sheet.Range("A3:H3").Value = Array("循环周期(s)", "平均节拍(s/件)", _
        "小时产能", "平均人员负荷", "最大人员负荷", "设备等待(s)", _
        "瓶颈人员", "瓶颈设备")
    sheet.Range("A4:H4").Value = Array(metrics("CycleTimeSec"), metrics("TaktSec"), _
        metrics("HourlyCapacity"), metrics("AveragePersonLoad"), _
        metrics("MaxPersonLoad"), metrics("TotalDeviceWaitSec"), _
        metrics("BottleneckPerson"), metrics("BottleneckDevice"))
    sheet.Range("A3:H3").Font.Bold = True
    sheet.Range("D4:E4").NumberFormat = "0.0%"
    sheet.Range("A4:C4").NumberFormat = "0.00"
End Sub

Private Function ImprovementAdvice(ByVal metrics As Object) As String
    Dim advice As String
    If CDbl(metrics("MaxPersonLoad")) > 0.9 Then
        advice = advice & "人员负荷超过90%，建议分担高负荷人工步骤或增加技能覆盖。" & vbLf
    End If
    If CDbl(metrics("AnalysisSec")) > 0 And _
            CDbl(metrics("TotalDeviceWaitSec")) / CDbl(metrics("AnalysisSec")) > 0.1 Then
        advice = advice & "设备等待人员占比较高，建议调整服务顺序、缩短移动距离。" & vbLf
    End If
    If CDbl(metrics("TotalMoveSec")) > 0 Then
        advice = advice & "存在人员移动时间，建议优化设备布局或固定服务区域。" & vbLf
    End If
    If Len(advice) = 0 Then advice = "当前方案未发现明显过载，可继续观察现场波动并验证稳定循环。"
    ImprovementAdvice = advice
End Function

Private Function DisplayStepType(ByVal stepType As String) As String
    Select Case stepType
        Case STEP_MANUAL: DisplayStepType = "人工"
        Case STEP_AUTO: DisplayStepType = "自动运行"
        Case STEP_JOINT: DisplayStepType = "人机协同"
        Case STEP_WAIT: DisplayStepType = "等待"
    End Select
End Function

Private Function StepTypeColor(ByVal stepType As String) As Long
    Select Case stepType
        Case STEP_MANUAL: StepTypeColor = RGB(68, 114, 196)
        Case STEP_AUTO: StepTypeColor = RGB(112, 173, 71)
        Case STEP_JOINT: StepTypeColor = RGB(237, 125, 49)
        Case STEP_WAIT: StepTypeColor = RGB(192, 0, 0)
        Case Else: StepTypeColor = RGB(191, 191, 191)
    End Select
End Function
