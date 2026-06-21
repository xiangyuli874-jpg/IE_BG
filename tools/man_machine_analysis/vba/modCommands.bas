Attribute VB_Name = "modCommands"
Option Explicit

Public Sub CmdNewBlankAnalysis()
    If MsgBox("将清空当前输入、排程和改善方案。是否继续？", _
            vbYesNo + vbQuestion, "新建空白分析") <> vbYes Then Exit Sub
    ResetToBlankState
    ThisWorkbook.Worksheets("基础设置").Activate
End Sub

Public Sub CmdLoadExampleAndRun()
    If MsgBox("将覆盖当前输入并载入1人3机示例。是否继续？", _
            vbYesNo + vbQuestion, "载入示例") <> vbYes Then Exit Sub
    LoadExampleDataAndRun
    ThisWorkbook.Worksheets("自动排程").Activate
End Sub

Public Sub ResetToBlankState()
    Dim settings As Worksheet, stepsSheet As Worksheet
    Dim peopleTable As ListObject, devicesTable As ListObject
    Dim moveTable As ListObject, stepsTable As ListObject
    Dim fromIndex As Long, toIndex As Long

    Set settings = ThisWorkbook.Worksheets("基础设置")
    Set stepsSheet = ThisWorkbook.Worksheets("作业步骤")
    Set peopleTable = settings.ListObjects("tblPeople")
    Set devicesTable = settings.ListObjects("tblDevices")
    Set moveTable = settings.ListObjects("tblMoveTime")
    Set stepsTable = stepsSheet.ListObjects("tblSteps")

    peopleTable.DataBodyRange.ClearContents
    devicesTable.DataBodyRange.ClearContents
    stepsTable.DataBodyRange.ClearContents

    With settings.ListObjects("tblParameters")
        .ListColumns("方案名称").DataBodyRange.Cells(1, 1).Value2 = "新分析"
        .ListColumns("人员数量").DataBodyRange.Cells(1, 1).ClearContents
        .ListColumns("设备数量").DataBodyRange.Cells(1, 1).ClearContents
        .ListColumns("计划分析循环数").DataBodyRange.Cells(1, 1).Value2 = 3
        .ListColumns("优化目标").DataBodyRange.Cells(1, 1).Value2 = TARGET_COMPOSITE
        .ListColumns("搜索迭代次数").DataBodyRange.Cells(1, 1).Value2 = 500
        .ListColumns("周期权重").DataBodyRange.Cells(1, 1).Value2 = 0.4
        .ListColumns("均衡权重").DataBodyRange.Cells(1, 1).Value2 = 0.2
        .ListColumns("等待权重").DataBodyRange.Cells(1, 1).Value2 = 0.3
        .ListColumns("移动权重").DataBodyRange.Cells(1, 1).Value2 = 0.1
    End With

    For fromIndex = 1 To MAX_DEVICES
        moveTable.DataBodyRange.Cells(fromIndex, 1).Value2 = "M" & fromIndex
        For toIndex = 1 To MAX_DEVICES
            moveTable.DataBodyRange.Cells(fromIndex, toIndex + 1).Value2 = 0#
        Next toIndex
    Next fromIndex

    Set gCurrentSchedule = Nothing
    Set gCurrentMetrics = Nothing
    ClearVisualOutputs
    ThisWorkbook.Worksheets("自动排程").Range("A6:M6").Merge
    SetStatus "空白模板已就绪：请先填写“基础设置”和“作业步骤”。", False
End Sub

Public Sub LoadExampleDataAndRun()
    Dim settings As Worksheet, peopleTable As ListObject
    Dim devicesTable As ListObject, stepsTable As ListObject
    Dim deviceIndex As Long, stepIndex As Long, rowIndex As Long
    Dim stepNames As Variant, stepTypes As Variant
    Dim durations As Variant, skills As Variant

    ResetToBlankState
    Set settings = ThisWorkbook.Worksheets("基础设置")
    Set peopleTable = settings.ListObjects("tblPeople")
    Set devicesTable = settings.ListObjects("tblDevices")
    Set stepsTable = ThisWorkbook.Worksheets("作业步骤").ListObjects("tblSteps")

    With settings.ListObjects("tblParameters")
        .ListColumns("方案名称").DataBodyRange.Cells(1, 1).Value2 = "1人3机示例"
        .ListColumns("人员数量").DataBodyRange.Cells(1, 1).Value2 = 1
        .ListColumns("设备数量").DataBodyRange.Cells(1, 1).Value2 = 3
        .ListColumns("优化目标").DataBodyRange.Cells(1, 1).Value2 = TARGET_THROUGHPUT
    End With
    SetTableRowValues peopleTable, 1, Array("P1", "操作员1", "通用", "M1,M2,M3", "是")
    For deviceIndex = 1 To 3
        SetTableRowValues devicesTable, deviceIndex, Array("M" & deviceIndex, _
            "设备" & deviceIndex, "案例产品", "独立循环", "", "是")
    Next deviceIndex

    stepNames = Array("上料", "自动运行1", "翻面", "自动运行2", "下料")
    stepTypes = Array("人工", "自动运行", "人工", "自动运行", "人工")
    durations = Array(5#, 20#, 6#, 15#, 5#)
    skills = Array("通用", "", "通用", "", "通用")
    rowIndex = 0
    For deviceIndex = 1 To 3
        For stepIndex = 1 To 5
            rowIndex = rowIndex + 1
            SetTableRowValues stepsTable, rowIndex, Array("M" & deviceIndex, _
                stepIndex, stepNames(stepIndex - 1), stepTypes(stepIndex - 1), _
                durations(stepIndex - 1), IIf(stepIndex = 1, "", stepIndex - 1), _
                skills(stepIndex - 1), "", "", "是", "")
        Next stepIndex
    Next deviceIndex

    CmdBuildInitial
    CmdSaveBefore
    CmdOptimize
    CmdSaveAfter
    RefreshComparisonReport
    SetStatus "示例已载入并完成试算：56秒完成3件，平均18.67秒/件。", False
End Sub

Public Sub CmdCheckData()
    Dim errors As String
    On Error GoTo Failed
    errors = ValidateCurrentModel()
    If Len(errors) = 0 Then
        SetStatus "数据检查通过。", False
    Else
        SetStatus "数据检查失败：" & vbLf & errors, True
    End If
    Exit Sub
Failed:
    SetStatus "数据检查异常：" & Err.Description, True
End Sub

Public Sub CmdBuildInitial()
    On Error GoTo Failed
    If Not gCurrentSchedule Is Nothing Then SaveScheduleSnapshot "UNDO", gCurrentSchedule
    Set gCurrentSchedule = BuildInitialSchedule(False, 1)
    Set gCurrentMetrics = CalculateMetrics(gCurrentSchedule)
    SaveScheduleSnapshot "UNDO", gCurrentSchedule
    WriteScheduleOutput gCurrentSchedule
    WriteMetricCards gCurrentMetrics
    RefreshGanttView "全部"
    SetStatus "初始方案已生成。", False
    Exit Sub
Failed:
    SetStatus "生成初始方案失败：" & Err.Description, True
End Sub

Public Sub CmdOptimize()
    Dim target As String, iterations As Long
    On Error GoTo Failed
    If Not gCurrentSchedule Is Nothing Then SaveScheduleSnapshot "UNDO", gCurrentSchedule
    target = CStr(ThisWorkbook.Names("nmOptimizationTarget").RefersToRange.Value2)
    iterations = CLng(ThisWorkbook.Worksheets("基础设置") _
        .ListObjects("tblParameters").ListColumns("搜索迭代次数") _
        .DataBodyRange.Cells(1, 1).Value2)
    Set gCurrentSchedule = OptimizeSchedule(target, iterations, 12345)
    Set gCurrentMetrics = CalculateMetrics(gCurrentSchedule)
    WriteScheduleOutput gCurrentSchedule
    WriteMetricCards gCurrentMetrics
    RefreshGanttView "全部"
    SetStatus "优化完成，目标：" & target, False
    Exit Sub
Failed:
    SetStatus "优化排程失败：" & Err.Description, True
End Sub

Public Sub CmdToggleLock()
    Dim table As ListObject, rowIndex As Long, taskId As String
    Dim task As Object
    On Error GoTo Failed
    Set table = ThisWorkbook.Worksheets("自动排程").ListObjects("tblSchedule")
    If Intersect(Selection, table.DataBodyRange) Is Nothing Then
        Err.Raise vbObjectError + 2500, "CmdToggleLock", "请先选择排程明细中的任务行。"
    End If
    rowIndex = Selection.Row - table.DataBodyRange.Row + 1
    taskId = CStr(table.ListColumns("任务编号").DataBodyRange.Cells(rowIndex, 1).Value2)
    For Each task In gCurrentSchedule
        If StrComp(CStr(task("TaskId")), taskId, vbTextCompare) = 0 Then
            task("IsLocked") = Not CBool(task("IsLocked"))
            Exit For
        End If
    Next task
    WriteScheduleOutput gCurrentSchedule
    SetStatus "任务锁定状态已切换：" & taskId, False
    Exit Sub
Failed:
    SetStatus "锁定操作失败：" & Err.Description, True
End Sub

Public Sub CmdUndo()
    On Error GoTo Failed
    If Not SnapshotExists("UNDO") Then Err.Raise vbObjectError + 2501, , "没有可撤回方案。"
    Set gCurrentSchedule = RestoreScheduleSnapshot("UNDO")
    Set gCurrentMetrics = CalculateMetrics(gCurrentSchedule)
    WriteScheduleOutput gCurrentSchedule
    WriteMetricCards gCurrentMetrics
    RefreshGanttView "全部"
    SetStatus "已恢复上一方案。", False
    Exit Sub
Failed:
    SetStatus "撤回失败：" & Err.Description, True
End Sub

Public Sub CmdSaveBefore()
    If gCurrentSchedule Is Nothing Then CmdBuildInitial
    If Not gCurrentSchedule Is Nothing Then
        SaveScheduleSnapshot "BEFORE", gCurrentSchedule
        SetStatus "已保存为改善前方案。", False
    End If
End Sub

Public Sub CmdSaveAfter()
    If gCurrentSchedule Is Nothing Then CmdBuildInitial
    If Not gCurrentSchedule Is Nothing Then
        SaveScheduleSnapshot "AFTER", gCurrentSchedule
        SetStatus "已保存为改善后方案。", False
    End If
End Sub

Public Sub CmdRefreshReport()
    On Error GoTo Failed
    RefreshComparisonReport
    SetStatus "改善对比报告已刷新。", False
    Exit Sub
Failed:
    SetStatus "刷新报告失败：" & Err.Description, True
End Sub

Public Sub CmdViewAll()
    RefreshGanttView "全部"
End Sub

Public Sub CmdViewPeople()
    RefreshGanttView "仅人员"
End Sub

Public Sub CmdViewDevices()
    RefreshGanttView "仅设备"
End Sub

Public Sub CmdGoSettings()
    ThisWorkbook.Worksheets("基础设置").Activate
End Sub

Public Sub CmdGoSteps()
    ThisWorkbook.Worksheets("作业步骤").Activate
End Sub

Public Sub CmdGoSchedule()
    ThisWorkbook.Worksheets("自动排程").Activate
End Sub

Private Sub SetTableRowValues(ByVal table As ListObject, ByVal rowIndex As Long, _
        ByVal values As Variant)
    Dim columnIndex As Long
    For columnIndex = LBound(values) To UBound(values)
        table.DataBodyRange.Cells(rowIndex, columnIndex + 1).Value2 = values(columnIndex)
    Next columnIndex
End Sub

Private Sub SetStatus(ByVal message As String, ByVal isError As Boolean)
    With ThisWorkbook.Worksheets("自动排程").Range("A6:M6")
        .Merge
        .Value2 = message
        .WrapText = True
        .Font.Bold = True
        If isError Then
            .Interior.Color = RGB(255, 199, 206)
            .Font.Color = RGB(156, 0, 6)
        Else
            .Interior.Color = RGB(226, 239, 218)
            .Font.Color = RGB(0, 97, 0)
        End If
    End With
End Sub
