Attribute VB_Name = "modCommands"
Option Explicit

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
