Attribute VB_Name = "modScenario"
Option Explicit

Public Sub SaveScheduleSnapshot(ByVal snapshotName As String, ByVal schedule As Collection)
    Dim sheet As Worksheet, startColumn As Long, task As Object
    Dim rowIndex As Long
    Set sheet = ThisWorkbook.Worksheets("自动排程")
    startColumn = SnapshotColumn(snapshotName)
    sheet.Range(sheet.Cells(1, startColumn), sheet.Cells(1000, startColumn + 11)).ClearContents
    sheet.Cells(1, startColumn).Value2 = snapshotName
    rowIndex = 2
    For Each task In schedule
        sheet.Cells(rowIndex, startColumn).Value2 = task("TaskId")
        sheet.Cells(rowIndex, startColumn + 1).Value2 = task("CycleNo")
        sheet.Cells(rowIndex, startColumn + 2).Value2 = task("DeviceId")
        sheet.Cells(rowIndex, startColumn + 3).Value2 = task("StepNo")
        sheet.Cells(rowIndex, startColumn + 4).Value2 = task("StepName")
        sheet.Cells(rowIndex, startColumn + 5).Value2 = task("StepType")
        sheet.Cells(rowIndex, startColumn + 6).Value2 = task("PersonId")
        sheet.Cells(rowIndex, startColumn + 7).Value2 = task("StartSec")
        sheet.Cells(rowIndex, startColumn + 8).Value2 = task("EndSec")
        sheet.Cells(rowIndex, startColumn + 9).Value2 = task("MoveSec")
        sheet.Cells(rowIndex, startColumn + 10).Value2 = task("WaitSec")
        sheet.Cells(rowIndex, startColumn + 11).Value2 = task("IsLocked")
        rowIndex = rowIndex + 1
    Next task
    sheet.Range(sheet.Columns(startColumn), sheet.Columns(startColumn + 11)).EntireColumn.Hidden = True
End Sub

Public Function RestoreScheduleSnapshot(ByVal snapshotName As String) As Collection
    Dim result As New Collection, sheet As Worksheet, startColumn As Long
    Dim rowIndex As Long, task As Object
    Set sheet = ThisWorkbook.Worksheets("自动排程")
    startColumn = SnapshotColumn(snapshotName)
    rowIndex = 2
    Do While Len(CStr(sheet.Cells(rowIndex, startColumn).Value2)) > 0
        Set task = CreateObject("Scripting.Dictionary")
        task.CompareMode = vbTextCompare
        task("TaskId") = CStr(sheet.Cells(rowIndex, startColumn).Value2)
        task("CycleNo") = CLng(sheet.Cells(rowIndex, startColumn + 1).Value2)
        task("DeviceId") = CStr(sheet.Cells(rowIndex, startColumn + 2).Value2)
        task("StepNo") = CLng(sheet.Cells(rowIndex, startColumn + 3).Value2)
        task("StepName") = CStr(sheet.Cells(rowIndex, startColumn + 4).Value2)
        task("StepType") = CStr(sheet.Cells(rowIndex, startColumn + 5).Value2)
        task("PersonId") = CStr(sheet.Cells(rowIndex, startColumn + 6).Value2)
        task("StartSec") = CDbl(sheet.Cells(rowIndex, startColumn + 7).Value2)
        task("EndSec") = CDbl(sheet.Cells(rowIndex, startColumn + 8).Value2)
        task("DurationSec") = CDbl(task("EndSec")) - CDbl(task("StartSec"))
        task("MoveSec") = CDbl(sheet.Cells(rowIndex, startColumn + 9).Value2)
        task("WaitSec") = CDbl(sheet.Cells(rowIndex, startColumn + 10).Value2)
        task("IsLocked") = CBool(sheet.Cells(rowIndex, startColumn + 11).Value2)
        task("PredecessorIds") = ""
        result.Add task
        rowIndex = rowIndex + 1
    Loop
    Set RestoreScheduleSnapshot = result
End Function

Public Function SnapshotExists(ByVal snapshotName As String) As Boolean
    SnapshotExists = Len(CStr(ThisWorkbook.Worksheets("自动排程") _
        .Cells(2, SnapshotColumn(snapshotName)).Value2)) > 0
End Function

Private Function SnapshotColumn(ByVal snapshotName As String) As Long
    Select Case UCase$(snapshotName)
        Case "UNDO": SnapshotColumn = 27
        Case "BEFORE": SnapshotColumn = 40
        Case "AFTER": SnapshotColumn = 53
        Case Else
            Err.Raise vbObjectError + 2300, "SnapshotColumn", "未知方案快照：" & snapshotName
    End Select
End Function
