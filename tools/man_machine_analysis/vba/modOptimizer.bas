Attribute VB_Name = "modOptimizer"
Option Explicit

Public Function OptimizeSchedule(ByVal target As String, Optional ByVal iterations As Long = 0, _
        Optional ByVal randomSeed As Long = 12345) As Collection
    Dim best As Collection, candidate As Collection
    Dim bestScore As Double, candidateScore As Double
    Dim iteration As Long, startedAt As Single, reason As String

    If iterations <= 0 Then iterations = ReadIterationCount()
    If Not gCurrentSchedule Is Nothing Then
        If HasLockedTasks(gCurrentSchedule) Then Set best = gCurrentSchedule
    End If
    If best Is Nothing Then Set best = BuildInitialSchedule(False, randomSeed)
    bestScore = ScoreSchedule(best, target)
    startedAt = Timer

    For iteration = 1 To iterations
        Set candidate = BuildInitialSchedule(True, iteration)
        reason = ""
        If IsScheduleFeasible(candidate, reason) And LocksPreserved(candidate, best) Then
            candidateScore = ScoreSchedule(candidate, target)
            If candidateScore + 0.000001 < bestScore Then
                Set best = candidate
                bestScore = candidateScore
            End If
        End If
        If ElapsedSeconds(startedAt, Timer) > 8# Then Exit For
        If target = TARGET_THROUGHPUT Then
            If CDbl(CalculateMetrics(best)("CycleTimeSec")) <= 56.000001 Then Exit For
        End If
    Next iteration
    Set OptimizeSchedule = best
End Function

Private Function HasLockedTasks(ByVal schedule As Collection) As Boolean
    Dim task As Object
    For Each task In schedule
        If CBool(task("IsLocked")) Then
            HasLockedTasks = True
            Exit Function
        End If
    Next task
End Function

Private Function LocksPreserved(ByVal candidate As Collection, _
        ByVal referenceSchedule As Collection) As Boolean
    Dim referenceTask As Object, candidateTask As Object, found As Boolean
    For Each referenceTask In referenceSchedule
        If CBool(referenceTask("IsLocked")) Then
            found = False
            For Each candidateTask In candidate
                If StrComp(CStr(candidateTask("TaskId")), CStr(referenceTask("TaskId")), _
                        vbTextCompare) = 0 Then
                    found = True
                    If StrComp(CStr(candidateTask("PersonId")), CStr(referenceTask("PersonId")), _
                            vbTextCompare) <> 0 Then Exit Function
                    If Abs(CDbl(candidateTask("StartSec")) - _
                            CDbl(referenceTask("StartSec"))) > 0.000001 Then Exit Function
                    Exit For
                End If
            Next candidateTask
            If Not found Then Exit Function
        End If
    Next referenceTask
    LocksPreserved = True
End Function

Private Function ReadIterationCount() As Long
    On Error GoTo Defaults
    ReadIterationCount = CLng(ThisWorkbook.Worksheets("基础设置") _
        .ListObjects("tblParameters").ListColumns("搜索迭代次数") _
        .DataBodyRange.Cells(1, 1).Value2)
    If ReadIterationCount < 1 Then GoTo Defaults
    Exit Function
Defaults:
    ReadIterationCount = 200
End Function

Private Function ElapsedSeconds(ByVal startedAt As Single, ByVal endedAt As Single) As Double
    If endedAt >= startedAt Then
        ElapsedSeconds = endedAt - startedAt
    Else
        ElapsedSeconds = (86400# - startedAt) + endedAt
    End If
End Function
