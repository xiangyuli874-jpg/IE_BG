Attribute VB_Name = "modScheduler"
Option Explicit

Public gCurrentSchedule As Collection
Public gCurrentMetrics As Object

Public Function BuildInitialSchedule(Optional ByVal randomized As Boolean = False, _
        Optional ByVal randomSeed As Long = 1) As Collection
    Dim validationErrors As String
    validationErrors = ValidateCurrentModel()
    If Len(validationErrors) > 0 Then
        Err.Raise vbObjectError + 2100, "BuildInitialSchedule", validationErrors
    End If

    Dim people As Collection, devices As Collection, steps As Collection
    Dim moveMatrix As Object, tasks As Collection
    Set people = ReadPeople()
    Set devices = ReadDevices()
    Set steps = ReadSteps()
    Set moveMatrix = ReadMoveMatrix()
    Set tasks = ExpandScheduleTasks(devices, steps, ReadCycleCount())
    Set BuildInitialSchedule = ScheduleTasks(tasks, people, moveMatrix, randomized, randomSeed)
End Function

Public Function ExpandScheduleTasks(ByVal devices As Collection, ByVal steps As Collection, _
        ByVal cycleCount As Long) As Collection
    Dim result As New Collection
    Dim maxStepByDevice As Object
    Dim stepItem As Object, device As Object, task As Object
    Dim cycleNo As Long, stepNo As Long, predecessorIds As String

    Set maxStepByDevice = CreateObject("Scripting.Dictionary")
    maxStepByDevice.CompareMode = vbTextCompare
    For Each device In devices
        maxStepByDevice(device("DeviceId")) = 0
    Next device
    For Each stepItem In steps
        If stepItem("StepNo") > maxStepByDevice(stepItem("DeviceId")) Then
            maxStepByDevice(stepItem("DeviceId")) = stepItem("StepNo")
        End If
    Next stepItem

    For cycleNo = 1 To cycleCount
        For Each device In devices
            For stepNo = 1 To CLng(maxStepByDevice(device("DeviceId")))
                Set stepItem = FindStep(steps, CStr(device("DeviceId")), stepNo)
                If Not stepItem Is Nothing Then
                    Set task = CreateObject("Scripting.Dictionary")
                    task.CompareMode = vbTextCompare
                    task("TaskId") = TaskKey(device("DeviceId"), cycleNo, stepNo)
                    task("DeviceId") = CStr(device("DeviceId"))
                    task("CycleNo") = cycleNo
                    task("StepNo") = stepNo
                    task("StepName") = CStr(stepItem("StepName"))
                    task("StepType") = CStr(stepItem("StepType"))
                    task("DurationSec") = CDbl(stepItem("DurationSec"))
                    task("RequiredSkill") = CStr(stepItem("RequiredSkill"))
                    task("LockedPersonId") = CStr(stepItem("LockedPersonId"))
                    task("HasLockedStart") = (CBool(stepItem("HasLockedStart")) And cycleNo = 1)
                    task("LockedStartSec") = CDbl(stepItem("LockedStartSec"))
                    task("SourceRow") = CLng(stepItem("SourceRow"))

                    predecessorIds = ""
                    If Len(CStr(stepItem("PredecessorText"))) > 0 Then
                        predecessorIds = TaskKey(device("DeviceId"), cycleNo, _
                            CLng(CDbl(stepItem("PredecessorText"))))
                    ElseIf stepNo > 1 Then
                        predecessorIds = TaskKey(device("DeviceId"), cycleNo, stepNo - 1)
                    ElseIf cycleNo > 1 Then
                        predecessorIds = TaskKey(device("DeviceId"), cycleNo - 1, _
                            CLng(maxStepByDevice(device("DeviceId"))))
                    End If

                    If stepNo = 1 And StrComp(CStr(device("RelationType")), _
                            "有先后顺序", vbTextCompare) = 0 Then
                        If Len(CStr(device("PredecessorDeviceId"))) > 0 Then
                            AppendPredecessor predecessorIds, TaskKey( _
                                device("PredecessorDeviceId"), cycleNo, _
                                CLng(maxStepByDevice(device("PredecessorDeviceId"))))
                        End If
                    End If
                    task("PredecessorIds") = predecessorIds
                    result.Add task
                End If
            Next stepNo
        Next device
    Next cycleNo
    Set ExpandScheduleTasks = result
End Function

Private Function ScheduleTasks(ByVal tasks As Collection, ByVal people As Collection, _
        ByVal moveMatrix As Object, ByVal randomized As Boolean, _
        ByVal randomSeed As Long) As Collection
    Dim result As New Collection
    Dim scheduledEnd As Object, deviceAvailable As Object, personAvailable As Object
    Dim personLocation As Object, done As Object
    Dim candidateTask() As Object, candidatePerson() As String
    Dim candidateStart() As Double, candidateMove() As Double
    Dim candidatePredecessorEnd() As Double
    Dim task As Object, scheduled As Object, person As Object
    Dim taskIndex As Long, candidateCount As Long, selectedIndex As Long
    Dim earliest As Double, bestEarliest As Double, moveSec As Double
    Dim personId As String, predecessorEnd As Double, finishTime As Double
    Dim key As Variant, nearIndexes() As Long, nearCount As Long
    Dim ignoredRandom As Single

    Set scheduledEnd = CreateObject("Scripting.Dictionary")
    scheduledEnd.CompareMode = vbTextCompare
    Set deviceAvailable = CreateObject("Scripting.Dictionary")
    deviceAvailable.CompareMode = vbTextCompare
    Set personAvailable = CreateObject("Scripting.Dictionary")
    personAvailable.CompareMode = vbTextCompare
    Set personLocation = CreateObject("Scripting.Dictionary")
    personLocation.CompareMode = vbTextCompare
    Set done = CreateObject("Scripting.Dictionary")
    done.CompareMode = vbTextCompare

    For Each person In people
        personAvailable(person("PersonId")) = 0#
        personLocation(person("PersonId")) = ""
    Next person
    If randomized Then
        ignoredRandom = Rnd(-1)
        Randomize randomSeed
    End If

    Do While done.Count < tasks.Count
        candidateCount = 0
        bestEarliest = 1E+99
        ReDim candidateTask(1 To tasks.Count)
        ReDim candidatePerson(1 To tasks.Count)
        ReDim candidateStart(1 To tasks.Count)
        ReDim candidateMove(1 To tasks.Count)
        ReDim candidatePredecessorEnd(1 To tasks.Count)

        For taskIndex = 1 To tasks.Count
            Set task = tasks(taskIndex)
            If Not done.Exists(task("TaskId")) Then
                If PredecessorsComplete(CStr(task("PredecessorIds")), scheduledEnd, predecessorEnd) Then
                    personId = ""
                    moveSec = 0#
                    earliest = MaxValue(predecessorEnd, DictionaryNumber(deviceAvailable, task("DeviceId")))
                    If UsesPerson(CStr(task("StepType"))) Then
                        personId = SelectPerson(people, task, personAvailable, _
                            personLocation, moveMatrix, earliest, moveSec)
                        If Len(personId) = 0 Then
                            Err.Raise vbObjectError + 2101, "ScheduleTasks", _
                                "任务无可用人员：" & task("TaskId")
                        End If
                        earliest = MaxValue(earliest, _
                            DictionaryNumber(personAvailable, personId) + moveSec)
                    End If
                    If CBool(task("HasLockedStart")) Then
                        If CDbl(task("LockedStartSec")) + 0.000001 < earliest Then
                            Err.Raise vbObjectError + 2102, "ScheduleTasks", _
                                "锁定开始时间与资源约束冲突：" & task("TaskId")
                        End If
                        earliest = CDbl(task("LockedStartSec"))
                    End If
                    candidateCount = candidateCount + 1
                    Set candidateTask(candidateCount) = task
                    candidatePerson(candidateCount) = personId
                    candidateStart(candidateCount) = earliest
                    candidateMove(candidateCount) = moveSec
                    candidatePredecessorEnd(candidateCount) = predecessorEnd
                    If earliest < bestEarliest Then bestEarliest = earliest
                End If
            End If
        Next taskIndex

        If candidateCount = 0 Then
            Err.Raise vbObjectError + 2103, "ScheduleTasks", "无法找到可执行任务，前置关系可能无解。"
        End If

        selectedIndex = 1
        If randomized Then
            ReDim nearIndexes(1 To candidateCount)
            nearCount = 0
            For taskIndex = 1 To candidateCount
                If candidateStart(taskIndex) <= bestEarliest + 10# Then
                    nearCount = nearCount + 1
                    nearIndexes(nearCount) = taskIndex
                End If
            Next taskIndex
            selectedIndex = nearIndexes(1 + Int(Rnd() * nearCount))
        Else
            For taskIndex = 1 To candidateCount
                If candidateStart(taskIndex) < candidateStart(selectedIndex) Then
                    selectedIndex = taskIndex
                ElseIf candidateStart(taskIndex) = candidateStart(selectedIndex) Then
                    If candidateTask(taskIndex)("CycleNo") < candidateTask(selectedIndex)("CycleNo") Then
                        selectedIndex = taskIndex
                    ElseIf candidateTask(taskIndex)("CycleNo") = candidateTask(selectedIndex)("CycleNo") _
                            And candidateTask(taskIndex)("StepNo") < candidateTask(selectedIndex)("StepNo") Then
                        selectedIndex = taskIndex
                    End If
                End If
            Next taskIndex
        End If

        Set task = candidateTask(selectedIndex)
        finishTime = candidateStart(selectedIndex) + CDbl(task("DurationSec"))
        Set scheduled = CreateObject("Scripting.Dictionary")
        scheduled.CompareMode = vbTextCompare
        For Each key In task.Keys
            scheduled(key) = task(key)
        Next key
        scheduled("PersonId") = candidatePerson(selectedIndex)
        scheduled("StartSec") = candidateStart(selectedIndex)
        scheduled("EndSec") = finishTime
        scheduled("MoveSec") = candidateMove(selectedIndex)
        scheduled("WaitSec") = MaxValue(0#, candidateStart(selectedIndex) - _
            candidatePredecessorEnd(selectedIndex) - candidateMove(selectedIndex))
        scheduled("IsLocked") = (Len(CStr(task("LockedPersonId"))) > 0 Or _
            CBool(task("HasLockedStart")))
        result.Add scheduled

        done(task("TaskId")) = True
        scheduledEnd(task("TaskId")) = finishTime
        deviceAvailable(task("DeviceId")) = finishTime
        If UsesPerson(CStr(task("StepType"))) Then
            personAvailable(candidatePerson(selectedIndex)) = finishTime
            personLocation(candidatePerson(selectedIndex)) = task("DeviceId")
        End If
    Loop
    Set ScheduleTasks = result
End Function

Public Function IsScheduleFeasible(ByVal schedule As Collection, _
        Optional ByRef reason As String) As Boolean
    Dim firstTask As Object, secondTask As Object
    Dim firstIndex As Long, secondIndex As Long
    Dim endById As Object, predecessor As Variant
    Set endById = CreateObject("Scripting.Dictionary")
    endById.CompareMode = vbTextCompare

    For Each firstTask In schedule
        If CDbl(firstTask("EndSec")) < CDbl(firstTask("StartSec")) Then
            reason = "任务结束早于开始：" & firstTask("TaskId")
            Exit Function
        End If
        endById(firstTask("TaskId")) = CDbl(firstTask("EndSec"))
    Next firstTask

    For firstIndex = 1 To schedule.Count
        Set firstTask = schedule(firstIndex)
        For secondIndex = firstIndex + 1 To schedule.Count
            Set secondTask = schedule(secondIndex)
            If IntervalsOverlap(firstTask, secondTask) Then
                If StrComp(CStr(firstTask("DeviceId")), CStr(secondTask("DeviceId")), _
                        vbTextCompare) = 0 Then
                    reason = "设备任务重叠：" & firstTask("TaskId") & " / " & secondTask("TaskId")
                    Exit Function
                End If
                If Len(CStr(firstTask("PersonId"))) > 0 And _
                        StrComp(CStr(firstTask("PersonId")), CStr(secondTask("PersonId")), _
                        vbTextCompare) = 0 Then
                    reason = "人员任务重叠：" & firstTask("TaskId") & " / " & secondTask("TaskId")
                    Exit Function
                End If
            End If
        Next secondIndex
        For Each predecessor In Split(CStr(firstTask("PredecessorIds")), ",")
            If Len(Trim$(CStr(predecessor))) > 0 Then
                If Not endById.Exists(Trim$(CStr(predecessor))) Then
                    reason = "前置任务不存在：" & predecessor
                    Exit Function
                End If
                If CDbl(endById(Trim$(CStr(predecessor)))) > CDbl(firstTask("StartSec")) + 0.000001 Then
                    reason = "前置任务未完成：" & firstTask("TaskId")
                    Exit Function
                End If
            End If
        Next predecessor
    Next firstIndex
    IsScheduleFeasible = True
End Function

Private Function SelectPerson(ByVal people As Collection, ByVal task As Object, _
        ByVal personAvailable As Object, ByVal personLocation As Object, _
        ByVal moveMatrix As Object, ByVal taskReady As Double, _
        ByRef selectedMove As Double) As String
    Dim person As Object, personId As String, moveSec As Double
    Dim candidateStart As Double, bestStart As Double
    bestStart = 1E+99
    For Each person In people
        personId = CStr(person("PersonId"))
        If Len(CStr(task("LockedPersonId"))) = 0 Or _
                StrComp(personId, CStr(task("LockedPersonId")), vbTextCompare) = 0 Then
            If PersonCanPerform(person, task) Then
                moveSec = LookupMove(moveMatrix, CStr(personLocation(personId)), _
                    CStr(task("DeviceId")))
                candidateStart = MaxValue(taskReady, _
                    DictionaryNumber(personAvailable, personId) + moveSec)
                If candidateStart < bestStart Then
                    bestStart = candidateStart
                    SelectPerson = personId
                    selectedMove = moveSec
                End If
            End If
        End If
    Next person
End Function

Private Function PersonCanPerform(ByVal person As Object, ByVal task As Object) As Boolean
    Dim requiredSkill As String, allowedDevices As String
    requiredSkill = Trim$(CStr(task("RequiredSkill")))
    allowedDevices = Trim$(CStr(person("AllowedDevices")))
    If Len(requiredSkill) > 0 Then
        If Not TokenListContains(CStr(person("Skill")), requiredSkill) Then Exit Function
    End If
    If Len(allowedDevices) > 0 Then
        If Not TokenListContains(allowedDevices, CStr(task("DeviceId"))) Then Exit Function
    End If
    PersonCanPerform = True
End Function

Private Function TokenListContains(ByVal listText As String, ByVal token As String) As Boolean
    Dim normalized As String
    normalized = "," & Replace(Replace(Replace(listText, "；", ","), ";", ","), "，", ",") & ","
    TokenListContains = (InStr(1, normalized, "," & token & ",", vbTextCompare) > 0)
End Function

Private Function LookupMove(ByVal matrix As Object, ByVal fromId As String, _
        ByVal toId As String) As Double
    If Len(fromId) = 0 Or StrComp(fromId, toId, vbTextCompare) = 0 Then Exit Function
    If matrix.Exists(fromId & "|" & toId) Then
        If IsNumeric(matrix(fromId & "|" & toId)) Then LookupMove = CDbl(matrix(fromId & "|" & toId))
    End If
End Function

Private Function PredecessorsComplete(ByVal predecessorIds As String, ByVal scheduledEnd As Object, _
        ByRef latestEnd As Double) As Boolean
    Dim predecessor As Variant, key As String
    latestEnd = 0#
    If Len(predecessorIds) = 0 Then
        PredecessorsComplete = True
        Exit Function
    End If
    For Each predecessor In Split(predecessorIds, ",")
        key = Trim$(CStr(predecessor))
        If Len(key) > 0 Then
            If Not scheduledEnd.Exists(key) Then Exit Function
            If CDbl(scheduledEnd(key)) > latestEnd Then latestEnd = CDbl(scheduledEnd(key))
        End If
    Next predecessor
    PredecessorsComplete = True
End Function

Private Function FindStep(ByVal steps As Collection, ByVal deviceId As String, _
        ByVal stepNo As Long) As Object
    Dim item As Object
    For Each item In steps
        If StrComp(CStr(item("DeviceId")), deviceId, vbTextCompare) = 0 And _
                CLng(item("StepNo")) = stepNo Then
            Set FindStep = item
            Exit Function
        End If
    Next item
End Function

Private Function ReadCycleCount() As Long
    On Error GoTo InvalidCycle
    ReadCycleCount = CLng(ThisWorkbook.Names("nmCycleCount").RefersToRange.Value2)
    If ReadCycleCount < 1 Then GoTo InvalidCycle
    Exit Function
InvalidCycle:
    Err.Raise vbObjectError + 2104, "ReadCycleCount", "计划分析循环数无效。"
End Function

Private Function TaskKey(ByVal deviceId As String, ByVal cycleNo As Long, _
        ByVal stepNo As Long) As String
    TaskKey = deviceId & "-C" & cycleNo & "-S" & stepNo
End Function

Private Sub AppendPredecessor(ByRef predecessorIds As String, ByVal predecessorId As String)
    If Len(predecessorIds) > 0 Then predecessorIds = predecessorIds & ","
    predecessorIds = predecessorIds & predecessorId
End Sub

Private Function UsesPerson(ByVal stepType As String) As Boolean
    UsesPerson = (stepType = STEP_MANUAL Or stepType = STEP_JOINT)
End Function

Private Function DictionaryNumber(ByVal dictionary As Object, ByVal key As String) As Double
    If dictionary.Exists(key) Then DictionaryNumber = CDbl(dictionary(key))
End Function

Private Function MaxValue(ByVal firstValue As Double, ByVal secondValue As Double) As Double
    If firstValue >= secondValue Then MaxValue = firstValue Else MaxValue = secondValue
End Function

Private Function IntervalsOverlap(ByVal firstTask As Object, ByVal secondTask As Object) As Boolean
    IntervalsOverlap = (CDbl(firstTask("StartSec")) < CDbl(secondTask("EndSec")) And _
        CDbl(secondTask("StartSec")) < CDbl(firstTask("EndSec")))
End Function
