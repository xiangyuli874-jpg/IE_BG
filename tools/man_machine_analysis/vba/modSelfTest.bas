Attribute VB_Name = "modSelfTest"
Option Explicit

Private mFixturePath As String
Private mBaselineTasks() As TaskDef
Private mBaselineTaskCount As Long

Public Function RunAllSelfTests(Optional ByVal fixturePath As String = "") As String
    On Error GoTo TestFailed
    mFixturePath = fixturePath

    Test_DomainConstants
    Test_DomainTypes
    Test_BaselineFixtureExpansion

    RunAllSelfTests = "Test_DomainConstants PASS; Test_DomainTypes PASS; " & _
        "Test_BaselineFixtureExpansion PASS"
    Exit Function

TestFailed:
    RunAllSelfTests = "SELF_TESTS FAIL: " & Err.Description
End Function

Public Sub Test_DomainConstants()
    AssertEqual STEP_MANUAL, "MANUAL", "manual step constant"
    AssertEqual STEP_AUTO, "AUTO", "automatic step constant"
    AssertEqual STEP_JOINT, "JOINT", "joint step constant"
    AssertEqual STEP_WAIT, "WAIT", "wait step constant"
    AssertEqual MAX_PEOPLE, 5, "maximum people"
    AssertEqual MAX_DEVICES, 10, "maximum devices"
    AssertEqual MAX_STEPS_PER_DEVICE, 20, "maximum steps per device"
End Sub

Public Sub Test_DomainTypes()
    Dim definition As TaskDef
    Dim scheduled As ScheduledTask

    definition.TaskId = "M2-C3-S5"
    definition.DeviceId = "M2"
    definition.CycleNo = 3
    definition.StepNo = 5
    definition.StepName = "下料"
    definition.StepType = STEP_MANUAL
    definition.DurationSec = 5#
    definition.RequiredSkill = "通用"
    definition.PredecessorId = "M2-C3-S4"
    definition.LockedPersonId = "P1"
    definition.LockedStartSec = 120#
    definition.HasLockedStart = True

    scheduled.Definition = definition
    scheduled.PersonId = "P1"
    scheduled.StartSec = 120#
    scheduled.EndSec = 125#
    scheduled.MoveSec = 2#
    scheduled.WaitSec = 1#
    scheduled.IsLocked = True

    AssertEqual scheduled.Definition.TaskId, "M2-C3-S5", "scheduled task definition"
    AssertEqual scheduled.Definition.DeviceId, "M2", "task device"
    AssertEqual scheduled.Definition.CycleNo, 3, "task cycle"
    AssertEqual scheduled.Definition.StepNo, 5, "task step"
    AssertEqual scheduled.Definition.StepType, STEP_MANUAL, "task step type"
    AssertEqual scheduled.Definition.DurationSec, 5#, "task duration"
    AssertEqual scheduled.PersonId, "P1", "scheduled person"
    AssertEqual scheduled.StartSec, 120#, "scheduled start"
    AssertEqual scheduled.EndSec, 125#, "scheduled end"
    AssertEqual scheduled.MoveSec, 2#, "scheduled move"
    AssertEqual scheduled.WaitSec, 1#, "scheduled wait"
    AssertEqual scheduled.IsLocked, True, "scheduled lock"
End Sub

Public Sub Test_BaselineFixtureExpansion()
    Dim deviceNumber As Long
    Dim cycleNumber As Long

    LoadBaselineFixture mFixturePath

    AssertEqual mBaselineTaskCount, 45, "expanded task count"
    AssertEqual CountDistinctDevices(), 3, "expanded device count"
    AssertEqual CountDistinctCycles(), 3, "expanded cycle count"

    For deviceNumber = 1 To 3
        For cycleNumber = 1 To 3
            AssertEqual CountTasks("M" & CStr(deviceNumber), cycleNumber), 5, _
                "five steps per device and cycle"
        Next cycleNumber
    Next deviceNumber
End Sub

Private Sub LoadBaselineFixture(ByVal fixturePath As String)
    Dim stream As Object
    Dim content As String
    Dim lines() As String
    Dim fields() As String
    Dim baseTasks() As TaskDef
    Dim baseCount As Long
    Dim lineNumber As Long
    Dim deviceNumber As Long
    Dim cycleNumber As Long
    Dim baseIndex As Long
    Dim taskIndex As Long

    If Len(fixturePath) = 0 Or Len(Dir$(fixturePath)) = 0 Then
        Err.Raise vbObjectError + 1100, "LoadBaselineFixture", _
            "baseline fixture not found: " & fixturePath
    End If

    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile fixturePath
    content = stream.ReadText
    stream.Close

    content = Replace(content, vbCrLf, vbLf)
    content = Replace(content, vbCr, vbLf)
    lines = Split(content, vbLf)
    ReDim baseTasks(1 To 5)

    For lineNumber = 1 To UBound(lines)
        If Len(Trim$(lines(lineNumber))) > 0 Then
            fields = Split(lines(lineNumber), ",")
            If UBound(fields) <> 5 Then
                Err.Raise vbObjectError + 1101, "LoadBaselineFixture", _
                    "invalid baseline CSV row: " & CStr(lineNumber + 1)
            End If

            baseCount = baseCount + 1
            If baseCount > 5 Then
                Err.Raise vbObjectError + 1102, "LoadBaselineFixture", _
                    "baseline must contain exactly five M1 steps"
            End If
            If fields(0) <> "M1" Then
                Err.Raise vbObjectError + 1103, "LoadBaselineFixture", _
                    "baseline source device must be M1"
            End If

            baseTasks(baseCount).DeviceId = fields(0)
            baseTasks(baseCount).StepNo = CLng(fields(1))
            baseTasks(baseCount).StepName = fields(2)
            baseTasks(baseCount).StepType = fields(3)
            baseTasks(baseCount).DurationSec = CDbl(fields(4))
            baseTasks(baseCount).RequiredSkill = fields(5)
        End If
    Next lineNumber

    If baseCount <> 5 Then
        Err.Raise vbObjectError + 1104, "LoadBaselineFixture", _
            "baseline must contain exactly five M1 steps"
    End If

    mBaselineTaskCount = baseCount * 3 * 3
    ReDim mBaselineTasks(1 To mBaselineTaskCount)

    For deviceNumber = 1 To 3
        For cycleNumber = 1 To 3
            For baseIndex = 1 To baseCount
                taskIndex = taskIndex + 1
                With mBaselineTasks(taskIndex)
                    .TaskId = "M" & CStr(deviceNumber) & "-C" & _
                        CStr(cycleNumber) & "-S" & CStr(baseTasks(baseIndex).StepNo)
                    .DeviceId = "M" & CStr(deviceNumber)
                    .CycleNo = cycleNumber
                    .StepNo = baseTasks(baseIndex).StepNo
                    .StepName = baseTasks(baseIndex).StepName
                    .StepType = baseTasks(baseIndex).StepType
                    .DurationSec = baseTasks(baseIndex).DurationSec
                    .RequiredSkill = baseTasks(baseIndex).RequiredSkill
                    If baseIndex > 1 Then
                        .PredecessorId = "M" & CStr(deviceNumber) & "-C" & _
                            CStr(cycleNumber) & "-S" & _
                            CStr(baseTasks(baseIndex - 1).StepNo)
                    End If
                End With
            Next baseIndex
        Next cycleNumber
    Next deviceNumber
End Sub

Private Function CountDistinctDevices() As Long
    Dim seen As Object
    Dim taskIndex As Long

    Set seen = CreateObject("Scripting.Dictionary")
    For taskIndex = 1 To mBaselineTaskCount
        seen(mBaselineTasks(taskIndex).DeviceId) = True
    Next taskIndex
    CountDistinctDevices = seen.Count
End Function

Private Function CountDistinctCycles() As Long
    Dim seen As Object
    Dim taskIndex As Long

    Set seen = CreateObject("Scripting.Dictionary")
    For taskIndex = 1 To mBaselineTaskCount
        seen(CStr(mBaselineTasks(taskIndex).CycleNo)) = True
    Next taskIndex
    CountDistinctCycles = seen.Count
End Function

Private Function CountTasks(ByVal deviceId As String, ByVal cycleNumber As Long) As Long
    Dim taskIndex As Long

    For taskIndex = 1 To mBaselineTaskCount
        If mBaselineTasks(taskIndex).DeviceId = deviceId And _
                mBaselineTasks(taskIndex).CycleNo = cycleNumber Then
            CountTasks = CountTasks + 1
        End If
    Next taskIndex
End Function

Private Sub AssertEqual(ByVal actual As Variant, ByVal expected As Variant, ByVal message As String)
    If actual <> expected Then
        Err.Raise vbObjectError + 1001, "AssertEqual", _
            message & ": expected [" & CStr(expected) & "], actual [" & CStr(actual) & "]"
    End If
End Sub
