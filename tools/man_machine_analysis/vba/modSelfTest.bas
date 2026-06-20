Attribute VB_Name = "modSelfTest"
Option Explicit

Private mFixturePath As String
Private mBaselineTasks() As TaskDef
Private mBaselineTaskCount As Long

Public Function RunAllSelfTests(Optional ByVal fixturePath As String = "") As String
    On Error GoTo TestFailed
    mFixturePath = fixturePath

    Test_ChineseSourceRoundTrip
    Test_DomainConstants
    Test_DomainTypes
    Test_BaselineFixtureExpansion
    Test_WorkbookStructure
    Test_WorkbookReaders
    Test_ValidationRejectsBadInputs

    RunAllSelfTests = ExpectedPassOutput()
    Exit Function

TestFailed:
    RunAllSelfTests = "SELF_TESTS FAIL: " & Err.Description
End Function

Public Sub Test_ChineseSourceRoundTrip()
    AssertEqual SourceChineseSample(), ChineseUnload() & "|" & ChineseGeneralSkill(), _
        "Chinese VBA source round trip"
End Sub

Public Sub Test_WorkbookReaders()
    Dim people As Collection
    Dim devices As Collection
    Dim steps As Collection
    Dim moveMatrix As Object
    Dim stepItem As Object

    LoadValidationFixture "VALID_MODEL"
    Set people = ReadPeople()
    Set devices = ReadDevices()
    Set steps = ReadSteps()
    Set moveMatrix = ReadMoveMatrix()

    AssertEqual people.Count, 1, "disabled people are skipped"
    AssertEqual devices.Count, 2, "disabled devices are skipped"
    AssertEqual steps.Count, 4, "blank steps are skipped"
    Set stepItem = steps(1)
    AssertEqual stepItem("StepType"), STEP_MANUAL, "manual display type mapping"
    Set stepItem = steps(2)
    AssertEqual stepItem("StepType"), STEP_AUTO, "automatic display type mapping"
    Set stepItem = steps(3)
    AssertEqual stepItem("StepType"), STEP_JOINT, "joint display type mapping"
    Set stepItem = steps(4)
    AssertEqual stepItem("StepType"), STEP_WAIT, "wait display type mapping"
    AssertEqual moveMatrix("M1|M2"), 2#, "move matrix read"
End Sub

Public Sub Test_ValidationRejectsBadInputs()
    LoadValidationFixture "DUPLICATE_STEP"
    AssertContains ValidateCurrentModel(), "步骤号重复", "duplicate step"

    LoadValidationFixture "NO_SKILLED_PERSON"
    AssertContains ValidateCurrentModel(), "无合格人员", "no skilled person"

    LoadValidationFixture "CYCLIC_PREDECESSOR"
    AssertContains ValidateCurrentModel(), "步骤前置循环依赖", "cyclic predecessor"

    LoadValidationFixture "NEGATIVE_MOVE"
    AssertContains ValidateCurrentModel(), "移动时间不能为负数", "negative move"

    LoadValidationFixture "LIMITS"
    AssertContains ValidateCurrentModel(), "人员数量超过上限", "people limit"
    AssertContains ValidateCurrentModel(), "设备数量超过上限", "device limit"
    AssertContains ValidateCurrentModel(), "每台设备步骤数超过上限", "step limit"

    LoadValidationFixture "BAD_DURATION"
    AssertContains ValidateCurrentModel(), "工时不能为空", "empty duration"
    AssertContains ValidateCurrentModel(), "工时必须为正数", "non-positive duration"

    LoadValidationFixture "MISSING_PREDECESSOR"
    AssertContains ValidateCurrentModel(), "前置步骤不存在", "missing predecessor"

    LoadValidationFixture "LOCKED_PERSON_SKILL"
    AssertContains ValidateCurrentModel(), "锁定人员技能不符", "locked person skill"

    LoadValidationFixture "LOCKED_START_PREDECESSOR"
    AssertContains ValidateCurrentModel(), "锁定开始早于前置完成", "locked start"

    LoadValidationFixture "MISSING_MOVE"
    AssertContains ValidateCurrentModel(), "移动时间缺值", "missing move"

    LoadValidationFixture "CYCLIC_DEVICE"
    AssertContains ValidateCurrentModel(), "设备关系循环依赖", "cyclic device relation"

    LoadValidationFixture "MULTIPLE_ERRORS"
    AssertContains ValidateCurrentModel(), "步骤号重复", "collect duplicate error"
    AssertContains ValidateCurrentModel(), "移动时间不能为负数", "collect move error"
    AssertContains ValidateCurrentModel(), vbLf, "errors are newline separated"
End Sub

Public Sub Test_WorkbookStructure()
    Dim expectedSheets As Variant
    Dim sheetName As Variant
    Dim visibleCount As Long
    Dim worksheet As Worksheet

    expectedSheets = Array("基础设置", "作业步骤", "自动排程", _
        "人机作业图", "改善对比报告")

    For Each worksheet In ThisWorkbook.Worksheets
        If worksheet.Visible = xlSheetVisible Then visibleCount = visibleCount + 1
    Next worksheet
    AssertEqual visibleCount, 5, "visible worksheet count"

    For Each sheetName In expectedSheets
        AssertSheetExists CStr(sheetName)
    Next sheetName

    AssertTableExists "基础设置", "tblParameters"
    AssertTableExists "基础设置", "tblPeople"
    AssertTableExists "基础设置", "tblDevices"
    AssertTableExists "基础设置", "tblMoveTime"
    AssertTableExists "作业步骤", "tblSteps"

    AssertTableHeaders "基础设置", "tblParameters", Array( _
        "方案名称", "人员数量", "设备数量", "计划分析循环数", "优化目标", "搜索迭代次数", _
        "周期权重", "均衡权重", "等待权重", "移动权重")
    AssertTableHeaders "基础设置", "tblPeople", Array( _
        "人员编号", "人员名称", "技能", "可操作设备", "启用")
    AssertTableHeaders "基础设置", "tblDevices", Array( _
        "设备编号", "设备名称", "产品", "关系类型", "前置设备", "启用")
    AssertTableHeaders "基础设置", "tblMoveTime", Array( _
        "起点", "M1", "M2", "M3", "M4", "M5", "M6", "M7", "M8", "M9", "M10")
    AssertTableHeaders "作业步骤", "tblSteps", Array( _
        "设备", "步骤号", "名称", "类型", "工时(s)", "前置步骤", _
        "所需技能", "锁定人员", "锁定开始(s)", "允许等待", "备注")

    AssertNameExists "nmOptimizationTarget"
    AssertNameExists "nmCycleCount"
    AssertNameExists "nmWeightCycle"
    AssertNameExists "nmWeightBalance"
    AssertNameExists "nmWeightWait"
    AssertNameExists "nmWeightMove"
    AssertNameExists "nmDeviceIds"
    AssertNameExists "nmPersonIds"

    AssertWorkbookNameTargetsCell "nmOptimizationTarget", "基础设置", _
        "tblParameters", "优化目标"
    AssertWorkbookNameTargetsCell "nmCycleCount", "基础设置", _
        "tblParameters", "计划分析循环数"
    AssertWorkbookNameTargetsCell "nmWeightCycle", "基础设置", _
        "tblParameters", "周期权重"
    AssertWorkbookNameTargetsCell "nmWeightBalance", "基础设置", _
        "tblParameters", "均衡权重"
    AssertWorkbookNameTargetsCell "nmWeightWait", "基础设置", _
        "tblParameters", "等待权重"
    AssertWorkbookNameTargetsCell "nmWeightMove", "基础设置", _
        "tblParameters", "移动权重"
    AssertWorkbookNameTargetsColumn "nmDeviceIds", "基础设置", _
        "tblDevices", "设备编号"
    AssertWorkbookNameTargetsColumn "nmPersonIds", "基础设置", _
        "tblPeople", "人员编号"

    AssertValidationList "基础设置", "tblParameters", "优化目标", "最高产能"
    AssertValidationList "基础设置", "tblPeople", "启用", "是"
    AssertValidationList "基础设置", "tblDevices", "关系类型", "独立循环"
    AssertValidationList "基础设置", "tblDevices", "启用", "是"
    AssertValidationList "作业步骤", "tblSteps", "类型", "人工"
    AssertValidationList "作业步骤", "tblSteps", "允许等待", "是"
    AssertResolvableListValidation "作业步骤", "tblSteps", "设备", _
        "nmDeviceIds", "基础设置", "tblDevices", "设备编号"
    AssertResolvableListValidation "作业步骤", "tblSteps", "锁定人员", _
        "nmPersonIds", "基础设置", "tblPeople", "人员编号"
    AssertValidationAllowsBlank "作业步骤", "tblSteps", "锁定人员"

    AssertTableCapacity "基础设置", "tblPeople", MAX_PEOPLE
    AssertTableCapacity "基础设置", "tblDevices", MAX_DEVICES
    AssertTableCapacity "基础设置", "tblMoveTime", MAX_DEVICES
    AssertEqual ThisWorkbook.Worksheets("基础设置").ListObjects("tblMoveTime") _
        .ListColumns.Count, MAX_DEVICES + 1, "move matrix column capacity"
    AssertTableCapacity "作业步骤", "tblSteps", _
        MAX_DEVICES * MAX_STEPS_PER_DEVICE

    AssertCellValue "基础设置", "tblParameters", "人员数量", 1
    AssertCellValue "基础设置", "tblParameters", "设备数量", 3
    AssertCellFill "基础设置", "tblParameters", "人员数量", RGB(255, 242, 204)
    AssertCellFill "基础设置", "tblParameters", "设备数量", RGB(255, 242, 204)
    AssertCellFill "基础设置", "tblParameters", "计划分析循环数", RGB(255, 242, 204)
    AssertCellFill "基础设置", "tblPeople", "人员名称", RGB(221, 235, 247)
    AssertCellFill "自动排程", "", "", RGB(242, 242, 242)
    AssertWorkbookFont "微软雅黑"
End Sub

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
    definition.StepName = SourceChineseUnload()
    definition.StepType = STEP_MANUAL
    definition.DurationSec = 5#
    definition.RequiredSkill = SourceChineseGeneralSkill()
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
    AssertEqual scheduled.Definition.StepName, ChineseUnload(), "task Chinese step name"
    AssertEqual scheduled.Definition.StepType, STEP_MANUAL, "task step type"
    AssertEqual scheduled.Definition.DurationSec, 5#, "task duration"
    AssertEqual scheduled.Definition.RequiredSkill, ChineseGeneralSkill(), _
        "task Chinese skill"
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
            AssertBaselineTaskFields "M" & CStr(deviceNumber), cycleNumber
            If cycleNumber > 1 Then
                AssertEqual FindTask("M" & CStr(deviceNumber), cycleNumber, 1).PredecessorId, _
                    "M" & CStr(deviceNumber) & "-C" & CStr(cycleNumber - 1) & "-S5", _
                    "first step continues from previous cycle"
            End If
        Next cycleNumber
    Next deviceNumber
End Sub

Private Sub AssertBaselineTaskFields(ByVal deviceId As String, ByVal cycleNumber As Long)
    AssertTaskFields FindTask(deviceId, cycleNumber, 1), ChineseLoad(), _
        STEP_MANUAL, 5#, ChineseGeneralSkill(), ExpectedPredecessor(deviceId, cycleNumber, 1)
    AssertTaskFields FindTask(deviceId, cycleNumber, 2), ChineseAutoRun() & "1", _
        STEP_AUTO, 20#, "", ExpectedPredecessor(deviceId, cycleNumber, 2)
    AssertTaskFields FindTask(deviceId, cycleNumber, 3), ChineseTurnOver(), _
        STEP_MANUAL, 6#, ChineseGeneralSkill(), ExpectedPredecessor(deviceId, cycleNumber, 3)
    AssertTaskFields FindTask(deviceId, cycleNumber, 4), ChineseAutoRun() & "2", _
        STEP_AUTO, 15#, "", ExpectedPredecessor(deviceId, cycleNumber, 4)
    AssertTaskFields FindTask(deviceId, cycleNumber, 5), ChineseUnload(), _
        STEP_MANUAL, 5#, ChineseGeneralSkill(), ExpectedPredecessor(deviceId, cycleNumber, 5)
End Sub

Private Sub AssertTaskFields(ByRef task As TaskDef, ByVal stepName As String, _
        ByVal stepType As String, ByVal durationSec As Double, ByVal skill As String, _
        ByVal predecessorId As String)
    AssertEqual task.StepName, stepName, "baseline step name"
    AssertEqual task.StepType, stepType, "baseline step type"
    AssertEqual task.DurationSec, durationSec, "baseline duration"
    AssertEqual task.RequiredSkill, skill, "baseline skill"
    AssertEqual task.PredecessorId, predecessorId, "baseline predecessor"
End Sub

Private Function ExpectedPredecessor(ByVal deviceId As String, ByVal cycleNumber As Long, _
        ByVal stepNumber As Long) As String
    If stepNumber > 1 Then
        ExpectedPredecessor = deviceId & "-C" & CStr(cycleNumber) & _
            "-S" & CStr(stepNumber - 1)
    ElseIf cycleNumber > 1 Then
        ExpectedPredecessor = deviceId & "-C" & CStr(cycleNumber - 1) & "-S5"
    End If
End Function

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
                    ElseIf cycleNumber > 1 Then
                        .PredecessorId = "M" & CStr(deviceNumber) & "-C" & _
                            CStr(cycleNumber - 1) & "-S" & _
                            CStr(baseTasks(baseCount).StepNo)
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

Private Function FindTask(ByVal deviceId As String, ByVal cycleNumber As Long, _
        ByVal stepNumber As Long) As TaskDef
    Dim taskIndex As Long

    For taskIndex = 1 To mBaselineTaskCount
        If mBaselineTasks(taskIndex).DeviceId = deviceId And _
                mBaselineTasks(taskIndex).CycleNo = cycleNumber And _
                mBaselineTasks(taskIndex).StepNo = stepNumber Then
            FindTask = mBaselineTasks(taskIndex)
            Exit Function
        End If
    Next taskIndex

    Err.Raise vbObjectError + 1105, "FindTask", _
        "expanded task not found"
End Function

Private Function SourceChineseSample() As String
    SourceChineseSample = "下料|通用"
End Function

Private Function SourceChineseUnload() As String
    SourceChineseUnload = "下料"
End Function

Private Function SourceChineseGeneralSkill() As String
    SourceChineseGeneralSkill = "通用"
End Function

Private Function ChineseLoad() As String
    ChineseLoad = ChrW$(&H4E0A) & ChrW$(&H6599)
End Function

Private Function ChineseUnload() As String
    ChineseUnload = ChrW$(&H4E0B) & ChrW$(&H6599)
End Function

Private Function ChineseGeneralSkill() As String
    ChineseGeneralSkill = ChrW$(&H901A) & ChrW$(&H7528)
End Function

Private Function ChineseAutoRun() As String
    ChineseAutoRun = ChrW$(&H81EA) & ChrW$(&H52A8) & _
        ChrW$(&H8FD0) & ChrW$(&H884C)
End Function

Private Function ChineseTurnOver() As String
    ChineseTurnOver = ChrW$(&H7FFB) & ChrW$(&H9762)
End Function

Private Function ExpectedPassOutput() As String
    ExpectedPassOutput = "Test_ChineseSourceRoundTrip PASS; " & _
        "Test_DomainConstants PASS; Test_DomainTypes PASS; " & _
        "Test_BaselineFixtureExpansion PASS; Test_WorkbookStructure PASS; " & _
        "Test_WorkbookReaders PASS; Test_ValidationRejectsBadInputs PASS"
End Function

Private Sub LoadValidationFixture(ByVal fixtureName As String)
    ResetValidationModel

    Select Case fixtureName
        Case "VALID_MODEL"
            ' Baseline is already valid.
        Case "DUPLICATE_STEP"
            SetStepRow 5, "M1", 1, "重复上料", "人工", 3#, "", "通用"
        Case "NO_SKILLED_PERSON"
            ThisWorkbook.Worksheets("基础设置").ListObjects("tblPeople") _
                .ListColumns("技能").DataBodyRange.Cells(1, 1).ClearContents
            SetStepRow 1, "M1", 1, "上料", "人工", 5#, "", "焊接"
        Case "CYCLIC_PREDECESSOR"
            SetStepRow 1, "M1", 1, "上料", "人工", 5#, 2, "通用"
            SetStepRow 2, "M1", 2, "加工", "自动运行", 10#, 1, ""
        Case "NEGATIVE_MOVE"
            SetMoveValue "M1", "M2", -1#
        Case "LIMITS"
            SetParameter "人员数量", MAX_PEOPLE + 1
            SetParameter "设备数量", MAX_DEVICES + 1
            Dim rowIndex As Long
            For rowIndex = 1 To MAX_STEPS_PER_DEVICE + 1
                SetStepRow rowIndex, "M1", rowIndex, "步骤" & CStr(rowIndex), _
                    "自动运行", 1#, IIf(rowIndex = 1, "", rowIndex - 1), ""
            Next rowIndex
        Case "BAD_DURATION"
            SetStepRow 1, "M1", 1, "空工时", "人工", "", "", "通用"
            SetStepRow 2, "M1", 2, "零工时", "自动运行", 0#, 1, ""
        Case "MISSING_PREDECESSOR"
            SetStepRow 1, "M1", 1, "上料", "人工", 5#, 99, "通用"
        Case "LOCKED_PERSON_SKILL"
            SetStepRow 1, "M1", 1, "焊接", "人工", 5#, "", "焊接", "P1"
        Case "LOCKED_START_PREDECESSOR"
            SetStepRow 1, "M1", 1, "上料", "人工", 5#, "", "通用", "", 0#
            SetStepRow 2, "M1", 2, "下料", "人工", 5#, 1, "通用", "P1", 3#
        Case "MISSING_MOVE"
            SetMoveValue "M1", "M2", ""
        Case "CYCLIC_DEVICE"
            SetDevicePredecessor "M1", "M2"
            SetDevicePredecessor "M2", "M1"
        Case "MULTIPLE_ERRORS"
            SetStepRow 5, "M1", 1, "重复上料", "人工", 3#, "", "通用"
            SetMoveValue "M1", "M2", -1#
        Case Else
            Err.Raise vbObjectError + 1300, "LoadValidationFixture", _
                "unknown validation fixture: " & fixtureName
    End Select
End Sub

Private Sub ResetValidationModel()
    Dim peopleTable As ListObject
    Dim devicesTable As ListObject
    Dim moveTable As ListObject
    Dim stepsTable As ListObject

    Set peopleTable = ThisWorkbook.Worksheets("基础设置").ListObjects("tblPeople")
    Set devicesTable = ThisWorkbook.Worksheets("基础设置").ListObjects("tblDevices")
    Set moveTable = ThisWorkbook.Worksheets("基础设置").ListObjects("tblMoveTime")
    Set stepsTable = ThisWorkbook.Worksheets("作业步骤").ListObjects("tblSteps")

    peopleTable.DataBodyRange.ClearContents
    devicesTable.DataBodyRange.ClearContents
    stepsTable.DataBodyRange.ClearContents

    SetParameter "人员数量", 1
    SetParameter "设备数量", 2

    SetTableRow peopleTable, 1, Array("P1", "操作员1", "通用", "M1,M2", "是")
    SetTableRow peopleTable, 2, Array("P2", "停用人员", "焊接", "M1", "否")
    SetTableRow devicesTable, 1, Array("M1", "设备1", "", "独立循环", "", "是")
    SetTableRow devicesTable, 2, Array("M2", "设备2", "", "有先后顺序", "M1", "是")
    SetTableRow devicesTable, 3, Array("M3", "停用设备", "", "独立循环", "", "否")

    Dim fromIndex As Long
    Dim toIndex As Long
    For fromIndex = 1 To MAX_DEVICES
        moveTable.DataBodyRange.Cells(fromIndex, 1).Value2 = "M" & CStr(fromIndex)
        For toIndex = 1 To MAX_DEVICES
            moveTable.DataBodyRange.Cells(fromIndex, toIndex + 1).Value2 = 0#
        Next toIndex
    Next fromIndex
    SetMoveValue "M1", "M2", 2#
    SetMoveValue "M2", "M1", 2#

    SetStepRow 1, "M1", 1, "上料", "人工", 5#, "", "通用"
    SetStepRow 2, "M1", 2, "加工", "自动运行", 10#, 1, ""
    SetStepRow 3, "M2", 1, "协同", "人机协同", 4#, "", "通用"
    SetStepRow 4, "M2", 2, "冷却", "等待", 3#, 1, ""
End Sub

Private Sub SetParameter(ByVal columnName As String, ByVal value As Variant)
    ThisWorkbook.Worksheets("基础设置").ListObjects("tblParameters") _
        .ListColumns(columnName).DataBodyRange.Cells(1, 1).Value2 = value
End Sub

Private Sub SetTableRow(ByVal table As ListObject, ByVal rowIndex As Long, _
        ByVal values As Variant)
    Dim columnIndex As Long
    For columnIndex = LBound(values) To UBound(values)
        table.DataBodyRange.Cells(rowIndex, columnIndex + 1).Value2 = values(columnIndex)
    Next columnIndex
End Sub

Private Sub SetStepRow(ByVal rowIndex As Long, ByVal deviceId As String, _
        ByVal stepNumber As Variant, ByVal stepName As String, ByVal displayType As String, _
        ByVal durationSec As Variant, ByVal predecessor As Variant, ByVal skill As String, _
        Optional ByVal lockedPerson As String = "", _
        Optional ByVal lockedStart As Variant)
    Dim table As ListObject
    Dim values As Variant

    Set table = ThisWorkbook.Worksheets("作业步骤").ListObjects("tblSteps")
    values = Array(deviceId, stepNumber, stepName, displayType, durationSec, predecessor, _
        skill, lockedPerson, "", "是", "")
    If Not IsMissing(lockedStart) Then values(8) = lockedStart
    SetTableRow table, rowIndex, values
End Sub

Private Sub SetMoveValue(ByVal fromDevice As String, ByVal toDevice As String, _
        ByVal value As Variant)
    Dim table As ListObject
    Dim rowIndex As Long
    Dim columnIndex As Long

    Set table = ThisWorkbook.Worksheets("基础设置").ListObjects("tblMoveTime")
    rowIndex = CLng(Mid$(fromDevice, 2))
    columnIndex = CLng(Mid$(toDevice, 2)) + 1
    table.DataBodyRange.Cells(rowIndex, columnIndex).Value2 = value
End Sub

Private Sub SetDevicePredecessor(ByVal deviceId As String, ByVal predecessorId As String)
    Dim table As ListObject
    Dim rowIndex As Long

    Set table = ThisWorkbook.Worksheets("基础设置").ListObjects("tblDevices")
    For rowIndex = 1 To table.ListRows.Count
        If CStr(table.DataBodyRange.Cells(rowIndex, 1).Value2) = deviceId Then
            table.DataBodyRange.Cells(rowIndex, 4).Value2 = "有先后顺序"
            table.DataBodyRange.Cells(rowIndex, 5).Value2 = predecessorId
            Exit Sub
        End If
    Next rowIndex
End Sub

Private Sub AssertContains(ByVal actual As String, ByVal expectedText As String, _
        ByVal message As String)
    If InStr(1, actual, expectedText, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 1002, "AssertContains", _
            message & ": expected text [" & expectedText & "], actual [" & actual & "]"
    End If
End Sub

Private Sub AssertSheetExists(ByVal sheetName As String)
    Dim worksheet As Worksheet

    On Error Resume Next
    Set worksheet = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If worksheet Is Nothing Then
        Err.Raise vbObjectError + 1200, "AssertSheetExists", _
            sheetName & " sheet missing"
    End If
End Sub

Private Sub AssertTableExists(ByVal sheetName As String, ByVal tableName As String)
    Dim table As ListObject

    On Error Resume Next
    Set table = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
    On Error GoTo 0
    If table Is Nothing Then
        Err.Raise vbObjectError + 1201, "AssertTableExists", _
            tableName & " table missing"
    End If
End Sub

Private Sub AssertTableHeaders(ByVal sheetName As String, ByVal tableName As String, _
        ByVal expectedHeaders As Variant)
    Dim table As ListObject
    Dim headerIndex As Long

    Set table = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName)
    AssertEqual table.ListColumns.Count, UBound(expectedHeaders) + 1, _
        tableName & " column count"
    For headerIndex = 0 To UBound(expectedHeaders)
        AssertEqual table.HeaderRowRange.Cells(1, headerIndex + 1).Value2, _
            expectedHeaders(headerIndex), tableName & " header " & CStr(headerIndex + 1)
    Next headerIndex
End Sub

Private Sub AssertNameExists(ByVal definedName As String)
    Dim workbookName As Name

    On Error Resume Next
    Set workbookName = ThisWorkbook.Names(definedName)
    On Error GoTo 0
    If workbookName Is Nothing Then
        Err.Raise vbObjectError + 1202, "AssertNameExists", _
            definedName & " name missing"
    End If
End Sub

Private Sub AssertWorkbookNameTargetsCell(ByVal definedName As String, _
        ByVal sheetName As String, ByVal tableName As String, ByVal columnName As String)
    Dim targetCell As Range

    Set targetCell = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName) _
        .ListColumns(columnName).DataBodyRange.Cells(1, 1)
    AssertWorkbookNameTargetsRange definedName, targetCell
End Sub

Private Sub AssertWorkbookNameTargetsColumn(ByVal definedName As String, _
        ByVal sheetName As String, ByVal tableName As String, ByVal columnName As String)
    Dim targetRange As Range

    Set targetRange = ThisWorkbook.Worksheets(sheetName).ListObjects(tableName) _
        .ListColumns(columnName).DataBodyRange
    AssertWorkbookNameTargetsRange definedName, targetRange
End Sub

Private Sub AssertWorkbookNameTargetsRange(ByVal definedName As String, _
        ByVal expectedRange As Range)
    Dim workbookName As Name
    Dim actualRange As Range

    Set workbookName = ThisWorkbook.Names(definedName)
    If TypeName(workbookName.Parent) <> "Workbook" Then
        Err.Raise vbObjectError + 1205, "AssertWorkbookNameTargetsRange", _
            definedName & " must have workbook scope"
    End If
    Set actualRange = workbookName.RefersToRange
    AssertEqual actualRange.Address(True, True, xlA1, True), _
        expectedRange.Address(True, True, xlA1, True), _
        definedName & " target range"
End Sub

Private Sub AssertValidationList(ByVal sheetName As String, ByVal tableName As String, _
        ByVal columnName As String, ByVal expectedItem As String)
    Dim targetCell As Range
    Dim formulaText As String

    Set targetCell = ThisWorkbook.Worksheets(sheetName) _
        .ListObjects(tableName).ListColumns(columnName).DataBodyRange.Cells(1, 1)
    If targetCell.Validation.Type <> xlValidateList Then
        Err.Raise vbObjectError + 1203, "AssertValidationList", _
            tableName & "." & columnName & " validation missing"
    End If
    formulaText = targetCell.Validation.Formula1
    If InStr(1, formulaText, expectedItem, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 1204, "AssertValidationList", _
            tableName & "." & columnName & " validation item missing: " & expectedItem
    End If
End Sub

Private Sub AssertResolvableListValidation(ByVal sheetName As String, ByVal tableName As String, _
        ByVal columnName As String, ByVal expectedName As String, _
        ByVal sourceSheetName As String, ByVal sourceTableName As String, _
        ByVal sourceColumnName As String)
    Dim targetCell As Range
    Dim resolvedRange As Range
    Dim expectedRange As Range
    Dim formulaText As String

    Set targetCell = ThisWorkbook.Worksheets(sheetName) _
        .ListObjects(tableName).ListColumns(columnName).DataBodyRange.Cells(1, 1)
    If targetCell.Validation.Type <> xlValidateList Then
        Err.Raise vbObjectError + 1206, "AssertResolvableListValidation", _
            tableName & "." & columnName & " validation missing"
    End If
    formulaText = targetCell.Validation.Formula1
    AssertEqual formulaText, "=" & expectedName, _
        tableName & "." & columnName & " validation formula"

    On Error Resume Next
    Set resolvedRange = targetCell.Parent.Evaluate(formulaText)
    On Error GoTo 0
    If resolvedRange Is Nothing Then
        Err.Raise vbObjectError + 1207, "AssertResolvableListValidation", _
            tableName & "." & columnName & " validation source cannot be evaluated"
    End If

    Set expectedRange = ThisWorkbook.Worksheets(sourceSheetName) _
        .ListObjects(sourceTableName).ListColumns(sourceColumnName).DataBodyRange
    AssertEqual resolvedRange.Address(True, True, xlA1, True), _
        expectedRange.Address(True, True, xlA1, True), _
        tableName & "." & columnName & " resolved validation source"
End Sub

Private Sub AssertTableCapacity(ByVal sheetName As String, ByVal tableName As String, _
        ByVal expectedRows As Long)
    AssertEqual ThisWorkbook.Worksheets(sheetName).ListObjects(tableName) _
        .ListRows.Count, expectedRows, tableName & " row capacity"
End Sub

Private Sub AssertValidationAllowsBlank(ByVal sheetName As String, ByVal tableName As String, _
        ByVal columnName As String)
    Dim targetCell As Range

    Set targetCell = ThisWorkbook.Worksheets(sheetName) _
        .ListObjects(tableName).ListColumns(columnName).DataBodyRange.Cells(1, 1)
    AssertEqual targetCell.Validation.IgnoreBlank, True, _
        tableName & "." & columnName & " validation allows blank"
End Sub

Private Sub AssertCellValue(ByVal sheetName As String, ByVal tableName As String, _
        ByVal columnName As String, ByVal expectedValue As Variant)
    Dim targetCell As Range

    Set targetCell = ThisWorkbook.Worksheets(sheetName) _
        .ListObjects(tableName).ListColumns(columnName).DataBodyRange.Cells(1, 1)
    AssertEqual targetCell.Value2, expectedValue, _
        sheetName & " " & columnName & " default value"
End Sub

Private Sub AssertCellFill(ByVal sheetName As String, ByVal tableName As String, _
        ByVal columnName As String, ByVal expectedColor As Long)
    Dim targetCell As Range

    If Len(tableName) = 0 Then
        Set targetCell = ThisWorkbook.Worksheets(sheetName).Range("A2")
    Else
        Set targetCell = ThisWorkbook.Worksheets(sheetName) _
            .ListObjects(tableName).ListColumns(columnName).DataBodyRange.Cells(1, 1)
    End If
    AssertEqual targetCell.Interior.Color, expectedColor, _
        sheetName & " " & columnName & " fill"
End Sub

Private Sub AssertWorkbookFont(ByVal expectedFont As String)
    Dim worksheet As Worksheet

    For Each worksheet In ThisWorkbook.Worksheets
        AssertEqual worksheet.Cells.Font.Name, expectedFont, _
            worksheet.Name & " default font"
    Next worksheet
End Sub

Private Sub AssertEqual(ByVal actual As Variant, ByVal expected As Variant, ByVal message As String)
    If actual <> expected Then
        Err.Raise vbObjectError + 1001, "AssertEqual", _
            message & ": expected [" & CStr(expected) & "], actual [" & CStr(actual) & "]"
    End If
End Sub
