Attribute VB_Name = "modValidation"
Option Explicit

Public Function ValidateCurrentModel() As String
    Dim errors As String
    Dim people As Collection
    Dim devices As Collection
    Dim steps As Collection
    Dim moveMatrix As Object

    Set people = ReadPeople()
    Set devices = ReadDevices()
    Set steps = ReadSteps()
    Set moveMatrix = ReadMoveMatrix()

    ValidateLimits errors, people, devices, steps
    ValidateSteps errors, people, devices, steps
    ValidateMoveMatrix errors, devices, moveMatrix
    ValidateDeviceRelations errors, devices
    ValidateCurrentModel = errors
End Function

Private Sub ValidateLimits(ByRef errors As String, ByVal people As Collection, _
        ByVal devices As Collection, ByVal steps As Collection)
    Dim parameters As ListObject
    Dim configuredValue As Long
    Dim counts As Object
    Dim stepItem As Variant
    Dim deviceId As Variant

    Set parameters = ThisWorkbook.Worksheets("基础设置").ListObjects("tblParameters")

    If TryParsePositiveLong(ParameterValue(parameters, "人员数量"), configuredValue) Then
        If configuredValue > MAX_PEOPLE Then
            AddError errors, "[基础设置!tblParameters] 人员数量超过上限 " & _
                CStr(MAX_PEOPLE) & "。"
        End If
    Else
        AddError errors, "[基础设置!tblParameters] 人员数量必须为正整数。"
    End If
    If people.Count > MAX_PEOPLE Then
        AddError errors, "[基础设置!tblPeople] 人员数量超过上限 " & _
            CStr(MAX_PEOPLE) & "。"
    End If

    If TryParsePositiveLong(ParameterValue(parameters, "设备数量"), configuredValue) Then
        If configuredValue > MAX_DEVICES Then
            AddError errors, "[基础设置!tblParameters] 设备数量超过上限 " & _
                CStr(MAX_DEVICES) & "。"
        End If
    Else
        AddError errors, "[基础设置!tblParameters] 设备数量必须为正整数。"
    End If
    If devices.Count > MAX_DEVICES Then
        AddError errors, "[基础设置!tblDevices] 设备数量超过上限 " & _
            CStr(MAX_DEVICES) & "。"
    End If

    If Not TryParsePositiveLong(ParameterValue(parameters, "计划分析循环数"), _
            configuredValue) Then
        AddError errors, "[基础设置!tblParameters] 计划分析循环数必须为正整数。"
    End If
    If Not TryParsePositiveLong(ParameterValue(parameters, "搜索迭代次数"), _
            configuredValue) Then
        AddError errors, "[基础设置!tblParameters] 搜索迭代次数必须为正整数。"
    End If

    Set counts = CreateObject("Scripting.Dictionary")
    counts.CompareMode = vbTextCompare
    For Each stepItem In steps
        counts(stepItem("DeviceId")) = counts(stepItem("DeviceId")) + 1
    Next stepItem
    For Each deviceId In counts.Keys
        If counts(deviceId) > MAX_STEPS_PER_DEVICE Then
            AddError errors, "[作业步骤!设备 " & CStr(deviceId) & _
                "] 每台设备步骤数超过上限 " & CStr(MAX_STEPS_PER_DEVICE) & "。"
        End If
    Next deviceId
End Sub

Private Sub ValidateSteps(ByRef errors As String, ByVal people As Collection, _
        ByVal devices As Collection, ByVal steps As Collection)
    Dim stepByKey As Object
    Dim seenKeys As Object
    Dim deviceHasStep As Object
    Dim stepItem As Variant
    Dim device As Variant
    Dim key As String
    Dim predecessorKey As String
    Dim earliestCache As Object
    Dim parsedStepNo As Long
    Dim lockedStart As Double
    Dim cycleKey As String

    Set stepByKey = CreateObject("Scripting.Dictionary")
    stepByKey.CompareMode = vbTextCompare
    Set seenKeys = CreateObject("Scripting.Dictionary")
    seenKeys.CompareMode = vbTextCompare
    Set deviceHasStep = CreateObject("Scripting.Dictionary")
    deviceHasStep.CompareMode = vbTextCompare

    For Each stepItem In steps
        key = StepKey(stepItem("DeviceId"), stepItem("StepNoText"))
        deviceHasStep(stepItem("DeviceId")) = True

        If Not TryParsePositiveLong(stepItem("StepNoText"), parsedStepNo) Then
            AddStepError errors, stepItem, "步骤号必须为正整数。"
        End If

        If seenKeys.Exists(key) Then
            AddStepError errors, stepItem, "步骤号重复：" & stepItem("StepNoText") & "。"
        Else
            seenKeys(key) = True
            Set stepByKey(key) = stepItem
        End If

        If Not HasTextValue(stepItem("DurationRaw")) Then
            AddStepError errors, stepItem, "工时不能为空。"
        ElseIf Not IsNumeric(stepItem("DurationRaw")) Or stepItem("DurationSec") <= 0# Then
            AddStepError errors, stepItem, "工时必须为正数。"
        End If

        If Len(stepItem("StepType")) = 0 Then
            AddStepError errors, stepItem, "步骤类型无效：" & stepItem("DisplayType") & "。"
        End If
    Next stepItem

    ValidateContinuousStepNumbers errors, steps

    For Each device In devices
        If Not deviceHasStep.Exists(device("DeviceId")) Then
            AddError errors, "[基础设置!tblDevices 行" & CStr(device("SourceRow")) & _
                "，设备 " & device("DeviceId") & "] 已启用设备没有作业步骤。"
        End If
    Next device

    For Each stepItem In steps
        predecessorKey = ResolvePredecessorKey(stepItem)
        If Len(predecessorKey) > 0 And Not stepByKey.Exists(predecessorKey) Then
            AddStepError errors, stepItem, "前置步骤不存在：" & _
                stepItem("PredecessorText") & "。"
        End If

        If RequiresPerson(stepItem("StepType")) Then
            If Not AnyQualifiedPerson(people, stepItem) Then
                AddStepError errors, stepItem, "无合格人员可执行该步骤。"
            End If
        End If

        If Len(stepItem("LockedPersonId")) > 0 Then
            If Not PersonIsQualified(people, stepItem("LockedPersonId"), stepItem) Then
                AddStepError errors, stepItem, "锁定人员技能不符：" & _
                    stepItem("LockedPersonId") & "。"
            End If
        End If
    Next stepItem

    cycleKey = FindStepCycleNode(stepByKey)
    If Len(cycleKey) > 0 Then
        Set stepItem = stepByKey(cycleKey)
        AddStepError errors, stepItem, "步骤前置循环依赖。"
    End If

    Set earliestCache = CreateObject("Scripting.Dictionary")
    earliestCache.CompareMode = vbTextCompare
    For Each stepItem In steps
        predecessorKey = ResolvePredecessorKey(stepItem)
        If stepItem("HasLockedStart") Then
            If Not TryParseDouble(stepItem("LockedStartRaw"), lockedStart) Then
                AddStepError errors, stepItem, "锁定开始必须为数字。"
            ElseIf lockedStart < 0# Then
                AddStepError errors, stepItem, "锁定开始不能为负数。"
            ElseIf Len(predecessorKey) > 0 Then
                If stepByKey.Exists(predecessorKey) Then
                    If lockedStart < EarliestFinish(predecessorKey, _
                            stepByKey, earliestCache, CreateObject("Scripting.Dictionary")) Then
                        AddStepError errors, stepItem, "锁定开始早于前置完成。"
                    End If
                End If
            End If
        End If
    Next stepItem
End Sub

Private Sub ValidateMoveMatrix(ByRef errors As String, ByVal devices As Collection, _
        ByVal moveMatrix As Object)
    Dim fromDevice As Variant
    Dim toDevice As Variant
    Dim key As String
    Dim value As Variant

    For Each fromDevice In devices
        For Each toDevice In devices
            key = fromDevice("DeviceId") & "|" & toDevice("DeviceId")
            If Not moveMatrix.Exists(key) Then
                AddError errors, "[基础设置!tblMoveTime " & key & "] 移动时间缺值。"
            Else
                value = moveMatrix(key)
                If Not HasTextValue(value) Or Not IsNumeric(value) Then
                    AddError errors, "[基础设置!tblMoveTime " & key & "] 移动时间缺值。"
                ElseIf CDbl(value) < 0# Then
                    AddError errors, "[基础设置!tblMoveTime " & key & _
                        "] 移动时间不能为负数。"
                End If
            End If
        Next toDevice
    Next fromDevice
End Sub

Private Sub ValidateDeviceRelations(ByRef errors As String, ByVal devices As Collection)
    Dim deviceById As Object
    Dim graph As Object
    Dim device As Variant
    Dim predecessorId As String
    Dim cycleDeviceId As String

    Set deviceById = CreateObject("Scripting.Dictionary")
    deviceById.CompareMode = vbTextCompare
    Set graph = CreateObject("Scripting.Dictionary")
    graph.CompareMode = vbTextCompare
    For Each device In devices
        Set deviceById(device("DeviceId")) = device
    Next device

    For Each device In devices
        predecessorId = device("PredecessorDeviceId")
        If StrComp(device("RelationType"), "有先后顺序", vbTextCompare) = 0 Then
            graph(device("DeviceId")) = predecessorId
            If Len(predecessorId) = 0 Or Not deviceById.Exists(predecessorId) Then
                AddError errors, "[基础设置!tblDevices 行" & CStr(device("SourceRow")) & _
                    "，设备 " & device("DeviceId") & "] 前置设备不存在。"
            End If
        Else
            graph(device("DeviceId")) = ""
        End If
    Next device

    cycleDeviceId = FindStringGraphCycleNode(graph)
    If Len(cycleDeviceId) > 0 Then
        Set device = deviceById(cycleDeviceId)
        AddError errors, "[基础设置!tblDevices 行" & CStr(device("SourceRow")) & _
            "，设备 " & device("DeviceId") & "] 设备关系循环依赖。"
    End If
End Sub

Private Function FindStepCycleNode(ByVal stepByKey As Object) As String
    Dim state As Object
    Dim key As Variant
    Dim cycleKey As String

    Set state = CreateObject("Scripting.Dictionary")
    state.CompareMode = vbTextCompare
    For Each key In stepByKey.Keys
        cycleKey = VisitStepNode(CStr(key), stepByKey, state)
        If Len(cycleKey) > 0 Then
            FindStepCycleNode = cycleKey
            Exit Function
        End If
    Next key
End Function

Private Function VisitStepNode(ByVal key As String, ByVal stepByKey As Object, _
        ByVal state As Object) As String
    Dim predecessorKey As String
    If state.Exists(key) Then
        If state(key) = 1 Then VisitStepNode = key
        Exit Function
    End If

    state(key) = 1
    predecessorKey = ResolvePredecessorKey(stepByKey(key))
    If Len(predecessorKey) > 0 And stepByKey.Exists(predecessorKey) Then
        VisitStepNode = VisitStepNode(predecessorKey, stepByKey, state)
        If Len(VisitStepNode) > 0 Then Exit Function
    End If
    state(key) = 2
End Function

Private Function FindStringGraphCycleNode(ByVal graph As Object) As String
    Dim state As Object
    Dim key As Variant
    Dim cycleKey As String

    Set state = CreateObject("Scripting.Dictionary")
    state.CompareMode = vbTextCompare
    For Each key In graph.Keys
        cycleKey = VisitStringNode(CStr(key), graph, state)
        If Len(cycleKey) > 0 Then
            FindStringGraphCycleNode = cycleKey
            Exit Function
        End If
    Next key
End Function

Private Function VisitStringNode(ByVal key As String, ByVal graph As Object, _
        ByVal state As Object) As String
    Dim nextKey As String
    If state.Exists(key) Then
        If state(key) = 1 Then VisitStringNode = key
        Exit Function
    End If

    state(key) = 1
    nextKey = CStr(graph(key))
    If Len(nextKey) > 0 And graph.Exists(nextKey) Then
        VisitStringNode = VisitStringNode(nextKey, graph, state)
        If Len(VisitStringNode) > 0 Then Exit Function
    End If
    state(key) = 2
End Function

Private Sub ValidateContinuousStepNumbers(ByRef errors As String, ByVal steps As Collection)
    Dim numbersByDevice As Object
    Dim firstStepByDevice As Object
    Dim numbers As Object
    Dim stepItem As Variant
    Dim deviceId As Variant
    Dim parsedStepNo As Long
    Dim expectedNo As Long
    Dim isContinuous As Boolean

    Set numbersByDevice = CreateObject("Scripting.Dictionary")
    numbersByDevice.CompareMode = vbTextCompare
    Set firstStepByDevice = CreateObject("Scripting.Dictionary")
    firstStepByDevice.CompareMode = vbTextCompare

    For Each stepItem In steps
        If TryParsePositiveLong(stepItem("StepNoText"), parsedStepNo) Then
            If Not numbersByDevice.Exists(stepItem("DeviceId")) Then
                Set numbers = CreateObject("Scripting.Dictionary")
                Set numbersByDevice(stepItem("DeviceId")) = numbers
                Set firstStepByDevice(stepItem("DeviceId")) = stepItem
            End If
            Set numbers = numbersByDevice(stepItem("DeviceId"))
            numbers(CStr(parsedStepNo)) = True
        End If
    Next stepItem

    For Each deviceId In numbersByDevice.Keys
        Set numbers = numbersByDevice(deviceId)
        isContinuous = True
        For expectedNo = 1 To numbers.Count
            If Not numbers.Exists(CStr(expectedNo)) Then
                isContinuous = False
                Exit For
            End If
        Next expectedNo
        If Not isContinuous Then
            Set stepItem = firstStepByDevice(deviceId)
            AddError errors, "[作业步骤 行" & CStr(stepItem("SourceRow")) & _
                "，设备 " & CStr(deviceId) & "] 步骤号必须从1连续且无缺失。"
        End If
    Next deviceId
End Sub

Private Function EarliestFinish(ByVal key As String, ByVal stepByKey As Object, _
        ByVal cache As Object, ByVal visiting As Object) As Double
    Dim stepItem As Object
    Dim predecessorKey As String
    Dim startSec As Double

    If cache.Exists(key) Then
        EarliestFinish = cache(key)
        Exit Function
    End If
    If visiting.Exists(key) Then Exit Function
    visiting(key) = True

    Set stepItem = stepByKey(key)
    predecessorKey = ResolvePredecessorKey(stepItem)
    If Len(predecessorKey) > 0 And stepByKey.Exists(predecessorKey) Then
        startSec = EarliestFinish(predecessorKey, stepByKey, cache, visiting)
    End If
    If stepItem("HasLockedStart") Then
        If stepItem("LockedStartSec") > startSec Then startSec = stepItem("LockedStartSec")
    End If
    If stepItem("DurationSec") > 0# Then startSec = startSec + stepItem("DurationSec")
    cache(key) = startSec
    visiting.Remove key
    EarliestFinish = startSec
End Function

Private Function AnyQualifiedPerson(ByVal people As Collection, _
        ByVal stepItem As Object) As Boolean
    Dim person As Variant
    For Each person In people
        If PersonMatchesStep(person, stepItem) Then
            AnyQualifiedPerson = True
            Exit Function
        End If
    Next person
End Function

Private Function PersonIsQualified(ByVal people As Collection, ByVal personId As String, _
        ByVal stepItem As Object) As Boolean
    Dim person As Variant
    For Each person In people
        If StrComp(person("PersonId"), personId, vbTextCompare) = 0 Then
            PersonIsQualified = PersonMatchesStep(person, stepItem)
            Exit Function
        End If
    Next person
End Function

Private Function PersonMatchesStep(ByVal person As Object, ByVal stepItem As Object) As Boolean
    If Not TokenListContains(person("AllowedDevices"), stepItem("DeviceId"), True) Then Exit Function
    If Len(Trim$(stepItem("RequiredSkill"))) > 0 Then
        If Not TokenListContains(person("Skill"), stepItem("RequiredSkill"), False) Then Exit Function
    End If
    PersonMatchesStep = True
End Function

Private Function TokenListContains(ByVal listText As String, ByVal requiredText As String, _
        ByVal emptyMeansAll As Boolean) As Boolean
    Dim normalized As String
    Dim tokens As Variant
    Dim token As Variant

    If Len(Trim$(requiredText)) = 0 Then
        TokenListContains = emptyMeansAll
        Exit Function
    End If
    If Len(Trim$(listText)) = 0 Then
        TokenListContains = emptyMeansAll
        Exit Function
    End If

    normalized = Replace(listText, "，", ",")
    normalized = Replace(normalized, ";", ",")
    normalized = Replace(normalized, "；", ",")
    tokens = Split(normalized, ",")
    For Each token In tokens
        If StrComp(Trim$(CStr(token)), requiredText, vbTextCompare) = 0 Or _
                Trim$(CStr(token)) = "*" Or Trim$(CStr(token)) = "全部" Then
            TokenListContains = True
            Exit Function
        End If
    Next token
End Function

Private Function ResolvePredecessorKey(ByVal stepItem As Object) As String
    Dim text As String
    Dim parts As Variant

    text = Trim$(stepItem("PredecessorText"))
    If Len(text) = 0 Then Exit Function
    If IsNumeric(text) Then
        ResolvePredecessorKey = StepKey(stepItem("DeviceId"), text)
        Exit Function
    End If

    text = Replace(text, "：", ":")
    text = Replace(text, "-", ":")
    parts = Split(text, ":")
    If UBound(parts) = 1 Then
        ResolvePredecessorKey = StepKey(Trim$(CStr(parts(0))), Trim$(CStr(parts(1))))
    Else
        ResolvePredecessorKey = StepKey(stepItem("DeviceId"), text)
    End If
End Function

Private Function StepKey(ByVal deviceId As String, ByVal stepNumber As String) As String
    StepKey = Trim$(deviceId) & "|" & Trim$(stepNumber)
End Function

Private Function RequiresPerson(ByVal stepType As String) As Boolean
    RequiresPerson = (stepType = STEP_MANUAL Or stepType = STEP_JOINT)
End Function

Private Function HasTextValue(ByVal value As Variant) As Boolean
    If IsError(value) Or IsEmpty(value) Then Exit Function
    HasTextValue = (Len(Trim$(CStr(value))) > 0)
End Function

Private Function ParameterValue(ByVal parameters As ListObject, _
        ByVal columnName As String) As Variant
    ParameterValue = parameters.ListColumns(columnName).DataBodyRange.Cells(1, 1).Value2
End Function

Private Function TryParsePositiveLong(ByVal value As Variant, _
        ByRef parsedValue As Long) As Boolean
    Dim numericValue As Double

    On Error GoTo InvalidValue
    parsedValue = 0
    If Not HasTextValue(value) Or Not IsNumeric(value) Then Exit Function
    numericValue = CDbl(value)
    If numericValue <= 0# Or numericValue > 2147483647# Then Exit Function
    If numericValue <> Fix(numericValue) Then Exit Function
    parsedValue = CLng(numericValue)
    TryParsePositiveLong = True
InvalidValue:
End Function

Private Function TryParseDouble(ByVal value As Variant, _
        ByRef parsedValue As Double) As Boolean
    On Error GoTo InvalidValue
    parsedValue = 0#
    If Not HasTextValue(value) Or Not IsNumeric(value) Then Exit Function
    parsedValue = CDbl(value)
    TryParseDouble = True
InvalidValue:
End Function

Private Sub AddStepError(ByRef errors As String, ByVal stepItem As Object, _
        ByVal message As String)
    AddError errors, "[作业步骤 行" & CStr(stepItem("SourceRow")) & "，设备 " & _
        stepItem("DeviceId") & "，步骤 " & stepItem("StepNoText") & "] " & message
End Sub

Private Sub AddError(ByRef errors As String, ByVal message As String)
    If Len(errors) > 0 Then errors = errors & vbLf
    errors = errors & message
End Sub
