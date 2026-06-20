Attribute VB_Name = "modWorkbookIO"
Option Explicit

Public Function ReadPeople() As Collection
    Dim result As New Collection
    Dim table As ListObject
    Dim rowIndex As Long
    Dim person As Object

    Set table = ThisWorkbook.Worksheets("基础设置").ListObjects("tblPeople")
    For rowIndex = 1 To table.ListRows.Count
        If IsEnabledRow(table, rowIndex, "启用") Then
            If Len(CellText(table, rowIndex, "人员编号")) > 0 Then
                Set person = CreateObject("Scripting.Dictionary")
                person.CompareMode = vbTextCompare
                person("PersonId") = CellText(table, rowIndex, "人员编号")
                person("PersonName") = CellText(table, rowIndex, "人员名称")
                person("Skill") = CellText(table, rowIndex, "技能")
                person("AllowedDevices") = CellText(table, rowIndex, "可操作设备")
                person("SourceRow") = table.DataBodyRange.Rows(rowIndex).Row
                result.Add person
            End If
        End If
    Next rowIndex
    Set ReadPeople = result
End Function

Public Function ReadDevices() As Collection
    Dim result As New Collection
    Dim table As ListObject
    Dim rowIndex As Long
    Dim device As Object

    Set table = ThisWorkbook.Worksheets("基础设置").ListObjects("tblDevices")
    For rowIndex = 1 To table.ListRows.Count
        If IsEnabledRow(table, rowIndex, "启用") Then
            If Len(CellText(table, rowIndex, "设备编号")) > 0 Then
                Set device = CreateObject("Scripting.Dictionary")
                device.CompareMode = vbTextCompare
                device("DeviceId") = CellText(table, rowIndex, "设备编号")
                device("DeviceName") = CellText(table, rowIndex, "设备名称")
                device("Product") = CellText(table, rowIndex, "产品")
                device("RelationType") = CellText(table, rowIndex, "关系类型")
                device("PredecessorDeviceId") = CellText(table, rowIndex, "前置设备")
                device("SourceRow") = table.DataBodyRange.Rows(rowIndex).Row
                result.Add device
            End If
        End If
    Next rowIndex
    Set ReadDevices = result
End Function

Public Function ReadMoveMatrix() As Object
    Dim result As Object
    Dim enabledDevices As Object
    Dim devices As Collection
    Dim device As Variant
    Dim table As ListObject
    Dim rowIndex As Long
    Dim fromId As String
    Dim toId As Variant
    Dim columnIndex As Long
    Dim key As String

    Set result = CreateObject("Scripting.Dictionary")
    result.CompareMode = vbTextCompare
    Set enabledDevices = CreateObject("Scripting.Dictionary")
    enabledDevices.CompareMode = vbTextCompare
    Set devices = ReadDevices()
    For Each device In devices
        enabledDevices(device("DeviceId")) = True
    Next device

    Set table = ThisWorkbook.Worksheets("基础设置").ListObjects("tblMoveTime")
    For rowIndex = 1 To table.ListRows.Count
        fromId = Trim$(CStr(table.DataBodyRange.Cells(rowIndex, 1).Value2))
        If enabledDevices.Exists(fromId) Then
            For Each toId In enabledDevices.Keys
                columnIndex = FindTableColumn(table, CStr(toId))
                key = fromId & "|" & CStr(toId)
                If columnIndex = 0 Then
                    result(key) = Empty
                Else
                    result(key) = table.DataBodyRange.Cells(rowIndex, columnIndex).Value2
                End If
            Next toId
        End If
    Next rowIndex
    Set ReadMoveMatrix = result
End Function

Public Function ReadSteps() As Collection
    Dim result As New Collection
    Dim table As ListObject
    Dim enabledDevices As Object
    Dim devices As Collection
    Dim device As Variant
    Dim rowIndex As Long
    Dim deviceId As String
    Dim stepItem As Object
    Dim rawDuration As Variant
    Dim rawLockedStart As Variant

    Set enabledDevices = CreateObject("Scripting.Dictionary")
    enabledDevices.CompareMode = vbTextCompare
    Set devices = ReadDevices()
    For Each device In devices
        enabledDevices(device("DeviceId")) = True
    Next device

    Set table = ThisWorkbook.Worksheets("作业步骤").ListObjects("tblSteps")
    For rowIndex = 1 To table.ListRows.Count
        deviceId = CellText(table, rowIndex, "设备")
        If Len(deviceId) > 0 And enabledDevices.Exists(deviceId) Then
            Set stepItem = CreateObject("Scripting.Dictionary")
            stepItem.CompareMode = vbTextCompare
            rawDuration = CellValue(table, rowIndex, "工时(s)")
            rawLockedStart = CellValue(table, rowIndex, "锁定开始(s)")

            stepItem("DeviceId") = deviceId
            stepItem("StepNoText") = CellText(table, rowIndex, "步骤号")
            stepItem("StepNo") = NumericLong(CellValue(table, rowIndex, "步骤号"))
            stepItem("StepName") = CellText(table, rowIndex, "名称")
            stepItem("DisplayType") = CellText(table, rowIndex, "类型")
            stepItem("StepType") = MapStepType(stepItem("DisplayType"))
            stepItem("DurationRaw") = rawDuration
            stepItem("DurationSec") = NumericDouble(rawDuration)
            stepItem("PredecessorText") = CellText(table, rowIndex, "前置步骤")
            stepItem("RequiredSkill") = CellText(table, rowIndex, "所需技能")
            stepItem("LockedPersonId") = CellText(table, rowIndex, "锁定人员")
            stepItem("LockedStartRaw") = rawLockedStart
            stepItem("HasLockedStart") = HasValue(rawLockedStart)
            stepItem("LockedStartSec") = NumericDouble(rawLockedStart)
            stepItem("AllowWait") = CellText(table, rowIndex, "允许等待")
            stepItem("Notes") = CellText(table, rowIndex, "备注")
            stepItem("SourceRow") = table.DataBodyRange.Rows(rowIndex).Row
            result.Add stepItem
        End If
    Next rowIndex
    Set ReadSteps = result
End Function

Public Function MapStepType(ByVal displayType As String) As String
    Select Case Trim$(displayType)
        Case "人工": MapStepType = STEP_MANUAL
        Case "自动运行": MapStepType = STEP_AUTO
        Case "人机协同": MapStepType = STEP_JOINT
        Case "等待": MapStepType = STEP_WAIT
        Case Else: MapStepType = ""
    End Select
End Function

Private Function IsEnabledRow(ByVal table As ListObject, ByVal rowIndex As Long, _
        ByVal enabledColumn As String) As Boolean
    IsEnabledRow = (StrComp(CellText(table, rowIndex, enabledColumn), "否", _
        vbTextCompare) <> 0)
End Function

Private Function CellValue(ByVal table As ListObject, ByVal rowIndex As Long, _
        ByVal columnName As String) As Variant
    CellValue = table.ListColumns(columnName).DataBodyRange.Cells(rowIndex, 1).Value2
End Function

Private Function CellText(ByVal table As ListObject, ByVal rowIndex As Long, _
        ByVal columnName As String) As String
    Dim value As Variant
    value = CellValue(table, rowIndex, columnName)
    If Not IsError(value) And Not IsEmpty(value) Then CellText = Trim$(CStr(value))
End Function

Private Function FindTableColumn(ByVal table As ListObject, ByVal columnName As String) As Long
    Dim columnIndex As Long
    For columnIndex = 1 To table.ListColumns.Count
        If StrComp(CStr(table.HeaderRowRange.Cells(1, columnIndex).Value2), _
                columnName, vbTextCompare) = 0 Then
            FindTableColumn = columnIndex
            Exit Function
        End If
    Next columnIndex
End Function

Private Function NumericLong(ByVal value As Variant) As Long
    If HasValue(value) And IsNumeric(value) Then NumericLong = CLng(value)
End Function

Private Function NumericDouble(ByVal value As Variant) As Double
    If HasValue(value) And IsNumeric(value) Then NumericDouble = CDbl(value)
End Function

Private Function HasValue(ByVal value As Variant) As Boolean
    If IsError(value) Or IsEmpty(value) Then Exit Function
    HasValue = (Len(Trim$(CStr(value))) > 0)
End Function
