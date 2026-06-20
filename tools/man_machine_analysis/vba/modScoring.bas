Attribute VB_Name = "modScoring"
Option Explicit

Public Const TARGET_THROUGHPUT As String = "最高产能"
Public Const TARGET_BALANCE As String = "人员均衡"
Public Const TARGET_DEVICE_WAIT As String = "设备少等待"
Public Const TARGET_COMPOSITE As String = "综合优化"

Public Function CalculateMetrics(ByVal schedule As Collection) As Object
    Dim metrics As Object, personWork As Object, personMove As Object
    Dim deviceWork As Object, cycleAnchor As Object, deviceSet As Object
    Dim task As Object, key As Variant
    Dim analysisSec As Double, totalPersonWork As Double, totalMove As Double
    Dim totalDeviceWait As Double, maxPersonLoad As Double, sumLoad As Double
    Dim sumLoadSquared As Double, personCount As Long, lastCycle As Long
    Dim cycleTime As Double, outputCount As Long, loadValue As Double
    Dim bottleneckPerson As String, bottleneckDevice As String, maxDeviceWork As Double

    Set metrics = CreateObject("Scripting.Dictionary")
    metrics.CompareMode = vbTextCompare
    Set personWork = CreateObject("Scripting.Dictionary")
    Set personMove = CreateObject("Scripting.Dictionary")
    Set deviceWork = CreateObject("Scripting.Dictionary")
    Set cycleAnchor = CreateObject("Scripting.Dictionary")
    Set deviceSet = CreateObject("Scripting.Dictionary")
    personWork.CompareMode = vbTextCompare
    personMove.CompareMode = vbTextCompare
    deviceWork.CompareMode = vbTextCompare
    deviceSet.CompareMode = vbTextCompare

    For Each task In schedule
        If CDbl(task("EndSec")) > analysisSec Then analysisSec = CDbl(task("EndSec"))
        deviceSet(task("DeviceId")) = True
        deviceWork(task("DeviceId")) = DictionaryMetric(deviceWork, task("DeviceId")) + _
            CDbl(task("DurationSec"))
        totalDeviceWait = totalDeviceWait + CDbl(task("WaitSec"))
        If Len(CStr(task("PersonId"))) > 0 Then
            personWork(task("PersonId")) = DictionaryMetric(personWork, task("PersonId")) + _
                CDbl(task("DurationSec"))
            personMove(task("PersonId")) = DictionaryMetric(personMove, task("PersonId")) + _
                CDbl(task("MoveSec"))
            totalPersonWork = totalPersonWork + CDbl(task("DurationSec"))
            totalMove = totalMove + CDbl(task("MoveSec"))
        End If
        If CLng(task("StepNo")) = 1 Then
            key = CStr(task("CycleNo"))
            If Not cycleAnchor.Exists(key) Or CDbl(task("StartSec")) < CDbl(cycleAnchor(key)) Then
                cycleAnchor(key) = CDbl(task("StartSec"))
            End If
            If CLng(task("CycleNo")) > lastCycle Then lastCycle = CLng(task("CycleNo"))
        End If
    Next task

    If lastCycle >= 2 And cycleAnchor.Exists(CStr(lastCycle)) And _
            cycleAnchor.Exists(CStr(lastCycle - 1)) Then
        cycleTime = CDbl(cycleAnchor(CStr(lastCycle))) - _
            CDbl(cycleAnchor(CStr(lastCycle - 1)))
    Else
        cycleTime = analysisSec
    End If
    outputCount = deviceSet.Count
    If outputCount < 1 Then outputCount = 1

    For Each key In personWork.Keys
        personCount = personCount + 1
        If analysisSec > 0 Then loadValue = CDbl(personWork(key)) / analysisSec Else loadValue = 0#
        sumLoad = sumLoad + loadValue
        sumLoadSquared = sumLoadSquared + loadValue * loadValue
        If loadValue > maxPersonLoad Then
            maxPersonLoad = loadValue
            bottleneckPerson = CStr(key)
        End If
    Next key
    For Each key In deviceWork.Keys
        If CDbl(deviceWork(key)) > maxDeviceWork Then
            maxDeviceWork = CDbl(deviceWork(key))
            bottleneckDevice = CStr(key)
        End If
    Next key

    metrics("CycleTimeSec") = cycleTime
    metrics("OutputCount") = outputCount
    If outputCount > 0 Then metrics("TaktSec") = cycleTime / outputCount _
        Else metrics("TaktSec") = 0#
    If cycleTime > 0 Then metrics("HourlyCapacity") = 3600# * outputCount / cycleTime _
        Else metrics("HourlyCapacity") = 0#
    metrics("AnalysisSec") = analysisSec
    metrics("TotalPersonWorkSec") = totalPersonWork
    metrics("TotalMoveSec") = totalMove
    metrics("TotalDeviceWaitSec") = totalDeviceWait
    If personCount > 0 Then metrics("AveragePersonLoad") = sumLoad / personCount _
        Else metrics("AveragePersonLoad") = 0#
    metrics("MaxPersonLoad") = maxPersonLoad
    If personCount > 0 Then
        metrics("PersonLoadStdDev") = _
            Sqr(MaxZero(sumLoadSquared / personCount - (sumLoad / personCount) ^ 2))
    Else
        metrics("PersonLoadStdDev") = 0#
    End If
    metrics("BottleneckPerson") = bottleneckPerson
    metrics("BottleneckDevice") = bottleneckDevice
    Set metrics("PersonWork") = personWork
    Set metrics("PersonMove") = personMove
    Set metrics("DeviceWork") = deviceWork
    Set CalculateMetrics = metrics
End Function

Public Function ScoreSchedule(ByVal schedule As Collection, ByVal target As String, _
        Optional ByVal baseline As Object = Nothing) As Double
    Dim metrics As Object, cyclePenalty As Double, overloadPenalty As Double
    Dim wCycle As Double, wBalance As Double, wWait As Double, wMove As Double
    Set metrics = CalculateMetrics(schedule)
    cyclePenalty = CDbl(metrics("CycleTimeSec"))
    If CDbl(metrics("MaxPersonLoad")) > 0.9 Then
        overloadPenalty = 1000# * (CDbl(metrics("MaxPersonLoad")) - 0.9)
    End If

    Select Case target
        Case TARGET_THROUGHPUT
            ScoreSchedule = 10000# * cyclePenalty + 0.1 * CDbl(metrics("TotalDeviceWaitSec")) + _
                0.05 * CDbl(metrics("TotalMoveSec"))
        Case TARGET_BALANCE
            ScoreSchedule = 1000# * CDbl(metrics("PersonLoadStdDev")) + _
                cyclePenalty + overloadPenalty
        Case TARGET_DEVICE_WAIT
            ScoreSchedule = CDbl(metrics("TotalDeviceWaitSec")) + cyclePenalty
        Case Else
            ReadWeights wCycle, wBalance, wWait, wMove
            ScoreSchedule = wCycle * cyclePenalty + _
                wBalance * 1000# * CDbl(metrics("PersonLoadStdDev")) + _
                wWait * CDbl(metrics("TotalDeviceWaitSec")) + _
                wMove * CDbl(metrics("TotalMoveSec"))
    End Select
End Function

Private Sub ReadWeights(ByRef wCycle As Double, ByRef wBalance As Double, _
        ByRef wWait As Double, ByRef wMove As Double)
    On Error GoTo Defaults
    wCycle = CDbl(ThisWorkbook.Names("nmWeightCycle").RefersToRange.Value2)
    wBalance = CDbl(ThisWorkbook.Names("nmWeightBalance").RefersToRange.Value2)
    wWait = CDbl(ThisWorkbook.Names("nmWeightWait").RefersToRange.Value2)
    wMove = CDbl(ThisWorkbook.Names("nmWeightMove").RefersToRange.Value2)
    Exit Sub
Defaults:
    wCycle = 0.4: wBalance = 0.2: wWait = 0.3: wMove = 0.1
End Sub

Private Function DictionaryMetric(ByVal dictionary As Object, ByVal key As String) As Double
    If dictionary.Exists(key) Then DictionaryMetric = CDbl(dictionary(key))
End Function

Private Function MaxZero(ByVal value As Double) As Double
    If value > 0# Then MaxZero = value
End Function
