Attribute VB_Name = "modDomain"
Option Explicit

Public Const MAX_PEOPLE As Long = 5
Public Const MAX_DEVICES As Long = 10
Public Const MAX_STEPS_PER_DEVICE As Long = 20

Public Const STEP_MANUAL As String = "MANUAL"
Public Const STEP_AUTO As String = "AUTO"
Public Const STEP_JOINT As String = "JOINT"
Public Const STEP_WAIT As String = "WAIT"

Public Type TaskDef
    TaskId As String
    DeviceId As String
    CycleNo As Long
    StepNo As Long
    StepName As String
    StepType As String
    DurationSec As Double
    RequiredSkill As String
    PredecessorId As String
    LockedPersonId As String
    LockedStartSec As Double
    HasLockedStart As Boolean
End Type

Public Type ScheduledTask
    Definition As TaskDef
    PersonId As String
    StartSec As Double
    EndSec As Double
    MoveSec As Double
    WaitSec As Double
    IsLocked As Boolean
End Type
