Attribute VB_Name = "modSelfTest"
Option Explicit

Public Function RunAllSelfTests() As String
    On Error GoTo TestFailed
    Test_DomainConstants
    RunAllSelfTests = "Test_DomainConstants PASS"
    Exit Function

TestFailed:
    RunAllSelfTests = "Test_DomainConstants FAIL: " & Err.Description
End Function

Public Sub Test_DomainConstants()
    AssertEqual ReadPublicConstant("STEP_MANUAL"), "MANUAL", "manual step constant"
    AssertEqual ReadPublicConstant("STEP_AUTO"), "AUTO", "automatic step constant"
    AssertEqual ReadPublicConstant("MAX_PEOPLE"), 5, "maximum people"
    AssertEqual ReadPublicConstant("MAX_DEVICES"), 10, "maximum devices"
    AssertEqual ReadPublicConstant("MAX_STEPS_PER_DEVICE"), 20, "maximum steps per device"
End Sub

Private Sub AssertEqual(ByVal actual As Variant, ByVal expected As Variant, ByVal message As String)
    If actual <> expected Then
        Err.Raise vbObjectError + 1001, "AssertEqual", _
            message & ": expected [" & CStr(expected) & "], actual [" & CStr(actual) & "]"
    End If
End Sub

Private Function ReadPublicConstant(ByVal constantName As String) As Variant
    Dim codeModule As Object
    Dim lineNumber As Long
    Dim sourceLine As String
    Dim declarationPrefix As String
    Dim rawValue As String

    On Error GoTo ConstantMissing
    Set codeModule = ThisWorkbook.VBProject.VBComponents("modDomain").CodeModule
    declarationPrefix = "Public Const " & constantName & " "

    For lineNumber = 1 To codeModule.CountOfLines
        sourceLine = Trim$(codeModule.Lines(lineNumber, 1))
        If StrComp(Left$(sourceLine, Len(declarationPrefix)), declarationPrefix, vbTextCompare) = 0 Then
            rawValue = Trim$(Mid$(sourceLine, InStr(sourceLine, "=") + 1))
            If Left$(rawValue, 1) = """" And Right$(rawValue, 1) = """" Then
                ReadPublicConstant = Mid$(rawValue, 2, Len(rawValue) - 2)
            Else
                ReadPublicConstant = CDbl(rawValue)
            End If
            Exit Function
        End If
    Next lineNumber

ConstantMissing:
    Err.Raise vbObjectError + 1002, "ReadPublicConstant", _
        "domain constant not defined: " & constantName
End Function
