[CmdletBinding()]
param(
    [switch]$TestOnly
)

$ErrorActionPreference = 'Stop'

if (-not $TestOnly) {
    throw 'Task 1 test host only supports -TestOnly.'
}

$toolRoot = $PSScriptRoot
$repoRoot = Split-Path (Split-Path $toolRoot -Parent) -Parent
$vbaRoot = Join-Path $toolRoot 'vba'
$tempRoot = Join-Path $repoRoot '.codex_tmp\man_machine_analysis'
New-Item -ItemType Directory -Force -Path $tempRoot | Out-Null

$stamp = Get-Date -Format 'yyyyMMdd_HHmmss_fff'
$tempWorkbook = Join-Path $tempRoot "domain_self_test_$stamp.xlsm"
$excel = $null
$book = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false

    $book = $excel.Workbooks.Add()
    $book.SaveAs($tempWorkbook, 52)

    try {
        $vbProject = $book.VBProject
    }
    catch {
        throw 'Excel blocked VBA project access. Enable "Trust access to the VBA project object model".'
    }

    Get-ChildItem -LiteralPath $vbaRoot -Filter '*.bas' |
        Sort-Object Name |
        ForEach-Object {
            [void]$vbProject.VBComponents.Import($_.FullName)
        }

    $book.Save()
    $macroName = "'$($book.Name)'!RunAllSelfTests"
    $fixturePath = Join-Path $toolRoot 'fixtures\baseline_1p3m.csv'
    $result = $excel.Run($macroName, $fixturePath)
    Write-Output $result
    if ($result -match 'FAIL') {
        throw $result
    }
}
finally {
    if ($null -ne $book) {
        $book.Close($false)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)
    }
    if ($null -ne $excel) {
        $excel.Quit()
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
