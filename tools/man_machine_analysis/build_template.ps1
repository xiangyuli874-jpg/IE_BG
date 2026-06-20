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
$books = $null
$book = $null
$vbProject = $null
$vbComponents = $null
$expectedResult = 'Test_ChineseSourceRoundTrip PASS; Test_DomainConstants PASS; Test_DomainTypes PASS; Test_BaselineFixtureExpansion PASS'

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false

    $books = $excel.Workbooks
    $book = $books.Add()
    $book.SaveAs($tempWorkbook, 52)

    try {
        $vbProject = $book.VBProject
        $vbComponents = $vbProject.VBComponents
    }
    catch {
        throw 'Excel blocked VBA project access. Enable "Trust access to the VBA project object model".'
    }

    Get-ChildItem -LiteralPath $vbaRoot -Filter '*.bas' |
        Sort-Object Name |
        ForEach-Object {
            $sourceText = [IO.File]::ReadAllText(
                $_.FullName,
                (New-Object Text.UTF8Encoding($false, $true))
            )
            $importPath = Join-Path $tempRoot ("import_" + $_.Name)
            [IO.File]::WriteAllText($importPath, $sourceText, [Text.Encoding]::Default)

            $importedComponent = $null
            try {
                $importedComponent = $vbComponents.Import($importPath)
            }
            finally {
                if ($null -ne $importedComponent -and
                        [Runtime.InteropServices.Marshal]::IsComObject($importedComponent)) {
                    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($importedComponent)
                }
            }
        }

    $book.Save()
    $macroName = "'$($book.Name)'!RunAllSelfTests"
    $fixturePath = Join-Path $toolRoot 'fixtures\baseline_1p3m.csv'
    $result = $excel.Run($macroName, $fixturePath)
    Write-Output $result
    if ($null -eq $result -or [string]$result -cne $expectedResult) {
        throw "Unexpected self-test result. Expected [$expectedResult], actual [$result]"
    }
}
finally {
    if ($null -ne $book) {
        try {
            $book.Close($false)
        }
        catch {
            Write-Warning "Failed to close temporary workbook: $($_.Exception.Message)"
        }
    }
    if ($null -ne $excel) {
        try {
            $excel.Quit()
        }
        catch {
            Write-Warning "Failed to quit temporary Excel instance: $($_.Exception.Message)"
        }
    }

    foreach ($comObject in @($vbComponents, $vbProject, $book, $books, $excel)) {
        if ($null -ne $comObject -and
                [Runtime.InteropServices.Marshal]::IsComObject($comObject)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($comObject)
        }
    }

    $vbComponents = $null
    $vbProject = $null
    $book = $null
    $books = $null
    $excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
