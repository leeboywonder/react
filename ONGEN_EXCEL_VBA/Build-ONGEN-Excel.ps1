[CmdletBinding()]
param(
    [string]$OutputDirectory = (Join-Path $PSScriptRoot "output")
)

$ErrorActionPreference = "Stop"
$sourceDirectory = Join-Path $PSScriptRoot "source"
$masterSource = Join-Path $sourceDirectory "ONGEN PD_DEFINITION_MASTER_WORKFLOW_V5.1.6.5.xlsm"
$fieldSource = Join-Path $sourceDirectory "ONGEN Field Devices Specification -v2.1.7.1.xlsm"
$masterOutput = Join-Path $OutputDirectory "ONGEN Excel Engineering Suite V0.2.xlsm"
$fieldOutput = Join-Path $OutputDirectory "ONGEN Field Devices Specification -v2.1.7.1.xlsm"
$logPath = Join-Path $OutputDirectory "ONGEN-build.log"

function Release-ComObject([object]$Object) {
    if ($null -ne $Object) {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($Object)
    }
}

function Replace-StandardModule([object]$Workbook, [string]$ModulePath) {
    $moduleName = [IO.Path]::GetFileNameWithoutExtension($ModulePath)
    $project = $Workbook.VBProject
    $existing = $null
    try { $existing = $project.VBComponents.Item($moduleName) } catch { $existing = $null }
    if ($null -ne $existing) {
        if ($existing.Type -ne 1) {
            throw "VBA component '$moduleName' is not a standard module and was not replaced."
        }
        $project.VBComponents.Remove($existing)
        Release-ComObject $existing
    }
    $imported = $project.VBComponents.Import($ModulePath)
    if ($imported.Name -ne $moduleName) { $imported.Name = $moduleName }
    Write-Host "Imported $moduleName into $($Workbook.Name)"
    Release-ComObject $imported
    Release-ComObject $project
}

if (-not $IsWindows -and $PSVersionTable.PSEdition -eq "Core") {
    throw "This builder requires Microsoft Excel on Windows."
}
if (-not (Test-Path $masterSource) -or -not (Test-Path $fieldSource)) {
    throw "Required source workbooks are missing from '$sourceDirectory'."
}

New-Item -ItemType Directory -Force -Path $OutputDirectory | Out-Null
Start-Transcript -Path $logPath -Force | Out-Null
$excel = $null
$master = $null
$field = $null
$saved = $false

try {
    Copy-Item -LiteralPath $masterSource -Destination $masterOutput -Force
    Copy-Item -LiteralPath $fieldSource -Destination $fieldOutput -Force
    Get-ChildItem -LiteralPath $sourceDirectory -File | Where-Object {
        $_.FullName -ne $masterSource -and $_.FullName -ne $fieldSource
    } | Copy-Item -Destination $OutputDirectory -Force

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.EnableEvents = $false
    $excel.AutomationSecurity = 3

    $master = $excel.Workbooks.Open($masterOutput, 0, $false)
    foreach ($module in @(
        "mod09_GeneratePD2.bas",
        "mod14_ImportPD6Power.bas",
        "mod20_DeviceWorkflow.bas",
        "modONGEN_SuiteControl.bas"
    )) {
        Replace-StandardModule $master (Join-Path $PSScriptRoot "vba\master\$module")
    }

    $excel.AutomationSecurity = 1
    $excel.Run("'" + $master.Name.Replace("'", "''") + "'!ONGEN_InstallControlCenter")
    $master.Save()
    $master.Close($true)
    Release-ComObject $master
    $master = $null

    $excel.AutomationSecurity = 3
    $field = $excel.Workbooks.Open($fieldOutput, 0, $false)
    Replace-StandardModule $field (Join-Path $PSScriptRoot "vba\field\EXPORT_TO_IOGEN.bas")
    $field.Save()
    $field.Close($true)
    Release-ComObject $field
    $field = $null

    $files = Get-ChildItem -LiteralPath $OutputDirectory -File | Where-Object { $_.Name -ne "ONGEN-build.log" }
    $manifest = @{
        generated_utc = [DateTime]::UtcNow.ToString("o")
        suite = "ONGEN Excel Engineering Suite V0.2"
        controlled_limitation = "Power Calculation V4.6 was not supplied and Phase 4 is not runtime validated."
        files = @($files | ForEach-Object {
            @{ name = $_.Name; bytes = $_.Length; sha256 = (Get-FileHash -Algorithm SHA256 -LiteralPath $_.FullName).Hash }
        })
    }
    $manifest | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath (Join-Path $OutputDirectory "ONGEN-manifest.json") -Encoding UTF8
    $saved = $true
    Write-Host "ONGEN Excel suite created in: $OutputDirectory" -ForegroundColor Green
}
catch {
    Write-Error $_
    if ($_.Exception.Message -match "VBProject|programmatic access|1004") {
        Write-Host "In Excel, enable: File > Options > Trust Center > Trust Center Settings > Macro Settings > Trust access to the VBA project object model." -ForegroundColor Yellow
    }
}
finally {
    if ($null -ne $field) { $field.Close($false); Release-ComObject $field }
    if ($null -ne $master) { $master.Close($false); Release-ComObject $master }
    if ($null -ne $excel) { $excel.Quit(); Release-ComObject $excel }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    Stop-Transcript | Out-Null
    if (-not $saved) {
        Remove-Item -LiteralPath $masterOutput -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $fieldOutput -Force -ErrorAction SilentlyContinue
        throw "Build failed. No generated Excel workbooks were retained. See '$logPath'."
    }
}
