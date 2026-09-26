param(
    [Parameter(Mandatory=$true)][string]$LocalStatusPath,
    [Parameter(Mandatory=$true)][string]$ManagementExportPath,
    [Parameter(Mandatory=$true)][string]$OutputPath
)

$ErrorActionPreference = 'Stop'

function Normalize-BoolState {
    param([object]$Value)
    if ($null -eq $Value) { return $null }
    $text = ([string]$Value).Trim().ToLowerInvariant()
    if ($text -match '^(enabled|on|true|yes|presentandready)') { return $true }
    if ($text -match '^(disabled|off|false|no|notpresent)') { return $false }
    return $null
}

function Normalize-BitLocker {
    param([object]$Value)
    if ($null -eq $Value) { return $null }
    $text = ([string]$Value).Trim().ToLowerInvariant()
    if ($text -match '^(on|encrypted|true)') { return 'on' }
    if ($text -match '^(off|notencrypted|not encrypted|false)') { return 'off' }
    if ($text -match '^on \|') { return 'on' }
    if ($text -match '^off \|') { return 'off' }
    return 'unknown'
}

$local = Get-Content -LiteralPath $LocalStatusPath -Raw -Encoding UTF8 | ConvertFrom-Json
$mgmtRaw = Get-Content -LiteralPath $ManagementExportPath -Raw -Encoding UTF8 | ConvertFrom-Json
$records = @($mgmtRaw)

$deviceName = [string]$local.ComputerName
$match = $records | Where-Object {
    ([string]$_.deviceName -ieq $deviceName) -or
    ([string]$_.computerName -ieq $deviceName)
} | Select-Object -First 1

$result = [ordered]@{
    device = $deviceName
    matched = $null -ne $match
    differences = @()
    matches = @()
    managementOnly = @()
    missingEvidence = @()
    boundary = 'Read-only comparison of local endpoint evidence and public-safe management export. No remediation.'
}

if ($null -eq $match) {
    $result.missingEvidence += "No management record found for device '$deviceName'."
}
else {
    $localSecureBoot = Normalize-BoolState $local.SecureBoot
    $mgmtSecureBoot  = Normalize-BoolState $match.secureBoot
    if ($null -ne $localSecureBoot -and $null -ne $mgmtSecureBoot) {
        if ($localSecureBoot -eq $mgmtSecureBoot) {
            $result.matches += [ordered]@{ field='SecureBoot'; local=$local.SecureBoot; management=$match.secureBoot }
        } else {
            $result.differences += [ordered]@{
                field='SecureBoot'; local=$local.SecureBoot; management=$match.secureBoot
                safeNextCheck='Validate firmware state, latest management sync, and compliance-policy interpretation.'
            }
        }
    } else {
        $result.missingEvidence += 'Secure Boot state could not be normalized from both sources.'
    }

    $localBitLocker = Normalize-BitLocker $local.BitLocker
    $mgmtBitLocker  = Normalize-BitLocker $match.encryptionState
    if ($localBitLocker -ne 'unknown' -and $mgmtBitLocker -ne 'unknown' -and $null -ne $mgmtBitLocker) {
        if ($localBitLocker -eq $mgmtBitLocker) {
            $result.matches += [ordered]@{ field='BitLocker'; local=$local.BitLocker; management=$match.encryptionState }
        } else {
            $result.differences += [ordered]@{
                field='BitLocker'; local=$local.BitLocker; management=$match.encryptionState
                safeNextCheck='Confirm protector state, encryption progress, policy assignment, and last management sync.'
            }
        }
    } else {
        $result.missingEvidence += 'BitLocker/encryption state could not be normalized from both sources.'
    }

    if ($null -ne $match.complianceState) {
        $result.managementOnly += [ordered]@{
            field='ComplianceState'; value=$match.complianceState
            note='Management-plane evidence only; does not independently identify root cause.'
        }
    }
    if ($null -ne $match.lastSync) {
        $result.managementOnly += [ordered]@{
            field='LastSync'; value=$match.lastSync
            note='Use to judge evidence freshness before interpreting discrepancies.'
        }
    }
}

$parent = Split-Path -Parent $OutputPath
if ($parent -and -not (Test-Path -LiteralPath $parent)) {
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
}
$result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
$result | ConvertTo-Json -Depth 8
