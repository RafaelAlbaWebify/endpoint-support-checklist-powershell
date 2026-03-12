Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

<#
.SYNOPSIS
    Endpoint Support Checklist v1.0

.DESCRIPTION
    Local-first PowerShell WinForms utility to inspect device status,
    record technical interventions, and keep a local intervention history.

.NOTES
    Author : Rafael Alba
    Year   : 2026
    Version: v1.0

    Public portfolio/demo version.
    This version contains no customer-specific logic, hostnames,
    internal paths beyond local app storage, or sensitive operational data.

    Design goals:
    - Local-only storage
    - Works without network connectivity
    - Stable WinForms layout
    - Readable and maintainable code
#>

[System.Windows.Forms.Application]::EnableVisualStyles()

# -----------------------------------------------------------------------------
# Configuration
# -----------------------------------------------------------------------------
$ScriptVersion  = 'v1.0'
$AppTitle       = 'Endpoint Support Checklist'
$BaseFolder     = 'C:\ProgramData\EndpointSupportChecklist'
$StatusFile     = Join-Path $BaseFolder 'DeviceStatus.json'
$StatusTextFile = Join-Path $BaseFolder 'DeviceStatus.txt'
$LogFile        = Join-Path $BaseFolder 'MaintenanceLog.csv'
$BrandBlue      = [System.Drawing.Color]::FromArgb(0,120,215)

# -----------------------------------------------------------------------------
# Utility helpers
# -----------------------------------------------------------------------------
function Ensure-Folder {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -ItemType Directory -Force | Out-Null
    }
}

function Get-SafeValue {
    param($Value)
    if ($null -eq $Value) { return '' }
    return [string]$Value
}

function Show-Message {
    param(
        [string]$Text,
        [string]$Title = 'Information',
        [System.Windows.Forms.MessageBoxIcon]$Icon = [System.Windows.Forms.MessageBoxIcon]::Information
    )

    [System.Windows.Forms.MessageBox]::Show(
        $Text,
        $Title,
        [System.Windows.Forms.MessageBoxButtons]::OK,
        $Icon
    ) | Out-Null
}

# -----------------------------------------------------------------------------
# Device data collection
# -----------------------------------------------------------------------------
function Get-BiosDateSafe {
    try {
        $bios = Get-CimInstance Win32_BIOS -ErrorAction Stop
        $raw = [string]$bios.ReleaseDate

        if (-not [string]::IsNullOrWhiteSpace($raw)) {
            if ($raw -match '^\d{14}\.\d{6}[\+\-]\d{3}$') {
                try {
                    return [System.Management.ManagementDateTimeConverter]::ToDateTime($raw).ToString('yyyy-MM-dd')
                } catch {}
            }

            try {
                return ([datetime]::Parse($raw)).ToString('yyyy-MM-dd')
            } catch {}
        }

        return 'Unknown'
    }
    catch {
        return 'Unknown'
    }
}

function Get-SecureBootState {
    try {
        $result = Confirm-SecureBootUEFI -ErrorAction Stop
        if ($result) { return 'Enabled' }
        return 'Disabled'
    }
    catch {
        return 'Unknown/Unsupported'
    }
}

function Get-TpmState {
    try {
        $tpm = Get-Tpm -ErrorAction Stop
        if ($null -eq $tpm) { return 'Unknown' }
        if ($tpm.TpmPresent -and $tpm.TpmReady) { return 'PresentAndReady' }
        if ($tpm.TpmPresent) { return 'PresentNotReady' }
        return 'NotPresent'
    }
    catch {
        return 'Unknown'
    }
}

function Get-BitLockerState {
    try {
        $volume = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop
        if ($null -eq $volume) { return 'Unknown' }

        $protection = switch ($volume.ProtectionStatus) {
            0 { 'Off' }
            1 { 'On' }
            default { [string]$volume.ProtectionStatus }
        }

        $status = Get-SafeValue $volume.VolumeStatus
        $method = Get-SafeValue $volume.EncryptionMethod
        return "$protection | $status | $method"
    }
    catch {
        return 'Unknown/NotAvailable'
    }
}

function Get-OSDisplayName {
    try {
        $os = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        $caption = Get-SafeValue $os.Caption
        $version = Get-SafeValue $os.Version
        $build   = Get-SafeValue $os.BuildNumber

        $parts = @($caption)
        if ($version) { $parts += $version }
        if ($build)   { $parts += "Build $build" }
        return ($parts -join ' | ')
    }
    catch {
        return 'Unknown'
    }
}

function Get-CurrentDeviceData {
    try {
        $computerSystem = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
        $bios           = Get-CimInstance Win32_BIOS -ErrorAction Stop
        $csProduct      = Get-CimInstance Win32_ComputerSystemProduct -ErrorAction Stop
    }
    catch {
        throw "Unable to retrieve device data: $($_.Exception.Message)"
    }

    [pscustomobject]@{
        ComputerName = $env:COMPUTERNAME
        CurrentUser  = $env:USERNAME
        Manufacturer = Get-SafeValue $computerSystem.Manufacturer
        Model        = Get-SafeValue $computerSystem.Model
        SerialNumber = Get-SafeValue $csProduct.IdentifyingNumber
        OSName       = Get-OSDisplayName
        BIOSVersion  = Get-SafeValue $bios.SMBIOSBIOSVersion
        BIOSDate     = Get-BiosDateSafe
        SecureBoot   = Get-SecureBootState
        TPMState     = Get-TpmState
        BitLocker    = Get-BitLockerState
    }
}

# -----------------------------------------------------------------------------
# Persistence
# -----------------------------------------------------------------------------
function Load-LocalStatus {
    if (-not (Test-Path -LiteralPath $StatusFile)) { return $null }

    try {
        return Get-Content -LiteralPath $StatusFile -Raw -Encoding UTF8 | ConvertFrom-Json
    }
    catch {
        return $null
    }
}

function Save-StatusTextFile {
    param([Parameter(Mandatory)]$StatusObject)

    $lines = @(
        '===== DEVICE STATUS ====='
        "Computer Name        : $($StatusObject.ComputerName)"
        "Current User         : $($StatusObject.CurrentUser)"
        "Manufacturer         : $($StatusObject.Manufacturer)"
        "Model                : $($StatusObject.Model)"
        "Serial Number        : $($StatusObject.SerialNumber)"
        "Operating System     : $($StatusObject.OSName)"
        "Last Technician      : $($StatusObject.LastTouchedBy)"
        "Last Intervention    : $($StatusObject.LastTouchedDate)"
        "Current BIOS         : $($StatusObject.BIOSVersion)"
        "BIOS Date            : $($StatusObject.BIOSDate)"
        "Secure Boot          : $($StatusObject.SecureBoot)"
        "TPM                  : $($StatusObject.TPMState)"
        "BitLocker            : $($StatusObject.BitLocker)"
        "BIOS Updated         : $($StatusObject.BIOSUpdatedThisVisit)"
        "Docking Updated      : $($StatusObject.DockingUpdated)"
        "Scans Executed       : $($StatusObject.ScansExecuted)"
        "Summary              : $($StatusObject.LastActionSummary)"
        "Pending Actions      : $($StatusObject.PendingActions)"
        "Notes                : $($StatusObject.Notes)"
    )

    $lines | Set-Content -LiteralPath $StatusTextFile -Encoding UTF8
}

function Save-Intervention {
    param(
        [string]$Technician,
        [string]$BiosUpdatedThisVisit,
        [string]$DockingUpdated,
        [string]$ScansExecuted,
        [string]$PendingActions,
        [string]$Notes,
        [string]$ActionSummary
    )

    Ensure-Folder -Path $BaseFolder

    $deviceData = Get-CurrentDeviceData
    $timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')

    if ([string]::IsNullOrWhiteSpace($Technician)) { $Technician = $env:USERNAME }
    if ([string]::IsNullOrWhiteSpace($ActionSummary)) { $ActionSummary = 'Intervention recorded' }

    $statusObject = [ordered]@{
        ComputerName           = $deviceData.ComputerName
        CurrentUser            = $deviceData.CurrentUser
        Manufacturer           = $deviceData.Manufacturer
        Model                  = $deviceData.Model
        SerialNumber           = $deviceData.SerialNumber
        OSName                 = $deviceData.OSName
        LastTouchedBy          = $Technician
        LastTouchedDate        = $timestamp
        BIOSVersion            = $deviceData.BIOSVersion
        BIOSDate               = $deviceData.BIOSDate
        SecureBoot             = $deviceData.SecureBoot
        TPMState               = $deviceData.TPMState
        BitLocker              = $deviceData.BitLocker
        BIOSUpdatedThisVisit   = $BiosUpdatedThisVisit
        DockingUpdated         = $DockingUpdated
        ScansExecuted          = $ScansExecuted
        PendingActions         = $PendingActions
        Notes                  = $Notes
        LastActionSummary      = $ActionSummary
    }

    $statusObject | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $StatusFile -Encoding UTF8
    Save-StatusTextFile -StatusObject $statusObject

    $logObject = [pscustomobject]@{
        DateTime             = $timestamp
        ComputerName         = $deviceData.ComputerName
        User                 = $Technician
        Manufacturer         = $deviceData.Manufacturer
        Model                = $deviceData.Model
        SerialNumber         = $deviceData.SerialNumber
        OSName               = $deviceData.OSName
        BIOSVersion          = $deviceData.BIOSVersion
        BIOSDate             = $deviceData.BIOSDate
        SecureBoot           = $deviceData.SecureBoot
        TPMState             = $deviceData.TPMState
        BitLocker            = $deviceData.BitLocker
        BIOSUpdatedThisVisit = $BiosUpdatedThisVisit
        DockingUpdated       = $DockingUpdated
        ScansExecuted        = $ScansExecuted
        ActionSummary        = $ActionSummary
        PendingActions       = $PendingActions
        Notes                = $Notes
    }

    if (Test-Path -LiteralPath $LogFile) {
        $logObject | Export-Csv -LiteralPath $LogFile -NoTypeInformation -Append -Encoding UTF8
    }
    else {
        $logObject | Export-Csv -LiteralPath $LogFile -NoTypeInformation -Encoding UTF8
    }
}

function Open-LogsFolder {
    Ensure-Folder -Path $BaseFolder
    Start-Process explorer.exe $BaseFolder | Out-Null
}

function Show-HelpDialog {
    $helpText = @'
Endpoint Support Checklist

What is it for?
- Displays the current technical status of the device.
- Stores the latest intervention locally.
- Maintains a local CSV history.
- Lets you open the local logs folder.

What each button does:
- Refresh: reads the current device state and local history again.
- Save: records a new local technical intervention.
- Open Logs: opens the local logs folder.
- Help: shows this explanation.
- Exit: closes the application.

Generated local files:
- DeviceStatus.json : latest saved status
- DeviceStatus.txt  : plain-text status summary
- MaintenanceLog.csv: local intervention history

Manual fields:
- Technician
- BIOS Updated
- Docking Updated
- Scans Executed
- Summary
- Pending Actions
- Notes
'@

    Show-Message -Text $helpText -Title 'Help' -Icon ([System.Windows.Forms.MessageBoxIcon]::Information)
}

# -----------------------------------------------------------------------------
# WinForms control factory helpers
# -----------------------------------------------------------------------------
function New-HeaderButton {
    param(
        [Parameter(Mandatory)][string]$Text,
        [int]$Width = 120,
        [switch]$Primary
    )

    $button = New-Object System.Windows.Forms.Button
    $button.Text = $Text
    $button.Size = [System.Drawing.Size]::new($Width, 36)
    $button.Margin = New-Object System.Windows.Forms.Padding(0,0,12,0)
    $button.FlatStyle = 'Standard'

    if ($Primary) {
        $button.BackColor = $BrandBlue
        $button.ForeColor = [System.Drawing.Color]::White
        $button.FlatStyle = 'Flat'
        $button.FlatAppearance.BorderSize = 0
    }

    return $button
}

function New-SectionPanel {
    param([Parameter(Mandatory)][string]$Title)

    $container = New-Object System.Windows.Forms.Panel
    $container.Dock = 'Fill'
    $container.Margin = New-Object System.Windows.Forms.Padding(8)
    $container.Padding = New-Object System.Windows.Forms.Padding(10,28,10,10)
    $container.BackColor = [System.Drawing.Color]::White
    $container.BorderStyle = 'FixedSingle'

    $caption = New-Object System.Windows.Forms.Label
    $caption.Text = $Title
    $caption.AutoSize = $true
    $caption.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
    $caption.BackColor = [System.Drawing.Color]::White
    $caption.Location = [System.Drawing.Point]::new(12, -1)
    $caption.Padding = New-Object System.Windows.Forms.Padding(6,0,6,0)
    $container.Controls.Add($caption)

    return $container
}

function New-DataTableLayout {
    param(
        [int]$Rows,
        [int]$LabelWidth = 170,
        [int]$RowHeight = 34
    )

    $table = New-Object System.Windows.Forms.TableLayoutPanel
    $table.Dock = 'Fill'
    $table.ColumnCount = 2
    $table.RowCount = $Rows
    $table.Margin = '0,0,0,0'
    $table.Padding = '0,0,0,0'
    $table.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, $LabelWidth))) | Out-Null
    $table.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null

    for ($i = 0; $i -lt $Rows; $i++) {
        $table.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, $RowHeight))) | Out-Null
    }

    return $table
}

function New-FieldLabel {
    param([string]$Text)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Dock = 'Fill'
    $label.TextAlign = 'MiddleLeft'
    $label.Margin = New-Object System.Windows.Forms.Padding(4,3,10,3)
    return $label
}

function New-ReadOnlyTextBox {
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Dock = 'Fill'
    $textBox.ReadOnly = $true
    $textBox.BackColor = [System.Drawing.Color]::White
    $textBox.Margin = New-Object System.Windows.Forms.Padding(4,4,4,4)
    return $textBox
}

function Add-LabeledReadOnlyField {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.TableLayoutPanel]$Table,
        [Parameter(Mandatory)][int]$Row,
        [Parameter(Mandatory)][string]$LabelText
    )

    $label = New-FieldLabel -Text $LabelText
    $textBox = New-ReadOnlyTextBox
    $Table.Controls.Add($label, 0, $Row)
    $Table.Controls.Add($textBox, 1, $Row)
    return $textBox
}

function Add-InputLabel {
    param(
        [Parameter(Mandatory)][System.Windows.Forms.TableLayoutPanel]$Table,
        [Parameter(Mandatory)][string]$Text,
        [Parameter(Mandatory)][int]$Row
    )

    $Table.Controls.Add((New-FieldLabel -Text $Text), 0, $Row)
}

function New-Combo {
    $combo = New-Object System.Windows.Forms.ComboBox
    $combo.Dock = 'Fill'
    $combo.DropDownStyle = 'DropDownList'
    $combo.Margin = New-Object System.Windows.Forms.Padding(4,4,4,4)
    [void]$combo.Items.AddRange(@('Yes','No','N/A'))
    $combo.SelectedItem = 'N/A'
    return $combo
}

# -----------------------------------------------------------------------------
# Build UI
# -----------------------------------------------------------------------------
Ensure-Folder -Path $BaseFolder

$form = New-Object System.Windows.Forms.Form
$form.Text = "$AppTitle | $ScriptVersion"
$form.Size = [System.Drawing.Size]::new(1400, 930)
$form.MinimumSize = [System.Drawing.Size]::new(1220, 860)
$form.StartPosition = 'CenterScreen'
$form.BackColor = [System.Drawing.Color]::White
$form.Font = New-Object System.Drawing.Font('Segoe UI', 10)
$form.FormBorderStyle = 'Sizable'
$form.MaximizeBox = $true

$root = New-Object System.Windows.Forms.TableLayoutPanel
$root.Dock = 'Fill'
$root.ColumnCount = 1
$root.RowCount = 3
$root.Padding = New-Object System.Windows.Forms.Padding(14,14,14,10)
$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 72))) | Out-Null
$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 26))) | Out-Null
$form.Controls.Add($root)

# Header
$header = New-Object System.Windows.Forms.TableLayoutPanel
$header.Dock = 'Fill'
$header.ColumnCount = 2
$header.RowCount = 1
$header.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 55))) | Out-Null
$header.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 45))) | Out-Null
$root.Controls.Add($header, 0, 0)

$titlePanel = New-Object System.Windows.Forms.Panel
$titlePanel.Dock = 'Fill'
$header.Controls.Add($titlePanel, 0, 0)

$titleMain = New-Object System.Windows.Forms.Label
$titleMain.Text = 'Endpoint Support Checklist'
$titleMain.Font = New-Object System.Drawing.Font('Segoe UI', 24, [System.Drawing.FontStyle]::Bold)
$titleMain.AutoSize = $true
$titleMain.Location = [System.Drawing.Point]::new(6, 12)
$titlePanel.Controls.Add($titleMain)

$titleVersion = New-Object System.Windows.Forms.Label
$titleVersion.Text = " $ScriptVersion"
$titleVersion.Font = New-Object System.Drawing.Font('Segoe UI', 22, [System.Drawing.FontStyle]::Bold)
$titleVersion.ForeColor = $BrandBlue
$titleVersion.AutoSize = $true
$titleVersion.Location = [System.Drawing.Point]::new(530, 14)
$titlePanel.Controls.Add($titleVersion)

$buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$buttonPanel.Dock = 'Fill'
$buttonPanel.FlowDirection = 'RightToLeft'
$buttonPanel.WrapContents = $false
$buttonPanel.Padding = New-Object System.Windows.Forms.Padding(0,18,0,0)
$header.Controls.Add($buttonPanel, 1, 0)

$btnExit    = New-HeaderButton -Text 'Exit'      -Width 90
$btnHelp    = New-HeaderButton -Text 'Help'      -Width 100
$btnLogs    = New-HeaderButton -Text 'Open Logs' -Width 120
$btnSave    = New-HeaderButton -Text 'Save'      -Width 120 -Primary
$btnRefresh = New-HeaderButton -Text 'Refresh'   -Width 120
$buttonPanel.Controls.AddRange(@($btnExit,$btnHelp,$btnLogs,$btnSave,$btnRefresh))

# Main layout
$main = New-Object System.Windows.Forms.TableLayoutPanel
$main.Dock = 'Fill'
$main.ColumnCount = 2
$main.RowCount = 2
$main.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 52))) | Out-Null
$main.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 48))) | Out-Null
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 48))) | Out-Null
$main.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 52))) | Out-Null
$root.Controls.Add($main, 0, 1)

$secCurrent  = New-SectionPanel -Title 'Detected Current Status'
$secStored   = New-SectionPanel -Title 'Last Saved Status'
$secRegister = New-SectionPanel -Title 'Register Intervention'
$secHistory  = New-SectionPanel -Title 'Local History'
$main.Controls.Add($secCurrent,  0, 0)
$main.Controls.Add($secStored,   1, 0)
$main.Controls.Add($secRegister, 0, 1)
$main.Controls.Add($secHistory,  1, 1)

# Current status section
$currentTable = New-DataTableLayout -Rows 9 -LabelWidth 190 -RowHeight 34
$secCurrent.Controls.Add($currentTable)
$txtCurComputer  = Add-LabeledReadOnlyField -Table $currentTable -Row 0 -LabelText 'Computer Name'
$txtCurUser      = Add-LabeledReadOnlyField -Table $currentTable -Row 1 -LabelText 'Current User'
$txtCurModel     = Add-LabeledReadOnlyField -Table $currentTable -Row 2 -LabelText 'Model'
$txtCurSerial    = Add-LabeledReadOnlyField -Table $currentTable -Row 3 -LabelText 'Serial Number'
$txtCurOS        = Add-LabeledReadOnlyField -Table $currentTable -Row 4 -LabelText 'Operating System'
$txtCurBios      = Add-LabeledReadOnlyField -Table $currentTable -Row 5 -LabelText 'Current BIOS'
$txtCurBiosDate  = Add-LabeledReadOnlyField -Table $currentTable -Row 6 -LabelText 'BIOS Date'
$txtCurSecure    = Add-LabeledReadOnlyField -Table $currentTable -Row 7 -LabelText 'Secure Boot'
$txtCurTpm       = Add-LabeledReadOnlyField -Table $currentTable -Row 8 -LabelText 'TPM'

# Stored status section
$storedTable = New-DataTableLayout -Rows 9 -LabelWidth 180 -RowHeight 34
$secStored.Controls.Add($storedTable)
$txtStoTech      = Add-LabeledReadOnlyField -Table $storedTable -Row 0 -LabelText 'Last Technician'
$txtStoDate      = Add-LabeledReadOnlyField -Table $storedTable -Row 1 -LabelText 'Last Intervention'
$txtStoModel     = Add-LabeledReadOnlyField -Table $storedTable -Row 2 -LabelText 'Model'
$txtStoOS        = Add-LabeledReadOnlyField -Table $storedTable -Row 3 -LabelText 'Operating System'
$txtStoBios      = Add-LabeledReadOnlyField -Table $storedTable -Row 4 -LabelText 'Current BIOS'
$txtStoSecure    = Add-LabeledReadOnlyField -Table $storedTable -Row 5 -LabelText 'Secure Boot'
$txtStoTpm       = Add-LabeledReadOnlyField -Table $storedTable -Row 6 -LabelText 'TPM'
$txtStoBitlocker = Add-LabeledReadOnlyField -Table $storedTable -Row 7 -LabelText 'BitLocker'
$txtStoSummary   = Add-LabeledReadOnlyField -Table $storedTable -Row 8 -LabelText 'Summary'

# Register section
$registerTable = New-Object System.Windows.Forms.TableLayoutPanel
$registerTable.Dock = 'Fill'
$registerTable.ColumnCount = 2
$registerTable.RowCount = 8
$registerTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 190))) | Out-Null
$registerTable.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
for ($i = 0; $i -lt 7; $i++) {
    $registerTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 38))) | Out-Null
}
$registerTable.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100))) | Out-Null
$secRegister.Controls.Add($registerTable)

Add-InputLabel -Table $registerTable -Text 'Technician' -Row 0
Add-InputLabel -Table $registerTable -Text 'BIOS Updated' -Row 1
Add-InputLabel -Table $registerTable -Text 'Docking Updated' -Row 2
Add-InputLabel -Table $registerTable -Text 'Scans Executed' -Row 3
Add-InputLabel -Table $registerTable -Text 'Summary' -Row 4
Add-InputLabel -Table $registerTable -Text 'Pending Actions' -Row 5
Add-InputLabel -Table $registerTable -Text 'Notes' -Row 6

$txtTechnician = New-Object System.Windows.Forms.TextBox
$txtTechnician.Dock = 'Fill'
$txtTechnician.Margin = New-Object System.Windows.Forms.Padding(4,4,4,4)
$txtTechnician.Text = $env:USERNAME
$registerTable.Controls.Add($txtTechnician, 1, 0)

$cmbBiosUpdated = New-Combo
$cmbDocking     = New-Combo
$cmbScans       = New-Combo
$registerTable.Controls.Add($cmbBiosUpdated, 1, 1)
$registerTable.Controls.Add($cmbDocking,     1, 2)
$registerTable.Controls.Add($cmbScans,       1, 3)

$txtActionSummary = New-Object System.Windows.Forms.TextBox
$txtActionSummary.Dock = 'Fill'
$txtActionSummary.Margin = New-Object System.Windows.Forms.Padding(4,4,4,4)
$registerTable.Controls.Add($txtActionSummary, 1, 4)

$txtPending = New-Object System.Windows.Forms.TextBox
$txtPending.Dock = 'Fill'
$txtPending.Margin = New-Object System.Windows.Forms.Padding(4,4,4,4)
$registerTable.Controls.Add($txtPending, 1, 5)

$txtNotes = New-Object System.Windows.Forms.TextBox
$txtNotes.Dock = 'Fill'
$txtNotes.Margin = New-Object System.Windows.Forms.Padding(4,4,4,4)
$txtNotes.Multiline = $true
$txtNotes.ScrollBars = 'Vertical'
$registerTable.Controls.Add($txtNotes, 1, 6)
$registerTable.SetRowSpan($txtNotes, 2)

# History section
$historySplit = New-Object System.Windows.Forms.SplitContainer
$historySplit.Dock = 'Fill'
$historySplit.Orientation = 'Horizontal'
$historySplit.SplitterDistance = 220
$historySplit.Panel1MinSize = 120
$historySplit.Panel2MinSize = 100
$secHistory.Controls.Add($historySplit)

$gridHistory = New-Object System.Windows.Forms.DataGridView
$gridHistory.Dock = 'Fill'
$gridHistory.ReadOnly = $true
$gridHistory.AllowUserToAddRows = $false
$gridHistory.AllowUserToDeleteRows = $false
$gridHistory.RowHeadersVisible = $false
$gridHistory.AutoSizeColumnsMode = 'Fill'
$gridHistory.SelectionMode = 'FullRowSelect'
$gridHistory.MultiSelect = $false
$gridHistory.AutoGenerateColumns = $true
$historySplit.Panel1.Controls.Add($gridHistory)

$txtHistoryDetail = New-Object System.Windows.Forms.TextBox
$txtHistoryDetail.Dock = 'Fill'
$txtHistoryDetail.Multiline = $true
$txtHistoryDetail.ScrollBars = 'Vertical'
$txtHistoryDetail.ReadOnly = $true
$txtHistoryDetail.WordWrap = $true
$txtHistoryDetail.BackColor = [System.Drawing.Color]::White
$historySplit.Panel2.Controls.Add($txtHistoryDetail)

# Footer
$footerPanel = New-Object System.Windows.Forms.Panel
$footerPanel.Dock = 'Fill'
$root.Controls.Add($footerPanel, 0, 2)

$lblInfo = New-Object System.Windows.Forms.Label
$lblInfo.Dock = 'Left'
$lblInfo.AutoSize = $false
$lblInfo.Width = 500
$lblInfo.TextAlign = 'MiddleLeft'
$lblInfo.Text = 'Ready'
$footerPanel.Controls.Add($lblInfo)

$lblSignature = New-Object System.Windows.Forms.Label
$lblSignature.Dock = 'Right'
$lblSignature.AutoSize = $false
$lblSignature.Width = 240
$lblSignature.TextAlign = 'MiddleRight'
$lblSignature.Text = "Rafael Alba | 2026 | $ScriptVersion"
$footerPanel.Controls.Add($lblSignature)

# -----------------------------------------------------------------------------
# UI refresh routines
# -----------------------------------------------------------------------------
function Show-HistoryDetail {
    if ($gridHistory.SelectedRows.Count -eq 0) {
        $txtHistoryDetail.Text = ''
        return
    }

    $row = $gridHistory.SelectedRows[0]
    $txtHistoryDetail.Lines = @(
        "Date: $($row.Cells['Date'].Value)"
        "Computer: $($row.Cells['Computer'].Value)"
        "Technician: $($row.Cells['Technician'].Value)"
        "Model: $($row.Cells['Model'].Value)"
        "Operating System: $($row.Cells['OperatingSystem'].Value)"
        "BIOS: $($row.Cells['BIOSVersion'].Value)"
        "Summary: $($row.Cells['Summary'].Value)"
        "Pending Actions: $($row.Cells['PendingActions'].Value)"
        "Notes: $($row.Cells['Notes'].Value)"
        "Secure Boot: $($row.Cells['SecureBoot'].Value)"
        "TPM: $($row.Cells['TPM'].Value)"
        "BitLocker: $($row.Cells['BitLocker'].Value)"
    )
}

function Refresh-CurrentSection {
    $deviceData = Get-CurrentDeviceData
    $txtCurComputer.Text = $deviceData.ComputerName
    $txtCurUser.Text     = $deviceData.CurrentUser
    $txtCurModel.Text    = $deviceData.Model
    $txtCurSerial.Text   = $deviceData.SerialNumber
    $txtCurOS.Text       = $deviceData.OSName
    $txtCurBios.Text     = $deviceData.BIOSVersion
    $txtCurBiosDate.Text = $deviceData.BIOSDate
    $txtCurSecure.Text   = $deviceData.SecureBoot
    $txtCurTpm.Text      = $deviceData.TPMState
}

function Refresh-StoredSection {
    $status = Load-LocalStatus

    if ($null -eq $status) {
        foreach ($tb in @($txtStoTech,$txtStoDate,$txtStoModel,$txtStoOS,$txtStoBios,$txtStoSecure,$txtStoTpm,$txtStoBitlocker,$txtStoSummary)) {
            $tb.Text = ''
        }
        return
    }

    $txtStoTech.Text      = Get-SafeValue $status.LastTouchedBy
    $txtStoDate.Text      = Get-SafeValue $status.LastTouchedDate
    $txtStoModel.Text     = Get-SafeValue $status.Model
    $txtStoOS.Text        = Get-SafeValue $status.OSName
    $txtStoBios.Text      = Get-SafeValue $status.BIOSVersion
    $txtStoSecure.Text    = Get-SafeValue $status.SecureBoot
    $txtStoTpm.Text       = Get-SafeValue $status.TPMState
    $txtStoBitlocker.Text = Get-SafeValue $status.BitLocker
    $txtStoSummary.Text   = Get-SafeValue $status.LastActionSummary
}

function Refresh-HistorySection {
    if (-not (Test-Path -LiteralPath $LogFile)) {
        $gridHistory.DataSource = $null
        $txtHistoryDetail.Text = ''
        return
    }

    try {
        $history = @(Import-Csv -LiteralPath $LogFile | Sort-Object DateTime -Descending)
        if ($history.Count -eq 0) {
            $gridHistory.DataSource = $null
            $txtHistoryDetail.Text = ''
            return
        }

        $table = New-Object System.Data.DataTable
        foreach ($column in @('Date','Computer','Technician','Model','OperatingSystem','BIOSVersion','Summary','PendingActions','Notes','BitLocker','SecureBoot','TPM')) {
            [void]$table.Columns.Add($column)
        }

        foreach ($entry in $history) {
            $newRow = $table.NewRow()
            $newRow['Date']            = [string]$entry.DateTime
            $newRow['Computer']        = [string]$entry.ComputerName
            $newRow['Technician']      = [string]$entry.User
            $newRow['Model']           = [string]$entry.Model
            $newRow['OperatingSystem'] = [string]$entry.OSName
            $newRow['BIOSVersion']     = [string]$entry.BIOSVersion
            $newRow['Summary']         = [string]$entry.ActionSummary
            $newRow['PendingActions']  = [string]$entry.PendingActions
            $newRow['Notes']           = [string]$entry.Notes
            $newRow['BitLocker']       = [string]$entry.BitLocker
            $newRow['SecureBoot']      = [string]$entry.SecureBoot
            $newRow['TPM']             = [string]$entry.TPMState
            [void]$table.Rows.Add($newRow)
        }

        $gridHistory.DataSource = $null
        $gridHistory.DataSource = $table

        foreach ($colName in @('PendingActions','Notes','BitLocker','SecureBoot','TPM','Model','OperatingSystem','BIOSVersion')) {
            if ($gridHistory.Columns.Contains($colName)) {
                $gridHistory.Columns[$colName].Visible = $false
            }
        }

        if ($gridHistory.Rows.Count -gt 0) {
            $gridHistory.Rows[0].Selected = $true
            Show-HistoryDetail
        }
        else {
            $txtHistoryDetail.Text = ''
        }
    }
    catch {
        $gridHistory.DataSource = $null
        $txtHistoryDetail.Text = ''
    }
}

function Refresh-All {
    Refresh-CurrentSection
    Refresh-StoredSection
    Refresh-HistorySection
    $lblInfo.Text = 'Status refreshed'
}

# -----------------------------------------------------------------------------
# Event wiring
# -----------------------------------------------------------------------------
$gridHistory.add_SelectionChanged({ try { Show-HistoryDetail } catch {} })
$gridHistory.add_CellDoubleClick({ try { Show-HistoryDetail } catch {} })

$btnRefresh.Add_Click({
    try {
        Refresh-All
    }
    catch {
        Show-Message -Text "Error refreshing: $($_.Exception.Message)" -Title 'Error' -Icon ([System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$btnSave.Add_Click({
    try {
        Save-Intervention `
            -Technician $txtTechnician.Text `
            -BiosUpdatedThisVisit $cmbBiosUpdated.SelectedItem `
            -DockingUpdated $cmbDocking.SelectedItem `
            -ScansExecuted $cmbScans.SelectedItem `
            -PendingActions $txtPending.Text `
            -Notes $txtNotes.Text `
            -ActionSummary $txtActionSummary.Text

        Refresh-All
        $lblInfo.Text = 'Intervention saved successfully'
        Show-Message -Text 'Intervention saved successfully.' -Title 'OK' -Icon ([System.Windows.Forms.MessageBoxIcon]::Information)
    }
    catch {
        Show-Message -Text "Error saving intervention: $($_.Exception.Message)" -Title 'Error' -Icon ([System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$btnLogs.Add_Click({
    try {
        Open-LogsFolder
    }
    catch {
        Show-Message -Text "Could not open logs folder: $($_.Exception.Message)" -Title 'Error' -Icon ([System.Windows.Forms.MessageBoxIcon]::Error)
    }
})

$btnHelp.Add_Click({ Show-HelpDialog })
$btnExit.Add_Click({ $form.Close() })

# -----------------------------------------------------------------------------
# Launch
# -----------------------------------------------------------------------------
Refresh-All
[void]$form.ShowDialog()
