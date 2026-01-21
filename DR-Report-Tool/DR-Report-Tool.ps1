#Requires -Version 5.1
<#
.SYNOPSIS
    AssetCentre DR Report Tool - Main Entry Point
.DESCRIPTION
    PowerShell application with XAML GUI that generates and sends
    DR (Disaster Recovery) backup status reports from FactoryTalk AssetCentre.
.PARAMETER Silent
    Run in silent mode without GUI (for scheduled tasks)
.PARAMETER ConfigPath
    Path to settings.json configuration file
.PARAMETER PreviewOnly
    Generate HTML report and open in browser without sending email
.EXAMPLE
    .\DR-Report-Tool.ps1
    Launches the GUI application
.EXAMPLE
    .\DR-Report-Tool.ps1 -Silent -ConfigPath ".\Config\settings.json"
    Runs in silent mode using specified configuration
.EXAMPLE
    .\DR-Report-Tool.ps1 -PreviewOnly
    Generates report and opens in browser for preview
.NOTES
    All data stays within the tenant - no cloud services or external APIs.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [switch]$Silent,

    [Parameter()]
    [string]$ConfigPath,

    [Parameter()]
    [switch]$PreviewOnly
)

# Script configuration
$ErrorActionPreference = 'Stop'
$script:ScriptRoot = $PSScriptRoot
$script:LogPath = Join-Path -Path $script:ScriptRoot -ChildPath "Logs\execution.log"
$script:DefaultConfigPath = Join-Path -Path $script:ScriptRoot -ChildPath "Config\settings.json"

#region Logging Functions

function Write-Log {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Message,

        [Parameter()]
        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Level = 'Info'
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"

    # Ensure log directory exists
    $logDir = Split-Path -Path $script:LogPath -Parent
    if (-not (Test-Path -Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }

    # Append to log file
    Add-Content -Path $script:LogPath -Value $logEntry

    # Also write to console in silent mode
    if ($Silent) {
        switch ($Level) {
            'Warning' { Write-Warning $Message }
            'Error'   { Write-Error $Message -ErrorAction Continue }
            default   { Write-Host $logEntry }
        }
    }
}

#endregion

#region Module Loading

function Import-RequiredModules {
    $modulePath = Join-Path -Path $script:ScriptRoot -ChildPath "Modules"

    $modules = @(
        'Database.psm1',
        'HtmlGenerator.psm1',
        'EmailSender.psm1',
        'Scheduler.psm1'
    )

    foreach ($module in $modules) {
        $fullPath = Join-Path -Path $modulePath -ChildPath $module
        if (Test-Path -Path $fullPath) {
            Import-Module $fullPath -Force
            Write-Log "Loaded module: $module"
        }
        else {
            throw "Required module not found: $fullPath"
        }
    }
}

#endregion

#region Configuration Functions

function Get-Configuration {
    param(
        [Parameter()]
        [string]$Path
    )

    if (-not $Path) {
        $Path = $script:DefaultConfigPath
    }

    if (-not (Test-Path -Path $Path)) {
        return $null
    }

    try {
        $config = Get-Content -Path $Path -Raw | ConvertFrom-Json
        return $config
    }
    catch {
        Write-Log "Failed to load configuration: $($_.Exception.Message)" -Level Error
        return $null
    }
}

function Save-Configuration {
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Config,

        [Parameter()]
        [string]$Path
    )

    if (-not $Path) {
        $Path = $script:DefaultConfigPath
    }

    try {
        # Ensure directory exists
        $configDir = Split-Path -Path $Path -Parent
        if (-not (Test-Path -Path $configDir)) {
            New-Item -ItemType Directory -Path $configDir -Force | Out-Null
        }

        $Config | ConvertTo-Json -Depth 10 | Set-Content -Path $Path -Encoding UTF8
        Write-Log "Configuration saved to: $Path"
        return $true
    }
    catch {
        Write-Log "Failed to save configuration: $($_.Exception.Message)" -Level Error
        return $false
    }
}

function New-DefaultConfiguration {
    return [PSCustomObject]@{
        database = [PSCustomObject]@{
            server = ""
            database = "AssetCentre"
            useWindowsAuth = $true
        }
        report = [PSCustomObject]@{
            scope = "all"
            folderFilter = $null
            includeWarnings = $true
            includeSuccess = $false
            includeFullMessage = $false
            attachCsv = $true
        }
        email = [PSCustomObject]@{
            to = @()
            cc = @()
            subjectTemplate = "DR Report - {Date} - {Summary}"
        }
        schedule = [PSCustomObject]@{
            enabled = $false
            frequency = "Weekly"
            dayOfWeek = "Monday"
            time = "07:00"
        }
    }
}

#endregion

#region Report Execution

function Invoke-ReportGeneration {
    param(
        [Parameter(Mandatory = $true)]
        [PSCustomObject]$Config,

        [Parameter()]
        [switch]$PreviewOnly
    )

    Write-Log "Starting report generation..."

    try {
        # Get data from database or use sample data
        $data = $null

        if ($Config.database.server) {
            Write-Log "Connecting to database: $($Config.database.server)\$($Config.database.database)"

            $queryParams = @{
                Server = $Config.database.server
                Database = $Config.database.database
                UseWindowsAuth = $Config.database.useWindowsAuth
            }

            if ($Config.report.folderFilter) {
                $queryParams['FolderFilter'] = $Config.report.folderFilter
            }

            $data = Invoke-AssetCentreQuery @queryParams

            if (-not $data) {
                throw "Failed to retrieve data from database: $(Get-LastDatabaseError)"
            }

            Write-Log "Retrieved $($data.Rows.Count) records from database"
        }
        else {
            Write-Log "No database configured - using sample data"
            $data = Get-SampleData
        }

        # Check for empty results
        if ($data.Rows.Count -eq 0) {
            Write-Log "No data returned from query" -Level Warning
        }

        # Generate HTML report
        Write-Log "Generating HTML report..."

        $reportParams = @{
            Data = $data
            IncludeFullMessage = $Config.report.includeFullMessage
            ScheduleName = if ($Config.schedule.enabled) { "Scheduled" } else { "Manual" }
        }

        if ($Config.report.scope -eq "failedOnly") {
            $reportParams['FailedOnly'] = $true
        }
        if ($Config.report.includeWarnings) {
            $reportParams['IncludeWarnings'] = $true
        }
        if ($Config.report.includeSuccess) {
            $reportParams['IncludeSuccess'] = $true
        }

        $htmlReport = New-DRReport @reportParams

        # Generate subject
        $subject = Get-ReportSubject -SubjectTemplate $Config.email.subjectTemplate -Data $data

        # Export CSV if configured
        $csvPath = $null
        if ($Config.report.attachCsv) {
            $csvPath = Join-Path -Path $script:ScriptRoot -ChildPath "Logs\DR_Report_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
            Export-DataToCsv -Data $data -OutputPath $csvPath
            Write-Log "CSV exported to: $csvPath"
        }

        if ($PreviewOnly) {
            # Save HTML to temp file and open in browser
            $tempHtml = Join-Path -Path $env:TEMP -ChildPath "DR_Report_Preview_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
            $htmlReport | Set-Content -Path $tempHtml -Encoding UTF8
            Start-Process $tempHtml
            Write-Log "Preview opened in browser: $tempHtml"

            return @{
                Success = $true
                Message = "Preview generated"
                HtmlPath = $tempHtml
                CsvPath = $csvPath
            }
        }

        # Send email
        if ($Config.email.to -and $Config.email.to.Count -gt 0) {
            Write-Log "Sending email to: $($Config.email.to -join ', ')"

            $emailParams = @{
                To = $Config.email.to
                Subject = $subject
                HtmlBody = $htmlReport
            }

            if ($Config.email.cc -and $Config.email.cc.Count -gt 0) {
                $emailParams['CC'] = $Config.email.cc
            }

            if ($csvPath) {
                $emailParams['AttachmentPath'] = $csvPath
            }

            $result = Send-OutlookReport @emailParams

            if ($result.Success) {
                Write-Log "Email sent successfully"
                return @{
                    Success = $true
                    Message = $result.Message
                    SentTime = $result.SentTime
                    Recipients = $result.Recipients
                    CsvPath = $csvPath
                }
            }
            else {
                throw "Email send failed: $($result.Message)"
            }
        }
        else {
            Write-Log "No email recipients configured" -Level Warning
            return @{
                Success = $true
                Message = "Report generated (no recipients configured)"
                CsvPath = $csvPath
            }
        }
    }
    catch {
        Write-Log "Report generation failed: $($_.Exception.Message)" -Level Error

        # Try to send error notification
        if ($Config.email.to -and $Config.email.to.Count -gt 0 -and -not $PreviewOnly) {
            try {
                Send-ErrorNotification -To $Config.email.to -ErrorMessage $_.Exception.Message
            }
            catch {
                Write-Log "Failed to send error notification: $($_.Exception.Message)" -Level Error
            }
        }

        return @{
            Success = $false
            Message = $_.Exception.Message
        }
    }
}

#endregion

#region GUI Functions

function Show-MainWindow {
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName System.Web

    # Load XAML
    $xamlPath = Join-Path -Path $script:ScriptRoot -ChildPath "DR-Report-Tool.xaml"
    if (-not (Test-Path -Path $xamlPath)) {
        throw "XAML file not found: $xamlPath"
    }

    [xml]$xaml = Get-Content -Path $xamlPath
    $reader = New-Object System.Xml.XmlNodeReader $xaml
    $window = [Windows.Markup.XamlReader]::Load($reader)

    # Get control references
    $controls = @{}
    $controlNames = @(
        'txtServer', 'cboDatabase', 'rbWindowsAuth', 'rbSqlAuth', 'btnTestConnection',
        'pnlSqlAuth', 'txtUsername', 'txtPassword',
        'rbEntireTree', 'rbSpecificFolder', 'txtFolderFilter', 'btnBrowseFolder',
        'chkFailedOnly', 'chkWarnings', 'chkSuccess',
        'txtEmailTo', 'txtEmailCC', 'txtEmailSubject',
        'chkIncludeFullMessage', 'chkAttachCsv',
        'chkEnableSchedule', 'pnlScheduleOptions', 'cboFrequency', 'cboDay', 'txtTime',
        'btnCreateTask', 'btnRemoveTask', 'txtTaskStatus',
        'btnExecute', 'btnPreview', 'btnSaveConfig', 'btnLoadConfig',
        'txtStatus', 'txtLastRun'
    )

    foreach ($name in $controlNames) {
        $controls[$name] = $window.FindName($name)
    }

    # Helper function to update status
    $updateStatus = {
        param([string]$Text, [string]$Color = "#1A1A1A")
        $controls['txtStatus'].Text = $Text
        $controls['txtStatus'].Foreground = [System.Windows.Media.BrushConverter]::new().ConvertFromString($Color)
    }

    # Load existing configuration
    $config = Get-Configuration
    if ($config) {
        # Populate controls from config
        $controls['txtServer'].Text = $config.database.server
        $controls['cboDatabase'].Text = $config.database.database
        $controls['rbWindowsAuth'].IsChecked = $config.database.useWindowsAuth
        $controls['rbSqlAuth'].IsChecked = -not $config.database.useWindowsAuth

        if ($config.report.folderFilter) {
            $controls['rbSpecificFolder'].IsChecked = $true
            $controls['txtFolderFilter'].Text = $config.report.folderFilter
        }

        $controls['chkWarnings'].IsChecked = $config.report.includeWarnings
        $controls['chkSuccess'].IsChecked = $config.report.includeSuccess
        $controls['chkIncludeFullMessage'].IsChecked = $config.report.includeFullMessage
        $controls['chkAttachCsv'].IsChecked = $config.report.attachCsv

        if ($config.email.to) {
            $controls['txtEmailTo'].Text = ($config.email.to -join ', ')
        }
        if ($config.email.cc) {
            $controls['txtEmailCC'].Text = ($config.email.cc -join ', ')
        }
        $controls['txtEmailSubject'].Text = $config.email.subjectTemplate

        $controls['chkEnableSchedule'].IsChecked = $config.schedule.enabled
        $controls['txtTime'].Text = $config.schedule.time

        # Set combo box selections
        foreach ($item in $controls['cboFrequency'].Items) {
            if ($item.Content -eq $config.schedule.frequency) {
                $controls['cboFrequency'].SelectedItem = $item
                break
            }
        }
        foreach ($item in $controls['cboDay'].Items) {
            if ($item.Content -eq $config.schedule.dayOfWeek) {
                $controls['cboDay'].SelectedItem = $item
                break
            }
        }

        & $updateStatus "Configuration loaded"
    }

    # Event: SQL Auth radio button changes
    $controls['rbSqlAuth'].Add_Checked({
        $controls['pnlSqlAuth'].Visibility = 'Visible'
    })

    $controls['rbWindowsAuth'].Add_Checked({
        $controls['pnlSqlAuth'].Visibility = 'Collapsed'
    })

    # Event: Specific folder radio button
    $controls['rbSpecificFolder'].Add_Checked({
        $controls['txtFolderFilter'].IsEnabled = $true
        $controls['btnBrowseFolder'].IsEnabled = $true
    })

    $controls['rbEntireTree'].Add_Checked({
        $controls['txtFolderFilter'].IsEnabled = $false
        $controls['btnBrowseFolder'].IsEnabled = $false
    })

    # Event: Enable schedule checkbox
    $controls['chkEnableSchedule'].Add_Checked({
        $controls['pnlScheduleOptions'].IsEnabled = $true
        $controls['btnCreateTask'].IsEnabled = $true
        $controls['btnRemoveTask'].IsEnabled = $true
    })

    $controls['chkEnableSchedule'].Add_Unchecked({
        $controls['pnlScheduleOptions'].IsEnabled = $false
        $controls['btnCreateTask'].IsEnabled = $false
        $controls['btnRemoveTask'].IsEnabled = $false
    })

    # Event: Test Connection
    $controls['btnTestConnection'].Add_Click({
        & $updateStatus "Testing connection..."

        $testParams = @{
            Server = $controls['txtServer'].Text
            Database = $controls['cboDatabase'].Text
            UseWindowsAuth = $controls['rbWindowsAuth'].IsChecked
        }

        $result = Test-SqlConnection @testParams

        if ($result.Success) {
            & $updateStatus $result.Message "#16A34A"
        }
        else {
            & $updateStatus $result.Message "#DC2626"
        }
    })

    # Event: Execute Now
    $controls['btnExecute'].Add_Click({
        & $updateStatus "Executing report..."
        $window.Cursor = [System.Windows.Input.Cursors]::Wait

        try {
            $config = Get-ConfigFromControls -Controls $controls
            $result = Invoke-ReportGeneration -Config $config

            if ($result.Success) {
                & $updateStatus $result.Message "#16A34A"
                $controls['txtLastRun'].Text = "Last Run: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
            }
            else {
                & $updateStatus "Failed: $($result.Message)" "#DC2626"
            }
        }
        catch {
            & $updateStatus "Error: $($_.Exception.Message)" "#DC2626"
        }
        finally {
            $window.Cursor = [System.Windows.Input.Cursors]::Arrow
        }
    })

    # Event: Preview HTML
    $controls['btnPreview'].Add_Click({
        & $updateStatus "Generating preview..."

        try {
            $config = Get-ConfigFromControls -Controls $controls
            $result = Invoke-ReportGeneration -Config $config -PreviewOnly

            if ($result.Success) {
                & $updateStatus "Preview opened in browser" "#16A34A"
            }
            else {
                & $updateStatus "Preview failed: $($result.Message)" "#DC2626"
            }
        }
        catch {
            & $updateStatus "Error: $($_.Exception.Message)" "#DC2626"
        }
    })

    # Event: Save Config
    $controls['btnSaveConfig'].Add_Click({
        try {
            $config = Get-ConfigFromControls -Controls $controls
            if (Save-Configuration -Config $config) {
                & $updateStatus "Configuration saved" "#16A34A"
            }
            else {
                & $updateStatus "Failed to save configuration" "#DC2626"
            }
        }
        catch {
            & $updateStatus "Error: $($_.Exception.Message)" "#DC2626"
        }
    })

    # Event: Load Config
    $controls['btnLoadConfig'].Add_Click({
        $dialog = New-Object Microsoft.Win32.OpenFileDialog
        $dialog.Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*"
        $dialog.InitialDirectory = Join-Path -Path $script:ScriptRoot -ChildPath "Config"

        if ($dialog.ShowDialog()) {
            try {
                $config = Get-Configuration -Path $dialog.FileName
                if ($config) {
                    # Repopulate controls
                    $controls['txtServer'].Text = $config.database.server
                    $controls['cboDatabase'].Text = $config.database.database
                    # ... (similar to initial load)
                    & $updateStatus "Configuration loaded from: $($dialog.FileName)" "#16A34A"
                }
            }
            catch {
                & $updateStatus "Failed to load: $($_.Exception.Message)" "#DC2626"
            }
        }
    })

    # Event: Create Task
    $controls['btnCreateTask'].Add_Click({
        try {
            $config = Get-ConfigFromControls -Controls $controls

            # Save config first
            Save-Configuration -Config $config

            $taskResult = Register-DRReportTask `
                -Frequency $config.schedule.frequency `
                -DayOfWeek $config.schedule.dayOfWeek `
                -Time $config.schedule.time `
                -ConfigPath $script:DefaultConfigPath

            if ($taskResult.Success) {
                & $updateStatus "Task created successfully" "#16A34A"
                $controls['txtTaskStatus'].Text = "Next run: $($taskResult.TaskInfo.NextRunTime)"
            }
            else {
                & $updateStatus "Failed: $($taskResult.Message)" "#DC2626"
            }
        }
        catch {
            & $updateStatus "Error: $($_.Exception.Message)" "#DC2626"
        }
    })

    # Event: Remove Task
    $controls['btnRemoveTask'].Add_Click({
        try {
            $result = Unregister-DRReportTask

            if ($result.Success) {
                & $updateStatus $result.Message "#16A34A"
                $controls['txtTaskStatus'].Text = ""
            }
            else {
                & $updateStatus $result.Message "#DC2626"
            }
        }
        catch {
            & $updateStatus "Error: $($_.Exception.Message)" "#DC2626"
        }
    })

    # Check current task status on load
    $taskStatus = Get-DRReportTaskStatus
    if ($taskStatus.Exists) {
        $controls['txtTaskStatus'].Text = "Status: $($taskStatus.State) | Next: $($taskStatus.NextRunTime)"
    }

    # Show window
    $window.ShowDialog() | Out-Null
}

function Get-ConfigFromControls {
    param(
        [Parameter(Mandatory = $true)]
        [hashtable]$Controls
    )

    $config = New-DefaultConfiguration

    # Database
    $config.database.server = $Controls['txtServer'].Text
    $config.database.database = $Controls['cboDatabase'].Text
    $config.database.useWindowsAuth = $Controls['rbWindowsAuth'].IsChecked

    # Report scope
    if ($Controls['rbSpecificFolder'].IsChecked) {
        $config.report.scope = "folder"
        $config.report.folderFilter = $Controls['txtFolderFilter'].Text
    }
    else {
        $config.report.scope = "all"
        $config.report.folderFilter = $null
    }

    $config.report.includeWarnings = $Controls['chkWarnings'].IsChecked
    $config.report.includeSuccess = $Controls['chkSuccess'].IsChecked
    $config.report.includeFullMessage = $Controls['chkIncludeFullMessage'].IsChecked
    $config.report.attachCsv = $Controls['chkAttachCsv'].IsChecked

    # Email
    $toText = $Controls['txtEmailTo'].Text
    if ($toText) {
        $config.email.to = $toText -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    }

    $ccText = $Controls['txtEmailCC'].Text
    if ($ccText) {
        $config.email.cc = $ccText -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    }

    $config.email.subjectTemplate = $Controls['txtEmailSubject'].Text

    # Schedule
    $config.schedule.enabled = $Controls['chkEnableSchedule'].IsChecked
    $config.schedule.frequency = $Controls['cboFrequency'].SelectedItem.Content
    $config.schedule.dayOfWeek = $Controls['cboDay'].SelectedItem.Content
    $config.schedule.time = $Controls['txtTime'].Text

    return $config
}

#endregion

#region Main Execution

try {
    Write-Log "=== DR Report Tool Started ==="
    Write-Log "Mode: $(if ($Silent) { 'Silent' } elseif ($PreviewOnly) { 'Preview' } else { 'GUI' })"

    # Import modules
    Import-RequiredModules

    if ($Silent) {
        # Silent mode for scheduled execution
        if (-not $ConfigPath) {
            $ConfigPath = $script:DefaultConfigPath
        }

        $config = Get-Configuration -Path $ConfigPath
        if (-not $config) {
            throw "Configuration file not found: $ConfigPath"
        }

        $result = Invoke-ReportGeneration -Config $config

        if ($result.Success) {
            Write-Log "Report execution completed successfully"
            exit 0
        }
        else {
            Write-Log "Report execution failed: $($result.Message)" -Level Error
            exit 1
        }
    }
    elseif ($PreviewOnly) {
        # Preview mode
        $config = Get-Configuration -Path $ConfigPath
        if (-not $config) {
            $config = New-DefaultConfiguration
        }

        $result = Invoke-ReportGeneration -Config $config -PreviewOnly

        if ($result.Success) {
            Write-Log "Preview generated successfully"
        }
        else {
            Write-Log "Preview failed: $($result.Message)" -Level Error
        }
    }
    else {
        # GUI mode
        Show-MainWindow
    }
}
catch {
    Write-Log "Fatal error: $($_.Exception.Message)" -Level Error

    if (-not $Silent) {
        [System.Windows.MessageBox]::Show(
            "An error occurred: $($_.Exception.Message)",
            "DR Report Tool Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }

    exit 1
}
finally {
    Write-Log "=== DR Report Tool Ended ==="
}

#endregion
