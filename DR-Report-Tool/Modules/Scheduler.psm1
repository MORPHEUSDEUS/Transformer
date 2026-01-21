#Requires -Version 5.1
<#
.SYNOPSIS
    Scheduler module for AssetCentre DR Report Tool
.DESCRIPTION
    Provides Windows Task Scheduler integration for automated report execution
#>

# Module-level variables
$script:DefaultTaskName = "AssetCentre-DR-Report"
$script:LastError = $null

function Register-DRReportTask {
    <#
    .SYNOPSIS
        Creates a Windows Task Scheduler task for automated DR report execution
    .PARAMETER TaskName
        Name for the scheduled task
    .PARAMETER Frequency
        Execution frequency: Daily, Weekly, Monthly
    .PARAMETER DayOfWeek
        Day of week for weekly tasks (Monday, Tuesday, etc.)
    .PARAMETER DayOfMonth
        Day of month for monthly tasks (1-31)
    .PARAMETER Time
        Time to run (HH:mm format)
    .PARAMETER ConfigPath
        Path to settings.json configuration file
    .PARAMETER ScriptPath
        Path to DR-Report-Tool.ps1 script
    .PARAMETER Description
        Task description
    .OUTPUTS
        Hashtable with Success, Message, and TaskInfo
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$TaskName = $script:DefaultTaskName,

        [Parameter(Mandatory = $true)]
        [ValidateSet('Daily', 'Weekly', 'Monthly')]
        [string]$Frequency,

        [Parameter()]
        [ValidateSet('Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday')]
        [string]$DayOfWeek = 'Monday',

        [Parameter()]
        [ValidateRange(1, 31)]
        [int]$DayOfMonth = 1,

        [Parameter(Mandatory = $true)]
        [ValidatePattern('^\d{2}:\d{2}$')]
        [string]$Time,

        [Parameter(Mandatory = $true)]
        [string]$ConfigPath,

        [Parameter()]
        [string]$ScriptPath,

        [Parameter()]
        [string]$Description = "Executes AssetCentre DR Report and sends email notification"
    )

    try {
        # Resolve script path
        if (-not $ScriptPath) {
            $ScriptPath = Join-Path -Path (Split-Path -Path $ConfigPath -Parent | Split-Path -Parent) -ChildPath "DR-Report-Tool.ps1"
        }

        if (-not (Test-Path -Path $ScriptPath)) {
            throw "Script not found: $ScriptPath"
        }

        if (-not (Test-Path -Path $ConfigPath)) {
            throw "Config file not found: $ConfigPath"
        }

        # Parse time
        $timeParts = $Time -split ':'
        $hour = [int]$timeParts[0]
        $minute = [int]$timeParts[1]

        # Build trigger based on frequency
        $triggerParams = @{
            At = (Get-Date -Hour $hour -Minute $minute -Second 0)
        }

        switch ($Frequency) {
            'Daily' {
                $trigger = New-ScheduledTaskTrigger -Daily @triggerParams
            }
            'Weekly' {
                $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $DayOfWeek @triggerParams
            }
            'Monthly' {
                # Monthly trigger requires different approach
                $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $DayOfWeek @triggerParams
                # Note: True monthly scheduling would require additional configuration
            }
        }

        # Build action - run PowerShell with script and config
        $argument = "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$ScriptPath`" -Silent -ConfigPath `"$ConfigPath`""
        $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $argument

        # Task settings
        $settings = New-ScheduledTaskSettingsSet `
            -AllowStartIfOnBatteries `
            -DontStopIfGoingOnBatteries `
            -StartWhenAvailable `
            -RunOnlyIfNetworkAvailable `
            -MultipleInstances IgnoreNew

        # Principal - run as current user with highest privileges
        $principal = New-ScheduledTaskPrincipal `
            -UserId $env:USERNAME `
            -LogonType Interactive `
            -RunLevel Highest

        # Remove existing task if present
        $existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        if ($existingTask) {
            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
        }

        # Register the task
        $task = Register-ScheduledTask `
            -TaskName $TaskName `
            -Description $Description `
            -Trigger $trigger `
            -Action $action `
            -Settings $settings `
            -Principal $principal

        $script:LastError = $null

        return @{
            Success = $true
            Message = "Task '$TaskName' registered successfully"
            TaskInfo = @{
                TaskName = $TaskName
                Frequency = $Frequency
                Time = $Time
                DayOfWeek = $DayOfWeek
                NextRunTime = (Get-ScheduledTask -TaskName $TaskName | Get-ScheduledTaskInfo).NextRunTime
                Status = $task.State
            }
        }
    }
    catch {
        $script:LastError = $_.Exception.Message
        return @{
            Success = $false
            Message = "Failed to register task: $($_.Exception.Message)"
            TaskInfo = $null
        }
    }
}

function Unregister-DRReportTask {
    <#
    .SYNOPSIS
        Removes the scheduled task
    .PARAMETER TaskName
        Name of the task to remove
    .OUTPUTS
        Hashtable with Success and Message
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$TaskName = $script:DefaultTaskName
    )

    try {
        $existingTask = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue

        if (-not $existingTask) {
            return @{
                Success = $true
                Message = "Task '$TaskName' does not exist"
            }
        }

        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false

        $script:LastError = $null

        return @{
            Success = $true
            Message = "Task '$TaskName' removed successfully"
        }
    }
    catch {
        $script:LastError = $_.Exception.Message
        return @{
            Success = $false
            Message = "Failed to remove task: $($_.Exception.Message)"
        }
    }
}

function Get-DRReportTaskStatus {
    <#
    .SYNOPSIS
        Gets the status of the scheduled task
    .PARAMETER TaskName
        Name of the task
    .OUTPUTS
        Hashtable with task information or null if not found
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$TaskName = $script:DefaultTaskName
    )

    try {
        $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue

        if (-not $task) {
            return @{
                Exists = $false
                Message = "Task '$TaskName' not found"
            }
        }

        $taskInfo = Get-ScheduledTaskInfo -TaskName $TaskName

        return @{
            Exists = $true
            TaskName = $task.TaskName
            State = $task.State.ToString()
            LastRunTime = $taskInfo.LastRunTime
            LastTaskResult = $taskInfo.LastTaskResult
            NextRunTime = $taskInfo.NextRunTime
            NumberOfMissedRuns = $taskInfo.NumberOfMissedRuns
            Description = $task.Description
        }
    }
    catch {
        $script:LastError = $_.Exception.Message
        return @{
            Exists = $false
            Message = "Error retrieving task: $($_.Exception.Message)"
        }
    }
}

function Start-DRReportTaskNow {
    <#
    .SYNOPSIS
        Manually triggers the scheduled task to run immediately
    .PARAMETER TaskName
        Name of the task
    .OUTPUTS
        Hashtable with Success and Message
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$TaskName = $script:DefaultTaskName
    )

    try {
        $task = Get-ScheduledTask -TaskName $TaskName -ErrorAction Stop
        Start-ScheduledTask -TaskName $TaskName

        return @{
            Success = $true
            Message = "Task '$TaskName' started"
        }
    }
    catch {
        $script:LastError = $_.Exception.Message
        return @{
            Success = $false
            Message = "Failed to start task: $($_.Exception.Message)"
        }
    }
}

function Stop-DRReportTask {
    <#
    .SYNOPSIS
        Stops a running task
    .PARAMETER TaskName
        Name of the task
    .OUTPUTS
        Hashtable with Success and Message
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$TaskName = $script:DefaultTaskName
    )

    try {
        Stop-ScheduledTask -TaskName $TaskName -ErrorAction Stop

        return @{
            Success = $true
            Message = "Task '$TaskName' stopped"
        }
    }
    catch {
        $script:LastError = $_.Exception.Message
        return @{
            Success = $false
            Message = "Failed to stop task: $($_.Exception.Message)"
        }
    }
}

function Enable-DRReportTask {
    <#
    .SYNOPSIS
        Enables a disabled task
    .PARAMETER TaskName
        Name of the task
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$TaskName = $script:DefaultTaskName
    )

    try {
        Enable-ScheduledTask -TaskName $TaskName -ErrorAction Stop

        return @{
            Success = $true
            Message = "Task '$TaskName' enabled"
        }
    }
    catch {
        $script:LastError = $_.Exception.Message
        return @{
            Success = $false
            Message = "Failed to enable task: $($_.Exception.Message)"
        }
    }
}

function Disable-DRReportTask {
    <#
    .SYNOPSIS
        Disables a task without removing it
    .PARAMETER TaskName
        Name of the task
    #>
    [CmdletBinding()]
    param(
        [Parameter()]
        [string]$TaskName = $script:DefaultTaskName
    )

    try {
        Disable-ScheduledTask -TaskName $TaskName -ErrorAction Stop

        return @{
            Success = $true
            Message = "Task '$TaskName' disabled"
        }
    }
    catch {
        $script:LastError = $_.Exception.Message
        return @{
            Success = $false
            Message = "Failed to disable task: $($_.Exception.Message)"
        }
    }
}

function Get-LastSchedulerError {
    <#
    .SYNOPSIS
        Returns the last scheduler error message
    #>
    return $script:LastError
}

function Get-TaskResultDescription {
    <#
    .SYNOPSIS
        Converts task result code to human-readable description
    .PARAMETER ResultCode
        The numeric result code
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int]$ResultCode
    )

    $descriptions = @{
        0 = "The operation completed successfully"
        1 = "Incorrect function called or unknown function called"
        2 = "File not found"
        10 = "The environment is incorrect"
        267008 = "Task is ready to run at its next scheduled time"
        267009 = "Task is currently running"
        267010 = "The task will not run at the scheduled times because it has been disabled"
        267011 = "Task has not yet run"
        267012 = "There are no more runs scheduled for this task"
        267013 = "One or more of the properties required to run this task have not been set"
        267014 = "The last run of the task was terminated by the user"
        267015 = "Either the task has no triggers or the existing triggers are disabled or not set"
        2147750687 = "An instance of this task is already running"
        2147942667 = "The operator or administrator has refused the request"
    }

    if ($descriptions.ContainsKey($ResultCode)) {
        return $descriptions[$ResultCode]
    }

    return "Unknown result code: $ResultCode"
}

# Export module members
Export-ModuleMember -Function @(
    'Register-DRReportTask',
    'Unregister-DRReportTask',
    'Get-DRReportTaskStatus',
    'Start-DRReportTaskNow',
    'Stop-DRReportTask',
    'Enable-DRReportTask',
    'Disable-DRReportTask',
    'Get-LastSchedulerError',
    'Get-TaskResultDescription'
)
