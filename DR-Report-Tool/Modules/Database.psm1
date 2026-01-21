#Requires -Version 5.1
<#
.SYNOPSIS
    Database module for AssetCentre DR Report Tool
.DESCRIPTION
    Provides SQL Server connectivity and query execution for FactoryTalk AssetCentre database
#>

# Module-level variables
$script:ConnectionString = $null
$script:LastError = $null

function Test-SqlConnection {
    <#
    .SYNOPSIS
        Tests SQL Server connectivity
    .PARAMETER Server
        SQL Server instance name
    .PARAMETER Database
        Database name (default: AssetCentre)
    .PARAMETER UseWindowsAuth
        Use Windows Authentication instead of SQL Authentication
    .PARAMETER Username
        SQL username (if not using Windows Auth)
    .PARAMETER Password
        SQL password (if not using Windows Auth)
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server,

        [Parameter()]
        [string]$Database = "AssetCentre",

        [Parameter()]
        [switch]$UseWindowsAuth,

        [Parameter()]
        [string]$Username,

        [Parameter()]
        [SecureString]$Password
    )

    try {
        $connectionString = Build-ConnectionString -Server $Server -Database $Database -UseWindowsAuth:$UseWindowsAuth -Username $Username -Password $Password

        $connection = New-Object System.Data.SqlClient.SqlConnection
        $connection.ConnectionString = $connectionString
        $connection.Open()
        $connection.Close()

        $script:ConnectionString = $connectionString
        $script:LastError = $null

        return @{
            Success = $true
            Message = "Connection successful to $Server\$Database"
            ServerVersion = $connection.ServerVersion
        }
    }
    catch {
        $script:LastError = $_.Exception.Message
        return @{
            Success = $false
            Message = "Connection failed: $($_.Exception.Message)"
            ServerVersion = $null
        }
    }
}

function Build-ConnectionString {
    <#
    .SYNOPSIS
        Builds SQL Server connection string
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server,

        [Parameter()]
        [string]$Database = "AssetCentre",

        [Parameter()]
        [switch]$UseWindowsAuth,

        [Parameter()]
        [string]$Username,

        [Parameter()]
        [SecureString]$Password
    )

    $builder = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
    $builder["Data Source"] = $Server
    $builder["Initial Catalog"] = $Database
    $builder["Connection Timeout"] = 30

    if ($UseWindowsAuth) {
        $builder["Integrated Security"] = $true
    }
    else {
        $builder["Integrated Security"] = $false
        $builder["User ID"] = $Username
        if ($Password) {
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
            $builder["Password"] = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
        }
    }

    return $builder.ConnectionString
}

function Invoke-AssetCentreQuery {
    <#
    .SYNOPSIS
        Executes the DR Report query against AssetCentre database
    .PARAMETER Server
        SQL Server instance name
    .PARAMETER Database
        Database name (default: AssetCentre)
    .PARAMETER FolderFilter
        Optional folder name to filter results (null = entire tree)
    .PARAMETER UseWindowsAuth
        Use Windows Authentication
    .PARAMETER Username
        SQL username (if not using Windows Auth)
    .PARAMETER Password
        SQL password (if not using Windows Auth)
    .OUTPUTS
        System.Data.DataTable containing query results
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server,

        [Parameter()]
        [string]$Database = "AssetCentre",

        [Parameter()]
        [string]$FolderFilter = $null,

        [Parameter()]
        [switch]$UseWindowsAuth,

        [Parameter()]
        [string]$Username,

        [Parameter()]
        [SecureString]$Password
    )

    $query = Get-AssetCentreQuery

    try {
        $connectionString = Build-ConnectionString -Server $Server -Database $Database -UseWindowsAuth:$UseWindowsAuth -Username $Username -Password $Password

        $connection = New-Object System.Data.SqlClient.SqlConnection
        $connection.ConnectionString = $connectionString

        $command = New-Object System.Data.SqlClient.SqlCommand
        $command.Connection = $connection
        $command.CommandText = $query
        $command.CommandTimeout = 300  # 5 minutes for large datasets

        # Add folder filter parameter
        if ([string]::IsNullOrWhiteSpace($FolderFilter)) {
            $command.Parameters.AddWithValue("@FolderName", [DBNull]::Value) | Out-Null
        }
        else {
            $command.Parameters.AddWithValue("@FolderName", $FolderFilter) | Out-Null
        }

        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $adapter.SelectCommand = $command

        $dataTable = New-Object System.Data.DataTable

        $connection.Open()
        $adapter.Fill($dataTable) | Out-Null
        $connection.Close()

        $script:LastError = $null

        Write-Verbose "Query returned $($dataTable.Rows.Count) rows"
        return $dataTable
    }
    catch {
        $script:LastError = $_.Exception.Message
        Write-Error "Query execution failed: $($_.Exception.Message)"
        return $null
    }
    finally {
        if ($connection -and $connection.State -eq 'Open') {
            $connection.Close()
        }
    }
}

function Get-AssetCentreQuery {
    <#
    .SYNOPSIS
        Returns the embedded SQL query for DR Report
    #>
    return @"
-- AssetCentre DR Report Query
-- If @FolderName is NULL, query entire tree
-- If @FolderName is provided, start from that folder

DECLARE @FolderName NVARCHAR(255) = @FolderName;

-- Folder type GUID
DECLARE @FolderTypeId UNIQUEIDENTIFIER = 'F87F4416-504D-4924-B7B6-3D9A49BE766C';
-- Root AssetCentre node
DECLARE @RootId NVARCHAR(100) = 'AssetCentre:{00000000-0000-0000-0000-000000000000}';

-- Build folder hierarchy
;WITH FolderTree AS (
    -- Base case: Start from root or specified folder
    SELECT
        a.AssetId,
        a.AssetName,
        a.ParentId,
        a.AssetTypeId,
        CAST(a.AssetName AS NVARCHAR(MAX)) AS FolderPath,
        0 AS Level
    FROM dbo.Assets a
    WHERE (@FolderName IS NULL AND a.ParentId = @RootId)
       OR (@FolderName IS NOT NULL AND a.AssetName = @FolderName AND a.AssetTypeId = @FolderTypeId)

    UNION ALL

    -- Recursive case: Get child folders
    SELECT
        a.AssetId,
        a.AssetName,
        a.ParentId,
        a.AssetTypeId,
        CAST(ft.FolderPath + '/' + a.AssetName AS NVARCHAR(MAX)) AS FolderPath,
        ft.Level + 1
    FROM dbo.Assets a
    INNER JOIN FolderTree ft ON a.ParentId = ft.AssetId
),
-- Get all assets with their folder paths
AssetData AS (
    SELECT
        ft.FolderPath,
        a.AssetId,
        a.AssetName,
        a.FileDescription,
        at.AssetTypeName AS AssetType,
        a.AddressingInfo,
        a.CatalogNumber,
        a.HardwareType,
        a.HardwareRevision,
        a.FirmwareRevision,
        a.SerialNumber,
        CASE WHEN bs.ScheduleId IS NOT NULL THEN 'True' ELSE 'False' END AS BackupEnabled,
        bs.ScheduleName,
        bh.ExecutionTime AS LastExecutionTime,
        bh.ExecutionResult AS LastExecutionResult,
        bh.StatusText,
        bh.ErrorMessage,
        bh.ExtendedError,
        CASE WHEN bh.RetryCount > 0 THEN 'Yes' ELSE 'No' END AS HasRetryInfo,
        a.NetworkRoute,
        CASE
            WHEN bh.ExecutionResult = 'Failed' AND bh.ErrorMessage LIKE '%Agent%' THEN 'Agent Failure'
            ELSE 'Scheduled Event'
        END AS MessageType,
        bh.FullMessage
    FROM FolderTree ft
    INNER JOIN dbo.Assets a ON a.ParentId = ft.AssetId
    INNER JOIN dbo.AssetTypes at ON a.AssetTypeId = at.AssetTypeId
    LEFT JOIN dbo.BackupSchedules bs ON a.AssetId = bs.AssetId AND bs.IsActive = 1
    LEFT JOIN (
        SELECT
            AssetId,
            ExecutionTime,
            ExecutionResult,
            StatusText,
            ErrorMessage,
            ExtendedError,
            RetryCount,
            FullMessage,
            ROW_NUMBER() OVER (PARTITION BY AssetId ORDER BY ExecutionTime DESC) AS RowNum
        FROM dbo.BackupHistory
    ) bh ON a.AssetId = bh.AssetId AND bh.RowNum = 1
    WHERE a.AssetTypeId <> @FolderTypeId  -- Exclude folders from results
)

SELECT
    FolderPath,
    AssetName,
    FileDescription,
    AssetType,
    AddressingInfo,
    CatalogNumber,
    HardwareType,
    HardwareRevision,
    FirmwareRevision,
    SerialNumber,
    BackupEnabled,
    ScheduleName,
    LastExecutionTime,
    LastExecutionResult,
    StatusText,
    ErrorMessage,
    ExtendedError,
    HasRetryInfo,
    NetworkRoute,
    MessageType,
    FullMessage
FROM AssetData
ORDER BY
    CASE LastExecutionResult
        WHEN 'Failed' THEN 1
        WHEN 'Warning' THEN 2
        ELSE 3
    END,
    FolderPath,
    AssetName;
"@
}

function Get-LastDatabaseError {
    <#
    .SYNOPSIS
        Returns the last database error message
    #>
    return $script:LastError
}

function Get-SampleData {
    <#
    .SYNOPSIS
        Returns sample data for testing without database connection
    #>
    $dataTable = New-Object System.Data.DataTable

    # Add columns
    $columns = @(
        @{Name="FolderPath"; Type=[string]},
        @{Name="AssetName"; Type=[string]},
        @{Name="FileDescription"; Type=[string]},
        @{Name="AssetType"; Type=[string]},
        @{Name="AddressingInfo"; Type=[string]},
        @{Name="CatalogNumber"; Type=[string]},
        @{Name="HardwareType"; Type=[string]},
        @{Name="HardwareRevision"; Type=[string]},
        @{Name="FirmwareRevision"; Type=[string]},
        @{Name="SerialNumber"; Type=[string]},
        @{Name="BackupEnabled"; Type=[string]},
        @{Name="ScheduleName"; Type=[string]},
        @{Name="LastExecutionTime"; Type=[datetime]},
        @{Name="LastExecutionResult"; Type=[string]},
        @{Name="StatusText"; Type=[string]},
        @{Name="ErrorMessage"; Type=[string]},
        @{Name="ExtendedError"; Type=[string]},
        @{Name="HasRetryInfo"; Type=[string]},
        @{Name="NetworkRoute"; Type=[string]},
        @{Name="MessageType"; Type=[string]},
        @{Name="FullMessage"; Type=[string]}
    )

    foreach ($col in $columns) {
        $dataTable.Columns.Add($col.Name, $col.Type) | Out-Null
    }

    # Add sample rows
    $sampleData = @(
        @{
            FolderPath = "Line 6 - BIB/BIB_Filler 1"
            AssetName = "BIB_Filler1_PLC"
            FileDescription = "Main Filler Controller"
            AssetType = "Logix5000_Controller"
            AddressingInfo = "192.168.1.10, Slot 0"
            CatalogNumber = "1756-L83E"
            HardwareType = "ControlLogix"
            HardwareRevision = "32.011"
            FirmwareRevision = "32.011"
            SerialNumber = "ABC123456"
            BackupEnabled = "True"
            ScheduleName = "Weekly DR Backup"
            LastExecutionTime = (Get-Date).AddDays(-1)
            LastExecutionResult = "Success"
            StatusText = "Backup completed successfully"
            ErrorMessage = ""
            ExtendedError = ""
            HasRetryInfo = "No"
            NetworkRoute = "1\\192.168.1.10\\0"
            MessageType = "Scheduled Event"
            FullMessage = "Backup completed successfully at $(Get-Date)"
        },
        @{
            FolderPath = "Line 6 - BIB/BIB_Filler 1"
            AssetName = "BIB_Filler1_HMI"
            FileDescription = "Operator Interface"
            AssetType = "PanelView_Plus"
            AddressingInfo = "192.168.1.11"
            CatalogNumber = "2711P-T15C22D9P"
            HardwareType = "PanelView Plus 7"
            HardwareRevision = "10.00"
            FirmwareRevision = "10.00"
            SerialNumber = "DEF789012"
            BackupEnabled = "True"
            ScheduleName = "Weekly DR Backup"
            LastExecutionTime = (Get-Date).AddDays(-1)
            LastExecutionResult = "Warning"
            StatusText = "Differences found in project"
            ErrorMessage = "Project differs from last backup"
            ExtendedError = "Screen modifications detected"
            HasRetryInfo = "No"
            NetworkRoute = "1\\192.168.1.11"
            MessageType = "Scheduled Event"
            FullMessage = "Differences found: Screen modifications detected"
        },
        @{
            FolderPath = "Line 6 - BIB/BIB_Capper"
            AssetName = "BIB_Capper_PLC"
            FileDescription = "Capping Machine Controller"
            AssetType = "Logix5000_Controller"
            AddressingInfo = "192.168.1.20, Slot 0"
            CatalogNumber = "1756-L73"
            HardwareType = "ControlLogix"
            HardwareRevision = "30.011"
            FirmwareRevision = "30.011"
            SerialNumber = "GHI345678"
            BackupEnabled = "True"
            ScheduleName = "Weekly DR Backup"
            LastExecutionTime = (Get-Date).AddDays(-1)
            LastExecutionResult = "Failed"
            StatusText = "Connection timeout"
            ErrorMessage = "Failed to establish connection"
            ExtendedError = "Network timeout after 30 seconds"
            HasRetryInfo = "Yes"
            NetworkRoute = "1\\192.168.1.20\\0"
            MessageType = "Agent Failure"
            FullMessage = "Agent failure: Network timeout after 30 seconds - Failed to establish connection"
        },
        @{
            FolderPath = "Line 7 - Packaging"
            AssetName = "Pkg_Wrapper_PLC"
            FileDescription = "Wrapper Controller"
            AssetType = "SLC_500"
            AddressingInfo = "192.168.2.10"
            CatalogNumber = "1747-L552"
            HardwareType = "SLC 5/05"
            HardwareRevision = "C"
            FirmwareRevision = "OS502"
            SerialNumber = "JKL901234"
            BackupEnabled = "True"
            ScheduleName = "Weekly DR Backup"
            LastExecutionTime = (Get-Date).AddDays(-1)
            LastExecutionResult = "Success"
            StatusText = "Backup completed successfully"
            ErrorMessage = ""
            ExtendedError = ""
            HasRetryInfo = "No"
            NetworkRoute = "1\\192.168.2.10"
            MessageType = "Scheduled Event"
            FullMessage = "Backup completed successfully"
        },
        @{
            FolderPath = "Line 7 - Packaging"
            AssetName = "Pkg_Palletizer_PLC"
            FileDescription = "Palletizer Controller"
            AssetType = "Logix5000_Controller"
            AddressingInfo = "192.168.2.20, Slot 0"
            CatalogNumber = "1756-L85E"
            HardwareType = "ControlLogix"
            HardwareRevision = "33.011"
            FirmwareRevision = "33.011"
            SerialNumber = "MNO567890"
            BackupEnabled = "True"
            ScheduleName = "Weekly DR Backup"
            LastExecutionTime = (Get-Date).AddDays(-1)
            LastExecutionResult = "Failed"
            StatusText = "Authentication failed"
            ErrorMessage = "Invalid credentials"
            ExtendedError = "Username or password incorrect for device"
            HasRetryInfo = "Yes"
            NetworkRoute = "1\\192.168.2.20\\0"
            MessageType = "Scheduled Event"
            FullMessage = "Authentication failed: Invalid credentials - Username or password incorrect"
        }
    )

    foreach ($row in $sampleData) {
        $dataRow = $dataTable.NewRow()
        foreach ($key in $row.Keys) {
            $dataRow[$key] = $row[$key]
        }
        $dataTable.Rows.Add($dataRow)
    }

    return $dataTable
}

# Export module members
Export-ModuleMember -Function @(
    'Test-SqlConnection',
    'Invoke-AssetCentreQuery',
    'Get-LastDatabaseError',
    'Get-SampleData',
    'Get-AssetCentreQuery'
)
