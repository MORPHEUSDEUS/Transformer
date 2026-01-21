#Requires -Version 5.1
<#
.SYNOPSIS
    HTML Generator module for AssetCentre DR Report Tool
.DESCRIPTION
    Generates professional HTML email reports from AssetCentre backup data
#>

function New-DRReport {
    <#
    .SYNOPSIS
        Generates an HTML DR Report from data
    .PARAMETER Data
        DataTable containing backup results
    .PARAMETER TemplatePath
        Path to HTML template file (optional, uses embedded template if not specified)
    .PARAMETER FailedOnly
        Only include failed items in the report
    .PARAMETER IncludeWarnings
        Include warning items in the report
    .PARAMETER IncludeSuccess
        Include successful items in the report (detailed listing)
    .PARAMETER IncludeFullMessage
        Include full message details in the report
    .PARAMETER ScheduleName
        Name of the schedule for the report header
    .OUTPUTS
        String containing HTML report
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.DataTable]$Data,

        [Parameter()]
        [string]$TemplatePath,

        [Parameter()]
        [switch]$FailedOnly,

        [Parameter()]
        [switch]$IncludeWarnings,

        [Parameter()]
        [switch]$IncludeSuccess,

        [Parameter()]
        [switch]$IncludeFullMessage,

        [Parameter()]
        [string]$ScheduleName = "DR Backup"
    )

    # Calculate summary counts
    $successCount = ($Data | Where-Object { $_.LastExecutionResult -eq 'Success' }).Count
    $warningCount = ($Data | Where-Object { $_.LastExecutionResult -eq 'Warning' }).Count
    $failedCount = ($Data | Where-Object { $_.LastExecutionResult -eq 'Failed' }).Count

    # Group data by folder and status
    $failedByFolder = $Data | Where-Object { $_.LastExecutionResult -eq 'Failed' } | Group-Object -Property FolderPath
    $warningsByFolder = $Data | Where-Object { $_.LastExecutionResult -eq 'Warning' } | Group-Object -Property FolderPath
    $successByFolder = $Data | Where-Object { $_.LastExecutionResult -eq 'Success' } | Group-Object -Property FolderPath

    # Build HTML
    $html = Get-HtmlTemplate

    # Replace placeholders
    $html = $html -replace '{{GeneratedDate}}', (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
    $html = $html -replace '{{ScheduleName}}', $ScheduleName
    $html = $html -replace '{{SuccessCount}}', $successCount
    $html = $html -replace '{{WarningCount}}', $warningCount
    $html = $html -replace '{{FailedCount}}', $failedCount

    # Build Failed Section
    if ($failedCount -gt 0) {
        $failedHtml = Build-StatusSection -GroupedData $failedByFolder -StatusClass "failed" -SectionTitle "Failed Assets" -SectionIcon "X" -IncludeFullMessage:$IncludeFullMessage
        $html = $html -replace '{{FailedSection}}', $failedHtml
    }
    else {
        $html = $html -replace '{{FailedSection}}', ''
    }

    # Build Warnings Section
    if ($warningCount -gt 0 -and (-not $FailedOnly -or $IncludeWarnings)) {
        $warningsHtml = Build-StatusSection -GroupedData $warningsByFolder -StatusClass "warning" -SectionTitle "Differences Found" -SectionIcon "!" -IncludeFullMessage:$IncludeFullMessage
        $html = $html -replace '{{WarningsSection}}', $warningsHtml
    }
    else {
        $html = $html -replace '{{WarningsSection}}', ''
    }

    # Build Success Section
    if ($successCount -gt 0 -and $IncludeSuccess -and -not $FailedOnly) {
        $successHtml = Build-StatusSection -GroupedData $successByFolder -StatusClass "success" -SectionTitle "Successful ($successCount assets)" -SectionIcon "OK" -IncludeFullMessage:$IncludeFullMessage
        $html = $html -replace '{{SuccessSection}}', $successHtml
    }
    elseif ($successCount -gt 0) {
        $successSummary = @"
    <div class="section">
      <div class="section-title">OK Successful ($successCount assets)</div>
      <p style="color: #6B6B6B;">All assets backed up successfully. Details available in attached CSV.</p>
    </div>
"@
        $html = $html -replace '{{SuccessSection}}', $successSummary
    }
    else {
        $html = $html -replace '{{SuccessSection}}', ''
    }

    return $html
}

function Build-StatusSection {
    <#
    .SYNOPSIS
        Builds an HTML section for a specific status group
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        $GroupedData,

        [Parameter(Mandatory = $true)]
        [string]$StatusClass,

        [Parameter(Mandatory = $true)]
        [string]$SectionTitle,

        [Parameter(Mandatory = $true)]
        [string]$SectionIcon,

        [Parameter()]
        [switch]$IncludeFullMessage
    )

    $sectionHtml = @"
    <div class="section">
      <div class="section-title">$SectionIcon $SectionTitle</div>
"@

    foreach ($folder in $GroupedData) {
        $folderName = if ($folder.Name) { $folder.Name } else { "Root" }
        $sectionHtml += @"
      <div class="folder-group">
        <div class="folder-name">$folderName</div>
        <table>
          <tr>
            <th>Asset</th>
            <th>Type</th>
            <th>Address</th>
            <th>Status</th>
            <th>Last Run</th>
          </tr>
"@

        foreach ($asset in $folder.Group) {
            $errorDiv = ""
            if ($asset.ErrorMessage -and $asset.ErrorMessage.ToString().Trim()) {
                $errorMsg = [System.Web.HttpUtility]::HtmlEncode($asset.ErrorMessage.ToString())
                $errorDiv = "<div class=`"error-text`">$errorMsg</div>"
            }

            $fullMessageDiv = ""
            if ($IncludeFullMessage -and $asset.FullMessage -and $asset.FullMessage.ToString().Trim()) {
                $fullMsg = [System.Web.HttpUtility]::HtmlEncode($asset.FullMessage.ToString())
                $fullMessageDiv = "<div class=`"full-message`">$fullMsg</div>"
            }

            $assetName = [System.Web.HttpUtility]::HtmlEncode($asset.AssetName.ToString())
            $assetType = [System.Web.HttpUtility]::HtmlEncode($asset.AssetType.ToString())
            $addressingInfo = [System.Web.HttpUtility]::HtmlEncode($asset.AddressingInfo.ToString())
            $statusText = [System.Web.HttpUtility]::HtmlEncode($asset.StatusText.ToString())
            $lastExecTime = if ($asset.LastExecutionTime) { $asset.LastExecutionTime.ToString("yyyy-MM-dd HH:mm") } else { "N/A" }

            $sectionHtml += @"
          <tr>
            <td>
              <strong>$assetName</strong>
              $errorDiv
              $fullMessageDiv
            </td>
            <td class="asset-type">$assetType</td>
            <td>$addressingInfo</td>
            <td><span class="status-badge status-$StatusClass">$statusText</span></td>
            <td>$lastExecTime</td>
          </tr>
"@
        }

        $sectionHtml += @"
        </table>
      </div>
"@
    }

    $sectionHtml += @"
    </div>
"@

    return $sectionHtml
}

function Get-HtmlTemplate {
    <#
    .SYNOPSIS
        Returns the embedded HTML template
    #>
    return @"
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>AssetCentre DR Report</title>
  <style>
    body {
      font-family: 'Segoe UI', Arial, sans-serif;
      background: #F9F9F9;
      margin: 0;
      padding: 20px;
      line-height: 1.5;
    }
    .container {
      max-width: 900px;
      margin: 0 auto;
      background: #FFFFFF;
      border-radius: 8px;
      box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .header {
      background: #2D2D2D;
      color: #FFFFFF;
      padding: 24px;
      border-radius: 8px 8px 0 0;
    }
    .header h1 {
      margin: 0;
      font-size: 24px;
      font-weight: 600;
    }
    .header .subtitle {
      color: #A0A0A0;
      margin-top: 4px;
      font-size: 14px;
    }
    .summary {
      display: flex;
      gap: 16px;
      padding: 20px;
      border-bottom: 1px solid #E5E5E5;
    }
    .summary-card {
      flex: 1;
      padding: 16px;
      border-radius: 6px;
      text-align: center;
    }
    .summary-card.success {
      background: #F0FDF4;
      border: 1px solid #16A34A;
    }
    .summary-card.warning {
      background: #FFFBEB;
      border: 1px solid #D97706;
    }
    .summary-card.failed {
      background: #FEF2F2;
      border: 1px solid #DC2626;
    }
    .summary-card .count {
      font-size: 32px;
      font-weight: 700;
    }
    .summary-card.success .count {
      color: #16A34A;
    }
    .summary-card.warning .count {
      color: #D97706;
    }
    .summary-card.failed .count {
      color: #DC2626;
    }
    .summary-card .label {
      color: #6B6B6B;
      font-size: 14px;
    }
    .section {
      padding: 20px;
    }
    .section-title {
      font-size: 18px;
      font-weight: 600;
      color: #1A1A1A;
      margin-bottom: 16px;
      border-bottom: 2px solid #D97757;
      padding-bottom: 8px;
    }
    .folder-group {
      margin-bottom: 24px;
    }
    .folder-name {
      font-weight: 600;
      color: #1A1A1A;
      background: #F9F9F9;
      padding: 8px 12px;
      border-radius: 4px;
      margin-bottom: 8px;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 13px;
    }
    th {
      background: #F9F9F9;
      text-align: left;
      padding: 10px;
      border-bottom: 2px solid #E5E5E5;
      color: #6B6B6B;
      font-weight: 600;
    }
    td {
      padding: 10px;
      border-bottom: 1px solid #E5E5E5;
      color: #1A1A1A;
      vertical-align: top;
    }
    tr:hover {
      background: #FAFAFA;
    }
    .status-badge {
      display: inline-block;
      padding: 4px 8px;
      border-radius: 4px;
      font-size: 12px;
      font-weight: 500;
      white-space: nowrap;
    }
    .status-success {
      background: #F0FDF4;
      color: #16A34A;
    }
    .status-warning {
      background: #FFFBEB;
      color: #D97706;
    }
    .status-failed {
      background: #FEF2F2;
      color: #DC2626;
    }
    .asset-type {
      color: #6B6B6B;
      font-size: 11px;
    }
    .error-text {
      color: #DC2626;
      font-size: 12px;
      margin-top: 4px;
    }
    .full-message {
      color: #6B6B6B;
      font-size: 11px;
      margin-top: 4px;
      font-style: italic;
      max-width: 300px;
      word-wrap: break-word;
    }
    .footer {
      background: #F9F9F9;
      padding: 16px;
      text-align: center;
      color: #6B6B6B;
      font-size: 12px;
      border-radius: 0 0 8px 8px;
      border-top: 1px solid #E5E5E5;
    }
    .footer a {
      color: #D97757;
      text-decoration: none;
    }
    .footer a:hover {
      text-decoration: underline;
    }
    /* Email client compatibility */
    @media screen and (max-width: 600px) {
      .summary {
        flex-direction: column;
      }
      .summary-card {
        margin-bottom: 10px;
      }
      table {
        font-size: 12px;
      }
      th, td {
        padding: 8px 6px;
      }
    }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>AssetCentre DR Report</h1>
      <div class="subtitle">Generated: {{GeneratedDate}} | Schedule: {{ScheduleName}}</div>
    </div>

    <div class="summary">
      <div class="summary-card success">
        <div class="count">{{SuccessCount}}</div>
        <div class="label">Success</div>
      </div>
      <div class="summary-card warning">
        <div class="count">{{WarningCount}}</div>
        <div class="label">Differences</div>
      </div>
      <div class="summary-card failed">
        <div class="count">{{FailedCount}}</div>
        <div class="label">Failed</div>
      </div>
    </div>

    {{FailedSection}}

    {{WarningsSection}}

    {{SuccessSection}}

    <div class="footer">
      Report generated by AssetCentre DR Report Tool | <a href="mailto:support@company.com">Support</a>
    </div>
  </div>
</body>
</html>
"@
}

function Export-DataToCsv {
    <#
    .SYNOPSIS
        Exports DataTable to CSV file
    .PARAMETER Data
        DataTable to export
    .PARAMETER OutputPath
        Path for output CSV file
    .OUTPUTS
        Path to created CSV file
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Data.DataTable]$Data,

        [Parameter(Mandatory = $true)]
        [string]$OutputPath
    )

    try {
        # Ensure directory exists
        $directory = Split-Path -Path $OutputPath -Parent
        if (-not (Test-Path -Path $directory)) {
            New-Item -ItemType Directory -Path $directory -Force | Out-Null
        }

        # Export to CSV
        $Data | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

        return $OutputPath
    }
    catch {
        Write-Error "Failed to export CSV: $($_.Exception.Message)"
        return $null
    }
}

function Get-ReportSubject {
    <#
    .SYNOPSIS
        Generates email subject line from template
    .PARAMETER SubjectTemplate
        Template string with placeholders
    .PARAMETER Data
        DataTable for calculating counts
    .OUTPUTS
        Formatted subject string
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SubjectTemplate,

        [Parameter(Mandatory = $true)]
        [System.Data.DataTable]$Data
    )

    $successCount = ($Data | Where-Object { $_.LastExecutionResult -eq 'Success' }).Count
    $warningCount = ($Data | Where-Object { $_.LastExecutionResult -eq 'Warning' }).Count
    $failedCount = ($Data | Where-Object { $_.LastExecutionResult -eq 'Failed' }).Count

    $subject = $SubjectTemplate
    $subject = $subject -replace '\{Date\}', (Get-Date -Format "yyyy-MM-dd")
    $subject = $subject -replace '\{DateTime\}', (Get-Date -Format "yyyy-MM-dd HH:mm")
    $subject = $subject -replace '\{SuccessCount\}', $successCount
    $subject = $subject -replace '\{WarningCount\}', $warningCount
    $subject = $subject -replace '\{FailedCount\}', $failedCount
    $subject = $subject -replace '\{TotalCount\}', ($successCount + $warningCount + $failedCount)

    # Add summary indicator
    if ($failedCount -gt 0) {
        $subject = $subject -replace '\{Summary\}', "FAILED ($failedCount)"
    }
    elseif ($warningCount -gt 0) {
        $subject = $subject -replace '\{Summary\}', "WARNINGS ($warningCount)"
    }
    else {
        $subject = $subject -replace '\{Summary\}', "ALL OK"
    }

    return $subject
}

# Export module members
Export-ModuleMember -Function @(
    'New-DRReport',
    'Export-DataToCsv',
    'Get-ReportSubject',
    'Get-HtmlTemplate'
)
