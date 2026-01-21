#Requires -Version 5.1
<#
.SYNOPSIS
    Email Sender module for AssetCentre DR Report Tool
.DESCRIPTION
    Provides Outlook COM integration for sending HTML email reports
    All email stays on-premise through Outlook client
#>

# Module-level variables
$script:OutlookApp = $null
$script:LastError = $null

function Test-OutlookAvailable {
    <#
    .SYNOPSIS
        Tests if Outlook is installed and available
    .OUTPUTS
        Hashtable with Success and Message
    #>
    [CmdletBinding()]
    param()

    try {
        $outlook = New-Object -ComObject Outlook.Application
        $namespace = $outlook.GetNamespace("MAPI")
        $defaultFolder = $namespace.GetDefaultFolder(6)  # olFolderInbox

        # Release COM objects
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($namespace) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null

        $script:LastError = $null

        return @{
            Success = $true
            Message = "Outlook is available and configured"
        }
    }
    catch {
        $script:LastError = $_.Exception.Message
        return @{
            Success = $false
            Message = "Outlook not available: $($_.Exception.Message)"
        }
    }
}

function Send-OutlookReport {
    <#
    .SYNOPSIS
        Sends HTML report via Outlook
    .PARAMETER To
        Recipient email addresses (string or array)
    .PARAMETER CC
        CC recipients (string or array)
    .PARAMETER Subject
        Email subject line
    .PARAMETER HtmlBody
        HTML content for email body
    .PARAMETER AttachmentPath
        Optional path to file attachment (e.g., CSV export)
    .PARAMETER Importance
        Email importance: Low, Normal, High
    .PARAMETER SaveToSentItems
        Save a copy to Sent Items folder
    .OUTPUTS
        Hashtable with Success, Message, and SentTime
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$To,

        [Parameter()]
        [string[]]$CC,

        [Parameter(Mandatory = $true)]
        [string]$Subject,

        [Parameter(Mandatory = $true)]
        [string]$HtmlBody,

        [Parameter()]
        [string]$AttachmentPath,

        [Parameter()]
        [ValidateSet('Low', 'Normal', 'High')]
        [string]$Importance = 'Normal',

        [Parameter()]
        [switch]$SaveToSentItems = $true
    )

    $outlook = $null
    $mail = $null

    try {
        # Create Outlook COM object
        $outlook = New-Object -ComObject Outlook.Application

        # Create new mail item (0 = olMailItem)
        $mail = $outlook.CreateItem(0)

        # Set recipients
        $mail.To = ($To -join '; ')

        if ($CC -and $CC.Count -gt 0) {
            $mail.CC = ($CC -join '; ')
        }

        # Set subject and body
        $mail.Subject = $Subject
        $mail.HTMLBody = $HtmlBody

        # Set importance
        switch ($Importance) {
            'Low'    { $mail.Importance = 0 }  # olImportanceLow
            'Normal' { $mail.Importance = 1 }  # olImportanceNormal
            'High'   { $mail.Importance = 2 }  # olImportanceHigh
        }

        # Add attachment if specified
        if ($AttachmentPath -and (Test-Path -Path $AttachmentPath)) {
            $mail.Attachments.Add($AttachmentPath) | Out-Null
            Write-Verbose "Attached file: $AttachmentPath"
        }
        elseif ($AttachmentPath) {
            Write-Warning "Attachment file not found: $AttachmentPath"
        }

        # Send the email
        $mail.Send()

        $sentTime = Get-Date

        $script:LastError = $null

        return @{
            Success = $true
            Message = "Email sent successfully to $($To -join ', ')"
            SentTime = $sentTime
            Recipients = $To.Count + $(if ($CC) { $CC.Count } else { 0 })
        }
    }
    catch {
        $script:LastError = $_.Exception.Message
        return @{
            Success = $false
            Message = "Failed to send email: $($_.Exception.Message)"
            SentTime = $null
            Recipients = 0
        }
    }
    finally {
        # Clean up COM objects
        if ($mail) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($mail) | Out-Null
        }
        if ($outlook) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function Send-ErrorNotification {
    <#
    .SYNOPSIS
        Sends error notification email
    .PARAMETER To
        Recipient email addresses
    .PARAMETER ErrorMessage
        Error message to include
    .PARAMETER ErrorDetails
        Additional error details
    .PARAMETER TaskName
        Name of the failed task
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$To,

        [Parameter(Mandatory = $true)]
        [string]$ErrorMessage,

        [Parameter()]
        [string]$ErrorDetails,

        [Parameter()]
        [string]$TaskName = "DR Report Generation"
    )

    $htmlBody = @"
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Segoe UI', Arial, sans-serif; background: #F9F9F9; margin: 0; padding: 20px; }
    .container { max-width: 600px; margin: 0 auto; background: #FFFFFF; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
    .header { background: #DC2626; color: #FFFFFF; padding: 24px; border-radius: 8px 8px 0 0; }
    .header h1 { margin: 0; font-size: 20px; font-weight: 600; }
    .content { padding: 24px; }
    .error-box { background: #FEF2F2; border: 1px solid #DC2626; border-radius: 6px; padding: 16px; margin-bottom: 16px; }
    .error-title { color: #DC2626; font-weight: 600; margin-bottom: 8px; }
    .error-message { color: #1A1A1A; font-family: monospace; white-space: pre-wrap; }
    .details { background: #F9F9F9; border-radius: 6px; padding: 16px; }
    .details-title { color: #6B6B6B; font-size: 12px; text-transform: uppercase; margin-bottom: 8px; }
    .footer { background: #F9F9F9; padding: 16px; text-align: center; color: #6B6B6B; font-size: 12px; border-radius: 0 0 8px 8px; }
  </style>
</head>
<body>
  <div class="container">
    <div class="header">
      <h1>Error: $TaskName Failed</h1>
    </div>
    <div class="content">
      <div class="error-box">
        <div class="error-title">Error Message</div>
        <div class="error-message">$([System.Web.HttpUtility]::HtmlEncode($ErrorMessage))</div>
      </div>
      $(if ($ErrorDetails) {
        @"
      <div class="details">
        <div class="details-title">Additional Details</div>
        <div class="error-message">$([System.Web.HttpUtility]::HtmlEncode($ErrorDetails))</div>
      </div>
"@
      })
      <p style="color: #6B6B6B; margin-top: 16px;">
        <strong>Time:</strong> $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")<br>
        <strong>Server:</strong> $env:COMPUTERNAME
      </p>
    </div>
    <div class="footer">
      AssetCentre DR Report Tool - Error Notification
    </div>
  </div>
</body>
</html>
"@

    $subject = "ERROR: $TaskName Failed - $(Get-Date -Format 'yyyy-MM-dd HH:mm')"

    return Send-OutlookReport -To $To -Subject $subject -HtmlBody $htmlBody -Importance 'High'
}

function Get-OutlookProfiles {
    <#
    .SYNOPSIS
        Gets available Outlook profiles
    .OUTPUTS
        Array of profile names
    #>
    [CmdletBinding()]
    param()

    try {
        $profilesPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Profiles"

        if (-not (Test-Path $profilesPath)) {
            # Try Office 2013 path
            $profilesPath = "HKCU:\Software\Microsoft\Office\15.0\Outlook\Profiles"
        }

        if (Test-Path $profilesPath) {
            $profiles = Get-ChildItem -Path $profilesPath | Select-Object -ExpandProperty PSChildName
            return $profiles
        }

        return @()
    }
    catch {
        Write-Warning "Could not retrieve Outlook profiles: $($_.Exception.Message)"
        return @()
    }
}

function Get-LastEmailError {
    <#
    .SYNOPSIS
        Returns the last email error message
    #>
    return $script:LastError
}

function New-DraftEmail {
    <#
    .SYNOPSIS
        Creates a draft email without sending (for preview)
    .PARAMETER To
        Recipient email addresses
    .PARAMETER CC
        CC recipients
    .PARAMETER Subject
        Email subject line
    .PARAMETER HtmlBody
        HTML content for email body
    .PARAMETER AttachmentPath
        Optional attachment path
    .OUTPUTS
        Hashtable with Success and Message
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$To,

        [Parameter()]
        [string[]]$CC,

        [Parameter(Mandatory = $true)]
        [string]$Subject,

        [Parameter(Mandatory = $true)]
        [string]$HtmlBody,

        [Parameter()]
        [string]$AttachmentPath
    )

    $outlook = $null
    $mail = $null

    try {
        $outlook = New-Object -ComObject Outlook.Application
        $mail = $outlook.CreateItem(0)

        $mail.To = ($To -join '; ')
        if ($CC -and $CC.Count -gt 0) {
            $mail.CC = ($CC -join '; ')
        }
        $mail.Subject = $Subject
        $mail.HTMLBody = $HtmlBody

        if ($AttachmentPath -and (Test-Path -Path $AttachmentPath)) {
            $mail.Attachments.Add($AttachmentPath) | Out-Null
        }

        # Save as draft instead of sending
        $mail.Save()

        # Display the email for review
        $mail.Display()

        return @{
            Success = $true
            Message = "Draft email created and displayed"
        }
    }
    catch {
        return @{
            Success = $false
            Message = "Failed to create draft: $($_.Exception.Message)"
        }
    }
    finally {
        if ($outlook) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# Export module members
Export-ModuleMember -Function @(
    'Test-OutlookAvailable',
    'Send-OutlookReport',
    'Send-ErrorNotification',
    'Get-OutlookProfiles',
    'Get-LastEmailError',
    'New-DraftEmail'
)
