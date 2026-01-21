# AssetCentre DR Report Automation Tool

## Project Overview

Build a PowerShell application with XAML GUI that:
1. Executes SQL queries against FactoryTalk AssetCentre database
2. Generates professional HTML email reports
3. Sends via Outlook (on-premise, no external services)
4. Supports scheduled execution via Windows Task Scheduler

**Critical constraint:** All data stays within the tenant. No cloud services, no external APIs.

---

## Data Input Format

The tool receives CSV/DataTable output from SQL Server with these columns:

| Column | Type | Description |
|--------|------|-------------|
| FolderPath | string | Hierarchical path: "Line 6 - BIB/BIB_Filler 1" |
| AssetName | string | Device name |
| FileDescription | string | Asset description |
| AssetType | string | Logix5000_Controller, PanelView_Plus, SLC_500, etc. |
| AddressingInfo | string | IP/slot path |
| CatalogNumber | string | Device catalog number |
| HardwareType | string | Hardware type |
| HardwareRevision | string | Hardware revision |
| FirmwareRevision | string | Firmware version |
| SerialNumber | string | Device serial |
| BackupEnabled | string | True/False |
| ScheduleName | string | DR schedule name |
| LastExecutionTime | datetime | When last executed |
| LastExecutionResult | string | Failed/Warning/Success |
| StatusText | string | Status message (matches UI) |
| ErrorMessage | string | Error details |
| ExtendedError | string | Root cause details |
| HasRetryInfo | string | Yes/No |
| NetworkRoute | string | Network route |
| MessageType | string | Agent Failure/Scheduled Event |
| FullMessage | string | Complete raw log |

---

## Color Palette (Claude Light Theme)

```
Primary Background:  #FFFFFF
Secondary Background: #F9F9F9
Card Background:      #FFFFFF
Border:               #E5E5E5
Text Primary:         #1A1A1A
Text Secondary:       #6B6B6B
Success:              #16A34A (green)
Warning:              #D97706 (amber)
Failed:               #DC2626 (red)
Accent/Links:         #D97757 (claude terracotta)
Header Background:    #2D2D2D
Header Text:          #FFFFFF
```

---

## XAML GUI Requirements

### Main Window Layout

```
┌─────────────────────────────────────────────────────────┐
│  AssetCentre DR Report Tool                    [_][□][X]│
├─────────────────────────────────────────────────────────┤
│ ┌─────────────────────────────────────────────────────┐ │
│ │ DATABASE CONNECTION                                 │ │
│ │ Server: [________________] Database: [AssetCentre▼] │ │
│ │ Auth: ○ Windows  ○ SQL    [Test Connection]         │ │
│ └─────────────────────────────────────────────────────┘ │
│ ┌─────────────────────────────────────────────────────┐ │
│ │ REPORT SCOPE                                        │ │
│ │ ○ Entire AssetCentre Tree                           │ │
│ │ ○ Specific Folder: [____________________] [Browse]  │ │
│ │ Filter: □ Failed Only  □ Warnings  □ Success        │ │
│ └─────────────────────────────────────────────────────┘ │
│ ┌─────────────────────────────────────────────────────┐ │
│ │ EMAIL CONFIGURATION                                 │ │
│ │ To:  [________________________________________]     │ │
│ │ CC:  [________________________________________]     │ │
│ │ Subject: [DR Report - {Date} - {Summary}     ]     │ │
│ │ □ Include full message details                      │ │
│ │ □ Attach CSV export                                 │ │
│ └─────────────────────────────────────────────────────┘ │
│ ┌─────────────────────────────────────────────────────┐ │
│ │ SCHEDULING                                          │ │
│ │ □ Enable scheduled execution                        │ │
│ │ Frequency: [Weekly▼]  Day: [Monday▼]  Time: [07:00] │ │
│ │ [Create Task]  [Remove Task]                        │ │
│ └─────────────────────────────────────────────────────┘ │
│                                                         │
│ ┌─────────────────────────────────────────────────────┐ │
│ │ [Execute Now]  [Preview HTML]  [Save Config]        │ │
│ └─────────────────────────────────────────────────────┘ │
│ ┌─────────────────────────────────────────────────────┐ │
│ │ Status: Ready                                       │ │
│ │ Last Run: 2026-01-21 07:00:00 | Sent to 3 recipients│ │
│ └─────────────────────────────────────────────────────┘ │
└─────────────────────────────────────────────────────────┘
```

---

## HTML Email Template Structure

```html
<!DOCTYPE html>
<html>
<head>
  <style>
    body { font-family: 'Segoe UI', Arial, sans-serif; background: #F9F9F9; margin: 0; padding: 20px; }
    .container { max-width: 900px; margin: 0 auto; background: #FFFFFF; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
    .header { background: #2D2D2D; color: #FFFFFF; padding: 24px; border-radius: 8px 8px 0 0; }
    .header h1 { margin: 0; font-size: 24px; font-weight: 600; }
    .header .subtitle { color: #A0A0A0; margin-top: 4px; }
    .summary { display: flex; gap: 16px; padding: 20px; border-bottom: 1px solid #E5E5E5; }
    .summary-card { flex: 1; padding: 16px; border-radius: 6px; text-align: center; }
    .summary-card.success { background: #F0FDF4; border: 1px solid #16A34A; }
    .summary-card.warning { background: #FFFBEB; border: 1px solid #D97706; }
    .summary-card.failed { background: #FEF2F2; border: 1px solid #DC2626; }
    .summary-card .count { font-size: 32px; font-weight: 700; }
    .summary-card.success .count { color: #16A34A; }
    .summary-card.warning .count { color: #D97706; }
    .summary-card.failed .count { color: #DC2626; }
    .summary-card .label { color: #6B6B6B; font-size: 14px; }
    .section { padding: 20px; }
    .section-title { font-size: 18px; font-weight: 600; color: #1A1A1A; margin-bottom: 16px; border-bottom: 2px solid #D97757; padding-bottom: 8px; }
    .folder-group { margin-bottom: 24px; }
    .folder-name { font-weight: 600; color: #1A1A1A; background: #F9F9F9; padding: 8px 12px; border-radius: 4px; margin-bottom: 8px; }
    table { width: 100%; border-collapse: collapse; font-size: 13px; }
    th { background: #F9F9F9; text-align: left; padding: 10px; border-bottom: 2px solid #E5E5E5; color: #6B6B6B; font-weight: 600; }
    td { padding: 10px; border-bottom: 1px solid #E5E5E5; color: #1A1A1A; }
    tr:hover { background: #FAFAFA; }
    .status-badge { display: inline-block; padding: 4px 8px; border-radius: 4px; font-size: 12px; font-weight: 500; }
    .status-success { background: #F0FDF4; color: #16A34A; }
    .status-warning { background: #FFFBEB; color: #D97706; }
    .status-failed { background: #FEF2F2; color: #DC2626; }
    .asset-type { color: #6B6B6B; font-size: 11px; }
    .error-text { color: #DC2626; font-size: 12px; margin-top: 4px; }
    .footer { background: #F9F9F9; padding: 16px; text-align: center; color: #6B6B6B; font-size: 12px; border-radius: 0 0 8px 8px; }
    .footer a { color: #D97757; }
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
    
    <!-- FAILED SECTION (if any) -->
    {{#if HasFailed}}
    <div class="section">
      <div class="section-title">⛔ Failed Assets</div>
      {{#each FailedByFolder}}
      <div class="folder-group">
        <div class="folder-name">📁 {{FolderPath}}</div>
        <table>
          <tr>
            <th>Asset</th>
            <th>Type</th>
            <th>Address</th>
            <th>Status</th>
            <th>Last Run</th>
          </tr>
          {{#each Assets}}
          <tr>
            <td>
              <strong>{{AssetName}}</strong>
              {{#if ErrorMessage}}<div class="error-text">{{ErrorMessage}}</div>{{/if}}
            </td>
            <td class="asset-type">{{AssetType}}</td>
            <td>{{AddressingInfo}}</td>
            <td><span class="status-badge status-failed">{{StatusText}}</span></td>
            <td>{{LastExecutionTime}}</td>
          </tr>
          {{/each}}
        </table>
      </div>
      {{/each}}
    </div>
    {{/if}}
    
    <!-- WARNINGS SECTION (if any) -->
    {{#if HasWarnings}}
    <div class="section">
      <div class="section-title">⚠️ Differences Found</div>
      <!-- Similar structure -->
    </div>
    {{/if}}
    
    <!-- SUCCESS SECTION (collapsible or summary only) -->
    {{#if HasSuccess}}
    <div class="section">
      <div class="section-title">✅ Successful ({{SuccessCount}} assets)</div>
      <p style="color: #6B6B6B;">All assets backed up successfully. Details available in attached CSV.</p>
    </div>
    {{/if}}
    
    <div class="footer">
      Report generated by AssetCentre DR Report Tool | <a href="mailto:support@company.com">Support</a>
    </div>
  </div>
</body>
</html>
```

---

## PowerShell Module Structure

```
DR-Report-Tool/
├── DR-Report-Tool.ps1          # Main entry point
├── DR-Report-Tool.xaml         # XAML GUI definition
├── Modules/
│   ├── Database.psm1           # SQL connection and query execution
│   ├── HtmlGenerator.psm1      # HTML report generation
│   ├── EmailSender.psm1        # Outlook COM integration
│   └── Scheduler.psm1          # Task Scheduler integration
├── Templates/
│   ├── ReportTemplate.html     # HTML email template
│   └── Styles.css              # Embedded styles (for reference)
├── Config/
│   └── settings.json           # Saved configuration
└── Logs/
    └── execution.log           # Execution history
```

---

## Key Functions to Implement

### 1. Database Module
```powershell
function Invoke-AssetCentreQuery {
    param(
        [string]$Server,
        [string]$Database = "AssetCentre",
        [string]$FolderFilter = $null,  # null = entire tree
        [switch]$UseWindowsAuth
    )
    # Returns DataTable with all columns
}
```

### 2. HTML Generator Module
```powershell
function New-DRReport {
    param(
        [System.Data.DataTable]$Data,
        [string]$TemplatePath,
        [switch]$FailedOnly,
        [switch]$IncludeFullMessage
    )
    # Returns HTML string
}
```

### 3. Email Sender Module
```powershell
function Send-OutlookReport {
    param(
        [string]$To,
        [string]$CC,
        [string]$Subject,
        [string]$HtmlBody,
        [string]$AttachmentPath  # Optional CSV
    )
    # Uses Outlook COM object - stays on-premise
}
```

### 4. Scheduler Module
```powershell
function Register-DRReportTask {
    param(
        [string]$TaskName = "AssetCentre-DR-Report",
        [string]$Frequency,  # Daily, Weekly, Monthly
        [string]$DayOfWeek,
        [string]$Time,
        [string]$ConfigPath
    )
    # Creates Windows Task Scheduler task
}
```

---

## SQL Query (Embedded)

The tool should embed the full SQL query that accepts a folder filter parameter:

```sql
-- If @FolderName is NULL, query entire tree
-- If @FolderName is provided, start from that folder

DECLARE @FolderName NVARCHAR(255) = {{FolderFilter}};  -- NULL for all

WITH FolderTree AS (
    -- Dynamic start point based on parameter
    SELECT ... 
    WHERE (@FolderName IS NULL AND ParentId = 'AssetCentre:{00000000-0000-0000-0000-000000000000}')
       OR (@FolderName IS NOT NULL AND AssetName = @FolderName AND AssetTypeId = 'F87F4416-504D-4924-B7B6-3D9A49BE766C')
    ...
)
-- Rest of query unchanged
```

---

## Configuration File Format (settings.json)

```json
{
  "database": {
    "server": "SERVERNAME",
    "database": "AssetCentre",
    "useWindowsAuth": true
  },
  "report": {
    "scope": "all",
    "folderFilter": null,
    "includeSuccess": false,
    "includeFullMessage": false,
    "attachCsv": true
  },
  "email": {
    "to": ["team@company.com"],
    "cc": ["manager@company.com"],
    "subjectTemplate": "DR Report - {Date} - {FailedCount} Failed, {WarningCount} Warnings"
  },
  "schedule": {
    "enabled": true,
    "frequency": "Weekly",
    "dayOfWeek": "Monday",
    "time": "07:00"
  }
}
```

---

## Execution Modes

1. **GUI Mode (default):** `.\DR-Report-Tool.ps1`
2. **Silent Mode (for scheduler):** `.\DR-Report-Tool.ps1 -Silent -ConfigPath ".\Config\settings.json"`
3. **Preview Only:** `.\DR-Report-Tool.ps1 -PreviewOnly` (generates HTML, opens in browser, no email)

---

## Error Handling Requirements

- Log all executions to `Logs/execution.log`
- On failure, send error notification email to configured recipients
- Validate SQL connection before query execution
- Validate Outlook availability before send attempt
- Graceful handling of empty result sets

---

## Build Instructions

1. Create all module files with proper error handling
2. Build XAML GUI with data binding
3. Implement template engine for HTML generation (simple string replacement or use a library)
4. Test with sample data before SQL connection
5. Package as single folder, portable deployment

---

## Deliverables

1. `DR-Report-Tool.ps1` - Main script with XAML GUI
2. All module files
3. HTML template file
4. Sample `settings.json`
5. README with deployment instructions
