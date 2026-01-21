# AssetCentre DR Report Tool

A PowerShell application with XAML GUI that generates and sends professional HTML email reports for FactoryTalk AssetCentre disaster recovery (DR) backup status.

## Features

- **SQL Database Integration**: Connects to AssetCentre database to retrieve backup status
- **Professional HTML Reports**: Generates styled email reports with summary statistics
- **Outlook Integration**: Sends emails via on-premise Outlook (no external services)
- **Task Scheduler Support**: Automate report generation with Windows Task Scheduler
- **Flexible Filtering**: Filter by folder, status (Failed/Warning/Success)
- **CSV Export**: Optionally attach detailed CSV exports to emails

## Requirements

- Windows 10/11 or Windows Server 2016+
- PowerShell 5.1 or later
- Microsoft Outlook (desktop application)
- SQL Server access to AssetCentre database
- .NET Framework 4.7.2 or later

## Installation

1. Copy the entire `DR-Report-Tool` folder to your desired location
2. Ensure PowerShell execution policy allows running scripts:
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```
3. No additional installation required - the tool is portable

## Directory Structure

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
│   └── ReportTemplate.html     # HTML email template reference
├── Config/
│   └── settings.json           # Saved configuration
├── Logs/
│   └── execution.log           # Execution history
└── README.md                   # This file
```

## Usage

### GUI Mode (Default)

Launch the graphical interface:

```powershell
.\DR-Report-Tool.ps1
```

### Silent Mode (Scheduled Tasks)

Run without GUI using a configuration file:

```powershell
.\DR-Report-Tool.ps1 -Silent -ConfigPath ".\Config\settings.json"
```

### Preview Mode

Generate report and open in browser without sending email:

```powershell
.\DR-Report-Tool.ps1 -PreviewOnly
```

## Configuration

### Database Connection

- **Server**: SQL Server instance name (e.g., `SERVERNAME\SQLEXPRESS`)
- **Database**: Database name (default: `AssetCentre`)
- **Authentication**: Windows Authentication (recommended) or SQL Authentication

### Report Scope

- **Entire Tree**: Query all assets in AssetCentre
- **Specific Folder**: Filter to a specific folder and its children
- **Status Filters**: Include/exclude Failed, Warning, and Success items

### Email Configuration

- **To**: Primary recipients (comma-separated)
- **CC**: Carbon copy recipients (optional)
- **Subject Template**: Use placeholders:
  - `{Date}` - Current date (yyyy-MM-dd)
  - `{DateTime}` - Date and time
  - `{FailedCount}` - Number of failed assets
  - `{WarningCount}` - Number of warnings
  - `{SuccessCount}` - Number of successful backups
  - `{Summary}` - Status summary (FAILED/WARNINGS/ALL OK)

### Scheduling

Create a Windows Task Scheduler task to run reports automatically:

1. Enable scheduled execution in the GUI
2. Select frequency (Daily, Weekly, Monthly)
3. Choose day of week and time
4. Click "Create Task"

## Configuration File Format

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
    "includeWarnings": true,
    "includeSuccess": false,
    "includeFullMessage": false,
    "attachCsv": true
  },
  "email": {
    "to": ["team@company.com"],
    "cc": ["manager@company.com"],
    "subjectTemplate": "DR Report - {Date} - {Summary}"
  },
  "schedule": {
    "enabled": true,
    "frequency": "Weekly",
    "dayOfWeek": "Monday",
    "time": "07:00"
  }
}
```

## Color Palette

The report uses the Claude Light Theme color palette:

| Element | Color |
|---------|-------|
| Primary Background | #FFFFFF |
| Secondary Background | #F9F9F9 |
| Border | #E5E5E5 |
| Text Primary | #1A1A1A |
| Text Secondary | #6B6B6B |
| Success | #16A34A (green) |
| Warning | #D97706 (amber) |
| Failed | #DC2626 (red) |
| Accent/Links | #D97757 (terracotta) |
| Header Background | #2D2D2D |
| Header Text | #FFFFFF |

## Troubleshooting

### "Outlook not available" Error

- Ensure Microsoft Outlook desktop application is installed
- Outlook must be configured with at least one email account
- The tool uses Outlook COM automation - web versions are not supported

### Database Connection Issues

- Verify SQL Server is accessible from your machine
- Check firewall rules allow SQL Server connections
- For Windows Auth, ensure your user has database access
- For SQL Auth, verify username and password

### Task Scheduler Issues

- Scheduled tasks run as the current user
- Ensure the user remains logged in or use "Run whether user is logged on or not"
- Check Windows Event Viewer for task execution errors

### Empty Reports

- Verify the database query returns results
- Check folder filter is correct (case-sensitive)
- Use "Preview HTML" to test without sending email

## Security Notes

- **All data stays within your network** - No cloud services or external APIs
- Credentials are not stored in configuration files
- Windows Authentication is recommended for database access
- Email is sent through your organization's Outlook/Exchange infrastructure

## Logs

Execution logs are stored in `Logs/execution.log`:

```
[2026-01-21 07:00:00] [Info] === DR Report Tool Started ===
[2026-01-21 07:00:00] [Info] Mode: Silent
[2026-01-21 07:00:01] [Info] Connecting to database: SERVERNAME\AssetCentre
[2026-01-21 07:00:05] [Info] Retrieved 150 records from database
[2026-01-21 07:00:06] [Info] Generating HTML report...
[2026-01-21 07:00:06] [Info] CSV exported to: Logs\DR_Report_20260121_070006.csv
[2026-01-21 07:00:07] [Info] Sending email to: team@company.com
[2026-01-21 07:00:08] [Info] Email sent successfully
[2026-01-21 07:00:08] [Info] === DR Report Tool Ended ===
```

## Support

For issues or feature requests, contact your system administrator.

## License

Internal use only. All rights reserved.
