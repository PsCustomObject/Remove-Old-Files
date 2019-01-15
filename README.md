# Remove Old Files

*Remove-OldFiles.ps1* is a complete script to help SysAdmins rotating *stale* files in any directory either local or remote.

## Script Components

Script support both PowerShell and PowerShell core and is designed to be run as either a scheduled task or as part of any other automation framework like *System Center Orchestrator*.

*Remove-OldFiles.ps1* uses an external *comma separated file (csv)* file named **FilePaths.csv** as a centralized configuration file for the maximum flexibility, fields and their usage are described in the table below

| Column Name         | **Description**                     | **Values**                                | Notes                                                        |
| ------------------- | ----------------------------------- | ----------------------------------------- | ------------------------------------------------------------ |
| *CleanupPath*       | Path to scan for old files          | Local or UNC Path                         | If empty script will log an exception                        |
| *FileExtension*     | Files to search for cleanup         | Any string like **log**, **pdf**, **txt** | If empty will match all files (\*.\*)                        |
| *AgeTolerance*      | Tolerance in days                   | Any number like 30 or 90                  | If empty will default to **90 days**                         |
| *IncludeSubFolders* | Defines if recursion should be used | **0** or **1**                            | If not specified it will default to **0** (no recursion)<br /> Any other value will cause script to log an exception |

**Note:** Configuration file name can be changed updating **$cleanupPath** variable

## First Run and Customization

In addition to updating the CSV file with path and options script needs a few customization before being ready for production.

### Logging

Script integrated a full [logging function](https://github.com/PsCustomObject/New-LogEntry) to make troubleshooting easier in case of issues, all log messages have been already configured[^1 Of course all messages can be customized and additional ones can be defined.] the only required step is setting the correct path where log will be written via the ***$logPath*** variable

```powershell
[string]$logPath = '\\MyServer\ScriptLog\Remove-OldFiles\'
```

Log file name can be customized via the ***$logName*** variable

```powershell
# Define log name
[string]$logFilePath = $logPath + $dateString + '-RemoveOldFiles.log'
```

### Notifications

In addition to log file creation script supports mail notifications via Email **only** when an exception is thrown during execution.

All mail related configuration parameters, included HTML code for mail body, are defined in a region named ***Mail Settings Configuration***, before script execution at least the following parameters should be set to allow correct operation

```powershell
# Set mail settings
[string]$mailSender = # Mail Sender
[string[]]$mailRecipients = # 'somebody@somewhere.com', 'admin@somewhere.com'
[string]$smtpRelay = #smtpRelay
```

### Ignore Folders

Folders can be ignored, useful when using recursion, creating a file named **ignore** underneath the path that should be skipped. File name can be customized updating the **$ignoreFile** variable value

```powershell
# Define Ignore control file
[string]$ignoreFile = 'ignore'
```

