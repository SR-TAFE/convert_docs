# Convert Docs Tools

This repository contains two PowerShell scripts designed to assist organizations in preparing for and implementing aspects of the Essential Eight security strategies, particularly focusing on application control and patch management.

## Scripts Overview

1. **Macro Discovery Script**: This script scans specified directories, identifying and cataloging Microsoft Office Documents with macros active. It creates an inventory of Word, PowerPoint and Excel files including details such as file path for all files identified with macros. This tool is helpful for organizations implementing identifying the impact of disabling macros at at an OU level.

2. **Office File Format Converter**: This script automates the conversion of legacy Microsoft Office file formats (.doc, .xls, .ppt) to their current equivalents (.docx, .xlsx, .pptx).

### Signing Note
This script is not signed, which means you will need to run it as either unsigned or with a local signature added to it.  As such it should only be used in test environments, and is not production ready. Take a read of Microsoft's guidance onExecutionPolicy before making any changes and tread with care. https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-executionpolicy?view=powershell-7.4

## Office File Format Converter

### Description

This PowerShell script converts old Microsoft Office file formats (.doc, .ppt, .xls) to their modern equivalents (.docx, .pptx, .xlsx) in a specified directory and its subdirectories. It's designed to help organizations modernize their document libraries efficiently.

### Features

- Converts Word (.doc), Excel (.xls), and PowerPoint (.ppt) files to their modern formats
- Processes files in the selected directory and all its subdirectories
- Provides a graphical interface for directory selection
- Logs all operations and errors for easy troubleshooting
- Displays a progress bar during conversion
- Ensures proper release of COM objects to prevent memory leaks

### Prerequisites

- Windows operating system
- PowerShell V3 or later
- Microsoft Office (Word, Excel, and PowerPoint) installed on the system

### Installation

1. Clone this repository or download the `convert_office_files.ps1` script.
2. Ensure that your PowerShell execution policy allows running scripts. You may need to run `Set-ExecutionPolicy RemoteSigned` in an elevated PowerShell prompt.

### Usage

1. Right-click on the `convert_office_files.ps1` file and select "Run with PowerShell", or
2. Open PowerShell, navigate to the script's directory, and run:
.\convert_office_files.ps1
3. When prompted, select the directory containing the files you want to convert.
4. The script will process all files and display progress in the console.
5. Once complete, check the `conversion_log.txt` file in the script's directory for detailed results.

### Logging

The script creates a log file named `conversion_log.txt` in the same directory as the script. This log contains information about each converted file and any errors encountered during the process.

### Error Handling

The script includes comprehensive error handling to manage issues such as:
- Invalid directory selection
- File access problems
- Conversion failures

All errors are logged in the `conversion_log.txt` file.

### Performance

Performance may vary depending on the number and size of files being converted. The script processes files sequentially to ensure stability.

### License

This script is released under the GNU General Public License v3.0. See the [LICENSE](LICENSE) file for details.

### Contributing

Contributions to improve the script are welcome. Please feel free to submit a Pull Request.

### Disclaimer

Always ensure you have backups of your files before running any conversion process. While this script has been designed to be safe, unforeseen issues can occur.

### Support

For bug reports or feature requests, please open an issue in the GitHub repository.


## Macro Scanner for Microsoft Office Files

### Description

This PowerShell script scans Microsoft Office files (.doc, .ppt, and .xls) in a specified directory and its subdirectories for the presence of macros. It reports all files containing macros and exports the list to a CSV file.

### Features

- Scans .doc, .ppt, and .xls files recursively in a selected directory
- Detects the presence of macros in each file
- Provides real-time console output of files containing macros
- Generates a summary report of all files with macros
- Exports the list of files with macros to a CSV file

### Prerequisites

- Windows operating system
- PowerShell V3 or later
- Microsoft Office (Word, Excel, and PowerPoint) installed on the machine

### Usage

1. Clone or download this repository to your local machine.
2. Open PowerShell as an administrator.
3. Navigate to the directory containing the script.
4. Run the script by typing: .\scan_for_macros.ps1
5. When prompted, select the folder you want to scan in the folder browser dialog.
6. Wait for the scan to complete. The script will display progress and results in the console.
7. Check the console output for the summary and the location of the exported CSV file.

### Output

- Console output showing files with macros as they are discovered
- A summary in the console listing all files containing macros
- A CSV file named `files_with_macros.csv` in the scanned directory, containing the full paths of all files with macros

## Office Macro Security Scanner

### Overview
This PowerShell script analyzes Microsoft Office files (Word, Excel, and PowerPoint) for potentially malicious macro content. It scans files for known high-risk and low-risk patterns commonly associated with malicious macros and generates a detailed security report.

### Features
- Scans .docm, .xlsm, and .pptm files (including legacy .doc, .xls, .ppt formats)
- Identifies high-risk patterns including:
  - Shell commands and process operations
  - Registry modifications
  - File system operations
  - Network communications
  - Anti-analysis techniques
  - Encryption and encoding
  - Security software interference
  - Persistence mechanisms
- Generates CSV reports with risk assessments
- Provides detailed pattern matching results
- Supports batch processing of multiple files

## Risk Analysis Table

### Pattern Categories and Associated Risks

| Pattern Category | Example Patterns | Risk Level | Security Impact | Common Usage |
|-----------------|------------------|------------|-----------------|--------------|
| Shell Commands | `Shell`, `WScript.Shell`, `PowerShell` | HIGH | • Command execution<br>• System modification<br>• Backdoor creation | • Process automation<br>• System administration |
| File Operations | `FileSystemObject`, `Kill`, `Open`, `Binary` | HIGH | • Data theft<br>• File deletion<br>• Malware dropping | • Document management<br>• File backup<br>• Data export |
| Registry Access | `RegRead`, `RegWrite`, `RegDelete` | HIGH | • Persistence<br>• System modification<br>• Security bypass | • Settings storage<br>• User preferences |
| Network Operations | `URLDownloadToFile`, `XMLHTTP`, `WinHttp` | HIGH | • Data exfiltration<br>• Malware download<br>• C2 communication | • Web API integration<br>• Data updates |
| Process Manipulation | `CreateProcess`, `Shell`, `Win32_Process` | HIGH | • Malware execution<br>• System compromise<br>• Privilege escalation | • Application launching<br>• System integration |
| Anti-Analysis | `Application.Visible`, `DisplayAlerts`, `ScreenUpdating` | HIGH | • Detection evasion<br>• Analysis prevention | • UI optimization<br>• Performance improvement |
| Encryption/Encoding | `Chr`, `Base64`, `XOR`, `Environ` | HIGH | • Code obfuscation<br>• Evasion technique | • Data protection<br>• String handling |
| Security Tools | `AutoExec`, `Auto_Open`, `Document_Open` | HIGH | • Automatic execution<br>• User circumvention | • Document initialization<br>• Setup procedures |
| Clipboard Access | `GetFromClipboard`, `SetClipboard` | MEDIUM | • Data interception<br>• Information theft | • Data transfer<br>• Copy/paste operations |
| Document Modification | `VBProject`, `CodeModule`, `AddFromString` | MEDIUM | • Self-modification<br>• Code injection | • Template generation<br>• Code updates |
| Email Operations | `Outlook.Application`, `MailItem`, `Send` | MEDIUM | • Data exfiltration<br>• Spam sending | • Email automation<br>• Notifications |
| Form Controls | `UserForm`, `TextBox`, `CommandButton` | LOW | • User interaction<br>• Data input | • User interface<br>• Data entry |
| Cell Operations | `Range`, `Cells`, `Selection` | LOW | • Content modification | • Data formatting<br>• Calculations |
| Basic Functions | `Len`, `Mid`, `Left`, `Right` | LOW | • String manipulation | • Text processing<br>• Data validation |
| Time/Date Functions | `Now`, `Date`, `Time` | LOW | • Timing operations | • Date calculations<br>• Scheduling |

### Risk Level Definitions

| Risk Level | Description | Recommended Action |
|------------|-------------|-------------------|
| HIGH | Patterns commonly associated with malicious activity | Immediate investigation required; Block execution unless explicitly authorized |
| MEDIUM | Patterns that could be misused but have legitimate uses | Review context and purpose; Monitor usage |
| LOW | Patterns typically associated with normal operations | Regular monitoring; No immediate action required |

### Contextual Risk Factors

| Factor | Description | Risk Multiplier |
|--------|-------------|-----------------|
| Pattern Combinations | Multiple high-risk patterns in single macro | Increases overall risk score |
| Code Obfuscation | Unclear or intentionally obscured code | Significantly increases risk |
| Auto-Execution | Macros that run automatically on document open | Increases risk level |
| External References | Links to external files or resources | Increases risk level |
| Origin | Source/author of the document | Contextual risk factor |

### Detection Confidence Levels

| Confidence Level | Criteria | False Positive Rate |
|-----------------|----------|-------------------|
| High | Multiple high-risk patterns with known malicious combinations | < 5% |
| Medium | Individual high-risk patterns or suspicious combinations | 5-15% |
| Low | Common patterns in unusual configurations | 15-30% |


### Prerequisites
- Windows PowerShell 5.1 or later
- Microsoft Office installed (Word, Excel, PowerPoint)
- Administrative privileges recommended
- Office Trust Center settings configured:
  - Trust access to the VBA project object model enabled
  - Macro settings appropriately configured

## Installation
1. Download `Macros_Security_Scan.ps1`
2. Place in desired directory
3. Configure Office applications' Trust Center settings

### Usage

# Basic usage
.\Macros_Security_Scan.ps1 -Path "C:\Documents"

# Scan specific file
.\Macros_Security_Scan.ps1 -Path "C:\Documents\suspicious.docm"

# Specify custom output location
.\Macros_Security_Scan.ps1 -Path "C:\Documents" -OutputPath "C:\Reports"


### License

This project is licensed under the GNU General Public License v3.0.

### Contributing

Contributions to improve the script are welcome. Please feel free to submit a Pull Request.

### Disclaimer

This script is provided as-is, without any warranty. Always ensure you have appropriate permissions before scanning files, especially in a corporate environment.

### Author

Ant Congdon

## Version History

- 1.0: Initial release

