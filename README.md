# Convert Docs Tools

This repository contains two PowerShell scripts designed to assist organizations in preparing for and implementing aspects of the Essential Eight security strategies, particularly focusing on application control and patch management.

## Scripts Overview

1. **Macro Discovery Script**: This script scans specified directories, identifying and cataloging Microsoft Office Documents with macros active. It creates an inventory of Word, PowerPoint and Excel files including details such as file path for all files identified with macros. This tool is helpful for organizations implementing identifying the impact of disabling macros at at an OU level.

2. **Office File Format Converter**: This script automates the conversion of legacy Microsoft Office file formats (.doc, .xls, .ppt) to their current equivalents (.docx, .xlsx, .pptx).

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

## Logging

The script creates a log file named `conversion_log.txt` in the same directory as the script. This log contains information about each converted file and any errors encountered during the process.

## Error Handling

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

## Disclaimer

Always ensure you have backups of your files before running any conversion process. While this script has been designed to be safe, unforeseen issues can occur.

## Support

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

