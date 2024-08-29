# Macro Scanner for Microsoft Office Files

## Description

This PowerShell script scans Microsoft Office files (.doc, .ppt, and .xls) in a specified directory and its subdirectories for the presence of macros. It reports all files containing macros and exports the list to a CSV file.

## Features

- Scans .doc, .ppt, and .xls files recursively in a selected directory
- Detects the presence of macros in each file
- Provides real-time console output of files containing macros
- Generates a summary report of all files with macros
- Exports the list of files with macros to a CSV file

## Prerequisites

- Windows operating system
- PowerShell V3 or later
- Microsoft Office (Word, Excel, and PowerPoint) installed on the machine

## Usage

1. Clone or download this repository to your local machine.
2. Open PowerShell as an administrator.
3. Navigate to the directory containing the script.
4. Run the script by typing: .\scan_for_macros.ps1
5. When prompted, select the folder you want to scan in the folder browser dialog.
6. Wait for the scan to complete. The script will display progress and results in the console.
7. Check the console output for the summary and the location of the exported CSV file.

## Output

- Console output showing files with macros as they are discovered
- A summary in the console listing all files containing macros
- A CSV file named `files_with_macros.csv` in the scanned directory, containing the full paths of all files with macros

## License

This project is licensed under the GNU General Public License v3.0.

## Contributing

Contributions to improve the script are welcome. Please feel free to submit a Pull Request.

## Disclaimer

This script is provided as-is, without any warranty. Always ensure you have appropriate permissions before scanning files, especially in a corporate environment.

## Author

Ant Congdon

## Version History

- 1.0: Initial release

