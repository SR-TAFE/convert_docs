<#
.SYNOPSIS
    Macro Scanner for Microsoft Office Files

.DESCRIPTION
    This script scans .doc, .ppt, and .xls files in a specified directory and its subdirectories
    for the presence of macros. It reports all files containing macros and exports the list to a CSV file.

.NOTES
    File Name      : macro_scan.ps1
    Author         : Ant Congdon
    Prerequisite   : PowerShell V3 or later, Microsoft Office (Word, Excel, and PowerPoint)

.LICENSE
    This program is free software: you can redistribute it and/or modify
    it under the terms of the GNU General Public License as published by
    the Free Software Foundation, either version 3 of the License, or
    (at your option) any later version.

    This program is distributed in the hope that it will be useful,
    but WITHOUT ANY WARRANTY; without even the implied warranty of
    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
    GNU General Public License for more details.

    You should have received a copy of the GNU General Public License
    along with this program.  If not, see <https://www.gnu.org/licenses/>.

.LINK
    For the full license text, see <https://www.gnu.org/licenses/gpl-3.0.html>
#>

# Prompt the user to choose the directory
Add-Type -AssemblyName System.Windows.Forms
$folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
$folderBrowser.Description = "Select the folder to scan for macros"
$folderBrowser.RootFolder = "MyComputer"

if ($folderBrowser.ShowDialog() -eq "OK") {
    $sourceDir = $folderBrowser.SelectedPath
} else {
    Write-Host "No folder selected. Exiting script."
    exit
}

# Create application objects
$word = New-Object -ComObject Word.Application
$excel = New-Object -ComObject Excel.Application
$powerpoint = New-Object -ComObject PowerPoint.Application

# Make applications invisible
$word.Visible = $false
$excel.Visible = $false
$powerpoint.Visible = $false

# Get all .doc, .ppt, and .xls files in the specified directory and all subdirectories
$files = Get-ChildItem -Path $sourceDir -Include *.doc, *.ppt, *.xls -Recurse

# Create an array to store files with macros
$filesWithMacros = @()

foreach ($file in $files) {
    try {
        $hasMacros = $false

        switch ($file.Extension.ToLower()) {
            ".doc" {
                $doc = $word.Documents.Open($file.FullName, $false, $true)
                $hasMacros = $doc.HasVBProject
                $doc.Close()
            }
            ".xls" {
                $workbook = $excel.Workbooks.Open($file.FullName, $false, $true)
                $hasMacros = $workbook.HasVBProject
                $workbook.Close()
            }
            ".ppt" {
                $presentation = $powerpoint.Presentations.Open($file.FullName, $false, $true, $false)
                $hasMacros = $presentation.HasVBProject
                $presentation.Close()
            }
        }

        if ($hasMacros) {
            $filesWithMacros += $file.FullName
            Write-Host "Macros found in: $($file.FullName)"
        }
    }
    catch {
        Write-Host "Error scanning $($file.FullName): $_"
    }
}

# Quit applications
$word.Quit()
$excel.Quit()
$powerpoint.Quit()

# Release COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerpoint) | Out-Null

# Report files with macros
if ($filesWithMacros.Count -gt 0) {
    Write-Host "`nThe following files contain macros:"
    foreach ($file in $filesWithMacros) {
        Write-Host $file
    }
    
    # Export the list to a CSV file
    $csvPath = Join-Path $sourceDir "files_with_macros.csv"
    $filesWithMacros | ForEach-Object { [PSCustomObject]@{ FilePath = $_ } } | Export-Csv -Path $csvPath -NoTypeInformation
    Write-Host "`nList of files with macros has been exported to: $csvPath"
} else {
    Write-Host "`nNo files with macros were found."
}

Write-Host "`nScan completed."
