<#
.SYNOPSIS
    DOC to DOCX Converter and Macro Detector

.DESCRIPTION
    This script converts .doc files to .docx format in a specified directory and its subdirectories.
    It also detects the presence of macros in each file and reports on files containing macros.

.NOTES
    File Name      : doc_to_docx_converter.ps1
    Author         : Ant Congdon
    Prerequisite   : PowerShell V3 or later, Microsoft Word

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
$folderBrowser.Description = "Select the folder containing .doc files"
$folderBrowser.RootFolder = "MyComputer"

if ($folderBrowser.ShowDialog() -eq "OK") {
    $sourceDir = $folderBrowser.SelectedPath
} else {
    Write-Host "No folder selected. Exiting script."
    exit
}

# Create a Word application object
$word = New-Object -ComObject Word.Application

# Make Word invisible
$word.Visible = $false

# Get all .doc files (but not .docx) in the specified directory and all subdirectories
$files = Get-ChildItem -Path $sourceDir -Filter *.doc -Recurse | Where-Object { $_.Extension -eq ".doc" }

# Create an array to store files with macros
$filesWithMacros = @()

foreach ($file in $files) {
    try {
        # Check if a .docx version already exists
        $docxVersion = [System.IO.Path]::ChangeExtension($file.FullName, ".docx")
        if (Test-Path $docxVersion) {
            Write-Host "Skipped: $($file.FullName) (DOCX version already exists)"
            continue
        }

        # Open the document
        $doc = $word.Documents.Open($file.FullName)
        
        # Check for macros
        $hasMacros = $doc.HasVBProject

        # Create the new file name
        $newFileName = [System.IO.Path]::ChangeExtension($file.FullName, ".docx")
        
        # Save as .docx
        $doc.SaveAs([ref] $newFileName, [ref] 16) # 16 is the value for .docx format
        
        # Close the document
        $doc.Close()
        
        if ($hasMacros) {
            $filesWithMacros += $file.FullName
            Write-Host "Converted (Contains Macros): $($file.FullName) to $newFileName"
        } else {
            Write-Host "Converted: $($file.FullName) to $newFileName"
        }
    }
    catch {
        Write-Host "Error converting $($file.FullName): $_"
    }
}

# Quit Word
$word.Quit()

# Release the COM object
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null

Write-Host "Conversion process completed."

# Report files with macros
if ($filesWithMacros.Count -gt 0) {
    Write-Host "`nThe following files contain macros:"
    foreach ($file in $filesWithMacros) {
        Write-Host $file
    }
} else {
    Write-Host "`nNo files with macros were found."
}
