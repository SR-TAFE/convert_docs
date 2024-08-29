<#
.SYNOPSIS
    Office File Format Converter

.DESCRIPTION
    This script converts old Microsoft Office file formats (.doc, .ppt, .xls) to their modern
    equivalents (.docx, .pptx, .xlsx) in a specified directory and its subdirectories.

.NOTES
    File Name      : convert_office_files.ps1
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
$folderBrowser.Description = "Select the folder containing Office files to convert"
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

# Get all .doc, .ppt, and .xls files in the specified directory
$files = Get-ChildItem -Path $sourceDir -Include *.doc, *.ppt, *.xls -Recurse

foreach ($file in $files) {
    try {
        switch ($file.Extension.ToLower()) {
            ".doc" {
                $doc = $word.Documents.Open($file.FullName)
                $newFileName = [System.IO.Path]::ChangeExtension($file.FullName, ".docx")
                $doc.SaveAs([ref] $newFileName, [ref] 16) # 16 is the value for .docx format
                $doc.Close()
                Write-Host "Converted: $($file.Name) to $([System.IO.Path]::GetFileName($newFileName))"
            }
            ".xls" {
                $workbook = $excel.Workbooks.Open($file.FullName)
                $newFileName = [System.IO.Path]::ChangeExtension($file.FullName, ".xlsx")
                $workbook.SaveAs($newFileName, 51) # 51 is the value for .xlsx format
                $workbook.Close()
                Write-Host "Converted: $($file.Name) to $([System.IO.Path]::GetFileName($newFileName))"
            }
            ".ppt" {
                $presentation = $powerpoint.Presentations.Open($file.FullName)
                $newFileName = [System.IO.Path]::ChangeExtension($file.FullName, ".pptx")
                $presentation.SaveAs($newFileName, 24) # 24 is the value for .pptx format
                $presentation.Close()
                Write-Host "Converted: $($file.Name) to $([System.IO.Path]::GetFileName($newFileName))"
            }
        }
    }
    catch {
        Write-Host "Error converting $($file.Name): $_"
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

Write-Host "Conversion process completed."
