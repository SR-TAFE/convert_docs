<#
.SYNOPSIS
    Office File Format Converter

.DESCRIPTION
    This script converts old Microsoft Office file formats (.doc, .ppt, .xls) to their modern
    equivalents (.docx, .pptx, .xlsx) in a specified directory and its subdirectories.

.NOTES
    File Name      : convert_office_files.ps1
    Author         : [Your Name]
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

# Add necessary assemblies
Add-Type -AssemblyName "Microsoft.Office.Interop.Word"
Add-Type -AssemblyName "Microsoft.Office.Interop.Excel"
Add-Type -AssemblyName "Microsoft.Office.Interop.PowerPoint"
Add-Type -AssemblyName System.Windows.Forms

# Set strict mode and error action preference
Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# Initialize logging
$logFile = Join-Path $PSScriptRoot "conversion_log.txt"
function Write-Log {
    param([string]$message)
    $logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $message"
    Add-Content -Path $logFile -Value $logMessage
    Write-Host $logMessage
}

# Function to safely release COM objects
function Release-ComObject {
    param($comObject)
    if ($null -ne $comObject) {
        try {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($comObject) | Out-Null
        }
        catch {
            Write-Log "Error releasing COM object: $($_.Exception.Message)"
        }
    }
}

# Function to convert a file
function Convert-OfficeFile {
    param(
        [System.IO.FileInfo]$file,
        $wordApp,
        $excelApp,
        $powerPointApp
    )

    $newFileName = [System.IO.Path]::Combine(
        [System.IO.Path]::GetDirectoryName($file.FullName),
        [System.IO.Path]::GetFileNameWithoutExtension($file.FullName)
    )
    
    try {
        switch ($file.Extension.ToLower()) {
            ".doc" {
                $newFileName += ".docx"
                $doc = $wordApp.Documents.Open($file.FullName)
                $doc.SaveAs([ref] $newFileName, [ref] 16) # 16 is the value for .docx format
                $doc.Close()
                Release-ComObject $doc
            }
            ".xls" {
                $newFileName += ".xlsx"
                $workbook = $excelApp.Workbooks.Open($file.FullName)
                $workbook.SaveAs($newFileName, 51) # 51 is the value for .xlsx format
                $workbook.Close()
                Release-ComObject $workbook
            }
            ".ppt" {
                $newFileName += ".pptx"
                $presentation = $powerPointApp.Presentations.Open($file.FullName)
                $presentation.SaveAs($newFileName, 24) # 24 is the value for .pptx format
                $presentation.Close()
                Release-ComObject $presentation
            }
        }
        Write-Log "Converted: $($file.Name) to $([System.IO.Path]::GetFileName($newFileName))"
        return $true
    }
    catch {
        Write-Log "Error converting $($file.Name): $($_.Exception.Message)"
        return $false
    }
}

# Main script execution
try {
    Write-Log "Script started."

    # Prompt user to choose directory
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "Select the folder containing Office files to convert"
    $folderBrowser.RootFolder = "MyComputer"

    if ($folderBrowser.ShowDialog() -ne "OK") {
        throw "No folder selected. Exiting script."
    }

    $sourceDir = $folderBrowser.SelectedPath

    # Validate input path
    if (-not (Test-Path $sourceDir -PathType Container)) {
        throw "Invalid directory path: $sourceDir"
    }

    Write-Log "Selected directory: $sourceDir"

    # Create application objects
    $word = New-Object -ComObject Word.Application
    $excel = New-Object -ComObject Excel.Application
    $powerpoint = New-Object -ComObject PowerPoint.Application

    # Make Word and Excel applications invisible
    $word.Visible = $false
    $excel.Visible = $false
    # PowerPoint visibility is not set as it's not allowed

    # Get all .doc, .ppt, and .xls files in the specified directory
    $files = @(Get-ChildItem -Path $sourceDir -Include *.doc, *.ppt, *.xls -Recurse)

    $totalFiles = $files.Count
    $convertedFiles = 0

    if ($totalFiles -eq 0) {
        Write-Log "No files found to convert in the selected directory."
    } else {
        foreach ($file in $files) {
            $success = Convert-OfficeFile -file $file -wordApp $word -excelApp $excel -powerPointApp $powerpoint
            if ($success) {
                $convertedFiles++
            }
            $percentComplete = ($convertedFiles / $totalFiles) * 100
            Write-Progress -Activity "Converting Files" -Status "Progress" -PercentComplete $percentComplete
        }
    }
}
catch {
    Write-Log "Critical error: $($_.Exception.Message)"
}
finally {
    # Ensure COM objects are released
    if ($null -ne $word) { 
        $word.Quit()
        Release-ComObject $word
    }
    if ($null -ne $excel) {
        $excel.Quit()
        Release-ComObject $excel
    }
    if ($null -ne $powerpoint) {
        $powerpoint.Quit()
        Release-ComObject $powerpoint
    }

    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    if ($null -eq $totalFiles) {
        $totalFiles = 0
    }
    if ($null -eq $convertedFiles) {
        $convertedFiles = 0
    }

    Write-Log "Conversion process completed. Total files processed: $totalFiles, Successfully converted: $convertedFiles"
}
