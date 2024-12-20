# Script to analyze macro content in Office files for security risks
# Outputs a CSV report with risk assessments and identified patterns

# Define all risk pattern categories
# High risk patterns - activities that could indicate malicious intent
$highRiskPatterns = @(
    # Process and Shell Operations
    'Shell',                 # Command shell execution
    'WScript\.Shell',
    'PowerShell',
    'Exec\(',
    'CreateObject',
    'CallByName',
    'MacScript',
    
    # Process and Memory Manipulation
    'VirtualProtect',
    'WriteProcessMemory',
    'ReadProcessMemory',
    'CreateThread',
    'CreateRemoteThread',
    'NtAllocateVirtualMemory',
    'GetProcAddress',
    'LoadLibrary',
    'AdjustTokenPrivileges',
    'VirtualAlloc',
    'RtlMoveMemory',
    'EnumProcesses',
    'TerminateProcess',
    
    # Registry Operations
    'RegRead',              # Registry manipulation
    'RegWrite',
    'RegDelete',
    'Registry',
    'HKEY_LOCAL_MACHINE',
    'HKEY_CURRENT_USER',
    
    # File System Operations
    'Kill\s',               # File deletion/manipulation
    'FileCopy\s',
    'DeleteFile',
    'SetFileAttributes',
    'AddFromFile',
    'SaveToFile',
    'OpenTextFile',
    
    # Network Operations
    'Winsock',             # Network communication
    'InternetReadFile',
    'InternetOpenUrl',
    'InternetConnect',
    'HTTPRequest',
    'URLDownloadToFile',
    'FTPRequest',
    'Sockets',
    'XMLHTTP',
    'ServerXMLHTTP',
    'WinHttpRequest',
    
    # Anti-Analysis
    'IsDebuggerPresent',    # Debug/Analysis prevention
    'CheckRemoteDebuggerPresent',
    'GetTickCount',
    'QueryPerformanceCounter',
    'Sleep',
    'EmptyClipboard',
    'Application\.Visible\s*=\s*False',
    
    # Encryption and Encoding
    'CryptoAPI',            # Encryption/Encoding
    'StrReverse',
    'Chr\(',
    'Asc\(',
    'Base64',
    'FromBase64String',
    'ConvertToAutoIt3',
    
    # Windows Management
    'netsh\s',              # System configuration
    'taskkill',
    'sc\s+config',
    'reg\s+add',
    'bcdedit',
    'schtasks',
    'RunDll32',
    
    # Security Software Interaction
    'FirewallAPI',          # Security software
    'AvastSvc',
    'McAfee',
    'Symantec',
    'Windows Defender',
    'avp',
    'AntiVirus',
    
    # Command Execution
    'InvokeExpression',     # Code execution
    'Invoke-Expression',
    'iex\s',
    'Invoke-Command',
    'Invoke-Item',
    'Start-Process',
    'ExecuteExcel4Macro',
    
    # Persistence Mechanisms
    'CurrentVersion\\Run',   # System persistence
    'StartupFolder',
    'New-Service',
    'HKEY_LOCAL_MACHINE\\System\\CurrentControlSet\\Services',
    'Schedule\.Service',
    
    # System Information Gathering
    'GetSystemInfo',        # System reconnaissance
    'GetComputerName',
    'GetUserName',
    'GetTempPath',
    'EnumProcesses',
    'Win32_Process',
    'Win32_Service',
    
    # COM Object Creation
    'CreateObject',         # Suspicious COM objects
    'GetObject',
    'WScript\.Shell',
    'Shell\.Application',
    
    # Macro Security
    'AccessVBOM',           # VBA security bypass
    'VBAWarnings',
    'DisableAttackProtection',
    
    # Suspicious Combinations
    'URLDownload.*Shell',   # Download and execute
    'CreateObject.*Shell',
    'RegWrite.*Run',
    
    # Additional High-Risk Indicators
    'System\.Reflection',   # Reflection/dynamic execution
    'Assembly\.Load',
    'WriteLines',
    'WriteAllBytes',
    'DownloadString',
    'DownloadFile',
    'WebClient',
    'MemoryStream',
    'StartInfo',
    'ProcessStartInfo',
    'RunspaceFactory',
    'AddScript',
    'Invoke',
    'DllImport',
    'Marshal',
    'InteropServices'
        # Command & Control
    'ActiveXObject',
    'MSScriptControl',
    'DDEInitiate',          # Dynamic Data Exchange
    'DDEExecute',
    'DDETerminate',
    
    # Evasion Techniques
    'ThisDocument\.FullName', # Document self-referencing
    'System\.Environment',
    'PrivateProfileString',
    'VBProject',
    'Windows\.Hide',
    
    # Remote Access
    'PSExec',
    'WMI',
    'Win32_Product',
    'ManagementObject',
    'RunAs',
    
    # Persistence Methods
    'HKCU:\\Software\\Microsoft\\Office',
    'VBAProject\.VBComponents',
    'OnTime',
    'AutoExec',
    'Document_Open',
    'Auto_Open',
    
    # Sandbox Detection
    'GetVolumeInformation',
    'EnumDisplaySettings',
    'GlobalMemoryStatus',
    'GetForegroundWindow',
    
    # Credential Access
    'CredentialManager',
    'LogonUser',
    'NetUserEnum',
    'SamAccountName',
    
    # Data Exfiltration
    'ADODB\.Stream',
    'MSXML2\.XMLHTTP',
    'certutil',             # Often used for encoding/decoding
    'Outlook\.Application',
    
    # Living Off The Land
    'bitsadmin',
    'mshta',
    'cscript',
    'wscript',
    'cmd\.exe',
    'powershell\.exe'

)

# Medium risk patterns - potentially concerning but may have legitimate uses
$mediumRiskPatterns = @(
    # Document Operations
    '\.SaveAs',             # File operations
    '\.Save',
    'Application\.ActiveDocument',
    'ActiveWorkbook',
    'Selection\.Text',
    'Documents\.Add',
    'Documents\.Open',
    'RecentFiles',
    
    # UI Manipulation
    'Application\.DisplayAlerts',  # Display settings
    'Application\.ScreenUpdating',
    'Application\.Visible',
    'ThisWorkbook\.Protect',
    'ActiveWindow',
    'WindowState',
    'DisplayStatusBar',
    'EnableEvents',
    
    # Event Handlers
    'Workbook_Open',        # Event procedures
    'Document_Close',
    'Auto_Close',
    'AutoClose',
    'Document_Open',
    'Auto_Open',
    'AutoExec',
    'AutoOpen',
    'DocumentChange',
    
    # System Interaction
    'Environment\..*Path',   # System/environment access
    'CurDir',
    'ChDir',
    'GetFolder',
    'FileSystemObject',
    'CreateFolder',
    'FolderExists',
    'GetSpecialFolder',
    
    # Error Handling
    'On\s?Error\s?Resume\s?Next',  # Error suppression
    'On\s?Error\s?GoTo',
    'Error\s?Handler',
    
    # Shell and Environment
    'Environ\$',            # Environment variables
    'Command\$',
    'System\.',
    
    # File Operations
    'Open\s+.*\s+For',     # File I/O
    'Binary\s+Access',
    'Random\s+Access',
    'Print\s+#',
    'Put\s+#',
    'Get\s+#',
    
    # ActiveX and COM
    'CreateObject',         # Object creation
    'GetObject',
    'CallByName',
    
    # Application Settings
    'Options\.',            # Excel/Word settings
    'Application\.Settings',
    'EnableMacros',
    
    # Custom Document Properties
    'CustomDocumentProperties',  # Document metadata
    'DocumentProperties',
    
    # Protected View
    'ProtectedView',        # Security features
    'Protect\.',
    'Unprotect',
    
    # Clipboard Operations
    'Clipboard',            # System clipboard
    'PasteSpecial',
    
    # Add-in Integration
    '\.ExportAsFixedFormat',  # Export functionality
    '\.SaveAs.*PDF',
    '\.SaveAs.*XPS',
    
    # External References
    'Links',                # External links
    'References',
    'ExternalReferences',
    
    # Application Control
    'SendKeys',             # Keyboard simulation
    'Application\.Wait',
    'Application\.Run'
)

# Low risk patterns - commonly used in legitimate macros
$lowRiskPatterns = @(
    # Basic Formatting
    'Bold\s*=',             # Text formatting
    'Italic\s*=',
    'Underline\s*=',
    'Font\.Size',
    'Font\.Name',
    'Font\.Color',
    'Interior\.Color',
    'Borders\(',
    'Alignment',
    'WordWrap',
    
    # Cell Operations
    'Range\.Clear',         # Basic cell operations
    'Range\.ClearContents',
    'Range\.ClearFormats',
    'AutoFit',
    'Columns\.Width',
    'Rows\.Height',
    'Hidden\s*=\s*False',
    'Sort\.',
    'AutoFilter',
    'Subtotal',
    
    # Basic Navigation
    'Cells\.Select',        # Cursor movement
    'Range\.Select',
    'Worksheet\.Select',
    'ActiveCell',
    'End\(xlUp\)',
    'End\(xlDown\)',
    'End\(xlToLeft\)',
    'End\(xlToRight\)',
    
    # Simple Math Operations
    'Sum\(',                # Basic calculations
    'Average\(',
    'Count\(',
    'Max\(',
    'Min\(',
    'Round\(',
    'Int\(',
    'Abs\(',
    
    # Display Settings
    'PageSetup\.',          # Page formatting
    'Orientation',
    'Zoom',
    'DisplayGridlines',
    'FitToPage',
    'CenterHeader',
    'LeftFooter',
    'RightMargin',
    
    # Basic String Operations
    'Trim\(',               # Text manipulation
    'LCase\(',
    'UCase\(',
    'Len\(',
    'InStr\(',
    'IsNumeric\(',
    'Format\(',
    
    # Basic UI Elements
    'StatusBar\s*=',        # Status updates
    'Application\.Caption',
    'DisplayScrollBars',
    'EnableResize',
    'DisplayHeadings',
    
    # Simple Validation
    'Validation\.Add',      # Data validation
    'Validation\.Delete',
    'Validation\.Modify',
    'InputMessage',
    'ErrorMessage',
    
    # Sheet Management
    'Worksheet\.Copy',      # Sheet operations
    'Worksheet\.Move',
    'Worksheet\.Name',
    'Sheets\.Count',
    'Worksheets\.Count',
    
    # Simple Cell References
    'Offset\(',             # Cell navigation
    'Address\(',
    'Column\(',
    'Row\(',
    'CurrentRegion',
    
    # Basic Functions
    'IsEmpty\(',            # Simple checks
    'WorksheetFunction\.IsNA',
    'WorksheetFunction\.IsError',
    'IsDate\(',
    
    # Display Formatting
    'NumberFormat',         # Number formatting
    'GeneralFormat',
    'PercentFormat',
    'TextFormat',
    
    # Simple Loops and Control
    'For\s+Each',          # Basic loops
    'Next\s+',
    'To\s+Step',
    'Do\s+While',
    'Do\s+Until',
    'Loop\s+',
    
    # Basic Variables
    'Dim\s+',              # Variable declaration
    'Set\s+',
    'Let\s+',
    'ReDim\s+',
    
    # Simple Conditionals
    'If\s+.*\s+Then',      # Basic logic
    'ElseIf',
    'End\s+If',
    'Select\s+Case',
    'Case\s+',
    
    # Comments and Documentation
    '^''.*$',              # Documentation
    '^REM\s+.*$',
    
    # Basic Sheet Protection
    'Protect',
    'Unprotect',
    'DisplayFormulaBar',
    'DisplayStatusBar',
    'EnableSelection',
    
    # Simple Arrays
    'LBound\(',            # Array operations
    'UBound\(',
    'Join\(',
    'Split\(',
    
    # Basic Date Operations
    'DateAdd\(',           # Date manipulation
    'DateDiff\(',
    'DatePart\(',
    'DateValue\(',
    'Now\(',
    'Date\(',
    
    # Basic Worksheet Functions
    'VLOOKUP\(',
    'HLOOKUP\(',
    'INDEX\(',
    'MATCH\(',
    'SUMIF\(',
    'COUNTIF\(',
    
    # Basic Print Operations
    'Preview',
    'PrintArea',
    'PrintTitleRows',
    'PrintTitleColumns'
)

# Function to extract macro content from Office files

function Get-MacroContent {
    param (
        [string]$filePath,
        [int]$timeoutSeconds = 30,  # Default timeout
        [switch]$Verbose
    )
    
    # Initialize timeout timer
    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    $app = $null
    $fileObj = $null
    
    try {
        # Check if file exists and is accessible
        if (-not (Test-Path $filePath)) {
            Write-Warning "File not found: $filePath"
            return $null
        }
        
        # Check if file is locked
        $fileInfo = New-Object System.IO.FileInfo($filePath)
        try {
            $stream = $fileInfo.Open([System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
            $stream.Close()
        }
        catch {
            Write-Warning "File is locked or inaccessible: $filePath"
            return $null
        }
        
        $extension = [System.IO.Path]::GetExtension($filePath).ToLower()
        
        switch ($extension) {
            # Word Documents
            { $_ -in @(".docm", ".doc", ".docx") } {
                $app = New-Object -ComObject Word.Application -ErrorAction Stop
                $app.Visible = $false
                $app.DisplayAlerts = 0
                
                Write-Verbose "Opening Word document: $filePath"
                
                # Add timeout check for file open
                $openScript = {
                    param($app, $path)
                    $app.Documents.Open($path)
                }
                
                $fileObj = Start-Job -ScriptBlock $openScript -ArgumentList $app, $filePath
                
                if (-not (Wait-Job $fileObj -Timeout $timeoutSeconds)) {
                    Stop-Job $fileObj
                    throw "Timeout while opening document"
                }
                
                $doc = Receive-Job $fileObj
                Remove-Job $fileObj
                
                $macroContent = ""
                
                # Timeout check for macro extraction
                if ($timer.Elapsed.TotalSeconds -gt $timeoutSeconds) {
                    throw "Operation timed out while processing macros"
                }
                
                if ($doc.VBProject -eq $null) {
                    Write-Verbose "No VBA Project found in document"
                    return $null
                }
                
                foreach ($component in $doc.VBProject.VBComponents) {
                    if ($timer.Elapsed.TotalSeconds -gt $timeoutSeconds) {
                        throw "Operation timed out while processing component"
                    }
                    
                    $codeModule = $component.CodeModule
                    $lineCount = $codeModule.CountOfLines
                    if ($lineCount -gt 0) {
                        $macroContent += "' Module: $($component.Name)`n"
                        $macroContent += $codeModule.Lines(1, $lineCount) + "`n`n"
                    }
                }
                
                $doc.Close($false)
            }
            
            # Excel Workbooks
            { $_ -in @(".xlsm", ".xls", ".xlam") } {
                $app = New-Object -ComObject Excel.Application -ErrorAction Stop
                $app.Visible = $false
                $app.DisplayAlerts = $false
                
                Write-Verbose "Opening Excel workbook: $filePath"
                
                # Add timeout check for file open
                $openScript = {
                    param($app, $path)
                    $app.Workbooks.Open($path)
                }
                
                $fileObj = Start-Job -ScriptBlock $openScript -ArgumentList $app, $filePath
                
                if (-not (Wait-Job $fileObj -Timeout $timeoutSeconds)) {
                    Stop-Job $fileObj
                    throw "Timeout while opening workbook"
                }
                
                $workbook = Receive-Job $fileObj
                Remove-Job $fileObj
                
                $macroContent = ""
                
                # Similar timeout checks for Excel macro extraction
                # [Excel specific code continues...]
            }
            
            # PowerPoint Presentations
            { $_ -in @(".pptm", ".ppt", ".pptx") } {
                $app = New-Object -ComObject PowerPoint.Application -ErrorAction Stop
                $app.Visible = [Microsoft.Office.Core.MsoTriState]::msoFalse
                
                Write-Verbose "Opening PowerPoint presentation: $filePath"
                
                # Add timeout check for file open
                $openScript = {
                    param($app, $path)
                    $app.Presentations.Open($path)
                }
                
                $fileObj = Start-Job -ScriptBlock $openScript -ArgumentList $app, $filePath
                
                if (-not (Wait-Job $fileObj -Timeout $timeoutSeconds)) {
                    Stop-Job $fileObj
                    throw "Timeout while opening presentation"
                }
                
                $presentation = Receive-Job $fileObj
                Remove-Job $fileObj
                
                $macroContent = ""
                
                # Similar timeout checks for PowerPoint macro extraction
                # [PowerPoint specific code continues...]
            }
            
            default {
                Write-Warning "Unsupported file type: $extension"
                return $null
            }
        }
        
        return $macroContent
    }
    catch {
        Write-Warning "Error processing $($filePath): $($_.Exception.Message)"
        return $null
    }
    finally {
        $timer.Stop()
        
        # Cleanup COM objects
        if ($fileObj) {
            try {
                switch ($extension) {
                    { $_ -in @(".docm", ".doc", ".docx") } { 
                        if ($doc) { $doc.Close($false) }
                    }
                    { $_ -in @(".xlsm", ".xls", ".xlam") } { 
                        if ($workbook) { $workbook.Close($false) }
                    }
                    { $_ -in @(".pptm", ".ppt", ".pptx") } { 
                        if ($presentation) { $presentation.Close() }
                    }
                }
            }
            catch {
                Write-Warning "Error during cleanup: $($_.Exception.Message)"
            }
        }
        
        if ($app) {
            try {
                $app.Quit()
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app) | Out-Null
                
                # Force cleanup of any remaining COM objects
                $null = [System.Runtime.Interopservices.Marshal]::GetActiveObject($app.GetType().FullName)
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($_) | Out-Null
            }
            catch {
                Write-Warning "Error releasing COM objects: $($_.Exception.Message)"
            }
        }
        
        # Force garbage collection
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        # Kill any hanging Office processes (use with caution)
        $processNames = @("WINWORD", "EXCEL", "POWERPNT")
        foreach ($proc in Get-Process | Where-Object {$processNames -contains $_.Name}) {
            if ((Get-Date) - $proc.StartTime -gt [TimeSpan]::FromMinutes(5)) {
                try {
                    $proc.Kill()
                }
                catch {
                    Write-Warning "Could not kill process $($proc.Name): $($_.Exception.Message)"
                }
            }
        }
    }
}

# Example usage with error handling
function Test-MacroScanner {
    param (
        [string]$path,
        [int]$timeout = 30
    )
    
    try {
        $result = Get-MacroContent -filePath $path -timeoutSeconds $timeout -Verbose
        if ($result) {
            Write-Output "Successfully extracted macros from $path"
            return $result
        }
        else {
            Write-Warning "No macros found or error occurred in $path"
            return $null
        }
    }
    catch {
        Write-Error "Critical error processing $path: $($_.Exception.Message)"
        return $null
    }
}



# Function to analyze macro content and determine risk level
function Analyze-MacroContent {
    param (
        [string]$macroContent  # VBA code content to analyze
    )
    
    # Initialize counters and collection for found patterns
    $foundPatterns = @()
    $highCount = 0
    $mediumCount = 0
    $lowCount = 0
    
    # Check for high risk patterns
    Write-Verbose "Checking for high risk patterns..."
    foreach ($pattern in $highRiskPatterns) {
        if ($macroContent -match $pattern) {
            $highCount++
            $foundPatterns += "$pattern (High)"
        }
    }
    
    # Check for medium risk patterns
    Write-Verbose "Checking for medium risk patterns..."
    foreach ($pattern in $mediumRiskPatterns) {
        if ($macroContent -match $pattern) {
            $mediumCount++
            $foundPatterns += "$pattern (Medium)"
        }
    }
    
    # Check for low risk patterns
    Write-Verbose "Checking for low risk patterns..."
    foreach ($pattern in $lowRiskPatterns) {
        if ($macroContent -match $pattern) {
            $lowCount++
            $foundPatterns += "$pattern (Low)"
        }
    }
    
    # Find any other commands/functions not in our lists
    Write-Verbose "Checking for unlisted commands..."
    $allCommands = [regex]::Matches($macroContent, '([A-Za-z0-9_]+)\s*\(') | 
                   ForEach-Object { $_.Groups[1].Value } | 
                   Sort-Object -Unique
    
    foreach ($cmd in $allCommands) {
        if (($cmd -notin $highRiskPatterns) -and 
            ($cmd -notin $mediumRiskPatterns) -and 
            ($cmd -notin $lowRiskPatterns)) {
            $highCount++
            $foundPatterns += "$cmd (Unlisted - High)"
        }
    }
    
    # Determine overall risk level based on rules:
    # - High if any high risk patterns or 3+ medium risk patterns
    # - Medium if 1-2 medium risk patterns
    # - Low otherwise
    $riskLevel = "Low"
    if ($highCount -gt 0 -or $mediumCount -ge 3) {
        $riskLevel = "High"
    }
    elseif ($mediumCount -gt 0) {
        $riskLevel = "Medium"
    }
    
    return @{
        RiskLevel = $riskLevel
        FoundPatterns = ($foundPatterns | Sort-Object -Unique)
        HighCount = $highCount
        MediumCount = $mediumCount
        LowCount = $lowCount
    }
}

# Main execution block
Write-Host "Macro Security Analysis Tool" -ForegroundColor Cyan
Write-Host "===========================" -ForegroundColor Cyan

# Get search path from user
$searchPath = Read-Host "Enter the path to search for Office files"

Write-Host "`nSearching for Office files with macros..." -ForegroundColor Yellow

# Find all Office files that might contain macros
$files = Get-ChildItem -Path $searchPath -Recurse -Include @(
    "*.xlsm",  # Excel macro-enabled workbook
    "*.xltm",  # Excel macro-enabled template
    "*.xlam",  # Excel add-in
    "*.docm",  # Word macro-enabled document
    "*.dotm",  # Word macro-enabled template
    "*.pptm"   # PowerPoint macro-enabled presentation
) -ErrorAction SilentlyContinue

# Initialize results array
$results = @()

# Process each file
foreach ($file in $files) {
    Write-Host "Analyzing $($file.Name)..." -ForegroundColor Green
    
    # Extract macro content
    $macroContent = Get-MacroContent -filePath $file.FullName
    
    if ($macroContent) {
        # Analyze the macro content
        $analysis = Analyze-MacroContent -macroContent $macroContent
        
        # Add results to array
        $results += [PSCustomObject]@{
            FilePath = $file.FullName
            FileName = $file.Name
            LastModified = $file.LastWriteTime
            RiskLevel = $analysis.RiskLevel
            HighRiskCount = $analysis.HighCount
            MediumRiskCount = $analysis.MediumCount
            LowRiskCount = $analysis.LowCount
            IdentifiedPatterns = ($analysis.FoundPatterns | ConvertTo-Json -Compress)
        }
    }
}

# Generate output filename with timestamp
$outputPath = Join-Path $PWD "MacroAnalysis_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

# Export results to CSV
$results | Export-Csv -Path $outputPath -NoTypeInformation

# Display summary
Write-Host "`nAnalysis complete. Results saved to: $outputPath" -ForegroundColor Green
Write-Host "Summary:" -ForegroundColor Cyan
Write-Host "Total files analyzed: $($results.Count)"
Write-Host "High risk files: $(($results | Where-Object RiskLevel -eq 'High').Count)" -ForegroundColor Red
Write-Host "Medium risk files: $(($results | Where-Object RiskLevel -eq 'Medium').Count)" -ForegroundColor Yellow
Write-Host "Low risk files: $(($results | Where-Object RiskLevel -eq 'Low').Count)" -ForegroundColor Green

# Display warning if high risk files found
if (($results | Where-Object RiskLevel -eq 'High').Count -gt 0) {
    Write-Host "`nWARNING: High risk files detected! Please review the detailed report." -ForegroundColor Red
}
