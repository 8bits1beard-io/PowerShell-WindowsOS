<#
.SYNOPSIS
Detects and reports corrupt Windows Management Instrumentation (WMI) components.

.DESCRIPTION
This script performs basic WMI health checks and reports the number of corruption issues found.
It validates WMI repository integrity, checks WMI services, and tests core WMI namespaces and classes
without making any system changes.

.PARAMETER LogPath
The directory path where diagnostic logs will be stored. Defaults to C:\Windows\Logs.

.PARAMETER DetailedOutput
Shows detailed information about each issue found. Defaults to $false.

.EXAMPLE
.\Test-WMICorruption.ps1
Runs basic WMI corruption check and reports issue count.

.EXAMPLE
.\Test-WMICorruption.ps1 -DetailedOutput
Runs WMI corruption check with detailed output of all issues.

.NOTES
Author: 8bits1beard
Date: 2025-01-15
Version: v1.0.0
Source: ../PoSh-Best-Practice/

.LINK
../PoSh-Best-Practice/
#>

[CmdletBinding()]
param(
    [Parameter()]
    [ValidateNotNullOrEmpty()]
    [string]$LogPath = "C:\Windows\Logs",
    
    [Parameter()]
    [switch]$DetailedOutput
)

# Requires administrative privileges
#Requires -RunAsAdministrator

# Import required modules for system diagnostics
Import-Module CimCmdlets -ErrorAction SilentlyContinue

# Define script-level variables
$script:LogFileName = "WMI-CorruptionCheck.log"
$script:IssuesFound = @()

# Standardized logging function as per repository instructions
function Write-LogMessage {
    <#
    .SYNOPSIS
    Logs messages to a specified file with structured levels and auto-rotation.

    .DESCRIPTION
    Writes structured JSON log entries, emits relevant output, and rotates files if oversized.

    .PARAMETER LogLevel
    Valid values: Verbose, Warning, Error, Information, Debug.

    .PARAMETER Message
    The message content.

    .PARAMETER LogPath
    Folder for logs. Default: C:\Windows\Logs

    .PARAMETER LogFileName
    Default: WMI-CorruptionCheck.log

    .PARAMETER MaxFileSizeMB
    Threshold for file rotation. Default: 5 MB

    .EXAMPLE
    Write-LogMessage -LogLevel "Information" -Message "Started process."

    .NOTES
    Author: 8bits1beard
    Created: 2025-01-15
    #>
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'Low')]
    param (
        [Parameter(Mandatory, Position = 0)]
        [ValidateSet("Verbose", "Warning", "Error", "Information", "Debug")]
        [string]$LogLevel,

        [Parameter(Mandatory, Position = 1)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,

        [Parameter(Position = 2)]
        [string]$LogPath = "C:\Windows\Logs",

        [Parameter(Position = 3)]
        [string]$LogFileName = "WMI-CorruptionCheck.log",

        [Parameter(Position = 4)]
        [int]$MaxFileSizeMB = 5
    )

    begin {
        $LogFile = Join-Path -Path $LogPath -ChildPath $LogFileName
        if (-not (Test-Path $LogPath)) {
            New-Item -ItemType Directory -Path $LogPath -Force | Out-Null
        }
        if (Test-Path $LogFile) {
            $FileSizeMB = (Get-Item $LogFile).Length / 1MB
            if ($FileSizeMB -ge $MaxFileSizeMB) {
                $Timestamp = Get-Date -Format "yyyyMMddHHmmss"
                Rename-Item -Path $LogFile -NewName "$LogPath\$LogFileName`_$Timestamp.log"
            }
        }
    }

    process {
        $LogEntry = @{
            Timestamp = Get-Date -Format "yyyy-MM-ddTHH:mm:ss"
            Level     = $LogLevel
            Message   = $Message
        }
        $LogEntryJSON = $LogEntry | ConvertTo-Json -Depth 2 -Compress

        try {
            Add-Content -Path $LogFile -Value $LogEntryJSON -ErrorAction Stop
        } catch {
            Write-Warning "Failed to log: $($_.Exception.Message)"
        }

        switch ($LogLevel) {
            "Verbose"     { if ($VerbosePreference -ne "SilentlyContinue") { Write-Verbose $Message } }
            "Warning"     { Write-Warning $Message }
            "Error"       { Write-Error $Message }
            "Information" { Write-Information -MessageData $LogEntryJSON -InformationAction Continue }
            "Debug"       { if ($DebugPreference -ne "SilentlyContinue") { Write-Debug $Message } }
        }
    }
}

# Function to check WMI service status
function Test-WMIService {
    <#
    .SYNOPSIS
    Tests WMI service status and adds issues to the global issues array.
    
    .DESCRIPTION
    Checks if the WMI service exists and is running properly.
    #>
    [CmdletBinding()]
    param()
    
    Write-LogMessage -LogLevel "Information" -Message "Checking WMI service status" -LogPath $LogPath -LogFileName $script:LogFileName
    
    # Check WMI service status
    $wmiService = Get-Service -Name 'Winmgmt' -ErrorAction SilentlyContinue
    if ($wmiService) {
        if ($wmiService.Status -ne 'Running') {
            $issue = "WMI Service is not running (Status: $($wmiService.Status))"
            $script:IssuesFound += $issue
            Write-LogMessage -LogLevel "Error" -Message $issue -LogPath $LogPath -LogFileName $script:LogFileName
        } else {
            Write-LogMessage -LogLevel "Information" -Message "WMI Service is running normally" -LogPath $LogPath -LogFileName $script:LogFileName
        }
    } else {
        $issue = "WMI Service not found on system"
        $script:IssuesFound += $issue
        Write-LogMessage -LogLevel "Error" -Message $issue -LogPath $LogPath -LogFileName $script:LogFileName
    }
}

# Function to test WMI repository consistency
function Test-WMIRepositoryConsistency {
    <#
    .SYNOPSIS
    Tests WMI repository consistency using winmgmt tool.
    
    .DESCRIPTION
    Runs winmgmt /verifyrepository to check repository integrity.
    #>
    [CmdletBinding()]
    param()
    
    Write-LogMessage -LogLevel "Information" -Message "Testing WMI repository consistency" -LogPath $LogPath -LogFileName $script:LogFileName
    
    try {
        $consistencyCheck = & winmgmt /verifyrepository 2>&1
        
        if ($consistencyCheck -like "*consistent*") {
            Write-LogMessage -LogLevel "Information" -Message "WMI repository consistency check passed" -LogPath $LogPath -LogFileName $script:LogFileName
        } else {
            $issue = "WMI repository consistency check failed: $consistencyCheck"
            $script:IssuesFound += $issue
            Write-LogMessage -LogLevel "Error" -Message $issue -LogPath $LogPath -LogFileName $script:LogFileName
        }
    }
    catch {
        $issue = "Failed to verify WMI repository: $_"
        $script:IssuesFound += $issue
        Write-LogMessage -LogLevel "Error" -Message $issue -LogPath $LogPath -LogFileName $script:LogFileName
    }
}

# Function to test core WMI namespaces
function Test-WMINamespaces {
    <#
    .SYNOPSIS
    Tests accessibility of core WMI namespaces.
    
    .DESCRIPTION
    Attempts to query essential WMI namespaces to verify they are accessible.
    #>
    [CmdletBinding()]
    param()
    
    Write-LogMessage -LogLevel "Information" -Message "Testing core WMI namespaces" -LogPath $LogPath -LogFileName $script:LogFileName
    
    # Test core WMI namespaces
    $coreNamespaces = @('root', 'root\cimv2', 'root\default', 'root\subscription')
    
    foreach ($namespace in $coreNamespaces) {
        try {
            $null = Get-CimInstance -Namespace $namespace -ClassName __NAMESPACE -ErrorAction Stop
            Write-LogMessage -LogLevel "Debug" -Message "Namespace '$namespace' is accessible" -LogPath $LogPath -LogFileName $script:LogFileName
        }
        catch {
            $issue = "Cannot access WMI namespace: $namespace"
            $script:IssuesFound += $issue
            Write-LogMessage -LogLevel "Error" -Message "$issue - $_" -LogPath $LogPath -LogFileName $script:LogFileName
        }
    }
}

# Function to test critical WMI classes
function Test-WMIClasses {
    <#
    .SYNOPSIS
    Tests functionality of critical WMI classes.
    
    .DESCRIPTION
    Attempts to query essential WMI classes to verify they are functional.
    #>
    [CmdletBinding()]
    param()
    
    Write-LogMessage -LogLevel "Information" -Message "Testing critical WMI classes" -LogPath $LogPath -LogFileName $script:LogFileName
    
    # Test critical WMI classes
    $criticalClasses = @(
        @{Namespace='root\cimv2'; ClassName='Win32_OperatingSystem'},
        @{Namespace='root\cimv2'; ClassName='Win32_ComputerSystem'},
        @{Namespace='root\cimv2'; ClassName='Win32_Process'}
    )
    
    foreach ($class in $criticalClasses) {
        try {
            $null = Get-CimInstance -Namespace $class.Namespace -ClassName $class.ClassName -ErrorAction Stop | Select-Object -First 1
            Write-LogMessage -LogLevel "Debug" -Message "WMI class '$($class.ClassName)' is functional" -LogPath $LogPath -LogFileName $script:LogFileName
        }
        catch {
            $issue = "Cannot query WMI class: $($class.ClassName)"
            $script:IssuesFound += $issue
            Write-LogMessage -LogLevel "Error" -Message "$issue - $_" -LogPath $LogPath -LogFileName $script:LogFileName
        }
    }
}

# Function to display results summary
function Show-WMICorruptionSummary {
    <#
    .SYNOPSIS
    Displays a summary of WMI corruption check results.
    
    .DESCRIPTION
    Shows the total number of issues found and optionally lists each issue.
    #>
    [CmdletBinding()]
    param()
    
    Write-Host "`n" -NoNewline
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "       WMI CORRUPTION CHECK RESULTS        " -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host ""
    
    # System Information
    Write-Host "System Information:" -ForegroundColor Yellow
    Write-Host "  Computer Name : $env:COMPUTERNAME"
    Write-Host "  Scan Time     : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Write-Host ""
    
    # Results
    if ($script:IssuesFound.Count -eq 0) {
        Write-Host "WMI Corruption Issues Found: " -ForegroundColor Yellow -NoNewline
        Write-Host "0" -ForegroundColor Green
        Write-Host ""
        Write-Host "✓ No WMI corruption detected!" -ForegroundColor Green
        Write-Host "  WMI appears to be functioning normally." -ForegroundColor Green
    }
    else {
        Write-Host "WMI Corruption Issues Found: " -ForegroundColor Yellow -NoNewline
        Write-Host "$($script:IssuesFound.Count)" -ForegroundColor Red
        Write-Host ""
        
        if ($DetailedOutput) {
            Write-Host "Issue Details:" -ForegroundColor Red
            $issueNumber = 1
            foreach ($issue in $script:IssuesFound) {
                Write-Host "  $issueNumber. $issue" -ForegroundColor Red
                $issueNumber++
            }
            Write-Host ""
        }
        
        Write-Host "⚠ WMI corruption detected!" -ForegroundColor Red
        Write-Host "  Consider running a WMI repair tool to fix these issues." -ForegroundColor Yellow
    }
    
    Write-Host ""
    Write-Host "Log file saved to:" -ForegroundColor Cyan
    Write-Host "  $(Join-Path -Path $LogPath -ChildPath $script:LogFileName)" -ForegroundColor Gray
    Write-Host ""
    Write-Host "============================================" -ForegroundColor Cyan
}

# Main execution block
try {
    Write-Host ""
    Write-Host "WMI Corruption Detection Tool" -ForegroundColor Green
    Write-Host "=============================" -ForegroundColor Green
    Write-Host ""
    
    # Initialize logging
    Write-LogMessage -LogLevel "Information" -Message "Starting WMI corruption detection scan" -LogPath $LogPath -LogFileName $script:LogFileName
    Write-LogMessage -LogLevel "Information" -Message "Parameters: LogPath=$LogPath, DetailedOutput=$DetailedOutput" -LogPath $LogPath -LogFileName $script:LogFileName
    
    # Run WMI corruption checks
    Write-Host "Scanning for WMI corruption..." -ForegroundColor Cyan
    
    # Check WMI service
    Test-WMIService
    
    # Check repository consistency
    Test-WMIRepositoryConsistency
    
    # Check namespaces
    Test-WMINamespaces
    
    # Check critical classes
    Test-WMIClasses
    
    # Display results
    Show-WMICorruptionSummary
    
    Write-LogMessage -LogLevel "Information" -Message "WMI corruption scan completed. Issues found: $($script:IssuesFound.Count)" -LogPath $LogPath -LogFileName $script:LogFileName
}
catch {
    Write-Host ""
    Write-Host "CRITICAL ERROR: $_" -ForegroundColor Red
    Write-LogMessage -LogLevel "Error" -Message "Critical error during WMI corruption scan: $_" -LogPath $LogPath -LogFileName $script:LogFileName
    
    Write-Host ""
    Write-Host "Log file: $(Join-Path -Path $LogPath -ChildPath $script:LogFileName)" -ForegroundColor Gray
    
    # Exit with error code
    exit 1
}

# Exit with appropriate code based on findings
if ($script:IssuesFound.Count -gt 0) {
    exit 1  # Issues found
} else {
    exit 0  # No issues found
}
