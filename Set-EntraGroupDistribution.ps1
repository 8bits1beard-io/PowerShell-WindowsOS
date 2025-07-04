<#
.SYNOPSIS
Interactively distributes device members from a source Entra group to target groups based on specified percentages.

.DESCRIPTION
This script provides an interactive interface to configure and execute the distribution of device members
from a source Azure AD/Entra group to multiple target groups based on user-defined percentages.
The script validates configuration, connects to Microsoft Graph, and uses the AzureGroupStuff module.

.PARAMETER LogPath
The directory path where diagnostic logs will be stored. Defaults to C:\Windows\Logs.

.PARAMETER SkipModuleInstall
Skip automatic installation of required modules. Defaults to $false.

.EXAMPLE
.\Set-EntraGroupDistribution.ps1
Runs the interactive group distribution configuration.

.EXAMPLE
.\Set-EntraGroupDistribution.ps1 -SkipModuleInstall
Runs without attempting to install required modules.

.NOTES
Author: 8bits1beard
Date: 2024-01-26
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
    [switch]$SkipModuleInstall
)

# Define script-level variables
$script:LogFileName = "EntraGroupDistribution.log"
$script:RequiredModules = @(
    'AzureGroupStuff',
    'Microsoft.Graph.Authentication',
    'Microsoft.Graph.Groups'
)

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
    Default: EntraGroupDistribution.log

    .PARAMETER MaxFileSizeMB
    Threshold for file rotation. Default: 5 MB

    .EXAMPLE
    Write-LogMessage -LogLevel "Information" -Message "Started process."

    .NOTES
    Author: 8bits1beard
    Created: 2025-01-26
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
        [string]$LogFileName = "EntraGroupDistribution.log",

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

# Function to install required PowerShell modules
function Install-RequiredModules {
    <#
    .SYNOPSIS
    Installs required PowerShell modules for Entra group management.
    
    .DESCRIPTION
    Checks for and installs the necessary modules for Azure AD/Entra group operations.
    #>
    [CmdletBinding()]
    param()
    
    Write-LogMessage -LogLevel "Information" -Message "Checking required modules" -LogPath $LogPath -LogFileName $script:LogFileName
    
    foreach ($module in $script:RequiredModules) {
        try {
            Write-Host "Checking module: $module..." -ForegroundColor Yellow
            
            if (-not (Get-Module -ListAvailable -Name $module)) {
                Write-Host "  Installing $module..." -ForegroundColor Cyan
                Install-Module -Name $module -Force -Scope CurrentUser
                Write-LogMessage -LogLevel "Information" -Message "Successfully installed module: $module" -LogPath $LogPath -LogFileName $script:LogFileName
            } else {
                Write-Host "  ✓ $module already installed" -ForegroundColor Green
                Write-LogMessage -LogLevel "Debug" -Message "Module already available: $module" -LogPath $LogPath -LogFileName $script:LogFileName
            }
        }
        catch {
            $errorMsg = "Failed to install module $module`: $($_.Exception.Message)"
            Write-Host "  ✗ $errorMsg" -ForegroundColor Red
            Write-LogMessage -LogLevel "Error" -Message $errorMsg -LogPath $LogPath -LogFileName $script:LogFileName
            throw $errorMsg
        }
    }
}

# Function to get user input for source group
function Get-SourceGroupConfiguration {
    <#
    .SYNOPSIS
    Prompts user for source group configuration.
    
    .DESCRIPTION
    Interactive prompt to gather source group ID from the user.
    #>
    [CmdletBinding()]
    param()
    
    Write-Host "`n" -NoNewline
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "         SOURCE GROUP CONFIGURATION        " -ForegroundColor Cyan  
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host ""
    
    do {
        $sourceGroup = Read-Host "Enter the source Entra group ID (GUID format)"
        
        # Validate GUID format
        if ($sourceGroup -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
            Write-Host "✓ Valid group ID format" -ForegroundColor Green
            Write-LogMessage -LogLevel "Information" -Message "Source group configured: $sourceGroup" -LogPath $LogPath -LogFileName $script:LogFileName
            return $sourceGroup
        } else {
            Write-Host "✗ Invalid GUID format. Please enter a valid group ID." -ForegroundColor Red
        }
    } while ($true)
}

# Function to get target groups configuration
function Get-TargetGroupsConfiguration {
    <#
    .SYNOPSIS
    Prompts user for target groups and percentages configuration.
    
    .DESCRIPTION
    Interactive prompt to gather target group IDs and their distribution percentages.
    #>
    [CmdletBinding()]
    param()
    
    Write-Host "`n" -NoNewline
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "        TARGET GROUPS CONFIGURATION        " -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Configure target groups and their percentage allocation." -ForegroundColor Yellow
    Write-Host "Note: Total percentages should equal 100%." -ForegroundColor Yellow
    Write-Host ""
    
    $targetGroups = [ordered]@{}
    $totalPercentage = 0
    $groupCount = 1
    
    do {
        Write-Host "Target Group #$groupCount" -ForegroundColor Cyan
        
        # Get group ID
        do {
            $groupId = Read-Host "  Enter group ID (GUID format)"
            if ($groupId -match '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') {
                break
            } else {
                Write-Host "  ✗ Invalid GUID format. Please try again." -ForegroundColor Red
            }
        } while ($true)
        
        # Get percentage
        do {
            $percentageInput = Read-Host "  Enter percentage for this group (1-100)"
            if ([int]::TryParse($percentageInput, [ref]$percentage) -and $percentage -ge 1 -and $percentage -le 100) {
                if (($totalPercentage + $percentage) -le 100) {
                    break
                } else {
                    Write-Host "  ✗ Total percentage would exceed 100%. Remaining: $(100 - $totalPercentage)%" -ForegroundColor Red
                }
            } else {
                Write-Host "  ✗ Please enter a valid percentage (1-100)." -ForegroundColor Red
            }
        } while ($true)
        
        # Add to configuration
        $targetGroups[$groupId] = $percentage
        $totalPercentage += $percentage
        
        Write-Host "  ✓ Added group with $percentage% allocation" -ForegroundColor Green
        Write-Host "  Current total: $totalPercentage%" -ForegroundColor Yellow
        
        # Check if user wants to add more groups
        if ($totalPercentage -lt 100) {
            $addMore = Read-Host "`n  Add another target group? (y/n)"
            if ($addMore -notmatch '^[yY]') {
                break
            }
            $groupCount++
        } else {
            Write-Host "`n✓ Configuration complete (100% allocated)" -ForegroundColor Green
            break
        }
    } while ($totalPercentage -lt 100)
    
    Write-LogMessage -LogLevel "Information" -Message "Target groups configured: $($targetGroups.Count) groups, Total percentage: $totalPercentage%" -LogPath $LogPath -LogFileName $script:LogFileName
    
    return $targetGroups
}

# Function to display configuration summary
function Show-ConfigurationSummary {
    <#
    .SYNOPSIS
    Displays a summary of the current configuration.
    
    .DESCRIPTION
    Shows the source group and target groups configuration for user confirmation.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SourceGroup,
        
        [Parameter(Mandatory)]
        [hashtable]$TargetGroups
    )
    
    Write-Host "`n" -NoNewline
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "          CONFIGURATION SUMMARY            " -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host ""
    
    Write-Host "Source Group:" -ForegroundColor Yellow
    Write-Host "  $SourceGroup" -ForegroundColor White
    Write-Host ""
    
    Write-Host "Target Groups:" -ForegroundColor Yellow
    $groupNumber = 1
    foreach ($group in $TargetGroups.GetEnumerator()) {
        Write-Host "  $groupNumber. $($group.Key) → $($group.Value)%" -ForegroundColor White
        $groupNumber++
    }
    
    $totalPercentage = ($TargetGroups.Values | Measure-Object -Sum).Sum
    Write-Host ""
    Write-Host "Total Allocation: $totalPercentage%" -ForegroundColor $(if ($totalPercentage -eq 100) { "Green" } else { "Red" })
    Write-Host ""
    Write-Host "============================================" -ForegroundColor Cyan
}

# Function to execute group distribution
function Invoke-GroupDistribution {
    <#
    .SYNOPSIS
    Executes the group distribution using the AzureGroupStuff module.
    
    .DESCRIPTION
    Connects to Microsoft Graph and executes the Set-AzureRingGroup function.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$SourceGroup,
        
        [Parameter(Mandatory)]
        [hashtable]$TargetGroups
    )
    
    Write-Host "`n" -NoNewline
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host "           EXECUTING DISTRIBUTION           " -ForegroundColor Cyan
    Write-Host "============================================" -ForegroundColor Cyan
    Write-Host ""
    
    try {
        # Connect to Microsoft Graph
        Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Yellow
        Write-LogMessage -LogLevel "Information" -Message "Attempting to connect to Microsoft Graph" -LogPath $LogPath -LogFileName $script:LogFileName
        
        $graphScopes = @(
            "Device.Read.All",
            "User.Read.All", 
            "Group.ReadWrite.All",
            "DeviceManagementManagedDevices.Read.All"
        )
        
        Connect-MgGraph -Scopes $graphScopes
        Write-Host "✓ Connected to Microsoft Graph" -ForegroundColor Green
        Write-LogMessage -LogLevel "Information" -Message "Successfully connected to Microsoft Graph" -LogPath $LogPath -LogFileName $script:LogFileName
        
        # Execute group distribution
        Write-Host "`nExecuting group distribution..." -ForegroundColor Yellow
        Write-LogMessage -LogLevel "Information" -Message "Starting group distribution execution" -LogPath $LogPath -LogFileName $script:LogFileName
        
        $setAzureRingGroupParams = @{
            rootGroup = $SourceGroup
            ringGroupConfig = $TargetGroups
            memberType = 'Device'
            forceRecalculate = $true
            skipUnderscoreInNameCheck = $true
            skipDescriptionUpdate = $true
            Verbose = $true
        }
        
        Set-AzureRingGroup @setAzureRingGroupParams
        
        Write-Host "✓ Group distribution completed successfully" -ForegroundColor Green
        Write-LogMessage -LogLevel "Information" -Message "Group distribution completed successfully" -LogPath $LogPath -LogFileName $script:LogFileName
        
    }
    catch {
        $errorMsg = "Failed to execute group distribution: $($_.Exception.Message)"
        Write-Host "✗ $errorMsg" -ForegroundColor Red
        Write-LogMessage -LogLevel "Error" -Message $errorMsg -LogPath $LogPath -LogFileName $script:LogFileName
        throw
    }
}

# Main execution block
try {
    Write-Host ""
    Write-Host "Entra Group Distribution Tool" -ForegroundColor Green
    Write-Host "=============================" -ForegroundColor Green
    Write-Host "Distributes device members from a source group to target groups based on percentages." -ForegroundColor Gray
    Write-Host ""
    
    # Initialize logging
    Write-LogMessage -LogLevel "Information" -Message "Starting Entra Group Distribution Tool" -LogPath $LogPath -LogFileName $script:LogFileName
    Write-LogMessage -LogLevel "Information" -Message "Parameters: LogPath=$LogPath, SkipModuleInstall=$SkipModuleInstall" -LogPath $LogPath -LogFileName $script:LogFileName
    
    # Install required modules if not skipped
    if (-not $SkipModuleInstall) {
        Write-Host "Installing required modules..." -ForegroundColor Cyan
        Install-RequiredModules
        Write-Host "✓ Module installation completed" -ForegroundColor Green
        Write-Host ""
    } else {
        Write-LogMessage -LogLevel "Information" -Message "Skipping module installation as requested" -LogPath $LogPath -LogFileName $script:LogFileName
    }
    
    # Get configuration from user
    $sourceGroup = Get-SourceGroupConfiguration
    $targetGroups = Get-TargetGroupsConfiguration
    
    # Show configuration summary
    Show-ConfigurationSummary -SourceGroup $sourceGroup -TargetGroups $targetGroups
    
    # Confirm execution
    $confirmation = Read-Host "`nProceed with group distribution? (y/n)"
    if ($confirmation -match '^[yY]') {
        Invoke-GroupDistribution -SourceGroup $sourceGroup -TargetGroups $targetGroups
        
        Write-Host "`n✓ Process completed successfully!" -ForegroundColor Green
        Write-Host "Log file: $(Join-Path -Path $LogPath -ChildPath $script:LogFileName)" -ForegroundColor Gray
    } else {
        Write-Host "Operation cancelled by user." -ForegroundColor Yellow
        Write-LogMessage -LogLevel "Information" -Message "Operation cancelled by user" -LogPath $LogPath -LogFileName $script:LogFileName
    }
    
}
catch {
    Write-Host ""
    Write-Host "CRITICAL ERROR: $_" -ForegroundColor Red
    Write-LogMessage -LogLevel "Error" -Message "Critical error in Entra Group Distribution Tool: $_" -LogPath $LogPath -LogFileName $script:LogFileName
    
    Write-Host ""
    Write-Host "Log file: $(Join-Path -Path $LogPath -ChildPath $script:LogFileName)" -ForegroundColor Gray
    
    # Exit with error code
    exit 1
}
    
    # Exit with error code
    exit 1
}
