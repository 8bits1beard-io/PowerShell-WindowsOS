<#
.SYNOPSIS
Adds system information text overlay to a wallpaper image.

.DESCRIPTION
This script overlays system information (computer name, serial number, uptime, etc.) 
onto a background image and optionally sets it as the desktop wallpaper. Uses readable
fonts optimized for wallpaper display.

.PARAMETER BackgroundImagePath
Path to the background image file.

.PARAMETER OutputImagePath
Path where the modified image will be saved.

.PARAMETER FontName
The font to be used for the text overlay. Default is Consolas for readability.

.PARAMETER Size
The font size for the text overlay. Default is 16.

.PARAMETER AntiAlias
Specifies whether to enable anti-aliasing for the text. Default is $true.

.PARAMETER SetAsDesktopBackground
Switch to set the modified image as the active desktop background.

.EXAMPLE
Add-TextToImage -BackgroundImagePath "C:\Images\background.jpg" -OutputImagePath "C:\Images\output.jpg"

.EXAMPLE
Add-TextToImage -BackgroundImagePath "C:\Images\bg.png" -OutputImagePath "C:\Images\info_bg.png" -SetAsDesktopBackground

.NOTES
Author: 8bits1beard
Date: 2025-01-26
Version: v1.0.0
Source: See project documentation or https://github.com/8bits1beard/PoSh-Best-Practice/ for more information.

.LINK
../PoSh-Best-Practice/
#>

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
    Default: WallpaperSystemInfo.log

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
        [string]$LogFileName = "WallpaperSystemInfo.log",

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

function Add-TextToImage {   
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0, HelpMessage = "Path to the background image.")]
        [ValidateScript({ Test-Path $_ -PathType Leaf })]
        [String]$BackgroundImagePath,

        [Parameter(Mandatory = $true, Position = 1, HelpMessage = "Path to save the modified image.")]
        [String]$OutputImagePath,

        [Parameter(Position = 2, HelpMessage = "The font to be used for the text.")]
        [String]$FontName = "Consolas",

        [Parameter(Position = 3, HelpMessage = "The font size.")]
        [int]$Size = 16,

        [Parameter(Position = 4, HelpMessage = "Specifies whether to enable anti-aliasing for the text.")]
        [bool]$AntiAlias = $true,

        [Parameter(Position = 5, HelpMessage = "Specifies whether to set the modified image as the active desktop background.")]
        [switch]$SetAsDesktopBackground
    )
    
    Write-LogMessage -LogLevel "Information" -Message "Starting Add-TextToImage function"
    
    # Get machine information
    $machineName = $env:COMPUTERNAME
    $serialNumber = (Get-CimInstance -Query "SELECT * FROM Win32_BIOS").SerialNumber
    $windowsBuild = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").CurrentBuild
    $roleType = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Workstation\Build").RoleType

    # Additional machine information variables
    $LastBootUpTime = (Get-CimInstance -ClassName Win32_OperatingSystem).LastBootUpTime
    $SystemUpTime = (Get-Date) - $LastBootUpTime
    $Days = $SystemUpTime.Days
    $Hours = $SystemUpTime.Hours
    $Minutes = $SystemUpTime.Minutes
    $Seconds = $SystemUpTime.Seconds
    $uptime = "$Days days, $Hours hours, $Minutes minutes"
    $domainName = (Get-CimInstance -ClassName Win32_ComputerSystem).Domain 
    $ipAddress = (Get-CimInstance -Query "SELECT IPAddress FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True").IPAddress | Where-Object { $_ -match '^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$' } | Select-Object -First 1

    # Get the current logon user
    $currentLogonUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

    # Create a hashtable to store the variable names and their corresponding values
    $variables = [ordered]@{
        "Uptime "        = $uptime
        "Current User "  = $currentLogonUser
        "Machine Name "  = $machineName
        "Serial Number " = $serialNumber
        "Windows Build " = $windowsBuild
        "Role Type "     = $roleType
        "Domain Name "   = $domainName
        "IP Address "    = $ipAddress
    }

  
    # Create a new Graphics object
    $image = [System.Drawing.Image]::FromFile($BackgroundImagePath)
    $graphic = [System.Drawing.Graphics]::FromImage($image)

    # Set the font properties
    $font = New-Object System.Drawing.Font($FontName, $Size, [System.Drawing.FontStyle]::Regular, [System.Drawing.GraphicsUnit]::Pixel)
    $brush = New-Object System.Drawing.SolidBrush([System.Drawing.Color]::White)

    # Set anti-aliasing
    if ($AntiAlias) {
        $graphic.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    }
   
    # Dynamically calculate starting positions based on image dimensions
    $marginRight = 50
    $marginTop = 50
    $positionX = $image.Width - $marginRight - 400
    if ($positionX -lt 0) { $positionX = $marginRight }
    $positionY = $marginTop

    foreach ($variable in $variables.GetEnumerator()) {
        $variableName = $variable.Key
        $variableValue = $variable.Value

        # Draw the text
        $position = New-Object System.Drawing.PointF($positionX, $positionY)
        $graphic.DrawString("${variableName}: $variableValue", $font, $brush, $position)

        # Increase the Y position for the next variable
        $positionY += 20
    }

    # Save the modified image
    $image.Save($OutputImagePath)

    # Clean up
    $graphic.Dispose()
    $font.Dispose()
    $brush.Dispose()

    # Set the wallpaper and refresh desktop only if -SetAsDesktopBackground is specified
    if ($SetAsDesktopBackground) {
        # Set the wallpaper path in the registry
        Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name Wallpaper -Value $OutputImagePath

        # Set the wallpaper style in the registry
        Set-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -Value 2

        # Refresh the desktop
        $user32Dll = Add-Type -MemberDefinition @"
            [DllImport("user32.dll", CharSet = CharSet.Auto)]
        $SPI_SETDESKWALLPAPER = 20
        $SPIF_UPDATEINIFILE = 0x01
        $SPIF_SENDCHANGE = 0x02
        $SPI_SETDESKWALLPAPER = 20
        $SPIF_UPDATEINIFILE = 0x01
        $SPIF_SENDCHANGE = 0x02
        $user32Dll::SystemParametersInfo($SPI_SETDESKWALLPAPER, 0, $OutputImagePath, $SPIF_UPDATEINIFILE -bor $SPIF_SENDCHANGE)
}

# Export module functions for public use
Export-ModuleMember -Function Add-TextToImage, Write-LogMessage
