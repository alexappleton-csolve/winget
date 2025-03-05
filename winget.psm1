<#
.SYNOPSIS
    PowerShell module for the winget package manager providing functions for installing, upgrading, and managing applications.

.DESCRIPTION
    This module provides PowerShell integration for the Windows Package Manager (winget), enabling automation of package management tasks.
    Includes functions for installing, upgrading, uninstalling, and listing applications using the winget CLI.
    
    Inspired by:
    - https://github.com/Romanitho/Winget-AutoUpdate
    - https://github.com/jdhitsolutions/WingetTools
    - https://docs.microsoft.com/en-us/windows/package-manager/winget/
    
    The module is provided "AS IS" without warranties or conditions of any kind.

.NOTES
    Functions:
    - Enable-WG: Installs winget, use -Preview switch for preview mode
    - Get-WGList: Lists installed applications
    - Get-WGUpgrade: Lists applications with available updates
    - Get-WGVer: Shows current winget version
    - Start-WGUpgrade: Updates specified application(s)
    - Start-WGInstall: Installs application by ID
    - Start-WGUninstall: Uninstalls application by ID
    - Test-WG: Tests winget availability
#>

#region Configuration

# Set TLS protocols
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12, [Net.SecurityProtocolType]::Tls13

# Module variables with configurable paths
$script:WingetPaths = @(
    "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe\winget.exe",
    "C:\Program Files\WindowsApps\Microsoft.DesktopAppInstaller_*_x86__8wekyb3d8bbwe\winget.exe"
)
$script:Winget = Get-ChildItem $WingetPaths -ErrorAction SilentlyContinue | 
                 Where-Object { $_.FullName -notlike "*deleted*" } | 
                 Sort-Object LastWriteTime -Descending | 
                 Select-Object -First 1 -ExpandProperty FullName

# Default paths with environment variable expansion
$script:LogPath = Join-Path $env:TEMP "winget-powershell"
$script:LogFile = Join-Path $script:LogPath "ps_winget.log"
$script:DownloadPath = Join-Path $env:TEMP "Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
$script:WingetDownloadUrl = "https://aka.ms/getwinget"

# Create log directory if it doesn't exist
if (-not (Test-Path -Path $script:LogPath)) {
    New-Item -Path $script:LogPath -ItemType Directory -Force | Out-Null
}

#endregion Configuration

#region Logging

# Define the Write-Log function
Function Write-Log {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true, ValueFromPipeline=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,

        [ValidateSet('Info', 'Warning', 'Error', 'Debug')]
        [string]$Severity = 'Info'
    )

    begin {
        # Archive log file if it exceeds 5MB
        if ((Test-Path -Path $script:LogFile) -and 
            ((Get-Item -Path $script:LogFile).Length -gt 5MB)) {
            $archiveLogName = "ps_winget_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').log"
            Move-Item -Path $script:LogFile -Destination (Join-Path $script:LogPath $archiveLogName) -Force -ErrorAction SilentlyContinue
        }
    }
    
    process {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logMessage = "[$timestamp] [$Severity] $Message"
        
        # Write to log file
        try {
            Add-Content -Path $script:LogFile -Value $logMessage -ErrorAction Stop
        }
        catch {
            Write-Warning "Failed to write to log file: $_"
        }
        
        # Output to console with appropriate color based on severity
        switch ($Severity) {
            'Error' { Write-Host $logMessage -ForegroundColor Red }
            'Warning' { Write-Host $logMessage -ForegroundColor Yellow }
            'Debug' { if ($VerbosePreference -eq 'Continue') { Write-Host $logMessage -ForegroundColor Cyan } }
            default { Write-Verbose $logMessage }
        }
    }
}

#endregion Logging

#region Winget Utility Functions

# Test winget availability
Function Test-WG {
    [CmdletBinding()]
    [OutputType([bool])]
    param()
    
    if (-not (Test-Path -Path $script:Winget)) {
        Write-Log "Winget not found at expected path" -Severity Warning
        return $false
    }
    
    try {
        $null = & $script:Winget --version 2>&1
        Write-Log "Winget is installed and accessible" -Severity Debug
        return $true
    }
    catch {
        Write-Log "Error accessing winget: $_" -Severity Error
        return $false
    }
}

# Get winget version
Function Get-WGVer {
    [CmdletBinding()]
    [OutputType([string])]
    param()
    
    if (-not (Test-WG)) {
        return "Not installed"
    }
    
    try {
        $versionOutput = & $script:Winget --version 2>&1
        if ($versionOutput -match 'v?(\d+\.\d+\.\d+)') {
            return $Matches[1]
        }
        
        # Fall back to file version if command output parsing fails
        $fileVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($script:Winget).FileVersion
        return $fileVersion
    }
    catch {
        Write-Log "Error retrieving winget version: $_" -Severity Error
        return "Error"
    }
}

#endregion Winget Utility Functions

#region Installation & Update Functions

# Install or update winget
Function Enable-WG {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [switch]$Preview
    )
    
    Write-Log "Starting winget installation/update process" -Severity Info
    
    try {
        # Check if winget is already installed
        $isInstalled = Test-WG
        if ($isInstalled) {
            $currentVersion = Get-WGVer
            Write-Log "Winget is already installed (version: $currentVersion)" -Severity Info
            
            if (-not $Preview) {
                Write-Log "No action needed for standard installation" -Severity Info
                return
            }
        }
        
        # Get latest release information from GitHub
        Write-Log "Retrieving latest winget release information from GitHub" -Severity Info
        $apiUrl = "https://api.github.com/repos/microsoft/winget-cli/releases"
        $apiUrl = if ($Preview) { $apiUrl } else { "$apiUrl/latest" }
        
        $releases = Invoke-RestMethod -Uri $apiUrl -UseBasicParsing
        $release = if ($Preview) { $releases | Where-Object { $_.prerelease -eq $true } | Select-Object -First 1 } else { $releases }
        
        if (-not $release) {
            Write-Log "No suitable winget release found" -Severity Error
            return
        }
        
        $latestVersion = $release.tag_name.TrimStart('v')
        Write-Log "Latest available version: $latestVersion" -Severity Info
        
        # Download and install if needed
        if (-not $isInstalled -or $currentVersion -ne $latestVersion) {
            $downloadUrl = ($release.assets | Where-Object name -eq "Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle").browser_download_url
            
            if (-not $downloadUrl) {
                Write-Log "Could not find download URL in release assets" -Severity Error
                return
            }
            
            Write-Log "Downloading winget from $downloadUrl" -Severity Info
            Invoke-WebRequest -Uri $downloadUrl -OutFile $script:DownloadPath -UseBasicParsing -ErrorAction Stop
            
            Write-Log "Installing winget package" -Severity Info
            Add-AppxProvisionedPackage -Online -PackagePath $script:DownloadPath -SkipLicense | Out-Null
            
            # Refresh winget path
            $script:Winget = Get-ChildItem $WingetPaths -ErrorAction SilentlyContinue | 
                             Where-Object { $_.FullName -notlike "*deleted*" } | 
                             Sort-Object LastWriteTime -Descending | 
                             Select-Object -First 1 -ExpandProperty FullName
            
            if (Test-WG) {
                $newVersion = Get-WGVer
                Write-Log "Winget successfully installed/updated to version $newVersion" -Severity Info
            }
            else {
                Write-Log "Winget installation failed or not found after installation" -Severity Error
            }
        }
    }
    catch {
        Write-Log "Error during winget installation: $_" -Severity Error
    }
    finally {
        # Clean up downloaded file
        if (Test-Path $script:DownloadPath) {
            Remove-Item -Path $script:DownloadPath -Force -ErrorAction SilentlyContinue
        }
    }
}

#endregion Installation & Update Functions

#region Output Processing Functions

# Process winget list output into objects
Function Process-WingetOutput {
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param (
        [Parameter(Mandatory=$true)]
        [string]$CommandOutput,
        
        [Parameter(Mandatory=$false)]
        [string]$CommandType = "list"
    )
    
    # Check if the output contains data
    if (-not $CommandOutput -or -not ($CommandOutput -match "-----")) {
        Write-Log "No parsable output from winget command" -Severity Warning
        return @()
    }
    
    # Split the output into lines and clean up special characters
    $lines = $CommandOutput.Split([Environment]::NewLine) | 
             ForEach-Object { $_.Replace("¦", "").Replace("", "").Trim() } | 
             Where-Object { $_ -ne "" }
    
    # Find the separator line
    $separatorIndex = $lines.IndexOf(($lines | Where-Object { $_ -match "^-+$" } | Select-Object -First 1))
    if ($separatorIndex -lt 2) {
        Write-Log "Invalid winget output format - separator line not found or at unexpected position" -Severity Warning
        return @()
    }
    
    # Get header line and column positions
    $headerLine = $lines[$separatorIndex - 1]
    
    # Determine column positions based on command type
    switch ($CommandType) {
        "list" {
            $columns = @{
                "Name" = 0
                "Id" = $headerLine.IndexOf("Id")
                "Version" = $headerLine.IndexOf("Version")
                "AvailableVersion" = $headerLine.IndexOf("Available")
                "Source" = $headerLine.IndexOf("Source")
            }
        }
        "upgrade" {
            $columns = @{
                "Name" = 0
                "Id" = $headerLine.IndexOf("Id")
                "Version" = $headerLine.IndexOf("Version")
                "AvailableVersion" = $headerLine.IndexOf("Available")
                "Source" = $headerLine.IndexOf("Source")
            }
        }
        default {
            Write-Log "Unknown command type: $CommandType" -Severity Warning
            return @()
        }
    }
    
    # Process data lines
    $results = @()
    for ($i = $separatorIndex + 1; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        
        # Skip empty lines or separator lines
        if ($line -match "^-+$" -or $line.Trim() -eq "") { continue }
        
        # Check if line has enough characters for parsing
        $lastColumnPos = $columns.Values | Sort-Object | Select-Object -Last 1
        if ($line.Length -lt $lastColumnPos) { continue }
        
        try {
            # Extract fields from the line
            $item = [PSCustomObject]@{
                Name = $line.Substring(0, $columns["Id"]).Trim()
                Id = $line.Substring($columns["Id"], $columns["Version"] - $columns["Id"]).Trim()
                Version = $line.Substring($columns["Version"], $columns["AvailableVersion"] - $columns["Version"]).Trim()
                AvailableVersion = $line.Substring($columns["AvailableVersion"], $columns["Source"] - $columns["AvailableVersion"]).Trim()
                Source = $line.Substring($columns["Source"]).Trim()
            }
            $results += $item
        }
        catch {
            Write-Log "Error parsing line: '$line'. Error: $_" -Severity Warning
        }
    }
    
    return $results
}

#endregion Output Processing Functions

#region Application Management Functions

# Get list of installed applications
Function Get-WGList {
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param (
        [Parameter(Mandatory=$false)]
        [string]$AppId
    )
    
    if (-not (Test-WG)) {
        Write-Log "Winget is not installed or not found" -Severity Error
        return @()
    }
    
    try {
        $wingetArgs = @("list", "--accept-source-agreements")
        if ($AppId) {
            $wingetArgs += "--id", $AppId
        }
        
        $output = & $script:Winget $wingetArgs | Out-String
        $results = Process-WingetOutput -CommandOutput $output -CommandType "list"
        
        return $results | Sort-Object -Property Name
    }
    catch {
        Write-Log "Error getting application list: $_" -Severity Error
        return @()
    }
}

# Get list of applications with available updates
Function Get-WGUpgrade {
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param()
    
    if (-not (Test-WG)) {
        Write-Log "Winget is not installed or not found" -Severity Error
        return @()
    }
    
    try {
        $output = & $script:Winget upgrade | Out-String
        $results = Process-WingetOutput -CommandOutput $output -CommandType "upgrade"
        
        return $results
    }
    catch {
        Write-Log "Error getting available upgrades: $_" -Severity Error
        return @()
    }
}

# Upgrade a specific application or all applications
Function Start-WGUpgrade {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory=$false)]
        [string]$AppId,
        
        [Parameter(Mandatory=$false)]
        [switch]$All
    )
    
    if (-not (Test-WG)) {
        Write-Log "Winget is not installed or not found" -Severity Error
        return $false
    }
    
    # Track success/failure
    $successCount = 0
    $failureCount = 0
    
    try {
        # Determine which applications to upgrade
        $appsToUpgrade = @()
        
        if ($All -or -not $AppId) {
            $appsToUpgrade = (Get-WGUpgrade).Id
            if (-not $appsToUpgrade -or $appsToUpgrade.Count -eq 0) {
                Write-Log "No applications found needing updates" -Severity Info
                return $true
            }
            
            Write-Log "Found $($appsToUpgrade.Count) applications needing updates" -Severity Info
        }
        else {
            # Check if the specific app needs updating
            $appInfo = Get-WGUpgrade | Where-Object { $_.Id -eq $AppId }
            if (-not $appInfo) {
                Write-Log "Application '$AppId' does not need updating or was not found" -Severity Warning
                return $false
            }
            
            $appsToUpgrade = @($AppId)
        }
        
        # Process each application
        foreach ($id in $appsToUpgrade) {
            $id = $id.Trim()
            
            # Get current app info for logging
            $appInfo = Get-WGList -AppId $id | Select-Object -First 1
            $appName = if ($appInfo) { $appInfo.Name } else { $id }
            
            Write-Log "Starting upgrade for $appName (ID: $id)" -Severity Info
            
            # Run the upgrade
            $wingetArgs = @(
                "upgrade",
                "--id", $id,
                "--accept-package-agreements",
                "--accept-source-agreements"
            )
            
            $output = & $script:Winget $wingetArgs 2>&1 | Out-String
            
            # Check if upgrade was successful
            $stillNeedsUpgrade = (Get-WGUpgrade | Where-Object { $_.Id -eq $id }).Count -gt 0
            
            if (-not $stillNeedsUpgrade) {
                Write-Log "Successfully upgraded $appName" -Severity Info
                $successCount++
            }
            else {
                Write-Log "Failed to upgrade $appName. Output: $output" -Severity Error
                $failureCount++
            }
        }
        
        # Report overall results
        Write-Log "Upgrade operation completed. Success: $successCount, Failed: $failureCount" -Severity Info
        
        return ($failureCount -eq 0 -and $successCount -gt 0)
    }
    catch {
        Write-Log "Error during upgrade process: $_" -Severity Error
        return $false
    }
}

# Install an application
Function Start-WGInstall {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$AppId,
        
        [Parameter(Mandatory=$false)]
        [switch]$Force
    )
    
    if (-not (Test-WG)) {
        Write-Log "Winget is not installed or not found" -Severity Error
        return $false
    }
    
    try {
        # Check if already installed
        $existingApp = Get-WGList -AppId $AppId
        if ($existingApp -and -not $Force) {
            Write-Log "Application '$AppId' is already installed. Version: $($existingApp.Version)" -Severity Info
            return $true
        }
        
        Write-Log "Starting installation for application: $AppId" -Severity Info
        
        # Install the application
        $wingetArgs = @(
            "install",
            "--id", $AppId,
            "--accept-package-agreements",
            "--accept-source-agreements"
        )
        
        if ($Force) {
            $wingetArgs += "--force"
        }
        
        $output = & $script:Winget $wingetArgs 2>&1 | Out-String
        Write-Log $output -Severity Debug
        
        # Verify installation
        $installedApp = Get-WGList -AppId $AppId
        if ($installedApp) {
            Write-Log "Application '$($installedApp.Name)' was successfully installed. Version: $($installedApp.Version)" -Severity Info
            return $true
        }
        else {
            Write-Log "Failed to install application '$AppId'. Output: $output" -Severity Error
            return $false
        }
    }
    catch {
        Write-Log "Error during installation process: $_" -Severity Error
        return $false
    }
}

# Uninstall an application
Function Start-WGUninstall {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$AppId,
        
        [Parameter(Mandatory=$false)]
        [switch]$Force
    )
    
    if (-not (Test-WG)) {
        Write-Log "Winget is not installed or not found" -Severity Error
        return $false
    }
    
    try {
        # Check if the application is installed
        $existingApp = Get-WGList -AppId $AppId
        if (-not $existingApp) {
            Write-Log "Application '$AppId' is not installed" -Severity Warning
            return $true
        }
        
        $appName = $existingApp.Name
        Write-Log "Starting uninstallation for $appName (ID: $AppId)" -Severity Info
        
        # Uninstall the application
        $wingetArgs = @(
            "uninstall",
            "--id", $AppId,
            "--accept-source-agreements"
        )
        
        if ($Force) {
            $wingetArgs += "--force"
        }
        
        $output = & $script:Winget $wingetArgs 2>&1 | Out-String
        Write-Log $output -Severity Debug
        
        # Verify uninstallation
        $stillInstalled = Get-WGList -AppId $AppId
        if ($stillInstalled) {
            Write-Log "Failed to uninstall application '$appName'" -Severity Error
            return $false
        }
        else {
            Write-Log "Application '$appName' was successfully uninstalled" -Severity Info
            return $true
        }
    }
    catch {
        Write-Log "Error during uninstallation process: $_" -Severity Error
        return $false
    }
}

#endregion Application Management Functions
