<#
.SYNOPSIS
    This is a PowerShell module for winget package manager

.DESCRIPTION
    Thank you https://github.com/Romanitho/Winget-AutoUpdate/tree/main/Winget-AutoUpdate
    Also https://github.com/jdhitsolutions/WingetTools

    https://docs.microsoft.com/en-us/windows/package-manager/winget/
        
    Module will try to make the winget package manager work for automation and integration with PowerShell
    
    Module is provided "AS IS". There are no warranties or conditions OF ANY KIND, expressed or implied, including,
    without limitation, any warranties or conditions of TITLE, NONINFRINGEMENT, MERCHANTABILITY, or FITNESS FOR ANY PARTICULAR
    PURPOSE.  You are solely responsible for determining the appropriateness of using or redistributing the module 
    and assume any risks associated with the use of the work. 

    This module does not grant, nor is responsible for, any licenses to third-party packages.  The applications distributed
    through this module are licensed to you by its owner.  

.NOTES
    Functions:
        Enable-WG: Installs winget, use -preview switch to install preview mode
        Get-WGList: Displays a list of applications currently installed
        Get-WGUpgrade: Displays a list of outdated apps
        Get-WGVer: Displays current version of winget
        Start-WGUpgrade: Updates individual application
        Start-WGInstall: Installs individual application based on application ID. appid parameter is mandatory
        Start-WGUninstall: Uninstalls individual application based on application ID.  appid parameter is mandatory
        Test-WG: Tests winget path
        Uninstall-WG: Uninstalls Winget
        More functions to come!


#>
#Set TLS protocols.
IF([Net.SecurityProtocolType]::Tls12) {[Net.ServicePointManager]::SecurityProtocol=[Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12}
IF([Net.SecurityProtocolType]::Tls13) {[Net.ServicePointManager]::SecurityProtocol=[Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls13}

#Set some global variables
$Winget = Get-ChildItem "C:\Program Files\WindowsApps" -Recurse -File | Where-Object name -like winget.exe | Where-Object fullname -notlike "*deleted*" | Select-Object -last 1 -ExpandProperty fullname
$logfile = "C:\Windows\Temp\ps_winget.log"
$dl = "C:\windows\Temp\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
$wgdl = "https://aka.ms/getwinget"

# Define the Write-Log function
Function Write-Log {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,

        [ValidateSet('Info', 'Warning', 'Error')]
        [string]$Severity = 'Info'
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logMessage = "[$timestamp] [$Severity] $Message"

    Write-Output $logMessage | Out-File -FilePath $logfile -Append
}

#archive existing logfile
if ([System.IO.File]::Exists($logfile)) {
    Rename-Item -Path $logfile -NewName "ps_winget$(get-date -f "yyyy-MM-dd HH-mm-ss").log" -ErrorAction SilentlyContinue
}

#Following function tests winget path
Function Test-WG {
    Try{
        Test-Path -Path $winget | Out-Null
        & $Winget list --accept-source-agreements | Out-Null
        $true
    }
    Catch{
        $false
    }
}

#Following function returns winget version
Function Get-WGver {
    if((Test-WG)){
        [System.Diagnostics.FileVersionInfo]::GetVersionInfo("$winget").FileVersion
    }
    else{
        Write-Output "Missing"    
    }
}

#Following function will enable winget.  Use the preview switch to install preview mode
Function Enable-WG {
    Write-Log -Message "Installing Winget..."
    If ((Test-Path -Path $Winget) -eq $true) {
        Write-Log -Message "Winget already installed"
        If ($PSBoundParameters['Preview']) {
            Write-Log -Message "Winget preview already installed"
        }
    }
    Else {
        Try {
            # Use the GitHub API to get the latest release of Winget
            $release = Invoke-RestMethod -Uri "https://api.github.com/repos/microsoft/winget-cli/releases/latest"

            # Check if the -Preview switch is specified
            If ($PSBoundParameters['Preview']) {
                # Check if the release is a preview release
                If ($release.prerelease -eq $true) {
                    # Get the download URL for the Winget preview installer
                    $previewUrl = ($release.assets | Where-Object name -eq "Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle").browser_download_url

                    Write-Log -Message "Winget preview not installed, downloading from $previewUrl"
                    Invoke-WebRequest -Uri $previewUrl -OutFile $dl -ErrorAction Stop
                    Write-Log -Message "Winget preview downloaded to $dl"
                    Start-Process -FilePath $dl -ArgumentList "/quiet", "/norestart" -Wait
                    Write-Log -Message "Winget preview installed successfully"
                }
                Else {
                    Write-Log -Message "No preview releases of Winget are available"
                }
            }
            Else {
                # Download and install the stable release of Winget
                Write-Log -Message "Downloading winget from $wgdl"
                Invoke-WebRequest -Uri $wgdl -OutFile $dl -ErrorAction Stop
                Write-Log -Message "Winget downloaded to $dl"
                Start-Process -FilePath $dl -ArgumentList "/quiet", "/norestart" -Wait
                Write-Log -Message "Winget installed successfully"
            }
        }
        Catch {
            Write-Log -Message "Error installing Winget: $_" -Severity "Error"
        }
    }
}

#Following function will list installed applications
Function Get-WGList {
  [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [string]$appid
    )

    # Get the output of the "winget list" command as a string
    if ($appid) {
        $listResult = & $Winget list --id $appid | out-string
    }
    else {
        $listResult = & $Winget list --accept-source-agreements | out-string
    }

    # Parse the output using the Parse-WingetListOutput function
    $softwareList = Parse-WingetListOutput -ListResult $listResult

     # Sort the list of software by the Name property
    $softwareList = $softwareList | Sort-Object -Property name

    # Return the parsed list of software
    return $softwareList
}

#following function lists only the apps that require updating
function Get-WGUpgrade {
	class Software {
        [string]$Name
        [string]$Id
        [string]$Version
        [string]$AvailableVersion
    }

    #Get list of available upgrades on winget format
    $upgradeResult = & $Winget upgrade | Out-String

    #Start Conversion of winget format to an array. Check if "-----" exists
    if (!($upgradeResult -match "-----")){
        return
    }

    #Split winget output to lines
    $lines = $upgradeResult.Split([Environment]::NewLine).Replace("¦ ","")

    # Find the line that starts with "------"
    $fl = 0
    while (-not $lines[$fl].StartsWith("-----")){
        $fl++
    }
    
    #Get header line
    $fl = $fl - 2

    #Get header titles
    $index = $lines[$fl] -split '\s+'

    # Line $i has the header, we can find char where we find ID and Version
    $idStart = $lines[$fl].IndexOf($index[1])
    $versionStart = $lines[$fl].IndexOf($index[2])
    $availableStart = $lines[$fl].IndexOf($index[3])
    $sourceStart = $lines[$fl].IndexOf($index[4])

    # Now cycle in real package and split accordingly
    $upgradeList = @()
    For ($i = $fl + 2; $i -le $lines.Length; $i++){
        $line = $lines[$i]
        if ($line.Length -gt ($sourceStart+5) -and -not $line.StartsWith('-')){
            $Application = [Application]::new()
            $Application.Name = $line.Substring(0, $idStart).TrimEnd(("[^\P{C}]+$"))
            $Application.Id = $line.Substring($idStart, $versionStart - $idStart).TrimStart("ª").TrimStart()
            $Application.Version = $line.Substring($versionStart, $availableStart - $versionStart).TrimStart().TrimStart('<').TrimStart("ª").TrimStart()
            $Application.AvailableVersion = $line.Substring($availableStart, $sourceStart - $availableStart).TrimStart()
            #add formated soft to list
            $upgradeList += $software
        }
    }

    return $upgradeList
}

#following function parses the results of winget
Function Parse-WingetListOutput {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$true)]
        [string]$ListResult
    )

    $ListResult = $ListResult | Out-String

    # Define a class for representing an application
    class Application {
        [string]$Name
        [string]$Id
        [string]$Version
        [string]$AvailableVersion
    }

    # Check if the output contains the separator line
    if (!($listResult -match "-----")){
        return
    }

    # Split the output into lines
    $lines = $listResult.Split([Environment]::NewLine).Replace("¦ ","")

    # Find the line that starts with "------"
    $fl = 0
    while (-not $lines[$fl].StartsWith("-----")){
        $fl++
    }
    
    # Get the header line
    $fl = $fl - 2

    # Get the header titles
    $index = $lines[$fl] -split '\s+'

    # Line $i has the header, we can find the characters where we find the ID and Version
    $idStart = $lines[$fl].IndexOf($index[1])
    $versionStart = $lines[$fl].IndexOf($index[2])
    $availableStart = $lines[$fl].IndexOf($index[3])
    $sourceStart = $lines[$fl].IndexOf($index[4])

    # Now cycle through the real package and split accordingly
    $softwarelist = @()
    For ($i = $fl + 2; $i -le $lines.Length; $i++){
        $line = $lines[$i]
        if ($line.Length -gt ($sourceStart+5) -and -not $line.StartsWith('-')){
            $Application = [Application]::new()
            $Application.Name = $line.Substring(0, $idStart).TrimEnd(("[^\P{C}]+$"))
            $Application.Id = $line.Substring($idStart, $versionStart - $idStart).TrimStart("ª").TrimStart()
            $Application.Version = $line.Substring($versionStart, $availableStart - $versionStart).TrimStart().TrimStart('<').TrimStart("ª").TrimStart()
            $Application.AvailableVersion = $line.Substring($availableStart, $sourceStart - $availableStart).TrimStart()
            #add formated soft to list
            $softwarelist += $Application
        }
    }
    return $softwarelist
}

#following function attempts to upgrade based on the application ID input parameter $appid  
Function Start-WGUpgrade {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=1)]
        [String]
        $appid
    )

    Write-Log -Message "Starting upgrade for application: $appid"

    if((Test-WG)){
        $app = & $Winget search -u $appid
        if($app){
            $appname = $app | Select-Object -first 1 -ExpandProperty displayName
            Write-Log -Message "Upgrading application: $appname"
            & $Winget install -e --accept-source-agreements $appid
        }
        else{
            Write-Log -Message "Application not found: $appid"
        }
    }
    else{
        Write-Log -Message "Winget not found. Installing Winget..."
        Enable-WG
        Write-Log -Message "Winget installed. Starting upgrade for application: $appid"
        $app = & $Winget search -u $appid
        if($app){
            $appname = $app | Select-Object -first 1 -ExpandProperty displayName
            Write-Log -Message "Upgrading application: $appname"
            & $Winget install -e --accept-source-agreements $appid
        }
        else{
            Write-Log -Message "Application not found: $appid"
        }
    }

    Write-Log -Message "Finished upgrading application: $appid"
}

#Following function will install individual application based on application ID
Function Start-WGInstall {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=1)]
        [String]
        $appid
    )

    Write-Log -Message "Starting installation for application: $appid"

    if((Test-WG)){
        $app = & $Winget search -u $appid
        if($app){
            $appname = $app | Select-Object -first 1 -ExpandProperty displayName
            Write-Log -Message "Installing application: $appname"
            & $Winget install -e --accept-source-agreements $appid
        }
        else{
            Write-Log -Message "Application not found: $appid"
        }
    }
    else{
        Write-Log -Message "Winget not found. Installing Winget..."
        Enable-WG
        Write-Log -Message "Winget installed. Starting installation for application: $appid"
        $app = & $Winget search -u $appid
        if($app){
            $appname = $app | Select-Object -first 1 -ExpandProperty displayName
            Write-Log -Message "Installing application: $appname"
            & $Winget install -e --accept-source-agreements $appid
        }
        else{
            Write-Log -Message "Application not found: $appid"
        }
    }

    Write-Log -Message "Finished installing application: $appid"
}


#Following function will uninstall individual application based on application ID
Function Start-WGUninstall {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=1)]
        [String]
        $appid
    )

    Write-Log -Message "Starting uninstallation for application: $appid"

    if((Test-WG)){
        $app = & $Winget search -u $appid
        if($app){
            $appname = $app | Select-Object -first 1 -ExpandProperty displayName
            Write-Log -Message "Uninstalling application: $appname"
            & $Winget uninstall -e --accept-source-agreements $appid
        }
        else{
            Write-Log -Message "Application not found: $appid"
        }
    }
    else{
        Write-Log -Message "Winget not found. Installing Winget..."
        Enable-WG
        Write-Log -Message "Winget installed. Starting uninstallation for application: $appid"
        $app = & $Winget search -u $appid
        if($app){
            $appname = $app | Select-Object -first 1 -ExpandProperty displayName
            Write-Log -Message "Uninstalling application: $appname"
            & $Winget uninstall -e --accept-source-agreements $appid
        }
        else{
            Write-Log -Message "Application not found: $appid"
        }
    }

    Write-Log -Message "Finished uninstalling application: $appid"
}

#Following function will uninstall winget from the system
Function Uninstall-WG {
    Write-Log -Message "Uninstalling Winget..."

    # Uninstall Winget
    $uninstalled = $false
    try {
        Get-AppxPackage -Name Microsoft.DesktopAppInstaller | Remove-AppxPackage
        $uninstalled = $true
    }
    catch {
        Write-Log -Message "Error uninstalling Winget: $_"
    }

    # Delete the installation files
    if([System.IO.File]::Exists($dl)){
        Remove-Item -Path $dl -Force
    }

    if($uninstalled){
        Write-Log -Message "Finished uninstalling Winget"
    }
    else{
        Write-Log -Message "Winget is already uninstalled"
    }
}
