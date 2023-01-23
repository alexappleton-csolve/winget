<#
.SYNOPSIS
    This is a PowerShell module for the winget package manager. 
    The module provides functions for installing, upgrading, and uninstalling applications using the winget command line interface. 
    The module also includes functions for testing the winget installation, displaying installed applications and available updates, and more.


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
        Upgrade-Application: Runs the actual process of upgrading and parsing the results of the upgrade
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

    Add-Content -Path $logfile -Value $logMessage
}

#archive existing logfile
if ([System.IO.File]::Exists($logfile)) {
    Rename-Item -Path $logfile -NewName "ps_winget$(get-date -f "yyyy-MM-dd HH-mm-ss").log" -ErrorAction SilentlyContinue
}

#Following function tests winget path
Function Test-WG {
    if (Test-Path -Path $Winget) {
    	& $Winget list --accept-source-agreements | out-null
        $true
    }
    else {
        $false
    }
}

#Following function returns winget version - need to bug squash here with invoke-expression
Function Get-WGver {
    if((Test-WG)){
        [System.Diagnostics.FileVersionInfo]::GetVersionInfo("$Winget").FileVersion
    }
    else{
        Write-Output "Missing"    
    }
}

#Following function will enable winget.  Use the preview switch to install preview mode.  Test more
Function Enable-WG {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [switch]$Preview
    )
    Write-Log -Message "Installing Winget..."

    # Use the GitHub API to get the latest release of Winget
    $release = Invoke-RestMethod -Uri "https://api.github.com/repos/microsoft/winget-cli/releases/latest"

    # Check if Winget is already installed
    If ((Test-Path -Path $Winget) -eq $true) {
        # Get the current version of Winget
        $currentVersion = (Get-AppxPackage -Name Microsoft.DesktopAppInstaller).Version

        # Check if the -Preview switch is specified
        If ($PSBoundParameters['Preview']) {
            # Get the version of the latest preview release
                $latestPreviewVersion = $release.tag_name.Substring(1)

                # Compare the current version to the latest preview version
                If ($currentVersion -lt $latestPreviewVersion) {
                    # Get the download URL for the Winget preview installer
                    $previewUrl = ($release.assets | Where-Object name -eq "Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle").browser_download_url

                    Write-Log -Message "Winget preview out of date, updating to $latestPreviewVersion"
                    Invoke-WebRequest -Uri $previewUrl -OutFile $dl -ErrorAction Stop
                    Write-Log -Message "Winget preview downloaded to $dl"
                    Add-AppxProvisionedPackage -Online -PackagePath $dl -SkipLicense | Out-File -FilePath $logfile -Append
                    Write-Log -Message "Winget preview updated successfully"
                }
                Else {
                    Write-Log -Message "Winget preview already up to date"
                }
        }
        Else {
            Write-Log -Message "Winget already installed"
        }
    }
    Else {
        # Check if the -Preview switch is specified
        If ($PSBoundParameters['Preview']) {
            # Check if the release is a preview release
            If ($release.prerelease -eq $true) {
                # Get the download URL for the Winget preview installer
                $previewUrl = ($release.assets | Where-Object name -eq "Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle").browser_download_url

                Write-Log -Message "Winget preview not installed, downloading from $previewUrl"
                Invoke-WebRequest -Uri $previewUrl -OutFile $dl -ErrorAction Stop
                Write-Log -Message "Winget preview downloaded to $dl"
                Add-AppxProvisionedPackage -Online -PackagePath $dl -SkipLicense
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
            Add-AppxProvisionedPackage -Online -PackagePath $dl -SkipLicense | Out-File -FilePath $logfile -Append
            Write-Log -Message "Winget installed successfully"
        }
    }
}

#Following function will search for applications
Function Get-WGSearch {
    [CmdletBinding()]
      param (
          [Parameter(Mandatory=$false)]
          [string]$appName
      )
  
      # Get the output of the "winget list" command as a string
      if ($appName) {
          $applistResult = & $Winget search $appName | out-string
      }
      else {
          $applistResult = & $Winget search --accept-source-agreements | out-string
      }
  
      # Parse the output using the Process-WingetListOutput function
      $softwareAvailList = Process-WingetListOutput -ListResult $applistResult
  
       # Sort the list of software by the Name property
      $softwareAvailList = $softwareAvailList | Sort-Object -Property name
  
      # Return the parsed list of software
      return $softwareAvailList
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

    # Parse the output using the Process-WingetListOutput function
    $softwareList = Process-WingetListOutput -ListResult $listResult

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
            $software = [Software]::new()
            $software.Name = $line.Substring(0, $idStart).TrimEnd(("[^\P{C}]+$"))
            $software.Id = $line.Substring($idStart, $versionStart - $idStart).TrimStart("ª").TrimStart()
            $software.Version = $line.Substring($versionStart, $availableStart - $versionStart).TrimStart().TrimStart('<').TrimStart("ª").TrimStart()
            $software.AvailableVersion = $line.Substring($availableStart, $sourceStart - $availableStart).TrimStart()
            #add formated soft to list
            $upgradeList += $software
        }
    }

    return $upgradeList
}

#following function parses the results of winget - still needs work
Function Process-WingetResults {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$Results
    )

    # Normalize the output to convert the whitespace characters to regular space characters
    $normalizedResults = $Results.Normalize()
    
    # Filter the output to select only the lines that match certain criteria
    $filteredResults = $normalizedResults | Where-Object {
        # Use a regular expression to match lines that contain words with basic Latin characters - still needs work
        # $_ -match '[A-Za-z].*[A-Za-z]' 
        -not ($_ -match '^\s*$' -or $_ -match '   -  \\  |  /  -  \\  ')
        
    }

    $filteredResults = [regex]::Replace($filteredResults, "[^\p{IsBasicLatin}]", "")

    return $filteredResults
}

#following function parses the results of winget
Function Process-WingetListOutput {
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
    Param(
        [Parameter(Mandatory=$false)]
        [string]$appid,
        [Parameter(Mandatory=$false)]
        [switch]$All
    )

    if ($All -or !$appid) {
        # Get a list of appids that need to be updated
        $appids = (Get-WGUpgrade).Id
        #need to trim trailing whitespace from ids so they dont send as part of id
        # Loop through each appid and call Upgrade-Application for each appid
        foreach ($id in $appids) {
            Upgrade-Application -appid $id
            continue
        }
    }
    else {
        Upgrade-Application -appid $appid
    }
}

#Following function will install individual application based on application ID
Function Start-WGInstall {
    [cmdletbinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$appid
    )

    Write-Log -Message "Starting installation for application: $appid" -Severity "Info"
    
    #Install the specified application
    $results = & $Winget install --id $appid | out-String
    $filteredResults = Process-WingetResults -results $Results

    # Output the filtered results to the log file
    $filteredResults | Out-File -Append -FilePath $logfile

    #Check if the installation was successful
    $installedApp = Get-WGList | Where-Object id -like "$appid*"
    if ($installedApp) {
        Write-Log -Message "Application '$($installedApp.name)' was installed successfully." -Severity "Info"
        $true
    }
    else {
        Write-Log -Message "Application '$appid' could not be installed." -Severity "Error"
        $false
    }
}

#Following function will uninstall individual application based on application ID
Function Start-WGUninstall {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true, Position=1)]
        [String]
        $appid
    )

    Write-Log -Message "Starting uninstallation for application: $appid" -Severity "Info"

    $app = Get-WGList | Where-Object id -like "$appid*"

        if($app) {
            $results = & $Winget uninstall --id $appid | out-string
            $filteredResults = Process-WingetResults -results $Results
           
            # Output the filtered results to the log file
            $filteredResults | Out-File -Append -FilePath $logfile

        }
        else {
            Write-Log -Message "Application not found: $appid" -Severity "Error"
            $false
            return
        }
        
    #Check if the uninstallation was successful
    $installedApp = Get-WGList | Where-Object id -like "$appid*"

        if ($installedApp) {
            Write-Log -Message "Application '$($installedApp.name)' is still installed." -Severity "Error"
            $false
        }
        else {
            Write-Log -Message "Application '$appid' uninstalled." -Severity "Info"
            $true
        }   
}

#Following function performs the winget upgrade procedure and captures the results
Function Upgrade-Application {
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory=$true)]
        [string]$appid
    )
    #future use
    $FailedToUpgrade = $false

    # Remove the trailing white space from the app ID
    $appID = $appid.TrimEnd()
        
    # Get the app version and available version from the Get-WGList function
    $appInfo = Get-WGList | Where-Object Id -match $appid
    $appversion = $appInfo.Version
    $availversion = $appInfo.AvailableVersion

    Write-Log -Message "UPGRADE START FOR APPLICATION ID: '$appid'" -Severity "Info"
    Write-Log -Message "Upgrading from $appversion to $availversion..." -Severity "Info"  

    #Run winget upgrade
    $results = & $Winget upgrade --id $appid --all --accept-package-agreements --accept-source-agreements -h 

    # Normalize the output to convert the whitespace characters to regular space characters
    $normalizedResults = $results.Normalize()
    
    # Filter the output to select only the lines that match certain criteria
    $filteredResults = $normalizedResults | Where-Object {
        # Use a regular expression to match lines that contain words with basic Latin characters - still needs work
         $_ -match '[A-Za-z].*[A-Za-z]'
    }

   $filteredResults = [regex]::Replace($filteredResults, "[^\p{IsBasicLatin}]", "")

    # Output the filtered results to the log file
    $filteredResults | Out-File -Append -FilePath $logfile

    #Check if application updated properly - this doesn't seem to be working correctly yet.  '
    if(Get-WGUpgrade| Where-Object id -eq $appid) {
        $FailedToUpgrade = $true
        Write-Log -Message "Update failed. Please review the log file." -Severity "Error"
        $InstallBAD += 1
        }
    else {
            Write-Log -Message "Update completed !" -Severity "Info"
            $InstallOK += 1
        }
    Write-Log -Message "UPGRADE FINISHED FOR APPLICATION ID: '$appid'" -Severity "Info"
}
