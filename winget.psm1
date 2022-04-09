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
        WG-EnablePreview: Installs the preview module as identified in the $previewurl variable
        WG-List: Displays a list of applications currently installed
        WG-Outdated: Displays a list of outdated apps
        WG-Update ($app): Updates individual application
        WG-Install: Installs individual application based on application ID. appid parameter is mandatory
    Creation date: April 8, 2022
    Last Update: April 8, 2022

#>

#Don't forget TLS stuff...
IF([Net.SecurityProtocolType]::Tls12) {[Net.ServicePointManager]::SecurityProtocol=[Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12}
IF([Net.SecurityProtocolType]::Tls13) {[Net.ServicePointManager]::SecurityProtocol=[Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls13}

#Set some global variables
$logfile = "C:\Windows\Temp\ps_winget.log"
$Winget = Get-ChildItem "C:\Program Files\WindowsApps" -Recurse -File | Where-Object name -like winget.exe | Where-Object fullname -notlike "*deleted*" | Select-Object -ExpandProperty fullname
$wgver = [System.Diagnostics.FileVersionInfo]::GetVersionInfo("$winget").FileVersion
$previewurl = "https://github.com/microsoft/winget-cli/releases/download/v1.3.431/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"

#archive existing logfile
if ([System.IO.File]::Exists($logfile)) {
    Rename-Item -Path $logfile -NewName "ps_winget$(get-date -f "yyyy-MM-dd HH-mm-ss").log"
}

#Run winget to list apps and accept source agrements (necessary on first run)
& $Winget list --accept-source-agreements | Out-Null

#Following function enables the preview build of winget
Function WG-EnablePreview {
    if ($wgver -lt '1.18.2202.12001') {
        "$(get-date -f "yyyy-MM-dd HH-mm-ss") [LOG]   Updating Winget to Preview version..." | Tee-Object -FilePath $logfile -Append
        $dl = "C:\windows\Temp\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
        #download the package
        (New-Object System.Net.WebClient).DownloadFile($previewurl, $dl)
        #install the package
        Add-AppxProvisionedPackage -Online -PackagePath $dl -SkipLicense | out-null
        #update the global variables
        $Winget = Get-ChildItem "C:\Program Files\WindowsApps" -Recurse -File | Where-Object name -like winget.exe | Where-Object fullname -notlike "*deleted*" | Select-Object -ExpandProperty fullname
		$wgver = [System.Diagnostics.FileVersionInfo]::GetVersionInfo("$winget").FileVersion
    }
    else{
        "$(get-date -f "yyyy-MM-dd HH-mm-ss") [ERR]   Winget already on Preview version!" | Tee-Object -FilePath $logfile -Append
    }
}

#Following function lists the available applications in winget
Function WG-List {
	class Application {
        [string]$Name
        [string]$Id
        [string]$Version
        [string]$AvailableVersion
    }
	
    $listResult = & $Winget list --accept-source-agreements | out-string
    if (!($listResult -match "-----")){
        return
    }

    #Split winget output to lines
    $lines = $listResult.Split([Environment]::NewLine).Replace("¦ ","")

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
    $softwarelist = @()
    For ($i = $fl + 2; $i -le $lines.Length; $i++){
        $line = $lines[$i]
        if ($line.Length -gt ($sourceStart+5) -and -not $line.StartsWith('-')){
            $Application = [Application]::new()
            $Application.Name = $line.Substring(0, $idStart).TrimEnd()
            $Application.Id = $line.Substring($idStart, $versionStart - $idStart).TrimEnd()
            $Application.Version = $line.Substring($versionStart, $availableStart - $versionStart).TrimEnd()
            $Application.AvailableVersion = $line.Substring($availableStart, $sourceStart - $availableStart).TrimEnd()
            #add formated soft to list
            $softwarelist += $Application
		}
    }

    return $softwarelist

}

#following function lists only the apps that require updating
function WG-Outdated {
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
            $software.Name = $line.Substring(0, $idStart).TrimEnd()
            $software.Id = $line.Substring($idStart, $versionStart - $idStart).TrimEnd()
            $software.Version = $line.Substring($versionStart, $availableStart - $versionStart).TrimEnd()
            $software.AvailableVersion = $line.Substring($availableStart, $sourceStart - $availableStart).TrimEnd()
            #add formated soft to list
            $upgradeList += $software
        }
    }

    return $upgradeList
}

#following function attempts to upgrade based on the application ID input variable $app.  I might change this a bit to remove the argument in the function and make it a parameter
Function WG-Update ($app) {
    
	$FailedToUpgrade = $false
    "$(get-date -f "yyyy-MM-dd HH-mm-ss") [LOG]   UPGRADE START FOR APPLICATION ID: '$($App.Id)' " | Tee-Object -FilePath $logfile -Append
    <#
	If ($($app.id -eq "Microsoft.Teams")) {
        $teamscount = (WG-List | Where-object {$_.Id -eq 'Microsoft.Teams'}).count
        If ($teamscount -gt 1) {
            Write-Host "[LOG]   Found $teamscount installations of Microsoft.Teams" | Tee-Object -file $logfile -Append
		}
		If((gp HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*).DisplayName -Contains "Teams Machine-Wide Installer" -eq $true) {
            Write-Host "[LOG]   Uninstalling Teams Machine-Wide Installer..." | Tee-Object -file $logfile -Append
            $MachineWide = Get-WmiObject -Class Win32_Product | Where-Object{$_.Name -eq "Teams Machine-Wide Installer"}
			$MachineWide.Uninstall() | out-null
            Write-Host "[LOG]   Installing Microsoft.Teams..." | Tee-Object -file $logfile -Append
			$upgradeResult = & $Winget install --id Microsoft.Teams -all --accept-package-agreements --accept-source-agreements -h | Out-String
         } 
	}
	#>
    "$(get-date -f "yyyy-MM-dd HH-mm-ss") [LOG]   Updating $($app.Name) from $($app.Version) to $($app.AvailableVersion)..." | Tee-Object -FilePath $logfile -Append
    
    #Run winget
    $results = & $Winget upgrade --id $($app.Id) --all --accept-package-agreements --accept-source-agreements -h 
    $results | Where-Object {$_ -notmatch "^\s*$|-.\\|\||^-|MB \/|KB \/|GB \/|B \/"} | Out-file -Append -FilePath $logfile 

        #Check if application updated properly

        if(WG-Outdated | Where-Object id -eq $($app.Id)) {
      			$FailedToUpgrade = $true
				"$(get-date -f "yyyy-MM-dd HH-mm-ss") [ERR]   $($app.Name) update failed. " | Tee-Object -FilePath $logfile -Append
				$InstallBAD += 1
            }
			else {
			"$(get-date -f "yyyy-MM-dd HH-mm-ss") [LOG]   $($app.Name) updated to $($app.AvailableVersion) !" | Tee-Object -FilePath $logfile -Append
			$InstallOK += 1
			}
	"$(get-date -f "yyyy-MM-dd HH-mm-ss") [LOG]   UPGRADE FINISHED FOR APPLICATION ID: '$($App.Id)' " | Tee-Object -FilePath $logfile -Append
}

#following function attemps to install based on application ID. appid parameter is mandatory
Function WG-Install {
    Param(
        [Parameter(Mandatory=$true)]
        [string]$appid,
        [Parameter()]
        [string]$scope
    )
    $failedtoinstall = $false
    "$(get-date -f "yyyy-MM-dd HH-mm-ss") [LOG]   INSTALL START FOR APPLICATION ID: '$appid' " | Tee-Object -FilePath $logfile -Append

    if($scope){
        $results = & $Winget install --id $appid --scope $scope --accept-package-agreements --accept-source-agreements -h
    }
    else{
        $results = & $Winget install --id $appid --accept-package-agreements --accept-source-agreements -h
    }
    
    $results | Where-Object {$_ -notmatch "^\s*$|-.\\|\||^-|MB \/|KB \/|GB \/|B \/"} | Out-file -Append -FilePath $logfile 

    #Check if application installed properly

    if(WG-List | Where-Object id -eq $appid) {
        "$(get-date -f "yyyy-MM-dd HH-mm-ss") [LOG]   $appid intalled. " | Tee-Object -FilePath $logfile -Append
    }
    else {

        if(!ProcessExists("msiexec.exe")) {
            "$(get-date -f "yyyy-MM-dd HH-mm-ss") [LOG]   Waiting for msiexec.exe to finish...' " | Tee-Object -FilePath $logfile -Append 
            $procid = (get-process msiexec.exe).Id
            Wait-process -id $procid 
        }
        else{
            start-sleep -Seconds 30
        }
        if(WG-List | Where-Object id -eq $appid) {
            "$(get-date -f "yyyy-MM-dd HH-mm-ss") [LOG]   $appid intalled. " | Tee-Object -FilePath $logfile -Append
        }
        else {
        $failedtoinstall=$true
        "$(get-date -f "yyyy-MM-dd HH-mm-ss") [ERR]   $appid install not completed or failed.  Please review the logs. " | Tee-Object -FilePath $logfile -Append
        }
    }
    "$(get-date -f "yyyy-MM-dd HH-mm-ss") [LOG]   INSTALL FINISHED FOR APPLICATION ID: '$appid)' " | Tee-Object -FilePath $logfile -Append    
}

#following function attempts to uninstall based on application ID.  appid parameter is mandatory
Function WG-UnInstall {
    Param(
        [Parameter(Mandatory=$true)]
        [string]$appid
    )
    $failedtouninstall = $false
    "$(get-date -f "yyyy-MM-dd HH-mm-ss") [LOG]   UNINSTALL START FOR APPLICATION ID: '$appid' " | Tee-Object -FilePath $logfile -Append

        $results = & $Winget uninstall --id $appid --accept-source-agreements -h
    
    $results | Where-Object {$_ -notmatch "^\s*$|-.\\|\||^-|MB \/|KB \/|GB \/|B \/"} | Out-file -Append -FilePath $logfile 

    #Check if application installed properly

    if(!(WG-List | Where-Object id -eq $appid)) {
        "$(get-date -f "yyyy-MM-dd HH-mm-ss") [LOG]   $appid uninstalled. " | Tee-Object -FilePath $logfile -Append
    }
    else {
        $failedtouninstall=$true
        "$(get-date -f "yyyy-MM-dd HH-mm-ss") [ERR]   $appid uninstall failed. " | Tee-Object -FilePath $logfile -Append
    }
    "$(get-date -f "yyyy-MM-dd HH-mm-ss") [LOG]   UNINSTALL FINISHED FOR APPLICATION ID: '$appid)' " | Tee-Object -FilePath $logfile -Append    
}