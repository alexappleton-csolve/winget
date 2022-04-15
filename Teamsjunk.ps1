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

    $major = (get-packageprovider nuget).verison.major
    $minor = (get-packageprovider nuget).verison.minor
    $build = (get-packageprovider nuget).verison.build
    $revision = (get-packageprovider nuget).verison.revision

    ($Major,$Minor,$Build,$Revision) -Join "."