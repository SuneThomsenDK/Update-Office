<#
.SYNOPSIS
	Install Microsoft Office Professional Plus 2010, 2013 and 2016 updates offline

.DESCRIPTION
	The purpose of this script is to install Office updates offline or during SCCM OS Deployment instead of WSUS, which takes forever to complete
	This script reduced my SCCM OS Deployment time by 20-30 minutes depending on the hardware

.PARAMETER UpdateRoot
	Changes the default path from "$PSScriptRoot\Updates\" to the path specified

.PARAMETER LogRoot
	Changes the default path from "$PSScriptRoot\Log\" to the path specified

.PARAMETER GridView
	Shows all available Office updates in GridView

.EXAMPLE
	Update-Office.ps1 -UpdateRoot "Add Custom Path Here"
	Changes the default path from "$PSScriptRoot\Updates\" to the path specified

.EXAMPLE
	Update-Office.ps1 -LogRoot "Add Custom Path Here"
	Changes the default path from "$PSScriptRoot\Log\" to the path specified

.EXAMPLE
	Update-Office.ps1 -GridView
	Shows all available Office updates in GridView

.EXAMPLE
	Update-Office.ps1 -UpdateRoot "Add Custom Path Here" -LogRoot "Add Custom Path Here" -GridView
	Changes the default path to the path specified and shows all available Office updates in GridView

.NOTES
	Version: 1.9.3.6
	Author: Sune Thomsen
	Creation date: 22-02-2019
	Last modified date: 06-03-2019

.LINK
	https://github.com/SuneThomsenDK
#>
	#===============================================================================
	#	Requirements
	#===============================================================================
	#Requires -Version 4
	#Requires -RunAsAdministrator

	Param (
		[System.IO.FileInfo][String]$UpdateRoot = "$PSScriptRoot\Updates\",
		[System.IO.FileInfo][String]$LogRoot = "$PSScriptRoot\Log\",
		[Switch]$GridView
	)

	Measure-Command -Expression {

		Function Write-Log {
			[CmdletBinding()]
			Param (
				[Parameter(Mandatory=$false)][String]$LogFile,
				[Parameter(Mandatory=$true)][String]$Message,
				[Parameter(Mandatory=$false)][ValidateSet("Information","Warning","Error")][String]$Type = "Information"
			)

			$LogTime = (Get-Date).toString("$DateFormat $TimeFormat")
			$LogLine = "$LogTime $($Type): $Message"

			if (($LogFile)) {
				Add-Content $LogFile -Value $LogLine
			}
			else {
				Write-Output $LogLine
			}
		}

		Function Get-MSPInfo {
			[CmdletBinding()]
			Param (
				[Parameter(Mandatory = $true)][System.IO.FileInfo][String]$MSPFile,
				[Parameter(Mandatory = $true)][ValidateSet("Classification", "DisplayName", "KBArticle Number", "TargetProductName", "CreationTimeUTC")][String]$Property
			)
			Try {
				#===============================================================================
				#	Get MSP Information
				#===============================================================================
				$WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
				$MSPDatabase = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $Null, $WindowsInstaller, @($MSPFile.FullName, 32))
				$MSPQuery = "SELECT Value FROM MsiPatchMetadata WHERE Property = '$($Property)'"
				$MSPView = $MSPDatabase.GetType().InvokeMember("OpenView", "InvokeMethod", $null, $MSPDatabase, ($MSPQuery))
				$MSPView.GetType().InvokeMember("Execute", "InvokeMethod", $null, $MSPView, $null)
				$MSPRecord = $MSPView.GetType().InvokeMember("Fetch", "InvokeMethod", $null, $MSPView, $null)
				$MSPValue = $MSPRecord.GetType().InvokeMember("StringData", "GetProperty", $null, $MSPRecord, 1)
				Return $MSPValue
			}
			Catch {
				Write-Host "Warning: Sune has created a awesome script, but something went wrong!" -foregroundcolor "Yellow"
				Write-Log -Message "Sune has created a awesome script, but something went wrong!" -Type Error -LogFile $LogPath
				Write-Log -Message "$_.Exception.Message" -Type Error -LogFile $LogPath
				Return $NULL
			}
		}

		Function Get-MSPPatchCode {
			[CmdletBinding()]
			Param (
				[Parameter(Mandatory = $true)][System.IO.FileInfo][String]$MSPFile
			)
			Try {
				#===============================================================================
				#	Get MSP PatchCode
				#===============================================================================
				$WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
				$MSPDatabase = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $Null, $WindowsInstaller, $($MSPFile.FullName, 32))
				$MSPSummary = $MSPDatabase.GetType().InvokeMember("SummaryInformation", "GetProperty", $Null, $MSPDatabase, $Null)
				[String]$MSPPatchCode = $MSPSummary.GetType().InvokeMember("Property", "GetProperty", $Null, $MSPSummary, 9)
				Return $MSPPatchCode
			}
			Catch {
				Write-Host "Warning: Sune has created a awesome script, but something went wrong!" -foregroundcolor "Yellow"
				Write-Log -Message "Sune has created a awesome script, but something went wrong!" -Type Error -LogFile $LogPath
				Write-Log -Message "$_.Exception.Message" -Type Error -LogFile $LogPath
				Return $NULL
			}
		}

		Function Check-Registry {
			Try {
				#===============================================================================
				#	Check PatchCode in Registry
				#===============================================================================
				$Office2010 = "HKLM:\SOFTWARE\Microsoft\Office\14.0\Outlook"
				$Office2013 = "HKLM:\SOFTWARE\Microsoft\Office\15.0\Outlook"
				$Office2016 = "HKLM:\SOFTWARE\Microsoft\Office\16.0\Outlook"
				$RegWin = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
				$RegWoW = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
				$IsOffice = $Null

				if ((Test-Path $Office2010)) {$IsOffice = Get-ItemProperty -Path $Office2010 -name Bitness -ErrorAction SilentlyContinue}
				if ((Test-Path $Office2013)) {$IsOffice = Get-ItemProperty -Path $Office2013 -name Bitness -ErrorAction SilentlyContinue}
				if ((Test-Path $Office2016)) {$IsOffice = Get-ItemProperty -Path $Office2016 -name Bitness -ErrorAction SilentlyContinue}

				if (([System.Environment]::Is64BitOperatingSystem)) {
					if (($IsOffice.Bitness -eq "x86")) {
						$CheckPatchCode = Get-ItemProperty -Path $RegWoW |
						Where-Object {$_.PSChildName -like "*$PatchCode*" -or $_.UninstallString -like "*$PatchCode*"} |
						Select-Object -Property PSChildName, DisplayName, UninstallString |
						Sort-Object -Property DisplayName -Unique
					}
					else {
						$CheckPatchCode = Get-ItemProperty -Path $RegWin |
						Where-Object {$_.PSChildName -like "*$PatchCode*" -or $_.UninstallString -like "*$PatchCode*"} |
						Select-Object -Property PSChildName, DisplayName, UninstallString |
						Sort-Object -Property DisplayName -Unique
					}
				}

				if (!([System.Environment]::Is64BitOperatingSystem)) {
					$CheckPatchCode = Get-ItemProperty -Path $RegWin |
					Where-Object {$_.PSChildName -like "*$PatchCode*" -or $_.UninstallString -like "*$PatchCode*"} |
					Select-Object -Property PSChildName, DisplayName, UninstallString |
					Sort-Object -Property DisplayName -Unique
				}
				Return $CheckPatchCode.DisplayName
			}
			Catch {
				Write-Host "Warning: Sune has created a awesome script, but something went wrong!" -foregroundcolor "Yellow"
				Write-Log -Message "Sune has created a awesome script, but something went wrong!" -Type Error -LogFile $LogPath
				Write-Log -Message "$_.Exception.Message" -Type Error -LogFile $LogPath
				Return $Null
			}
		}

		Function Install-MSPUpdate {
			[CmdletBinding()]
			Param (
				[Parameter(Mandatory = $true)][System.IO.FileInfo][String]$MSPFile
			)
			Try {
				#===============================================================================
				#	Install MSP Update
				#===============================================================================
				$KBNumber = $Update.KBNumber
				$DisplayName = $Update.DisplayName
				$PatchCode = $Update.PatchCode
				$Process = "msiexec.exe"
				$CheckPatchCode = Check-Registry

				$MSPArguments = @(
					"/p",
					"""$MSPFile""",
					"/qn",
					"REBOOT=ReallySuppress",
					"MSIRESTARTMANAGERCONTROL=Disable"
				)

				if (!($CheckPatchCode)) {
					$MSPInstall = Start-Process $process -ArgumentList $MSPArguments -PassThru -Wait
					$MSPInstall.WaitForExit()
					if (($MSPInstall.ExitCode -eq 0) -or ($MSPInstall.ExitCode -eq 3010)){
						$Script:CountInstall++
						Write-Host "Installing: $DisplayName ($($Update.BaseName))" -foregroundcolor "Green"
						Write-Log -Message "Installing $DisplayName ($($Update.BaseName))" -Type Information -LogFile $LogPath
					}
					else {
						$Script:CountNotInstalled++
						Write-Host "Attention: $DisplayName ($($Update.BaseName)) were not installed" -foregroundcolor "Cyan"
						Write-Host "Possible cause: The program to be updated might not be installed, or the patch may update a different version of the program."
						Write-Log -Message "$DisplayName ($($Update.BaseName)) were not installed" -Type Warning -LogFile $LogPath
						Write-Log -Message "Possible cause: The program to be updated might not be installed, or the patch may update a different version of the program." -Type Information -LogFile $LogPath
					}
				}
				else {
					$Script:CountNotInstalled++
					Write-Host "Attention: $DisplayName ($($Update.BaseName)) is already installed" -foregroundcolor "Cyan"
					Write-Log -Message "$DisplayName ($($Update.BaseName)) is already installed" -Type Information -LogFile $LogPath
				}
			}
			Catch {
				Write-Host "Warning: Sune has created a awesome script, but something went wrong!" -foregroundcolor "Yellow"
				Write-Log -Message "Sune has created a awesome script, but something went wrong!" -Type Error -LogFile $LogPath
				Write-Log -Message "$_.Exception.Message" -Type Error -LogFile $LogPath
				Return $NULL
			}
		}

		#===============================================================================
		#	Check That Update Root Exists
		#===============================================================================
		if (!(Test-Path -Path $UpdateRoot)) {
			$PSCommandPath = $PSCommandPath.Split("\")[2]
			Write-Host "$($PSCommandPath): Cannot find $UpdateRoot because it does not exist! Please verify that the path is correct and try again." -foregroundcolor "Yellow"
			Exit
		}

		#===============================================================================
		#	Set Variables
		#===============================================================================
		$LocalCulture = Get-Culture
		$LanguageCode = $LocalCulture.LCID
		$DateFormat = [System.Globalization.CultureInfo]::GetCultureInfo($LanguageCode).DateTimeFormat.ShortDatePattern
		$TimeFormat = [System.Globalization.CultureInfo]::GetCultureInfo($LanguageCode).DateTimeFormat.LongTimePattern

		$OfficeUpdates = Get-ChildItem $UpdateRoot -Recurse -File -Include *.msp
		$LogFileTime = (Get-Date -f "yyyy-MM-dd-HHmmss")
		$LogFile = "$($LogFileTime)-Update-Office.log"
		$LogPath = Join-Path "$LogRoot" "$LogFile"

		$Script:CountInstall = 0
		$Script:CountNotInstalled = 0

		$OfficeArrayList = @(
			"acewss-x-none",
			"ace-x-none",
			"chart-x-none",
			"csisyncclient-x-none",
			"csi-x-none",
			"dcf-x-none",
			"exppdf-x-none",
			"filterpack-x-none",
			"fonts-x-none",
			"gkall-x-none",
			"graph-x-none",
			"ieawsdc-x-none",
			"mscomctlocx-x-none",
			"msmipc-x-none",
			"msodll20-x-none",
			"msodll30-x-none",
			"msodll40ui-x-none",
			"msodll99l-x-none",
			"msohevi-x-none",
			"mtextra-x-none",
			"oart-x-none",
			"oleo-x-none",
			"orgidcrl-x-none",
			"otkruntimertl-x-none",
			"outexum-x-none",
			"outlfltr-x-none",
			"policytips-x-none",
			"ppaddin-x-none",
			"project-x-none",
			"protocolhndlr-x-none",
			"riched20-x-none",
			"seguiemj-x-none",
			"stslist-x-none",
			"stsupld-x-none",
			"vbe7-x-none",
			"visio-x-none",
			"wxpcore-x-none",
			"wxpnse-x-none",
			"xdext-x-none"
		)

		$OfficeLIPArrayList = @(
			"lip-af-za",
			"lip-am-et",
			"lip-as-in",
			"lip-az-latn-az",
			"lip-be-by",
			"lip-bn-bd",
			"lip-bn-in",
			"lip-bs-latn-ba",
			"lip-ca-es-valencia",
			"lip-ca-es",
			"lip-cy-gb",
			"lip-eu-es",
			"lip-fa-ir",
			"lip-fil-ph",
			"lip-ga-ie",
			"lip-gd-gb",
			"lip-gl-es",
			"lip-gu-in",
			"lip-ha-latn-ng",
			"lip-hy-am",
			"lip-id-id",
			"lip-ig-ng",
			"lip-is-is",
			"lip-ja-jp.pseudo",
			"lip-ka-ge",
			"lip-km-kh",
			"lip-kn-in",
			"lip-kok-in",
			"lip-ky-kg",
			"lip-lb-lu",
			"lip-mi-nz",
			"lip-mk-mk",
			"lip-ml-in",
			"lip-mn-mn",
			"lip-mr-in",
			"lip-ms-my",
			"lip-mt-mt",
			"lip-ne-np",
			"lip-nn-no",
			"lip-nso-za",
			"lip-or-in",
			"lip-pa-in",
			"lip-prs-af",
			"lip-ps-af",
			"lip-quz-pe",
			"lip-rw-rw",
			"lip-sd-arab-pk",
			"lip-si-lk",
			"lip-sq-al",
			"lip-sr-cyrl-ba",
			"lip-sr-cyrl-cs",
			"lip-sr-cyrl-rs",
			"lip-sw-ke",
			"lip-ta-in",
			"lip-te-in",
			"lip-tk-tm",
			"lip-tn-za",
			"lip-tt-ru",
			"lip-ug-cn",
			"lip-ur-pk",
			"lip-uz-latn-uz",
			"lip-vi-vn",
			"lip-wo-sn",
			"lip-xh-za",
			"lip-yo-ng",
			"lip-zu-za"
		)

		$OfficeLPArrayList = @(
			"access-x-none",
			"conv-x-none",
			"eqnedt32-x-none",
			"excelpp-x-none",
			"excel-x-none",
			"groove-x-none",
			"lync-x-none",
			"mso-x-none",
			"onenote-x-none",
			"ose-x-none",
			"osfclient-x-none",
			"outlook-x-none",
			"powerpoint-x-none",
			"publisher-x-none",
			"word-x-none"
		)

		$OfficePKArrayList = @(
			"kohhc-x-none",
			"osetup-x-none",
			"ospp-x-none",
			"proof-x-none"
		)

		#===============================================================================
		#	Create Log Folder
		#===============================================================================
		if (!(Test-Path -Path $LogRoot)) {New-Item $LogRoot -ItemType Directory | Out-Null}

		#===============================================================================
		#	Start Logging
		#===============================================================================
		Write-Log -Message "===============================================================================" -Type Information -LogFile $LogPath
		Write-Log -Message "- Start-Logging" -Type Information -LogFile $LogPath
		Write-Log -Message " " -Type Information -LogFile $LogPath
		Write-Log -Message "- Username: $Env:USERNAME" -Type Information -LogFile $LogPath
		Write-Log -Message "- Computername: $env:COMPUTERNAME" -Type Information -LogFile $LogPath
		Write-Log -Message "- Update path: $UpdateRoot" -Type Information -LogFile $LogPath
		Write-Log -Message "- Log path: $LogRoot" -Type Information -LogFile $LogPath

		ForEach ($Update in $OfficeUpdates) {
			#===============================================================================
			#	Get MSP Properties
			#===============================================================================
			$KBNumber = Get-MSPInfo -MSPFile $Update.FullName -Property "KBArticle Number"
			$Classification = Get-MSPInfo -MSPFile $Update.FullName -Property "Classification"
			$DisplayName = Get-MSPInfo -MSPFile $Update.FullName -Property "DisplayName"
			$ProductName = Get-MSPInfo -MSPFile $Update.FullName -Property "TargetProductName"
			$CreationDateUTC = Get-MSPInfo -MSPFile $Update.FullName -Property "CreationTimeUTC"
			$PatchCode = Get-MSPPatchCode -MSPFile $Update.FullName

			#===============================================================================
			#	Format CreationDateUTC
			#===============================================================================
			$SplitDate = ($CreationDateUTC[1] -split " ")[0]
			$SplitTime = ($CreationDateUTC[1] -split " ")[1]

			if (($DateFormat -like "D*")) {
				$CreationDateUTC = $SplitDate
				$CreationDateUTC = $CreationDateUTC.Split("/")
				$CreationDateUTC = "{0}/{1}/{2}" -f $CreationDateUTC[1],$CreationDateUTC[0],$CreationDateUTC[2]
				$CreationDateUTC = $CreationDateUTC+"`t"+$SplitTime
				$CreationDateUTC = Get-Date $CreationDateUTC -f "$DateFormat $TimeFormat"
				$CreationDateUTC = ([DateTime]::ParseExact($CreationDateUTC,"$DateFormat $TimeFormat",[Globalization.CultureInfo]::InvariantCulture))
			}

			if (($DateFormat -like "M*")) {
				$CreationDateUTC = $SplitDate
				$CreationDateUTC = $CreationDateUTC+"`t"+$SplitTime
				$CreationDateUTC = Get-Date $CreationDateUTC -f "$DateFormat $TimeFormat"
				$CreationDateUTC = ([DateTime]::ParseExact($CreationDateUTC,"$DateFormat $TimeFormat",[Globalization.CultureInfo]::InvariantCulture))
			}

			if (($DateFormat -like "Y*")) {
				$CreationDateUTC = $SplitDate
				$CreationDateUTC = $CreationDateUTC.Split("/")
				$CreationDateUTC = "{0}/{1}/{2}" -f $CreationDateUTC[2],$CreationDateUTC[0],$CreationDateUTC[1]
				$CreationDateUTC = $CreationDateUTC+"`t"+$SplitTime
				$CreationDateUTC = Get-Date $CreationDateUTC -f "$DateFormat $TimeFormat"
				$CreationDateUTC = ([DateTime]::ParseExact($CreationDateUTC,"$DateFormat $TimeFormat",[Globalization.CultureInfo]::InvariantCulture))
			}

			#===============================================================================
			#	Add MSP Properties to Updates
			#===============================================================================
			$Update = $Update | Add-Member @{KBNumber=$KBNumber[1]} -PassThru
			$Update = $Update | Add-Member @{Classification=$Classification[1]} -PassThru
			$Update = $Update | Add-Member @{DisplayName=$DisplayName[1]} -PassThru
			$Update = $Update | Add-Member @{ProductName=$ProductName[1]} -PassThru
			$Update = $Update | Add-Member @{CreationDateUTC=$CreationDateUTC} -PassThru
			$Update = $Update | Add-Member @{PatchCode=$PatchCode} -PassThru
		}

		#===============================================================================
		#	Sort Updates in Correct Install Order
		#===============================================================================
		$OfficeUpdates = $OfficeUpdates | Select-Object -Property CreationDateUTC, LastWriteTime, KBNumber, Classification, DisplayName, ProductName, PatchCode, FullName, BaseName, Extension, Length | Sort-Object -Property @{Expression = {$_.CreationDateUTC}; Ascending = $true}, Length -Descending
		if (($GridView.IsPresent)) {$OfficeUpdates | Out-GridView -Title "Available Office Updates"}

		#===============================================================================
		#	Installing Microsoft Office Updates
		#===============================================================================
		Write-Host "`n"
		Write-Host "===============================================================================" -ForegroundColor DarkGray
		Write-Host "Installing Microsoft Office Updates"
		Write-Host "===============================================================================" -ForegroundColor DarkGray
		Write-Log -Message " " -Type Information -LogFile $LogPath
		Write-Log -Message "===============================================================================" -Type Information -LogFile $LogPath
		Write-Log -Message "Installing Microsoft Office Updates" -Type Information -LogFile $LogPath
		Write-Log -Message "===============================================================================" -Type Information -LogFile $LogPath

		ForEach ($Update in $OfficeUpdates) {
			if (($Update.BaseName -in $OfficeArrayList)) {Install-MSPUpdate -MSPFile "$($Update.FullName)"}
		}

		#===============================================================================
		#	Installing Microsoft Office Language Interface Pack Updates
		#===============================================================================

		ForEach ($Update in $OfficeUpdates) {
			if (($Update.BaseName -in $OfficeLIPArrayList)) {Install-MSPUpdate -MSPFile "$($Update.FullName)"}
		}

		#===============================================================================
		#	Installing Microsoft Office Language Pack Updates
		#===============================================================================

		ForEach ($Update in $OfficeUpdates) {
			if (($Update.BaseName -in $OfficeLPArrayList)) {Install-MSPUpdate -MSPFile "$($Update.FullName)"}
		}

		#===============================================================================
		#	Installing Microsoft Office Proofing Tools Kit Updates
		#===============================================================================

		ForEach ($Update in $OfficeUpdates) {
			if (($Update.BaseName -in $OfficePKArrayList)) {Install-MSPUpdate -MSPFile "$($Update.FullName)"}
		}

		#===============================================================================
		#	Installation Summary
		#===============================================================================
		Write-Host "`n"
		Write-Host "===============================================================================" -ForegroundColor DarkGray
		Write-Host "Installation Summary"
		Write-Host "===============================================================================" -ForegroundColor DarkGray
		Write-Log -Message " " -Type Information -LogFile $LogPath
		Write-Log -Message "===============================================================================" -Type Information -LogFile $LogPath
		Write-Log -Message "Installation Summary" -Type Information -LogFile $LogPath
		Write-Log -Message "===============================================================================" -Type Information -LogFile $LogPath
		Write-Host $CountInstall "Updates were installed"
		Write-Host $CountNotInstalled "Updates were not installed"
		Write-Log -Message "$CountInstall Updates were installed" -Type Information -LogFile $LogPath
		Write-Log -Message "$CountNotInstalled Updates were not installed" -Type Information -LogFile $LogPath

		#===============================================================================
		#	End Logging
		#===============================================================================
		Write-Log -Message " " -Type Information -LogFile $LogPath
		Write-Log -Message "- End-Logging" -Type Information -LogFile $LogPath
		Write-Log -Message "===============================================================================" -Type Information -LogFile $LogPath
	} | ft @{n="Total installation time`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t";e={$_.Hours,"Hours",$_.Minutes,"Minutes",$_.Seconds,"Seconds",$_.Milliseconds,"Milliseconds" -join " "}}