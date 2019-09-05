<#
.SYNOPSIS
	Install Microsoft Office Professional Plus 2010, 2013 and 2016 updates offline

.DESCRIPTION
	The purpose of this script is to install Office updates offline or during SCCM OS Deployment instead of WSUS, which takes forever to complete
	This script reduced our SCCM OS Deployment time by 20-30 minutes depending on the hardware configuration

	If you are installing Language Packs, Language Interface Packs or Proofing Tools Kits, you have to install updates in a specific order and
	that's done by adding them to the following ArrayLists:

	"ArrayList_OfficeCore.txt"
	"ArrayList_OfficeLIP.txt"
	"ArrayList_OfficeLP.txt"
	"ArrayList_OfficePK.txt"

	If you do not install Language Packs, Language Interface Packs or Proofing Tools Kits, it´s best to leave these ArrayList files empty
	or just delete them

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
	Version: 1.9.3.15
	Author: Sune Thomsen
	Creation date: 22-02-2019
	Last modified date: 15-03-2019

.LINK
	https://github.com/SuneThomsenDK
#>

	#=========================================================================================
	#	Requirements
	#=========================================================================================
	#Requires -Version 4
	#Requires -RunAsAdministrator

	Param (
		[System.IO.FileInfo][String]$UpdateRoot = "$PSScriptRoot\Updates\",
		[System.IO.FileInfo][String]$LogRoot = "$PSScriptRoot\Log\",
		[Switch]$GridView
	)

	Function Write-Log {
		[CmdletBinding()]
		Param (
			[Parameter(Mandatory=$false)][String]$LogFile,
			[Parameter(Mandatory=$true)][String]$Message,
			[Parameter(Mandatory=$false)][ValidateSet("Information","Warning","Error")][String]$Type = "Information"
		)

		$LogTime = (Get-Date).toString("yyyy-MM-dd HH:mm:ss")
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
			#=========================================================================================
			#	Get MSP Information
			#=========================================================================================
			$WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
			$MSPDatabase = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $Null, $WindowsInstaller, @($MSPFile.FullName, 32))
			$MSPQuery = "SELECT Value FROM MsiPatchMetadata WHERE Property = '$($Property)'"
			$MSPView = $MSPDatabase.GetType().InvokeMember("OpenView", "InvokeMethod", $Null, $MSPDatabase, ($MSPQuery))
			$MSPView.GetType().InvokeMember("Execute", "InvokeMethod", $Null, $MSPView, $Null)
			$MSPRecord = $MSPView.GetType().InvokeMember("Fetch", "InvokeMethod", $Null, $MSPView, $Null)
			$MSPValue = $MSPRecord.GetType().InvokeMember("StringData", "GetProperty", $Null, $MSPRecord, 1)
			Return $MSPValue
		}
		Catch {
			Write-Host "Warning: Sune has created a awesome script, but something went wrong!" -ForegroundColor "Yellow"
			Write-Log -Message "Sune has created a awesome script, but something went wrong!" -Type Error -LogFile $LogPath
			Write-Log -Message "$_.Exception.Message" -Type Error -LogFile $LogPath
			Return $Null
		}
	}

	Function Get-MSPPatchCode {
		[CmdletBinding()]
		Param (
			[Parameter(Mandatory = $true)][System.IO.FileInfo][String]$MSPFile
		)
		Try {
			#=========================================================================================
			#	Get MSP PatchCode
			#=========================================================================================
			$WindowsInstaller = New-Object -ComObject WindowsInstaller.Installer
			$MSPDatabase = $WindowsInstaller.GetType().InvokeMember("OpenDatabase", "InvokeMethod", $Null, $WindowsInstaller, $($MSPFile.FullName, 32))
			$MSPSummary = $MSPDatabase.GetType().InvokeMember("SummaryInformation", "GetProperty", $Null, $MSPDatabase, $Null)
			[String]$MSPPatchCode = $MSPSummary.GetType().InvokeMember("Property", "GetProperty", $Null, $MSPSummary, 9)
			Return $MSPPatchCode
		}
		Catch {
			Write-Host "Warning: Sune has created a awesome script, but something went wrong!" -ForegroundColor "Yellow"
			Write-Log -Message "Sune has created a awesome script, but something went wrong!" -Type Error -LogFile $LogPath
			Write-Log -Message "$_.Exception.Message" -Type Error -LogFile $LogPath
			Return $Null
		}
	}

	Function Check-Registry {
		Try {
			#=========================================================================================
			#	Check PatchCode in Registry
			#=========================================================================================
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
			Write-Host "Warning: Sune has created a awesome script, but something went wrong!" -ForegroundColor "Yellow"
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
			#=========================================================================================
			#	Install MSP Update
			#=========================================================================================
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
				$MSPInstall = Start-Process $Process -ArgumentList $MSPArguments -PassThru -Wait
				$MSPInstall.WaitForExit()
				if (($MSPInstall.ExitCode -eq 0) -or ($MSPInstall.ExitCode -eq 3010)) {
					$Script:CountInstall++
					Write-Host "Installing: $DisplayName ($($Update.BaseName))" -ForegroundColor "Green"
					Write-Log -Message "Installing $DisplayName ($($Update.BaseName))" -Type Information -LogFile $LogPath
				}
				else {
					$Script:CountNotInstalled++
					Write-Host "Attention: $DisplayName ($($Update.BaseName)) were not installed" -ForegroundColor "Cyan"
					Write-Host "Possible cause: The program to be updated might not be installed, or the patch may update a different version of the program."
					Write-Log -Message "$DisplayName ($($Update.BaseName)) were not installed" -Type Warning -LogFile $LogPath
					Write-Log -Message "Possible cause: The program to be updated might not be installed, or the patch may update a different version of the program." -Type Information -LogFile $LogPath
				}
			}
			else {
				$Script:CountNotInstalled++
				Write-Host "Attention: $DisplayName ($($Update.BaseName)) is already installed" -ForegroundColor "Cyan"
				Write-Log -Message "$DisplayName ($($Update.BaseName)) is already installed" -Type Information -LogFile $LogPath
			}
		}
		Catch {
			Write-Host "Warning: Sune has created a awesome script, but something went wrong!" -ForegroundColor "Yellow"
			Write-Log -Message "Sune has created a awesome script, but something went wrong!" -Type Error -LogFile $LogPath
			Write-Log -Message "$_.Exception.Message" -Type Error -LogFile $LogPath
			Return $Null
		}
	}

	Measure-Command -Expression {

		#=========================================================================================
		#	Check That Update Root Exists
		#=========================================================================================
		if (!(Test-Path -Path $UpdateRoot)) {
			$PSCommandPath = $PSCommandPath.Split("\")[2]
			Write-Host "$($PSCommandPath): Cannot find $UpdateRoot because it does not exist! Please verify that the path is correct and try again." -ForegroundColor "Yellow"
			Exit
		}

		#=========================================================================================
		#	Set Variables
		#=========================================================================================
		$LocalCulture = Get-Culture
		$RegionFormat = [System.Globalization.CultureInfo]::GetCultureInfo($LocalCulture.LCID).DateTimeFormat.FullDateTimePattern

		$LogFileTime = (Get-Date).toString("yyyy-MM-dd-HHmmss")
		$LogFile = "$($LogFileTime)_Update-Office.log"
		$LogPath = Join-Path "$LogRoot" "$LogFile"

		$OfficeUpdates = Get-ChildItem $UpdateRoot -Recurse -File -Include *.msp

		$OfficeCorePath = "$PSScriptRoot\ArrayList_OfficeCore.txt"
		$OfficeLIPPath = "$PSScriptRoot\ArrayList_OfficeLIP.txt"
		$OfficeLPPath = "$PSScriptRoot\ArrayList_OfficeLP.txt"
		$OfficePKPath = "$PSScriptRoot\ArrayList_OfficePK.txt"

		if ((Test-Path -Path $OfficeCorePath)) {$OfficeCoreArrayList = Get-content -Path $OfficeCorePath}
		if ((Test-Path -Path $OfficeLIPPath)) {$OfficeLIPArrayList = Get-content -Path $OfficeLIPPath}
		if ((Test-Path -Path $OfficeLPPath)) {$OfficeLPArrayList = Get-content -Path $OfficeLPPath}
		if ((Test-Path -Path $OfficePKPath)) {$OfficePKArrayList = Get-content -Path $OfficePKPath}

		$OfficeArrayList = $OfficeCoreArrayList + $OfficeLIPArrayList + $OfficeLPArrayList + $OfficePKArrayList

		$Script:CountInstall = 0
		$Script:CountNotInstalled = 0

		#=========================================================================================
		#	Create Log Folder
		#=========================================================================================
		if (!(Test-Path -Path $LogRoot)) {New-Item $LogRoot -ItemType Directory | Out-Null}

		#=========================================================================================
		#	Start Logging
		#=========================================================================================
		Write-Log -Message "=========================================================================================" -Type Information -LogFile $LogPath
		Write-Log -Message "- Start-Logging" -Type Information -LogFile $LogPath
		Write-Log -Message " " -Type Information -LogFile $LogPath
		Write-Log -Message "- Username: $Env:USERNAME" -Type Information -LogFile $LogPath
		Write-Log -Message "- Computername: $env:COMPUTERNAME" -Type Information -LogFile $LogPath
		Write-Log -Message "- Update path: $UpdateRoot" -Type Information -LogFile $LogPath
		Write-Log -Message "- Log path: $LogRoot" -Type Information -LogFile $LogPath

		ForEach ($Update in $OfficeUpdates) {
			#=========================================================================================
			#	Get MSP Properties
			#=========================================================================================
			$KBNumber = Get-MSPInfo -MSPFile $Update.FullName -Property "KBArticle Number"
			$Classification = Get-MSPInfo -MSPFile $Update.FullName -Property "Classification"
			$DisplayName = Get-MSPInfo -MSPFile $Update.FullName -Property "DisplayName"
			$ProductName = Get-MSPInfo -MSPFile $Update.FullName -Property "TargetProductName"
			$CreationDateUTC = Get-MSPInfo -MSPFile $Update.FullName -Property "CreationTimeUTC"
			$PatchCode = Get-MSPPatchCode -MSPFile $Update.FullName

			#=========================================================================================
			#	Format CreationDateUTC
			#=========================================================================================
			$CreationDateUTC = $CreationDateUTC[1]
			$CreationDateUTC = [DateTime]::ParseExact($CreationDateUTC, "MM/dd/yy HH:mm", $Null)
			$CreationDateUTC = Get-Date $CreationDateUTC -f $RegionFormat
			$CreationDateUTC = [DateTime]::ParseExact($CreationDateUTC, $RegionFormat, $Null)

			#=========================================================================================
			#	Add MSP Properties to Updates
			#=========================================================================================
			$Update = $Update | Add-Member @{KBNumber=$KBNumber[1]} -PassThru
			$Update = $Update | Add-Member @{Classification=$Classification[1]} -PassThru
			$Update = $Update | Add-Member @{DisplayName=$DisplayName[1]} -PassThru
			$Update = $Update | Add-Member @{ProductName=$ProductName[1]} -PassThru
			$Update = $Update | Add-Member @{CreationDateUTC=$CreationDateUTC} -PassThru
			$Update = $Update | Add-Member @{PatchCode=$PatchCode} -PassThru
		}

		#=========================================================================================
		#	Sort Updates in Correct Install Order
		#=========================================================================================
		$OfficeUpdates = $OfficeUpdates | Select-Object -Property CreationDateUTC, LastWriteTime, KBNumber, Classification, DisplayName, ProductName, PatchCode, FullName, BaseName, Extension, Length | Sort-Object -Property @{Expression = {$_.CreationDateUTC}; Ascending = $true}, Length -Descending
		if (($GridView.IsPresent)) {$OfficeUpdates | Out-GridView -Title "Available Office Updates"}

		#=========================================================================================
		#	Installing Microsoft Office Updates
		#=========================================================================================
		Write-Host "`n"
		Write-Host "=========================================================================================" -ForegroundColor "DarkGray"
		Write-Host "Installing Microsoft Office Updates"
		Write-Host "=========================================================================================" -ForegroundColor "DarkGray"
		Write-Log -Message " " -Type Information -LogFile $LogPath
		Write-Log -Message "=========================================================================================" -Type Information -LogFile $LogPath
		Write-Log -Message "Installing Microsoft Office Updates" -Type Information -LogFile $LogPath
		Write-Log -Message "=========================================================================================" -Type Information -LogFile $LogPath

		if (($OfficeCoreArrayList -notlike $Null)) {Write-Host "Attention: Updates were found in Office Core ArrayList and will be installed in correct order." -ForegroundColor "DarkGray"}
		if (($OfficeLIPArrayList -notlike $Null)) {Write-Host "Attention: Updates were found in Office LIP ArrayList and will be installed in correct order." -ForegroundColor "DarkGray"}
		if (($OfficeLPArrayList -notlike $Null)) {Write-Host "Attention: Updates were found in Office LP ArrayList and will be installed in correct order." -ForegroundColor "DarkGray"}
		if (($OfficePKArrayList -notlike $Null)) {Write-Host "Attention: Updates were found in Office PK ArrayList and will be installed in correct order." -ForegroundColor "DarkGray"}
		if (($OfficeArrayList -notlike $Null)) {Write-Host "`n"}

		ForEach ($Update in $OfficeUpdates) {
			if (($Update.BaseName -in $OfficeCoreArrayList)) {Install-MSPUpdate -MSPFile $($Update.FullName)}
		}

		#=========================================================================================
		#	Installing Microsoft Office Updates Defined in ArrayList for Language Interface Pack
		#=========================================================================================

		ForEach ($Update in $OfficeUpdates) {
			if (($Update.BaseName -in $OfficeLIPArrayList)) {Install-MSPUpdate -MSPFile $($Update.FullName)}
		}

		#=========================================================================================
		#	Installing Microsoft Office Updates Defined in ArrayList for Language Pack
		#=========================================================================================

		ForEach ($Update in $OfficeUpdates) {
			if (($Update.BaseName -in $OfficeLPArrayList)) {Install-MSPUpdate -MSPFile $($Update.FullName)}
		}

		#=========================================================================================
		#	Installing Microsoft Office Updates Defined in ArrayList for Proofing Tools Kit
		#=========================================================================================

		ForEach ($Update in $OfficeUpdates) {
			if (($Update.BaseName -in $OfficePKArrayList)) {Install-MSPUpdate -MSPFile $($Update.FullName)}
		}

		#=========================================================================================
		#	Installing Microsoft Office Updates not Defined in ArrayList
		#=========================================================================================

		ForEach ($Update in $OfficeUpdates) {
			if (($Update.BaseName -notin $OfficeArrayList)) {Install-MSPUpdate -MSPFile $($Update.FullName)}
		}

		#=========================================================================================
		#	Installation Summary
		#=========================================================================================
		Write-Host "`n"
		Write-Host "=========================================================================================" -ForegroundColor "DarkGray"
		Write-Host "Installation Summary"
		Write-Host "=========================================================================================" -ForegroundColor "DarkGray"
		Write-Log -Message " " -Type Information -LogFile $LogPath
		Write-Log -Message "=========================================================================================" -Type Information -LogFile $LogPath
		Write-Log -Message "Installation Summary" -Type Information -LogFile $LogPath
		Write-Log -Message "=========================================================================================" -Type Information -LogFile $LogPath
		Write-Host $CountInstall "Updates were installed"
		Write-Host $CountNotInstalled "Updates were not installed"
		Write-Log -Message "$CountInstall Updates were installed" -Type Information -LogFile $LogPath
		Write-Log -Message "$CountNotInstalled Updates were not installed" -Type Information -LogFile $LogPath

		#=========================================================================================
		#	End Logging
		#=========================================================================================
		Write-Log -Message " " -Type Information -LogFile $LogPath
		Write-Log -Message "- End-Logging" -Type Information -LogFile $LogPath
		Write-Log -Message "=========================================================================================" -Type Information -LogFile $LogPath
	} | ft @{n="Total installation time`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t";e={$_.Hours,"Hours",$_.Minutes,"Minutes",$_.Seconds,"Seconds",$_.Milliseconds,"Milliseconds" -join " "}}