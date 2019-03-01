﻿<#
.SYNOPSIS
	Install Microsoft Office 2016 updates offline

.DESCRIPTION
	The purpose of this script is to install Microsoft Office 2016 updates offline or during SCCM OS Deployment instead of WSUS, which takes forever to complete
	This script reduced my SCCM OS Deployment time by 20-30 minutes depending on the hardware

.PARAMETER UpdatePath
	Used by script Update-Office.ps1
	Changes the default path from "$PSScriptRoot\Updates\" to the path specified

.PARAMETER GridView
	Used by the function Update-Office
	Shows all available Office updates in GridView

.EXAMPLE
	Update-Office.ps1 -UpdatePath
	Changes the default path from "$PSScriptRoot\Updates\" to the path specified
	
	Function:
	---------
	Update-Office -FilePath $UpdatePath -GridView
	Shows all available Office updates in GridView

.NOTES
	Version: 1.9.3.1
	Author: Sune Thomsen
	Creation date: 22-02-2019
	Last modified date: 01-03-2019

.LINK
	https://github.com/SuneThomsenDK
#>
	#===============================================================================
	#	Requirements
	#===============================================================================
	#Requires -Version 4
	#Requires -RunAsAdministrator

	Param (
		[IO.FileInfo][String]$UpdatePath = "$PSScriptRoot\Updates\"
	)

	Function Get-MSPInfo {
		Param (
			[Parameter(Mandatory = $true)][IO.FileInfo][String]$MSPFile,
			[Parameter(Mandatory = $true)][ValidateSet('Classification', 'DisplayName', 'KBArticle Number', 'TargetProductName', 'CreationTimeUTC')][String]$Property
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
			Write-Output $_.Exception.Message
			Return $NULL
		}
	}

	Function Get-MSPPatchCode {
		Param (
			[Parameter(Mandatory = $true)][IO.FileInfo][String]$MSPFile
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
			Write-Output $_.Exception.Message
			Return $NULL
		}
	}

	Function Install-MSPUpdate {
		Param (
			[Parameter(Mandatory = $true)][IO.FileInfo][String]$MSPFile
		)
		Try {
			#===============================================================================
			#	Install MSP Update
			#===============================================================================
			$Process = "C:\Windows\System32\msiexec.exe"
			$MSPInstall = Start-Process $process -ArgumentList "/p $MSPFile /qn REBOOT=ReallySuppress MSIRESTARTMANAGERCONTROL=Disable" -PassThru -Wait
			$MSPInstall.WaitForExit()
			if (($MSPInstall.ExitCode -eq 0) -or ($MSPInstall.ExitCode -eq 3010)){
				$Script:CountInstall++
				Write-Host "Installing: $DisplayName ($($Update.BaseName))" -foregroundcolor "Green"
			}
			else {
				$Script:CountNotInstalled++
				Write-Host "Attention: $DisplayName ($($Update.BaseName)) were not installed" -foregroundcolor "Cyan"
				Write-Host "Possible cause: The program to be updated might not be installed, or the patch may update a different version of the program."
			}
		}
		Catch {
			Write-Output $_.Exception.Message
			Return $NULL
		}
	}

	Function Update-Office {
		Param (
			[Parameter(Mandatory = $true)][IO.FileInfo][String]$FilePath,
			[Switch]$GridView
		)
		Measure-Command -Expression {
			#===============================================================================
			#	Set Variables
			#===============================================================================
			$OfficeUpdates = Get-ChildItem $FilePath -Recurse -File -Include *.msp
			$RegPath = "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
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

			Write-Host "`n"
			Write-Host "===============================================================================" -ForegroundColor DarkGray
			Write-Host "Installing Microsoft Office 2016 Updates"
			Write-Host "===============================================================================" -ForegroundColor DarkGray

			ForEach ($Update in $OfficeUpdates) {
				#===============================================================================
				#	Get MSP Properties
				#===============================================================================
				$KBNumber = Get-MSPInfo -MSPFile $Update.FullName -Property 'KBArticle Number'
				$Classification = Get-MSPInfo -MSPFile $Update.FullName -Property 'Classification'
				$DisplayName = Get-MSPInfo -MSPFile $Update.FullName -Property 'DisplayName'
				$ProductName = Get-MSPInfo -MSPFile $Update.FullName -Property 'TargetProductName'
				$CreationDateUTC = Get-MSPInfo -MSPFile $Update.FullName -Property 'CreationTimeUTC'
				$PatchCode = Get-MSPPatchCode -MSPFile $Update.FullName

				#===============================================================================
				#	Format CreationDateUTC
				#===============================================================================
				$CreationDateUTC = $CreationDateUTC[1]
				$CreationDateUTC = $CreationDateUTC.Split('/')
				$CreationDateUTC = "{0}/{1}/{2}" -f $CreationDateUTC[1],$CreationDateUTC[0],$CreationDateUTC[2]
				$CreationDateUTC = Get-Date $CreationDateUTC -f "dd-MM-yyyy HH:mm:ss"
				$CreationDateUTC = ([DateTime]::ParseExact($CreationDateUTC,"dd-MM-yyyy HH:mm:ss",[Globalization.CultureInfo]::InvariantCulture))
				$CreationDateUTC

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
			if ($GridView.IsPresent) {$OfficeUpdates | Out-GridView -Title "Available Office Updates"}

			#===============================================================================
			#	Update Office
			#===============================================================================
			ForEach ($Update in $OfficeUpdates) {
				if (($Update.BaseName -in $OfficeArrayList)) {
					$KBNumber = $Update.KBNumber
					$DisplayName = $Update.DisplayName
					$PatchCode = $Update.PatchCode
					$CheckPatchCode = Get-ItemProperty $RegPath | Where-Object {$_.PSChildName -like "*$PatchCode*" -or $_.UninstallString -like "*$PatchCode*"} | Select-Object -Property PSChildName, DisplayName, UninstallString | Sort-Object -Property DisplayName -Unique

					Try {
						if (!($CheckPatchCode)) {
							Install-MSPUpdate -MSPFile "$($Update.FullName)"
						}
						else {
							$Script:CountNotInstalled++
							Write-Host "Attention: $DisplayName ($($Update.BaseName)) is already installed" -foregroundcolor "Cyan"
						}
					}
					Catch {
						Write-Output $_.Exception.Message
						Write-Host "Warning: Sune has created a awesome script, but something went wrong!" -foregroundcolor "Yellow"
					}
				}
			}

			Write-Host "`n"
			Write-Host "===============================================================================" -ForegroundColor DarkGray
			Write-Host "Installing Microsoft Office 2016 Language Interface Pack Updates"
			Write-Host "===============================================================================" -ForegroundColor DarkGray

			ForEach ($Update in $OfficeUpdates) {
				if (($Update.BaseName -in $OfficeLIPArrayList)) {
					$KBNumber = $Update.KBNumber
					$DisplayName = $Update.DisplayName
					$PatchCode = $Update.PatchCode
					$CheckPatchCode = Get-ItemProperty $RegPath | Where-Object {$_.PSChildName -like "*$PatchCode*" -or $_.UninstallString -like "*$PatchCode*"} | Select-Object -Property PSChildName, DisplayName, UninstallString | Sort-Object -Property DisplayName -Unique

					Try {
						if (!($CheckPatchCode)) {
							Install-MSPUpdate -MSPFile "$($Update.FullName)"
						}
						else {
							$Script:CountNotInstalled++
							Write-Host "Attention: $DisplayName ($($Update.BaseName)) is already installed" -foregroundcolor "Cyan"
						}
					}
					Catch {
						Write-Output $_.Exception.Message
						Write-Host "Warning: Sune has created a awesome script, but something went wrong!" -foregroundcolor "Yellow"
					}
				}
			}

			Write-Host "`n"
			Write-Host "===============================================================================" -ForegroundColor DarkGray
			Write-Host "Installing Microsoft Office 2016 Language Pack Updates"
			Write-Host "===============================================================================" -ForegroundColor DarkGray

			ForEach ($Update in $OfficeUpdates) {
				if (($Update.BaseName -in $OfficeLPArrayList)) {
					$KBNumber = $Update.KBNumber
					$DisplayName = $Update.DisplayName
					$PatchCode = $Update.PatchCode
					$CheckPatchCode = Get-ItemProperty $RegPath | Where-Object {$_.PSChildName -like "*$PatchCode*" -or $_.UninstallString -like "*$PatchCode*"} | Select-Object -Property PSChildName, DisplayName, UninstallString | Sort-Object -Property DisplayName -Unique

					Try {
						if (!($CheckPatchCode)) {
							Install-MSPUpdate -MSPFile "$($Update.FullName)"
						}
						else {
							$Script:CountNotInstalled++
							Write-Host "Attention: $DisplayName ($($Update.BaseName)) is already installed" -foregroundcolor "Cyan"
						}
					}
					Catch {
						Write-Output $_.Exception.Message
						Write-Host "Warning: Sune has created a awesome script, but something went wrong!" -foregroundcolor "Yellow"
					}
				}
			}

			Write-Host "`n"
			Write-Host "===============================================================================" -ForegroundColor DarkGray
			Write-Host "Installing Microsoft Office 2016 Proofing Kit Updates"
			Write-Host "===============================================================================" -ForegroundColor DarkGray

			ForEach ($Update in $OfficeUpdates) {
				if (($Update.BaseName -in $OfficePKArrayList)) {
					$KBNumber = $Update.KBNumber
					$DisplayName = $Update.DisplayName
					$PatchCode = $Update.PatchCode
					$CheckPatchCode = Get-ItemProperty $RegPath | Where-Object {$_.PSChildName -like "*$PatchCode*" -or $_.UninstallString -like "*$PatchCode*"} | Select-Object -Property PSChildName, DisplayName, UninstallString | Sort-Object -Property DisplayName -Unique

					Try {
						if (!($CheckPatchCode)) {
							Install-MSPUpdate -MSPFile "$($Update.FullName)"
						}
						else {
							$Script:CountNotInstalled++
							Write-Host "Attention: $DisplayName ($($Update.BaseName)) is already installed" -foregroundcolor "Cyan"
						}
					}
					Catch {
						Write-Output $_.Exception.Message
						Write-Host "Warning: Sune has created a awesome script, but something went wrong!" -foregroundcolor "Yellow"
					}
				}
			}

			Write-Host "`n"
			Write-Host "===============================================================================" -ForegroundColor DarkGray
			Write-Host "Installation Summary"
			Write-Host "===============================================================================" -ForegroundColor DarkGray
			Write-Host $CountInstall "Updates were installed"
			Write-Host $CountNotInstalled "Updates were not installed"
		} | ft @{n="Total installation time`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t`t";e={$_.Hours,"Hours",$_.Minutes,"Minutes",$_.Seconds,"Seconds",$_.Milliseconds,"Milliseconds" -join " "}}
	}

Update-Office -FilePath $UpdatePath

	#Write-Host "`n"
	#Read-Host "Press any key to exit..."
	#Exit