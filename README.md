# Update-Office

The purpose of this script is to install Microsoft Office 2016 updates offline or during SCCM OS Deployment instead of WSUS, which takes forever to complete! 

This script reduced my SCCM OS Deployment time by 20-30 minutes depending on the hardware.

**SCRIPT EXAMPLE:** Update-Office.ps1 -UpdatePath "specify your own file path here"

This changes the default path from "$PSScriptRoot\Updates\" to the path specified


**FUNCTION EXAMPLE:** Update-Office -FilePath $UpdatePath -GridView

Add -GridView to the function "Update-Office" will show all available Office updates during install
