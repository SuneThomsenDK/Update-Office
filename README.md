# Update-Office

The purpose of this script is to install Microsoft Office updates offline or during SCCM OS Deployment instead of WSUS, which takes forever to complete! 

This script reduced my SCCM OS Deployment time by 20-30 minutes depending on the hardware.

**SCRIPT EXAMPLE:** Update-Office.ps1 -UpdateRoot "specify your own file path here"

This changes the default path from "$PSScriptRoot\Updates\" to the path specified


**FUNCTION EXAMPLE:** Update-Office -FilePath $UpdateRoot -GridView

Add -GridView to the function "Update-Office" will show all available Office updates during install

**Supported Versions:** Microsoft Office Professional Plus 2010, 2013 and 2016 (it should also work with 2019, but it's not been tested yet!)
