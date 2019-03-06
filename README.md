# Update-Office

The purpose of this script is to install Office updates offline or during SCCM OS Deployment instead of WSUS, which takes forever to complete! This script reduced our SCCM OS Deployment time by 20-30 minutes depending on the hardware

**EXAMPLE 1:** .\Update-Office.ps1 -UpdateRoot "Add Custom Path Here"

Changes the default path from "$PSScriptRoot\Updates\" to the path specified

**EXAMPLE 2:** .\Update-Office.ps1 -LogRoot "Add Custom Path Here"

Changes the default path from "$PSScriptRoot\Log\" to the path specified

**EXAMPLE 3:** .\Update-Office.ps1 -GridView

Shows all available Office updates in GridView

**EXAMPLE 4:** .\Update-Office.ps1 -UpdateRoot "Add Custom Path Here" -LogRoot "Add Custom Path Here" -GridView

Changes the default path to the path specified and shows all available Office updates in GridView

**Supported Versions:** Microsoft Office Professional Plus 2010, 2013 and 2016 (it should also work with 2019, but it's not been tested yet!)

**Attention:** Very few region format are not supported! In order for the script to work, please change region format to something else eg. English (United States)
