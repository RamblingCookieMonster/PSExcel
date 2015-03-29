[![Build status](https://ci.appveyor.com/api/projects/status/cew1v6k58hvfiseo/branch/master?svg=true)](https://ci.appveyor.com/project/RamblingCookieMonster/psexcel)

PSExcel: A Rudimentary Excel PowerShell Module
=============

This is a rudimentary PowerShell module for working with Excel via the [EPPlus](http://epplus.codeplex.com/) library.

* Thanks to Doug Finke for his [ImportExcel example](https://github.com/dfinke/ImportExcel/blob/master/ImportExcel) - hadn't seen EPPlus before this!
* Thanks to Philip Thompson for his [expansive module](https://excelpslib.codeplex.com/) illustrating how to work with EPPlus in PowerShell
* Thanks to the team and contributors behind [EPPlus](http://epplus.codeplex.com/) for a fantastic solution allowing .NET Excel interaction, without Excel.

Caveats:

* This covers limited functionality; contributions to this function or additional functions would be welcome!
* Minimal testing.  Contributions welcome!
* Naming conventions subject to change.  Suggestions welcome!

#Functionality

* Export random PowerShell output to Excel spreadsheets:
* Import Excel spreadsheets to PowerShell as objects
* No dependency on Excel being installed

![Example](/Media/Example.png)

#Instructions

```powershell
# One time setup
    # Download the repository
    # Unblock the zip
    # Extract the PSExcel folder to a module path (e.g. $env:USERPROFILE\Documents\WindowsPowerShell\Modules\)

# Import the module.
    Import-Module PSExcel    #Alternatively, Import-Module \\Path\To\PSExcel

# Get commands in the module
    Get-Command -Module PSExcel

# Get help for a command
    Get-Help Import-XLSX -Full

# Export data to an XLSX spreadsheet
    Get-ChildItem C:\ -File |
        Export-XLSX -Path C:\Files.xlsx

# Import data from an XLSX spreadsheet
    Import-XLSX -Path C:\Files.xlsx
```

#Notes

TODO:
* Export-XLSX should provide options similar to those Philip provides:
  * Allow various common styles.
     * Format headers.  Perhaps bold. Or highlight the background.  Maybe not.
     * Alignment
*Import-XLSX should handle dates, don't see an easy way around this?

Minimal testing:
* Import-XLSX
* Export-XLSX

Minimal to no testing:
* Everything else
