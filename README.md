[![Build status](https://ci.appveyor.com/api/projects/status/cew1v6k58hvfiseo/branch/master?svg=true)](https://ci.appveyor.com/project/RamblingCookieMonster/psexcel)

PSExcel
=============

This is a rudimentary PowerShell module for working with Excel via the [EPPlus](http://epplus.codeplex.com/) library, with no dependencies on Excel itself.

* Thanks to Doug Finke for his [ImportExcel example](https://github.com/dfinke/ImportExcel) - hadn't seen EPPlus before this!
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

#Create some demo data
    $DemoData = 1..10 | Foreach-Object{

        $EID = Get-Random -Minimum 1 -Maximum 1000
        $Date = (Get-Date).adddays(-$EID)

        New-Object -TypeName PSObject -Property @{
            Name = "jsmith$_"
            EmployeeID = $EID
            Date = $Date
        } | Select Name, EmployeeID, Date
    }

# Export it
    $DemoData | Export-XLSX -Path C:\temp\Demo.xlsx

# Import it back
    $Imported = Import-XLSX -Path C:\Temp\Demo.xlsx -Header samaccountname, EID, Date

# Open that Excel file...
    $Excel = New-Excel -Path C:\temp\Demo.xlsx

# Get a workbook
    $Workbook = $Excel | Get-Workbook

# Get a worksheet - can pipe ExcelPackage or Workbook.
# Filtering on Name is optional
    $Worksheet = $Excel | Get-Worksheet
    $Worksheet = $Workbook | Get-Worksheet -Name Worksheet1

# Freeze the top row
    $Worksheet | Set-FreezePane -Row 2

# Save and close!
    $Excel | Close-Excel -Save
```
