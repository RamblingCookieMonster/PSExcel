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

* Export random PowerShell output to Excel spreadsheets
* Import Excel spreadsheets to PowerShell as objects
* No dependency on Excel being installed

#Instructions

```powershell
# One time setup
    # Download the repository
    # Unblock the zip
    # Extract the PSExcel folder to a module path (e.g. $env:USERPROFILE\Documents\WindowsPowerShell\Modules\)

    #Simple alternative, if you have PowerShell 5, or the PowerShellGet module:
        Install-Module PSExcel

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

#Examples

Several examples are available on [the accompanying blog post](http://ramblingcookiemonster.github.io/PSExcel-Intro/) and the embedded [Gist](https://gist.github.com/RamblingCookieMonster/7f49beeaebb570204581#file-zpsexcel-intro-ps1).

Some highlights:

### Export and import data

```powershell
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
```

Verify that it exported:

![Excel](http://ramblingcookiemonster.github.io/images/psexcel-intro/export.png)


Check the data we imported back:

![Imported data](http://ramblingcookiemonster.github.io/images/psexcel-intro/imported.png)


### Fun with formatting

Freeze panes:

```powershell
# Open the previously created Excel file...
    $Excel = New-Excel -Path C:\temp\Demo.xlsx

# Get a Worksheet
    $Worksheet = $Excel | Get-Worksheet -Name Worksheet1

# Freeze the top row
    $Worksheet | Set-FreezePane -Row 2

# Save and close!
    $Excel | Close-Excel -Save
```


![Freeze panes](http://ramblingcookiemonster.github.io/images/psexcel-intro/frozenpane.png)


Format the header:

```powershell
# Re-open the file
    $Excel = New-Excel -Path C:\temp\Demo.xlsx

# Add bold, size 15 formatting to the header
    $Excel |
        Get-WorkSheet |
        Format-Cell -Header -Bold $True -Size 14

# Save and re-open the saved changes
    $Excel = $Excel | Save-Excel -Passthru
```


![Header format](http://ramblingcookiemonster.github.io/images/psexcel-intro/header.png)


Format the first column:

```powershell
#  Text was too large!  Set it to 11
    $Excel |
        Get-WorkSheet |
        Format-Cell -Header -Size 11

    $Excel |
        Get-WorkSheet |
        Format-Cell -StartColumn 1 -EndColumn 1 -Autofit -AutofitMinWidth -AutofitMaxWidth 7 -Color DarkRed

# Save and close
    $Excel | Save-Excel -Close
```


![First column](http://ramblingcookiemonster.github.io/images/psexcel-intro/format2.png)


### Create tables

Why format the columns yourself? Create a table (thanks to awiddersheim!):

```
# Add a table, autofit the data.  We use force to overwrite our previous demo.
    $DemoData | Export-XLSX -Path C:\Temp\Demo.xlsx -Table -Autofit -Force
```


![Table](http://ramblingcookiemonster.github.io/images/psexcel-intro/table.png)


### Pivot tables and charts

This is straight from Doug Finke's fantastic [ImportExcel module](https://github.com/dfinke/ImportExcel):

```powershell
# Fun with pivot tables and charts! Props to Doug Finke
    Get-ChildItem $env:USERPROFILE -Recurse -File |
        Export-XLSX -Path C:\Temp\Files.xlsx -PivotRows Extension -PivotValues Length -ChartType Pie
```


[![Pivot](http://ramblingcookiemonster.github.io/images/psexcel-intro/pivot.png)](http://ramblingcookiemonster.github.io/images/psexcel-intro/pivot.png)


Note that while some of these examples leverage PowerShell version 3 or later language, the module itself should work with PowerShell 2, and all Pester tests run against both PowerShell 2 and PowerShell 4.