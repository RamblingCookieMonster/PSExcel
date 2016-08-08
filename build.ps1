# Grab nuget bits, install modules, set build variables, start build.
Get-PackageProvider -Name NuGet -ForceBootstrap | Out-Null

Install-Module PSDepend -Force
Invoke-PSDepend -Force -verbose

# For PS2, after installing with PS5.
Move-Item C:\temp\pester\*\* -Destination C:\temp\pester -force