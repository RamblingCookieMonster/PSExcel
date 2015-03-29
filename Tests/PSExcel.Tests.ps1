#handle PS2
if(-not $PSScriptRoot)
{
    $PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent
}

$Verbose = @{}
if($env:APPVEYOR_REPO_BRANCH -and $env:APPVEYOR_REPO_BRANCH -notlike "master")
{
    $Verbose.add("Verbose",$True)
}

$PSVersion = $PSVersionTable.PSVersion.Major
Import-Module $PSScriptRoot\..\PSExcel -Force

#Set up some data we will use in testing
    $ExistingXLSXFile = "$PSScriptRoot\Working.xlsx"
    Remove-Item $ExistingXLSXFile  -force -ErrorAction SilentlyContinue
    Copy-Item $PSScriptRoot\Test.xlsx $ExistingXLSXFile -force

    $NewXLSXFile = "$PSScriptRoot\New.xlsx"
    Remove-Item $NewXLSXFile  -force -ErrorAction SilentlyContinue

    $Files = Get-ChildItem $PSScriptRoot | Where {-not $_.PSIsContainer}

Describe "New-Excel PS$PSVersion" {
    
    Context 'Strict mode' { 

        Set-StrictMode -Version latest

        It 'should create an ExcelPackage' {
            $Excel = New-Excel
            $Excel -is [OfficeOpenXml.ExcelPackage] | Should Be $True
            $Excel.Dispose()
            $Excel = $Null

            $Excel = New-Excel -Path $NewXLSXFile
            $Excel -is [OfficeOpenXml.ExcelPackage] | Should Be $True
            $Excel.Dispose()
            $Excel = $Null

        }

        It 'should reflect the correct path' {
            Remove-Item $NewXLSXFile -force -ErrorAction silentlycontinue
            $Excel = New-Excel -Path $NewXLSXFile
            $Excel.File | Should be $NewXLSXFile
            $Excel.Dispose()
            $Excel = $Null
        }

        It 'should not create a file' {
            Test-Path $NewXLSXFile | Should Be $False
        }
    }
}


Describe "Import-XLSX PS$PSVersion" {
    
    Context 'Strict mode' { 

        Set-StrictMode -Version latest

        It 'should import data with expected results' {
            $ExcelData = Import-XLSX -Path $ExistingXLSXFile
            $Props = $ExcelData[0].PSObject.Properties | Select -ExpandProperty Name

            $ExcelData.count | Should be 10
            $Props[0] | Should be 'Val'
            $Props[1] | Should be 'Name'
            $Exceldata[0].val | Should be '944041859'
            $Exceldata[0].name | Should be 'Prop1'
        }

        It 'should replace headers' {
            $ExcelData = Import-XLSX -Path $ExistingXLSXFile -Header one, two
            $Props = $ExcelData[0].PSObject.Properties | Select -ExpandProperty Name

            $Props[0] | Should be 'one'
            $Props[1] | Should be 'two'
        }
    }
}

Describe "Export-XLSX PS$PSVersion" {
    
    Context 'Strict mode' { 

        Set-StrictMode -Version latest

        It 'should create a file' {
            $Files | Export-XLSX -Path $NewXLSXFile

            Test-Path $NewXLSXFile | Should Be $True

        }

        It 'should add the correct number of rows' {
            $ExportedData = Import-XLSX -Path $NewXLSXFile
            $Files.Count | Should be $ExportedData.count
        }

    }
}

Describe "Close-Excel PS$PSVersion" {
    
    Context 'Strict mode' { 

        Set-StrictMode -Version latest

        It 'should close an excelpackage' {
            $Excel = New-Excel -Path $NewXLSXFile
            $File = $Excel.File
            $Excel | Close-Excel
            $Excel.File -like $File | Should be $False
        }

        It 'should save when requested' {
            Remove-Item $NewXLSXFile -Force -ErrorAction SilentlyContinue
            $Excel = New-Excel -Path $NewXLSXFile
            [void]$Excel.Workbook.Worksheets.Add(1)
            $Excel | Close-Excel -Save
            Test-Path $NewXLSXFile | Should be $True
        }

        It 'should save as a specified path' {
            $Excel = New-Excel -Path $NewXLSXFile
            $Excel | Close-Excel -Path "$NewXLSXFile`2"
            Test-Path "$NewXLSXFile`2" | Should be $True
            Remove-Item "$NewXLSXFile`2" -Force -ErrorAction SilentlyContinue
        }
    }
}


Remove-Item $NewXLSXFile -force -ErrorAction SilentlyContinue
Remove-Item $ExistingXLSXFile -force -ErrorAction SilentlyContinue