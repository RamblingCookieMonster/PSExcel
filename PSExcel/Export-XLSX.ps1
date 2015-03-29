Function Export-XLSX {
    <#
    .SYNOPSIS
        Export data to an XLSX file

    .DESCRIPTION
        Export data to an XLSX file

    .PARAMETER InputObject
        Data to export

    .PARAMETER Path
        Path to the file to export

    .PARAMETER WorksheetName
        Name the worksheet you are importing to

    .PARAMETER Header
        Header to use. Must match order and count of your data's properties

    .PARAMETER Force
        If file exists, overwrite it.

    .EXAMPLE
        $Files = Get-ChildItem C:\ -File

        Export-XLSX -Path C:\Files.xlsx -InputObject $Files

        Export file listing to C:\Files.xlsx

    .EXAMPLE

        1..10 | Foreach-Object {
            New-Object -typename PSObject -Property @{
                Something = "Prop$_"
                Value = Get-Random
            }
        } |
            Select-Object Something, Value |
            Export-XLSX -Path C:\Random.xlsx -Force -Header Name, Val

        # Generate data
        # Send it to export-xlsx
        # give it new headers
        # overwrite C:\random.xlsx if it exists

    .NOTES
        Thanks to Doug Finke for his example:
            https://github.com/dfinke/ImportExcel/blob/master/ImportExcel.psm1

        Thanks to Philip Thompson for an expansive set of examples on working with EPPlus in PowerShell:
            https://excelpslib.codeplex.com/

    .FUNCTIONALITY
        Excel
    #>
    [CmdletBinding()]
    param(
        [parameter( Position = 0,
                    Mandatory=$true )]
        [ValidateScript({
            $Parent = Split-Path $_ -Parent
            if( -not (Test-Path -Path $Parent -PathType Container) )
            {
                Throw "Specify a valid path.  Parent '$Parent' does not exist: $_"
            }
            $True
        })]
        [string]$Path,

        [parameter( Position = 1,
                    Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromRemainingArguments=$false)]
        $InputObject,

        [string[]]$Header,

        [string]$WorksheetName = "Worksheet1",

        [switch]$Force
    )
    begin
    {
        if ( Test-Path $Path )
        {
            if($Force)
            {
                Try
                {
                    Remove-Item -Path $Path -Force -Confirm:$False
                }
                Catch
                {
                    Throw "'$Path' exists and could not be removed: $_"
                }
            }
            else
            {
                Throw "'$Path' exists.  Use -Force to overwrite."
            }
        }

        #Resolve relative paths... Thanks Oisin! http://stackoverflow.com/a/3040982/3067642
        $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)

        Write-Verbose "PSBoundParameters = $($PSBoundParameters | Out-String)"

        $bound = $PSBoundParameters.keys -contains "InputObject"
        if(-not $bound)
        {
            [System.Collections.ArrayList]$AllData = @()
        }

    }
    process
    {
        #We write data by row, so need everything countable, not going to stream...
        if($bound)
        {
            $AllData = $InputObject
        }
        Else
        {
            foreach($Object in $InputObject)
            {
                [void]$AllData.add($Object)
            }
        }
    }
    end
    {

        #Deal with headers
            $ExistingHeader = @(

                # indexes might be an issue if we get array of strings, so select first
                ($AllData | Select -first 1).PSObject.Properties |
                    Select -ExpandProperty Name
            )

            $Columns = $ExistingHeader.count

            if($Header)
            {
                if($Header.count -ne $ExistingHeader.count)
                {
                    Throw "Found '$columns' columns, provided $($header.count) headers.  You must provide a header for every column."
                }
            }
            else
            {
                $Header = $ExistingHeader
            }

        #initialize stuff
            Try
            {
                $Excel = New-Object OfficeOpenXml.ExcelPackage($Path)
                $Workbook = $Excel.Workbook
                $WorkSheet = $Workbook.Worksheets.Add($WorkSheetName)
            }
            Catch
            {
                Throw "Failed to initialize Excel, Workbook, or Worksheet: $_"
            }

        #Set those headers
            for ($ColumnIndex = 1; $ColumnIndex -le $Header.count; $ColumnIndex++)
            {
                $WorkSheet.SetValue(1, $ColumnIndex, $Header[$ColumnIndex - 1])
            }

        #Write the data...
            $RowIndex = 2
            foreach($RowData in $AllData)
            {
                Write-Verbose "Working on object:`n$($RowData | Out-String)"
                for ($ColumnIndex = 1; $ColumnIndex -le $Header.count; $ColumnIndex++)
                {
                    $Object = @($RowData.PSObject.Properties)[$ColumnIndex - 1]
                    $Value = $Object.Value
                    $WorkSheet.SetValue($RowIndex, $ColumnIndex, $Value)

                    Try
                    {
                        #Nulls will error, catch them
                        $ThisType = $Null
                        $ThisType = $Value.GetType().FullName
                    }
                    Catch
                    {
                        Write-Verbose "Applying no style to null in row $RowIndex, column $ColumnIndex"
                    }

                    #Idea from Philip Thompson, thank you Philip!
                    $StyleName = $Null
                    $ExistingStyles = @($WorkBook.Styles.NamedStyles | Select -ExpandProperty Name)
                    Switch -regex ($ThisType)
                    {
                        "double|decimal|single" 
                        {
                            $StyleName = 'decimals'
                            $StyleFormat = "0.00"
                        }
                        "int\d\d$"
                        {
                            $StyleName = 'ints'
                            $StyleFormat = "0"
                        }
                        "datetime"
                        {
                            $StyleName = "dates"
                            $StyleFormat = "M/d/yyy h:mm"
                        }
                        default
                        {
                            #No default yet...
                        }
                    }

                    if($StyleName)
                    {
                        if($ExistingStyles -notcontains $StyleName)
                        {
                            $StyleSheet = $WorkBook.Styles.CreateNamedStyle($StyleName)
                            $StyleSheet.Style.Numberformat.Format = $StyleFormat
                        }

                        $WorkSheet.Cells.Item($RowIndex, $ColumnIndex).Stylename = $StyleName
                    }

                }
                Write-Verbose "Wrote row $RowIndex"
                $RowIndex++
            }

            $Excel.SaveAs($Path)
    }
}