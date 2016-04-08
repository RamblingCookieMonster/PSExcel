function Import-XLSX {
    <#
    .SYNOPSIS
        Import data from Excel

    .DESCRIPTION
        Import data from Excel

    .PARAMETER Path
        Path to an xlsx file to import

    .PARAMETER Sheet
        Index or name of Worksheet to import

    .PARAMETER Header
        Replacement headers.  Must match order and count of your data's properties.

    .PARAMETER RowStart
        First row to start reading from, typically the header. Default is 1

    .PARAMETER ColumnStart
        First column to start reading from. Default is 1

    .PARAMETER FirstRowIsData
        Indicates that the first row is data, not headers.  Must be used with -Header.

    .PARAMETER Text
        Extract cell text, rather than value.

        For example, if you have a cell with value 5:
            If the Number Format is '0', the text would be 5
            If the Number Format is 0.00, the text would be 5.00 

    .EXAMPLE
        Import-XLSX -Path "C:\Excel.xlsx"

        #Import data from C:\Excel.xlsx

    .EXAMPLE
        Import-XLSX -Path "C:\Excel.xlsx" -Header One, Two, Five

        # Import data from C:\Excel.xlsx
        # Replace headers with One, Two, Five

    .EXAMPLE
        Import-XLSX -Path "C:\Excel.xlsx" -Header One, Two, Five -FirstRowIsData -Sheet 2

        # Import data from C:\Excel.xlsx
        # Assume first row is data
        # Use headers One, Two, Five
        # Pull from sheet 2 (sheet 1 is default)

    .EXAMPLE
       #    A        B        C 
       # 1  Random text to mess with you!
       # 2  Header1  Header2  Header3
       # 3  data1    Data2    Data3

       # Your worksheet has data you don't care about in the first row or column
       # Use the ColumnStart or RowStart parameters to solve this.

       Import-XLSX -Path C:\RandomTextInRow1.xlsx -RowStart 2

    .NOTES
        Thanks to Doug Finke for his example:
            https://github.com/dfinke/ImportExcel/blob/master/ImportExcel.psm1

        Thanks to Philip Thompson for an expansive set of examples on working with EPPlus in PowerShell:
            https://excelpslib.codeplex.com/

    .LINK
        https://github.com/RamblingCookieMonster/PSExcel

    .FUNCTIONALITY
        Excel
    #>
    [cmdletbinding()]
    param(
        [parameter( Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        [validatescript({Test-Path $_})]
        [string[]]$Path,

        $Sheet = 1,

        [string[]]$Header,

        [switch]$FirstRowIsData,

        [switch]$Text,

        [int]$RowStart = 1,

        [int]$ColumnStart = 1
    )
    Process
    {
        foreach($file in $path)
        {
            #Resolve relative paths... Thanks Oisin! http://stackoverflow.com/a/3040982/3067642
            $file = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($file)

            write-verbose "target excel file $($file)"

            Try
            {
                $xl = New-Object OfficeOpenXml.ExcelPackage $file
                $workbook  = $xl.Workbook
            }
            Catch
            {
                Write-Error "Failed to open '$file':`n$_"
                continue
            }

            Try
            {
                if( @($workbook.Worksheets).count -eq 0)
                {
                    Throw "No worksheets found"
                }
                else
                {
                    $worksheet = $workbook.Worksheets[$Sheet]
                    $dimension = $worksheet.Dimension

                    $Rows = $dimension.Rows
                    $Columns = $dimension.Columns

                    $ColumnEnd = $Columns + $ColumnStart - 1
                    $RowEnd = $Rows + $RowStart - 1
                }

            }
            Catch
            {
                Write-Error "Failed to gather Worksheet '$Sheet' data for file '$file':`n$_"
                continue
            }
  
            if($Header -and $Header.count -gt 0)
            {
                if($Header.count -ne $Columns)
                {
                    Write-Error "Found '$columns' columns, provided $($header.count) headers.  You must provide a header for every column."
                }
                Write-Verbose "User defined headers: $Header"
            }
            else 
            {
                $Header = @( foreach ($Column in $ColumnStart..$ColumnEnd)
                {
                    if($Text)
                    {
                        $PotentialHeader = $worksheet.Cells.Item($RowStart,$Column).Text
                    }
                    else
                    {
                        $PotentialHeader = $worksheet.Cells.Item($RowStart,$Column).Value
                    }

                    if( -Not $PotentialHeader -Or $PotentialHeader.Trim().Equals("") )
                    {
                        Write-Warning "Header in column $Column is whitespace or empty, setting header to '<Column $Column>'"
                        $PotentialHeader = "<Column $Column>" # Use placeholder name
                    }
                    $PotentialHeader
                })
            }

            [string[]]$SelectedHeaders = @( $Header | select -Unique )
            Write-Verbose "Found $Rows rows, $Columns columns, with headers:`n$($Header | Out-String)"

            if(-not $FirstRowIsData)
            {
                $RowStart++
            }

            foreach ($Row in $RowStart..$RowEnd)
            {
                $RowData = @{}

                foreach ($Column in 0..($Columns - 1))
                {
                    $Name  = $Header[$Column]
                    if($Text)
                    {
                        $Value = $worksheet.Cells.Item($Row, ($Column + $ColumnStart)).Text
                    }
                    else
                    {
                        $Value = $worksheet.Cells.Item($Row, ($Column + $ColumnStart)).Value
                    }

                    Write-Debug "Row: $Row, Column: $Column, Name: $Name, Value = $Value"

                    #Handle dates, they're too common to overlook... Could use help, not sure if this is the best regex to use?
                    $Format = $worksheet.Cells.Item($Row, ($Column + $ColumnStart)).style.numberformat.format
                    if($Format -match '\w{1,4}/\w{1,2}/\w{1,4}( \w{1,2}:\w{1,2})?')
                    {
                        Try
                        {
                            $Value = [datetime]::FromOADate($Value)
                        }
                        Catch
                        {
                            Write-Verbose "Error converting '$Value' to datetime"
                        }
                    }

                    if($RowData.ContainsKey($Name) )
                    {
                        Write-Warning "Duplicate header for '$Name' found, with value '$Value', in row $Row"
                    }
                    else
                    {
                        $RowData.Add($Name, $Value)
                    }
                }
                New-Object -TypeName PSObject -Property $RowData | Select -Property $SelectedHeaders
            }

            $xl.Dispose()
            $xl = $null
        }
    }
}