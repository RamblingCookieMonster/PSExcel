function Get-CellValue {
    <#
    .SYNOPSIS
        Get cell data from Excel

    .DESCRIPTION
        Get cell data from Excel

    .PARAMETER Path
        Path to an xlsx file to get cells from

    .PARAMETER Excel
        An ExcelPackage to get cells from

    .PARAMETER WorkSheet
        An Excel WorkSheet to get cells from

    .PARAMETER WorksheetName
        Optional name of Worksheet to get cells from

    .PARAMETER Header
        Replacement headers.  Must match order and count of your data's columns

    .PARAMETER Coordinates
        Excel style coordinates specifying starting cell and final cell (e.g. A1:B2)

        If not specified, we get the dimension for the worksheet and return everything
            
    .EXAMPLE
        Get-CellValue -Path C:\temp\Demo.xlsx -Coordinates A2:A2

        #Get the value at column 1, row 2

    .EXAMPLE
        Get-CellValue -Path C:\temp\Demo.xlsx -Coordinates A2:B3 -Header One, Two

        #Get the values from cells in column one, row two through column two, row three.  Replace headers with One, Two

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
        [parameter( Position = 1,
                    ParameterSetName = 'Excel',
                    Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        [OfficeOpenXml.ExcelPackage]$Excel,

        [parameter( Position = 1,
                    ParameterSetName = 'File',
                    Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        [validatescript({Test-Path $_})]
        [string]$Path,

        [parameter( Position = 1,
                    ParameterSetName = 'Worksheet',
                    Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        [OfficeOpenXml.ExcelWorksheet]$WorkSheet,

        [validatescript({
            if( $_ -match "^[a-zA-Z]+[0-9]+:[a-zA-Z]+[0-9]+$" )
            {
                $True
            }
            else
            {
                Throw "'$_' is not a valid coordinate.  See help for 'Coordinates' parameter"
            }
        
        })]
        [string]$Coordinates,
       
        $WorkSheetName,

        [string[]]$Header

    )
    Process
    {
        Write-Verbose "PSBoundParameters: $($PSBoundParameters | Out-String)"    
        $WSParam = @{}
        if($PSBoundParameters.ContainsKey( 'WorkSheetName') )
        {
            $WSParam.Add('Name',$WorkSheetName)
        }
        Try
        {
            switch ($PSCmdlet.ParameterSetName)
            {
                'Excel'
                {
                    $WorkSheets = @( $Excel | Get-Worksheet @WSParam -ErrorAction Stop )
                }
                'File'
                {
                    $WorkSheets = @( New-Excel -Path $Path -ErrorAction Stop | Get-Worksheet @WSParam -ErrorAction Stop )
                }
                'Worksheet'
                {
                    $WorkSheets = @( $WorkSheet )
                }
            }
        }
        Catch
        {
            Throw "Could not get worksheets to search: $_"
        }

        If($WorkSheets.Count -eq 0)
        {
            Throw "Something went wrong, we didn't find a worksheet"
        }

        Foreach($Worksheet in $WorkSheets)
        {
            Write-Verbose "Working with worksheet $($Worksheet.Name)"
            if($PSBoundParameters.ContainsKey('Coordinates'))
            {
                Try
                {
                    $CellRange = $WorkSheet.Cells.item($Coordinates)
                }
                Catch
                {
                    Write-Error "Could not get cells from '$($WorkSheet.Name)' for coordinates '$Coordinates'"
                    Continue
                }
            }
            else
            {
                $CellRange = $Worksheet.Cells
                $Coordinates = $WorkSheet.Dimension.Address
            }


            $ColumnStart = ($($Coordinates -split ":")[0] -replace "[0-9]", "").ToUpperInvariant()
            $ColumnEnd = ($($Coordinates -split ":")[1] -replace "[0-9]", "").ToUpperInvariant()
            [int]$RowStart = $($Coordinates -split ":")[0] -replace "[a-zA-Z]", ""
            [int]$RowEnd = $($Coordinates -split ":")[1] -replace "[a-zA-Z]", ""
            
            Function Get-ExcelColumnInt 
            {   # http://stackoverflow.com/questions/667802/what-is-the-algorithm-to-convert-an-excel-column-letter-into-its-number
                [cmdletbinding()]
                param($ColumnName)
                [int]$Sum = 0
                for ($i = 0; $i -lt $ColumnName.Length; $i++)
                { 
                    $sum *= 26
                    $sum += ($ColumnName[$i] - 65 + 1)
                }
                $sum
                Write-Verbose "Translated $ColumnName to $Sum"
            }

            $ColumnStart = Get-ExcelColumnInt $ColumnStart
            $ColumnEnd = Get-ExcelColumnInt $ColumnEnd
            $Columns = $ColumnEnd - $ColumnStart + 1

            if($Header -and $Header.count -gt 0)
            {
                if($Header.count -ne $Columns)
                {
                    Write-Error "Found '$columns' columns, provided $($header.count) headers.  You must provide a header for every column."
                }
            }
            else
            {
                $Header = @( foreach ($Column in $ColumnStart..$ColumnEnd)
                {
                    $worksheet.Cells.Item(1,$Column).Value
                } )
            }

            [string[]]$SelectedHeaders = @( $Header | select -Unique )

            Write-Verbose "Found headers $Header"


            #Skip headers...
            if($RowStart -eq 1 -and $RowEnd -ne 1)
            {
                $RowStart += 1
            }
            foreach($Row in ($RowStart)..$RowEnd)
            {
                $RowData = @{}
                $HeaderCol = 0

                foreach($Column in $ColumnStart..$ColumnEnd)
                {
                    $Name  = $Header[$HeaderCol]
                    $Value = $WorkSheet.Cells.Item($Row,$Column).Value
                    $HeaderCol++

                    Write-Debug "Row: $Row, Column: $Column, HeaderCol: $HeaderCol, Name: $Name, Value = $Value"
                                   
                    #Handle dates, they're too common to overlook... Could use help, not sure if this is the best regex to use?
                    $Format = $WorkSheet.Cells.Item($Row,$Column).style.numberformat.format
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
        }
    }
}