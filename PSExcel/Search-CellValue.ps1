function Search-CellValue {
    <#
    .SYNOPSIS
        Find a value in a spreadsheet

    .DESCRIPTION
        Find a value in a spreadsheet

        Specify an xlsx path, an ExcelPackage object, or a WorkSheet to search, and a ScriptBlock you want to run the cell values against

    .PARAMETER Path
        Path to an xlsx file to search

    .PARAMETER Excel
        An ExcelPackage to search

    .PARAMETER WorkSheet
        An Excel WorkSheet to search

    .PARAMETER WorksheetName
        Optional name of Worksheet to search

    .PARAMETER FilterScript
        Scriptblock that we call with Where-Object against every cell value

        $_ would refer to the value of the cell.

    .PARAMETER As
        How the data should be returned
        
            Default   WorksheetName, Row, Column, Match data
            Raw       Return the value only
            Passthru  Return the cell item

    .EXAMPLE
        Search-CellValue -Path C:\Temp\Demo.xlsx {$_ -like 'jsmith*'}

        #Find any cell values like jsmith*

    .EXAMPLE
        Search-CellValue = Get-ExcelValue -Path C:\Temp\Demo.xlsx {$_ -like 'jsmith*'} -As Passthru

        #Returns an OfficeOpenXml.ExcelRangeBase that you can process (e.g. add formatting to manually)

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
    [cmdletbinding(DefaultParameterSetName = 'Excel')]
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

        [parameter( Mandatory = $True,
                    Position = 0)]
        [scriptblock]$FilterScript,
       
        $WorkSheetName,

        [validateset('Default','Raw','Passthru')]
        [string]$As = 'Default'
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

        Foreach($Sheet in $WorkSheets)
        {
            $Dimension = $Sheet.Dimension
            if($IgnoreHeader)
            {
                $RowStart = 2
            }
            else
            {
                $RowStart = $Dimension.Start.Row
            }
            $ColumnStart = $Dimension.Start.Column
            $RowEnd = $Dimension.End.Row
            $ColumnEnd = $Dimension.End.Column
            
            Write-Verbose "Searching $($Sheet.Name) over coordinates $RowStart, $ColumnStart through $RowEnd, $ColumnEnd"

            for ($Row = $RowStart; $Row -le $RowEnd; $Row++)
            {
                for ($Column = $ColumnStart; $Column -le $ColumnEnd; $Column++)
                {
                    $Value = $Sheet.Cells.Item($Row, $Column).Value
                    
                    #Handle dates, they're too common to overlook... Could use help, not sure if this is the best regex to use?
                    $Format = $Sheet.Cells.Item($Row, $Column).style.numberformat.format
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

                    $Data = $Null
                    $Data = $Value | Where-Object -FilterScript $FilterScript
                    if($Data -ne $Null)
                    {
                        Switch ($As)
                        {
                            'Raw'
                            {
                                $Data
                            }
                            'Default'
                            {
                                New-Object -TypeName PSObject -Property @{
                                    WorksheetName = $Sheet.Name
                                    Row = $Row
                                    Column = $Column
                                    Match = $Data
                                } | Select WorkSheetName, Row, Column, Match
                            }
                            'Passthru'
                            {
                                $Sheet.Cells.Item($Row, $Column)
                            }
                        }
                    }
                }
            }

            if($PSBoundParameters.ContainsKey('Path'))
            {
                $Sheet.Dispose()
            }
        }
    }
}