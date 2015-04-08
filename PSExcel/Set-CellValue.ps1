function Set-CellValue {
    <#
    .SYNOPSIS
        Set the value of a specific cell or range

    .DESCRIPTION
        Set the value of a specific cell or range

        NOTE:
            Each time you call this function, you need to save and re-create your Excel Object.
            If you attempt to modify the Excel object, save, modify, and save a second time, it will fail.
            See Save-Excel Passthru parameter for a workaround
    
    .PARAMETER CellRange
        CellRange to set value on.  This is an ExcelRangeBase

        See help for Search-CellValue, with the '-As Passthru' parameter.  This generates an ExcelRangeBase

    .PARAMETER Excel
        An Excel object to set values in. We do not save this.

    .PARAMETER Path
        A path to set values in. We save changes to this.

    .PARAMETER Worksheet
        A worksheet to set values in.
    
    .PARAMETER WorksheetName
        A specific worksheet to set values in, otherwise, assume all worksheets from the input object

    .PARAMETER Coordinates
        Excel style coordinates specifying starting cell and final cell (e.g. A1:B2)

        If not specified, we get the dimension for the worksheet and change everything.

    .PARAMETER Value
        The value to set cells to.

    .PARAMETER Passthru
        If specified, passthru the inputobject (Excel, Worksheet, or Cellrange)

    .EXAMPLE
        Set-CellValue -Path C:\Temp\Demo.xlsx -Coordinates a1:a1 -Value Header1

        #Set the first column header to 'Header1'

    .EXAMPLE
        
        #
        # Open an existing XLSX to search and set cells within
        $Excel = New-Excel -Path C:\Temp\Demo.xlsx

        #Search for any cells like 'jsmith*'.  Set them all to REDACTED
        $Excel | Search-CellValue {$_ -like 'jsmith*'} -As Passthru | Set-CellValue -Value "REDACTED"

        #Save your changes and close the ExcelPackage
        $Excel | Save-Excel -Close

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
    [cmdletbinding(DefaultParametersetName = 'CellRange')]
    param(
        [parameter( Position = 0,
                    ParameterSetName = 'CellRange',
                    Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        [OfficeOpenXml.ExcelRangeBase]$CellRange,

        [parameter( Position = 0,
                    ParameterSetName = 'Excel',
                    Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        [OfficeOpenXml.ExcelPackage]$Excel,

        [parameter( Position = 0,
                    ParameterSetName = 'File',
                    Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        [validatescript({Test-Path $_})]
        [string]$Path,

        [parameter( Position = 0,
                    ParameterSetName = 'Worksheet',
                    Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        [OfficeOpenXml.ExcelWorksheet]$WorkSheet,

        [parameter(ParametersetName = 'Excel')]
        [parameter(ParametersetName = 'File')]
        [parameter(ParametersetName = 'Worksheet')]        
        [string]$WorksheetName,

        [parameter(ParametersetName = 'Excel')]
        [parameter(ParametersetName = 'File')]
        [parameter(ParametersetName = 'Worksheet')]
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

        $Value,

        [parameter(ParametersetName = 'Excel')]
        [parameter(ParametersetName = 'CellRange')]
        [parameter(ParametersetName = 'Worksheet')]
        [switch]$Passthru

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
                    $Excel = New-Excel -Path $Path -ErrorAction Stop
                    $WorkSheets = @( $Excel | Get-Worksheet @WSParam -ErrorAction Stop )
                }
                'Worksheet'
                {
                    $WorkSheets = @( $WorkSheet )
                }
                'CellRange'
                {
                    $WorkSheets = @( $CellRange.Worksheet | Select -First 1 )
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


        foreach($Worksheet in $Worksheets)
        {
            if($PSCmdlet.ParameterSetName -notlike 'CellRange')
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
                }
            }

            $CellRange.Value = $Value
            $StyleName = $null
            $StyleFormat = $null
            Try
            {
                #Nulls will error, catch them
                $ThisType = $Null
                $ThisType = $Value.GetType().FullName
            }
            Catch
            {
                Write-Verbose "Applying no style to null in range $($CellRange.FullAddress)"
            }

            Switch -regex ($ThisType)
            {
                "double|decimal|single"
                {
                    $StyleFormat = "0.00"
                }
                "int\d\d$"
                {
                    $StyleFormat = "0"
                }
                "datetime"
                {
                    $StyleFormat = "M/d/yyy h:mm"
                }
                default
                {
                    #No default yet...
                }
            }

            if($StyleFormat)
            {
                $CellRange.Style.Numberformat.Format = $StyleFormat
            }
        }

        switch($PSCmdlet.ParameterSetName)
        {
            'File'
            {
                $Excel.save()
            }
            'Excel'
            {
                if($Passthru) {$Excel}
            }
            'Worksheet'
            {
                if($Passthru) {$Worksheet}
            }
            'CellRange'
            {
                if($Passthru) {$CellRange}
            }
        }
    }
}