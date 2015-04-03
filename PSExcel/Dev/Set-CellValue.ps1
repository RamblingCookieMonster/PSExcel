function Set-CellValue {
    <#
    .SYNOPSIS
        Set the value of a specific cell or range

    .DESCRIPTION
        Set the value of a specific cell or range

        BETA NOTE:
            This is not a fully fledged function yet.
            Ideally we should allow the specification of a Worksheet (or lower level), and the string based range to set.

        NOTE:
            Each time you call this function, you need to save and re-create your Excel Object.
            If you attempt to modify the Excel object, save, modify, and save a second time, it will fail.
            See Save-Excel Passthru parameter for a workaround
    
    .PARAMETER CellRange
        CellRange to set value on.  This is an ExcelRangeBase

        See help for Search-CellValue, with the '-As Passthru' parameter.  This generates an ExcelRangeBase

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
    [cmdletbinding()]
    param(
        [parameter( Position = 0,
                    ParameterSetName = 'CellRange',
                    Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        [OfficeOpenXml.ExcelRangeBase]$CellRange,

        $Value
    )
    
    Process
    {
        Write-Verbose "PSBoundParameters: $($PSBoundParameters | Out-String)"    

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
}