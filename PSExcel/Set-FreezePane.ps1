function Set-FreezePane {
    <#
    .SYNOPSIS
        Set FreezePanes on a specified worksheet

    .DESCRIPTION
        Set FreezePanes on a specified worksheet
    
    .PARAMETER Worksheet
        Worksheet to add FreezePanes to
    
    .PARAMETER Row
        The first row with live data.

        Examples and outcomes:
            -Row 2        Freeze row 1
            -Row 5        Freeze rows 1 through 4

    .PARAMETER Column
        Examples and outcomes:
            -Column 2     Freeze column 1
            -Column 5     Freeze columns 1 through 4   

    .PARAMETER Passthru
        If specified, pass the Worksheet back

    .EXAMPLE
        $WorkSheet | Set-FreezePane

        #Freeze the top row of $Worksheet (default parameter values handle this)

    .EXAMPLE
        $WorkSheet | Set-FreezePane -Row 2 -Column 4

        # Freeze the top row and top 3 columns of $Worksheet

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
        [OfficeOpenXml.ExcelWorksheet]$WorkSheet,

        [int]$Row = 2,

        [int]$Column = 1,

        [switch]$Passthru
    
    )
    Process
    {
        $WorkSheet.View.FreezePanes($Row, $Column)
        if($Passthru)
        {
            $WorkSheet
        }
    }
}