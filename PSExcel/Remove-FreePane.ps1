function Remove-FreezePane {
    <#
    .SYNOPSIS
        Remove FreezePanes on a specified worksheet
    
    .DESCRIPTION
        Remove FreezePanes on a specified worksheet

    .PARAMETER Worksheet
        Worksheet to remove FreezePanes from

    .PARAMETER Passthru
        If specified, pass the Worksheet back

    .EXAMPLE
        $WorkSheet | Remove-FreezePane

        # Remove frozen panes on $WorkSheet
        
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
    [OutputType([OfficeOpenXml.ExcelWorksheet])]
    [cmdletbinding()]
    param(
        [parameter( Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        [OfficeOpenXml.ExcelWorksheet]$WorkSheet,

        [switch]$Passthru
    )
    Process
    {
        $WorkSheet.View.UnFreezePanes()
        if($Passthru)
        {
            $WorkSheet
        }        
    }
}