function Remove-FreezePane {
    <#
    .SYNOPSIS
        Remove FreezePanes on a specified worksheet

    .DESCRIPTION
        Remove FreezePanes on a specified worksheet

    .EXAMPLE
        $WorkSheet | Remove-FreezePane

        # Remove frozen panes on $WorkSheet
        
    .NOTES
        Thanks to Doug Finke for his example:
            https://github.com/dfinke/ImportExcel/blob/master/ImportExcel.psm1

        Thanks to Philip Thompson for an expansive set of examples on working with EPPlus in PowerShell:
            https://excelpslib.codeplex.com/

    .FUNCTIONALITY
        Excel
    #>
    [cmdletbinding()]
    param(
        [parameter( Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        [OfficeOpenXml.ExcelWorksheet]$WorkSheet    
    )
    Process
    {
        $WorkSheet.View.UnFreezePanes()         
    }
}