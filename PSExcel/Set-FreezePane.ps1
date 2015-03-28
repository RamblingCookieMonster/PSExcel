function Set-FreezePane {
    <#
    .SYNOPSIS
        Set FreezePanes on a specified worksheet

    .DESCRIPTION
        Set FreezePanes on a specified worksheet

    .EXAMPLE
        $WorkSheet | Set-FreezePane

        #Freeze the top row of $Worksheet

    .EXAMPLE
        $WorkSheet | Set-FreezePane -Row 2 -Column 4

        # Freeze the top row and top 3 columns of $Worksheet

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
        [OfficeOpenXml.ExcelWorksheet]$WorkSheet,

        [int]$Row = 2,

        [int]$Column = 1
    
    )
    Process
    {
        $WorkSheet.View.FreezePanes($Row, $Column)
    }
}