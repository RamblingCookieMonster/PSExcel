function Set-FreezePane {
    <#
    .SYNOPSIS
        Set FreezePanes on a specified worksheet

    .DESCRIPTION
        Set FreezePanes on a specified worksheet

    .PARAMETER Worksheet
        A Worksheet to set FreezePanes on

    .PARAMETER Row
        First live row after the frozen pane

        Examples and outcomes:
            -Row 2      Freeze Row 1 only
            -Row 5      Freeze Rows 1 through 4

    .PARAMETER Column
        First live column after the frozen pane

        Examples and outcomes:
            -Column 2   Freeze Column 1 only
            -Column 5   Freeze Columns 1 through 4

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