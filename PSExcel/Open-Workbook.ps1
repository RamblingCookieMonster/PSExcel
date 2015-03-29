function Open-Workbook {
    <#
    .SYNOPSIS
        Open an ExcelPackage Workbook

    .DESCRIPTION
        Open an ExcelPackage Workbook

    .PARAMETER Excel
        Path to an ExcelPackage

    .EXAMPLE
        $Excel = New-Excel -Path "C:\Excel.xlsx"
        $WorkBook = Open-Workbook $Excel
        $WorkBook

        #Open C:\Excel.xlsx, view the workbook

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
        [OfficeOpenXml.ExcelPackage]$Excel
    )
    Process
    {
        $Excel.WorkBook
    }
}