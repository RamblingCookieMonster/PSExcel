function Get-Workbook {
    <#
    .SYNOPSIS
        Return a Workbook from an ExcelPackage

    .DESCRIPTION
        Return a Workbook from an ExcelPackage

    .PARAMETER Excel
        ExcelPackage to extract workbook from
    
    .EXAMPLE
        $Excel = New-Excel -Path "C:\Excel.xlsx"
        $WorkBook = Get-Workbook $Excel
        $WorkBook

        #Open C:\Excel.xlsx, view the workbook

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
    [OutputType([OfficeOpenXml.ExcelWorkbook])]
    [cmdletbinding()]
    param(
        [parameter( Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$false)]
        [OfficeOpenXml.ExcelPackage]$Excel
    )
    Process
    {
        $Excel.WorkBook
    }
}