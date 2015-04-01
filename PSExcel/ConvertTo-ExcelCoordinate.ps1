Function ConvertTo-ExcelCoordinate
{
    <#
    .SYNOPSIS
        Convert a row and column to an Excel coordinate

    .DESCRIPTION
        Convert a row and column to an Excel coordinate

    .PARAMETER Row
        Row number

    .PARAMETER Column
        Column number

    .EXAMPLE
        ConvertTo-ExcelCoordinate -Row 1 -Column 2

        #Get Excel coordinates for Row 1, Column 2.  B1.

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
    [OutputType([system.string])]
    [cmdletbinding()]
    param(
        [int]$Row,
        [int]$Column
    )

        #From http://stackoverflow.com/questions/297213/translate-a-column-index-into-an-excel-column-name
        Function Get-ExcelColumn
        {
            param([int]$ColumnIndex)

            [string]$Chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

            $ColumnIndex -= 1
            [int]$Quotient = [math]::floor($ColumnIndex / 26)

            if($Quotient -gt 0)
            {
                ( Get-ExcelColumn -ColumnIndex $Quotient ) + $Chars[$ColumnIndex % 26]
            }
            else
            {
                $Chars[$ColumnIndex % 26]
            }
        }

    $ColumnIndex = Get-ExcelColumn $Column
    "$ColumnIndex$Row"
}