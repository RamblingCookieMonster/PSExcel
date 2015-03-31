function Format-Cell {
    <#
    .SYNOPSIS
        Format cells in an Excel worksheet

    .DESCRIPTION
        Format cells in an Excel worksheet

        Note:
            Each time you call this function, you need to save and re-create your Excel Object.
            If you attempt to modify the Excel object, save, modify, and save a second time, it will fail.
            See Save-Excel Passthru parameter for a workaround
        
    .PARAMETER Worksheet
        Worksheet to format cells on
    
    .PARAMETER StartRow
        The top row to format.  If not specified, we use the dimensions start row

    .PARAMETER StartColumn
        The leftmost column to format.  If not specified, we use the dimensions start column

    .PARAMETER EndRow
        The bottom row to format.  If not specified, we use the dimensions' end row

    .PARAMETER EndColumn
        The rightmost column to format.  If not specified, we use the dimensions' end column

    .PARAMETER Header
        If specified, identify and apply formatting to the header row only.

    .PARAMETER Bold
        Add or remove bold font (boolean)

    .PARAMETER Italic
        Add or remove Italic font (boolean)

    .PARAMETER Underline
        Add or remove Underline font (boolean)

    .PARAMETER Size
        Set font size

    .PARAMETER Font
        Set font name
        
    .PARAMETER Color
        Set color

    .PARAMETER WrapText
        Add or remove WrapText property (boolean)

    .PARAMETER AutoFit
        Apply auto fit to cells

    .PARAMETER AutoFitMinWidth
        Minimum width to set autofit with
    
    .PARAMETER AutoFitMaxWidth
        Maximum width to set autofit with

    .PARAMETER Passthru
        If specified, pass the Worksheet back

    .EXAMPLE
        #
        # Create an Excel object to work with
            $Excel = New-Excel -Path C:\Temp\Demo.xlsx
        
        #Get the worksheet, format the header as bold, size 14
            $Excel |
                Get-WorkSheet |
                Format-Cell -Header -Bold $True -Size 14
        
        #Save your changes, re-open the excel file
            $Excel = $Excel | Save-Excel -Passthru

        #Oops, too big!  Get the worksheet, format the header as size 11
            $Excel |
                Get-WorkSheet |
                Format-Cell -Header -Size 11

            $Excel | Save-Excel -Close

    .EXAMPLE
        $WorkSheet | Format-Cell -StartRow 2 -StartColumn 1 -EndColumn 1 -Italic $True -Size 10

        # Set the first column, rows 2 through the end to size 10, italic

    .EXAMPLE
          
        # Get the worksheet
        # format all the cells (default if nothing specified)
        # Set autofit between minumum of 5 and maximum of 20
        $Excel |
            Get-WorkSheet |
            Format-Cell -Autofit -AutofitMinWidth  -AutofitMaxWidth 20

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
    [cmdletbinding(DefaultParameterSetname = 'Range')]
    param(
        [parameter( Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        [OfficeOpenXml.ExcelWorksheet]$WorkSheet,

        [parameter( ParameterSetName = 'Range',
                    Mandatory=$false,
                    ValueFromPipeline=$false,
                    ValueFromPipelineByPropertyName=$false)]
        [int]$StartRow,
        
        [parameter( ParameterSetName = 'Range',
                    Mandatory=$false,
                    ValueFromPipeline=$false,
                    ValueFromPipelineByPropertyName=$false)]
        [int]$StartColumn,
        
        [parameter( ParameterSetName = 'Range',
                    Mandatory=$false,
                    ValueFromPipeline=$false,
                    ValueFromPipelineByPropertyName=$false)]
        [int]$EndRow,

        [parameter( ParameterSetName = 'Range',
                    Mandatory=$false,
                    ValueFromPipeline=$false,
                    ValueFromPipelineByPropertyName=$false)]
        [int]$EndColumn,

        [parameter( ParameterSetName = 'Header',
                    Mandatory=$true,
                    ValueFromPipeline=$false,
                    ValueFromPipelineByPropertyName=$false)]
        [Switch]$Header,

        [boolean]$Bold,
        [boolean]$Italic,
        [boolean]$Underline,
        [int]$Size,
        [string]$Font,
        [System.Drawing.Color]$Color,
        [boolean]$WrapText,
        [switch]$Autofit,
        [double]$AutofitMinWidth,
        [double]$AutofitMaxWidth,

        [switch]$Passthru
    )
    Begin
    {
        #From http://stackoverflow.com/questions/297213/translate-a-column-index-into-an-excel-column-name
        Function Get-ExcelColumn
        {
            param([int]$ColumnIndex)

            [string]$Chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'

            $ColumnIndex -= 1
            [int]$Quotient = [math]::floor($ColumnIndex / 26)

            if($Quotient -gt 0)
            {
                ( GetExcelColumn -ColumnIndex $Quotient ) + $Chars[$ColumnIndex % 26]
            }
            else
            {
                $Chars[$ColumnIndex % 26]
            }
        }
        Function Get-ExcelCoordinate
        {
            param($RowIndex, $ColumnIndex)
            $ColumnIndex = Get-ExcelColumn $ColumnIndex
            "$ColumnIndex$RowIndex"
        }
    }
    Process
    {
        #Get the coordinates
            $dimension = $WorkSheet.Dimension
        
            if($PSCmdlet.ParameterSetName -like 'Range')
            {
                If(-not $StartRow)
                {
                    $StartRow = $dimension.Start.Row
                }
                If(-not $StartColumn)
                {
                    $StartColumn = $dimension.Start.Column
                }
                If(-not $EndRow)
                {
                    $EndRow = $dimension.End.Row
                }
                If(-not $EndColumn)
                {
                    $EndColumn = $dimension.End.Column
                }
            }
            Elseif($PSCmdlet.ParameterSetName -like 'Header')
            {
                $StartRow = $dimension.Start.Row
                $StartColumn = $dimension.Start.Column
                $EndRow = $dimension.Start.Row
                $EndColumn = $dimension.End.Column
            }

            $Start = Get-ExcelCoordinate -RowIndex $StartRow -ColumnIndex $StartColumn
            $End = Get-ExcelCoordinate -RowIndex $EndRow -ColumnIndex $EndColumn
            $RangeCoordinates = "$Start`:$End"

        # Apply the formatting
            $CellRange = $WorkSheet.Cells[$RangeCoordinates]
            
            switch ($PSBoundParameters.Keys)
            {
                'Bold'       { $CellRange.Style.Font.Bold = $Bold  }
                'Italic'     { $CellRange.Style.Font.Italic = $Italic  }
                'Underline'  { $CellRange.Style.Font.UnderLine = $Underline}
                'Size'       { $CellRange.Style.Font.Size = $Size }
                'Font'       { $CellRange.Style.Font.Name = $Font }
                'Color'      { $CellRange.Style.Font.Color.SetColor($Color) }
                'WrapText'   { $CellRange.Style.WrapText = $WrapText  }
                'Autofit'    {
                    #Probably a cleaner way to call this...
                    try
                    {
                        if($PSBoundParameters.ContainsKey('AutofitMaxWidth'))
                        {
                            $CellRange.AutoFitColumns($AutofitMinWidth, $AutofitMaxWidth)
                        }
                        elseif($PSBoundParameters.ContainsKey('AutofitMinWidth'))
                        {
                            $CellRange.AutoFitColumns($AutofitMinWidth)
                        }
                        else
                        {
                            $CellRange.AutoFitColumns()
                        }
                    }
                    Catch
                    {
                        Write-Error $_
                    }
                }
            }
        if($Passthru)
        {
            $WorkSheet
        }
    }
}