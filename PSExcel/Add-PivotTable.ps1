function Add-PivotTable {
    <#
    .SYNOPSIS
        Add a pivot table to an Excel worksheet

    .DESCRIPTION
        Add a pivot table to an Excel worksheet

        Note:
            Each time you call this function, you need to save and re-create your Excel Object.
            If you attempt to modify the Excel object, save, modify, and save a second time, it will fail.
            See Save-Excel Passthru parameter for a workaround
 
    .PARAMETER Path
        Path to an xlsx file to add the pivot table to

        If Path is specified and you do not use passthru, we save the file

    .PARAMETER Excel
        ExcelPackage to add the pivot table to

        We do not save the ExcelPackage upon completion.  See Save-Excel.
         
    .PARAMETER WorksheetName
        If specified, use this worksheet as the source.
            
    .PARAMETER StartRow
        The top row for pivottable data.  If not specified, we use the dimensions start row

    .PARAMETER StartColumn
        The leftmost column for pivottable data.  If not specified, we use the dimensions start column

    .PARAMETER EndRow
        The bottom row for pivottable data.  If not specified, we use the dimensions' end row

    .PARAMETER EndColumn
        The rightmost column for pivottable data.  If not specified, we use the dimensions' end column

    .PARAMETER PivotTableWorksheetName
        Name for the WorkSheet we create for the pivottable

    .PARAMETER PivotData
        Pivot data
        
    .PARAMETER PivotRows
        Pivot on these rows

    .PARAMETER PivotColumns
        Pivot on these columns

    .PARAMETER ChartType
        If specified, add a chart with this type

    .PARAMETER Passthru
        If specified, pass the ExcelPackage back

    .EXAMPLE
    
        Get-ChildItem C:\ -file | Export-XLSX -Path C:\temp\files.xlsx
        
        Add-PivotTable -Path C:\Temp\files.xlsx -PivotRows Extension -PivotValues Length -ChartType PieExploded3D

        # Get files, create an xlsx in C:\temp\ps.xlsx
        
        # Take existing xlsx and add a pivot chart
            # Pivot rows on 'Extension'
            # Pivot values on 'Length'
            # Add an exploding pie chart!

        #This example gives you a pie chart breaking down storage by file extension

    .EXAMPLE

    .NOTES
        Thanks to Doug Finke for his example
        This function borrows heavily if not everything from Doug:
            https://github.com/dfinke/ImportExcel

        Thanks to Philip Thompson for an expansive set of examples on working with EPPlus in PowerShell:
            https://excelpslib.codeplex.com/

    .LINK
        https://github.com/RamblingCookieMonster/PSExcel

    .FUNCTIONALITY
        Excel
    #>
    [OutputType([OfficeOpenXml.ExcelPackage])]
    [cmdletbinding(DefaultParameterSetName = 'Excel')]
    param(
        [parameter( Position = 0,
                    ParameterSetName = 'File',
                    Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        [validatescript({Test-Path $_})]
        [string]$Path,

        [parameter( Position = 0,
                    ParameterSetName = 'Excel',
                    Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        [OfficeOpenXml.ExcelPackage]$Excel,

        [parameter( Position = 1,
                    Mandatory=$false,
                    ValueFromPipeline=$false,
                    ValueFromPipelineByPropertyName=$false)]
        [string]$WorkSheetName,

        [parameter( Position = 2,
                    Mandatory=$false,
                    ValueFromPipeline=$false,
                    ValueFromPipelineByPropertyName=$false)]
        [string]$PivotTableWorksheetName = 'PivotTable1',

        [parameter(
                    Mandatory=$false,
                    ValueFromPipeline=$false,
                    ValueFromPipelineByPropertyName=$false)]
        [int]$StartRow,
        
        [parameter(
                    Mandatory=$false,
                    ValueFromPipeline=$false,
                    ValueFromPipelineByPropertyName=$false)]
        [int]$StartColumn,
        
        [parameter(
                    Mandatory=$false,
                    ValueFromPipeline=$false,
                    ValueFromPipelineByPropertyName=$false)]
        [int]$EndRow,

        [parameter(
                    Mandatory=$false,
                    ValueFromPipeline=$false,
                    ValueFromPipelineByPropertyName=$false)]
        [int]$EndColumn,

        [string[]]$PivotRows,
        [string[]]$PivotColumns,
        [string[]]$PivotValues,

        [OfficeOpenXml.Drawing.Chart.eChartType]$ChartType,

        [switch]$Passthru
    )
    Process
    {

        Write-Verbose "PSBoundParameters: $($PSBoundParameters | Out-String)"    
        $SourceWS = @{}
        if($PSBoundParameters.ContainsKey( 'WorkSheetName') )
        {
            $SourceWS.Add('Name',$WorkSheetName)
        }

        Try
        {
            if($PSCmdlet.ParameterSetName -like 'File')
            {
                $Excel = New-Excel -Path $Path -ErrorAction Stop
            }

            $WorkSheets = @( $Excel | Get-Worksheet @SourceWS -ErrorAction Stop )
        }
        Catch
        {
            Throw "Could not get worksheets to search: $_"
        }

        If($WorkSheets.Count -eq 0)
        {
            Throw "Something went wrong, we didn't find a worksheet"
        }

        Foreach($SourceWorkSheet in $WorkSheets)
        {
            if($WorkSheets.count -gt 1)
            {
                $PivotTableWorksheetName = "$PivotTableWorksheetName-$($SourceWorkSheet.Name)"
            }

            Try
            {
                $PivotWorkSheet = $Excel.Workbook.Worksheets.Add($PivotTableWorksheetName)
            }
            Catch
            {
                Throw "Could not add PivotTable: $_"
            }
            
            #Get the coordinates
                $dimension = $SourceWorkSheet.Dimension
            
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

                $Start = ConvertTo-ExcelCoordinate -Row $StartRow -Column $StartColumn
                $End = ConvertTo-ExcelCoordinate -Row $EndRow -Column $EndColumn
                $RangeCoordinates = "$Start`:$End"
                
            #Pivot! Borrowed from Doug Finke - thanks Doug!
                $PivotWorkSheet.View.TabSelected = $True
                $PivotTable = $PivotWorkSheet.PivotTables.Add($PivotWorkSheet.Cells["A1"], $SourceWorkSheet.Cells[$RangeCoordinates], "PivotTable1")
            
                if($PivotRows)
                {
                    foreach ($Row in @($PivotRows | Select -Unique))
                    {
                        [void]$PivotTable.RowFields.Add($PivotTable.Fields[$Row])
                    }
                }

                if($PivotColumns)
                {
                    foreach ($Column in @($PivotColumns | Select -Unique))
                    {
                        [void]$PivotTable.ColumnFields.Add($PivotTable.Fields[$Column])
                    }
                }

                if($PivotValues)
                {
                    foreach ($Value in @($PivotValues | Select -Unique))
                    {
                        [void]$PivotTable.DataFields.Add($PivotTable.Fields[$Value])
                    }
                }

                if($ChartType)
                {
                    Write-Verbose "Adding $ChartType chart"
                    $chart = $PivotWorkSheet.Drawings.AddChart('PivotChart1', $ChartType, $PivotTable)
                    $chart.SetPosition(1, 0, 6, 0)
                    $chart.SetSize(600, 400)
                }

            if($PSCmdlet.ParameterSetName -like 'File' -and -not $Passthru)
            {
                $Excel.save()
            }
            if($Passthru)
            {
                $Excel
            }
        }
    }
}