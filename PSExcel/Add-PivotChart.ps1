function Add-PivotChart {
    <#
    .SYNOPSIS
        Add a pivot chart to an Excel worksheet

    .DESCRIPTION
        Add a pivot chart to an Excel worksheet

        Note:
            Each time you call this function, you need to save and re-create your Excel Object.
            If you attempt to modify the Excel object, save, modify, and save a second time, it will fail.
            See Save-Excel Passthru parameter for a workaround
 
    .PARAMETER Path
        Path to an xlsx file to add the pivot chart to

        If Path is specified and you do not use passthru, we save the file

    .PARAMETER Excel
        ExcelPackage to add the pivot chart to

        We do not save the ExcelPackage upon completion.  See Save-Excel.
         
    .PARAMETER PivotTableName
        Pivot table for chart data. If not specified, we add a chart to all pivot tables

    .PARAMETER TargetWorkSheetName
        Optional target worksheet for the chart.  If not specified, we use the existing pivot table worksheet

    .PARAMETER ChartName
        Optional, use this to ensure chart names are unique.  Defaults to CT-<PivotTableName>

    .PARAMETER ChartType
        If specified, add a chart with this type

    .PARAMETER ChartTitle
        Optional chart title

    .PARAMETER ChartWidth
        Width of the chart

    .PARAMETER ChartHeight
        Height of the chart

    .PARAMETER Passthru
        If specified, pass the ExcelPackage back

    .EXAMPLE
    
        Get-ChildItem C:\ -file | Export-XLSX -Path C:\temp\files.xlsx -PivotRows Extension -PivotValues Length
        
        Add-PivotChart -Path C:\Temp\files.xlsx -ChartType Pie -ChartName CT1
        Add-PivotChart -Path C:\Temp\files.xlsx -ChartType Area3D -ChartName CT2

        # Get files, create an xlsx in C:\temp\ps.xlsx
            # Pivot rows on 'Extension'
            # Pivot values on 'Length'

        # Take the xlsx and add a pie pivot chart 
        # Take the xlsx and add an Area3D pivot chart
        
        #This example gives you a pie chart breaking down storage by file extension

    .EXAMPLE

        #Create an xlsx and pivot table
            Get-ChildItem C:\ -file | Export-XLSX -Path C:\temp\files.xlsx -PivotRows Extension -PivotValues Length

        # Open the excel file, add a pivot chart (this won't save), add another pivot chart (this won't save), save.
            New-Excel -Path C:\temp\files.xlsx |
                Add-PivotChart -ChartType Pie -ChartTitle "Space per Extension" -ChartWidth 800 -ChartHeight 600 -Passthru |
                Add-PivotChart -ChartType PieExploded3D -ChartTitle "Why Do I Want This?" -ChartName CT2 -Passthru |
                Save-Excel -Close

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
                    ParameterSetName = 'Excel',
                    Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$false)]
        [OfficeOpenXml.ExcelPackage]$Excel,

        [parameter( Position = 0,
                    ParameterSetName = 'File',
                    Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$false)]
        [validatescript({Test-Path $_})]
        [string]$Path,

        [parameter( Position = 1,
                    Mandatory=$false,
                    ValueFromPipeline=$false,
                    ValueFromPipelineByPropertyName=$false)]
        [string]$PivotTableName = '*',

        [parameter( Position = 2,
                    Mandatory=$false,
                    ValueFromPipeline=$false,
                    ValueFromPipelineByPropertyName=$false)]
        [string]$TargetWorkSheetName,

        [string]$ChartName = 'Chart1',

        [parameter( Mandatory=$true,
                    ValueFromPipeline=$false,
                    ValueFromPipelineByPropertyName=$false)]
        [OfficeOpenXml.Drawing.Chart.eChartType]$ChartType,

        [string]$ChartTitle,

        [int]$ChartWidth = 600,

        [int]$ChartHeight = 400,

        [switch]$Passthru
    )
    Process
    {

        Write-Verbose "PSBoundParameters: $($PSBoundParameters | Out-String)"    
        $SourceWS = @{}

        #Find sheets with pivot tables
            Try
            {
                if($PSCmdlet.ParameterSetName -like 'File')
                {
                    $Excel = New-Excel -Path $Path -ErrorAction Stop
                }

                $PivotTableWorkSheets = @( $Excel | Get-Worksheet -ErrorAction Stop | Where-Object {$_.PivotTables} )
            }
            Catch
            {
                Throw "Could not get ExcelPackage or Worksheets to search: $_"
            }

        #Filter those tables
            If($PivotTableWorkSheets.Count -eq 0)
            {
                Throw "Something went wrong, we didn't find any worksheets with a pivot table"
            }
            else
            {
                $PivotTables = @( $PivotTableWorkSheets | Select -ExpandProperty PivotTables | Where-Object {$_.Name -Like $PivotTableName})
            }

        if($PivotTables.count -gt 0)
        {
            Foreach($PivotTable in $PivotTables)
            {

                #No chart name? Take the pivottable name, prepend CT
                    if(-not $PSBoundParameters.ContainsKey('ChartName'))
                    {
                        $ChartName = "CT-$( $PivotTable.Name )"
                    }

                #We need a worksheet for the chart
                    if( @( $Excel.WorkBook.Worksheets | Select -ExpandProperty Name -ErrorAction SilentlyContinue) -notcontains $TargetWorkSheetName)
                    {
                        $TargetWorkSheet = $Excel.Workbook.Worksheets | Where-Object {$_.Name -like $PivotTable.Worksheet.Name}
                        Write-Verbose "Could not find target worksheet '$TargetWorkSheetName', picking $($TargetWorkSheet.Name)"
                    }
                    else
                    {
                        $TargetWorkSheet = $Excel.Workbook.Worksheets[$TargetWorkSheetName]
                    }

                #We need to avoid dupes
                    if( @( $TargetWorkSheet.Drawings | Select -ExpandProperty Name -ErrorAction SilentlyContinue) -contains $ChartName)
                    {
                        Write-Error "Duplicate drawing found for ChartName '$ChartName', please specify a unique chart name"
                        continue
                    }

                #We have all we need, create the chart!
                    Write-Verbose "Adding $ChartType chart"

                    $chart = $TargetWorkSheet.Drawings.AddChart("$ChartName", $ChartType, $PivotTable)
                    $chart.SetPosition(1, 0, 6, 0)
                    $chart.SetSize($ChartWidth, $ChartHeight)
                    if($ChartTitle)
                    {
                        $chart.title.text = $ChartTitle
                    }
            }
        }
        else
        {
            Throw "Found no pivot tables matching '$PivotTableName'.  Existing pivot tables:`n$($PivotTableWorkSheets | Select -ExpandProperty PivotTables | Select -ExpandProperty Name )"
        }
        
        #Clean up
            if($PSCmdlet.ParameterSetName -like 'File' -and -not $Passthru)
            {
                Write-Verbose "Saving '$($Excel.File)'"
                $Excel.save()
                $Excel.Dispose()
            }
            if($Passthru)
            {
                $Excel
            }
    }
}
