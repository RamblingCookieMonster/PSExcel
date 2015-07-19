function Add-Table {
    <#
    .SYNOPSIS
        Add a table to an Excel worksheet

    .DESCRIPTION
        Add a table to an Excel worksheet

        Note:
            Each time you call this function, you need to save and re-create your Excel Object.
            If you attempt to modify the Excel object, save, modify, and save a second time, it will fail.
            See Save-Excel Passthru parameter for a workaround

    .PARAMETER Path
        Path to an xlsx file to add the table to

        If Path is specified and you do not use passthru, we save the file

    .PARAMETER Excel
        ExcelPackage to add the table to

        We do not save the ExcelPackage upon completion.  See Save-Excel.

    .PARAMETER WorkSheetName
        If specified, use this worksheet as the source.

    .PARAMETER StartRow
        The top row for table data.  If not specified, we use the dimensions start row

    .PARAMETER StartColumn
        The leftmost column for table data.  If not specified, we use the dimensions start column

    .PARAMETER EndRow
        The bottom row for table data.  If not specified, we use the dimensions' end row

    .PARAMETER EndColumn
        The rightmost column for table data.  If not specified, we use the dimensions' end column

    .PARAMETER TableStyle
        Style of the table

    .PARAMETER TableName
        Name of the table, defaults to worksheet name if none provided

    .PARAMETER Passthru
        If specified, pass the ExcelPackage back

    .EXAMPLE

        Get-ChildItem C:\ -file | Export-XLSX -Path C:\temp\files.xlsx

        Add-Table -Path C:\Temp\files.xlsx -TableStyle Medium10

        # Get files, create an xlsx in C:\temp\ps.xlsx

        # Take existing xlsx and add a table with the Medium10 style

    .EXAMPLE
        # Create an xlsx.
            Get-ChildItem C:\ -file | Export-XLSX -Path C:\temp\files.xlsx

        # Open the excel file, add a table (this won't save), pass through the excel object, save.
            New-Excel -Path C:\temp\files.xlsx |
                Add-Table -TableStyle Medium10 -TableName "Files" -Passthru |
                Save-Excel -Close

    .NOTES
        Added by Andrew Widdersheim

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
        [string]$WorkSheetName,

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

        [OfficeOpenXml.Table.TableStyles]$TableStyle,

        [string]$TableName,

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
            # Get the coordinates
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

                if(-not $TableName)
                {
                    $TableWorksheetName = $SourceWorkSheet.Name
                }
                else
                {
                    $TableWorksheetName = $TableName
                }

                Write-Verbose "Adding table over data range '$RangeCoordinates' with name $TableWorksheetName"
                $Table = $SourceWorkSheet.Tables.Add($SourceWorkSheet.Cells[$RangeCoordinates], $TableWorksheetName)

                if($TableStyle)
                {
                    Write-Verbose "Adding $TableStyle table style"
                    $Table.TableStyle = $TableStyle
                }

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
}
