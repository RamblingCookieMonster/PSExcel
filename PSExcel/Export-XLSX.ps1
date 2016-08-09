Function Export-XLSX {
    <#
    .SYNOPSIS
        Export data to an XLSX file

    .DESCRIPTION
        Export data to an XLSX file

    .PARAMETER InputObject
        Data to export

    .PARAMETER Path
        Path to the file to export

    .PARAMETER WorksheetName
        Name the worksheet you are importing to

    .PARAMETER Header
        Header to use. Must match order and count of your data's properties

    .PARAMETER AutoFit
        If specified, autofit everything

    .PARAMETER PivotRows
        If specified, add pivot table pivoting on these rows

    .PARAMETER PivotColumns
        If specified, add pivot table pivoting on these columns

    .PARAMETER PivotValues
        If specified, add pivot table pivoting on these values

    .PARAMETER PivotFunction
        If specified, use this summary mode for pivot values (defaults to Count mode)

    .PARAMETER ChartType
        If specified, add pivot chart of this type

    .PARAMETER Table
        If specified, add table to all cells

    .PARAMETER TableStyle
        If specified, add table style

    .PARAMETER Append
        If specified, append to existing worksheet.

        We don't check header names, but header count must match.

    .PARAMETER Force
        If file exists, overwrite it.

    .PARAMETER ClearSheet
        If worksheet with the same name exists, clear it before filling it in

    .PARAMETER RemoveSheet
        If worksheet with the same name exists, remove it and re-add it

    .PARAMETER Passthru
        If specified, we re-open the ExcelPackage and return it

    .EXAMPLE
        $Files = Get-ChildItem C:\ -File

        Export-XLSX -Path C:\Files.xlsx -InputObject $Files

        Export file listing to C:\Files.xlsx

    .EXAMPLE
        $Files = Get-ChildItem C:\ -File

		$Worksheet = 'Files'

        Export-XLSX -Path C:\temp\Files.xlsx -InputObject $Files -WorksheetName $Worksheet -ClearSheet

        Export file listing to C:\temp\Files.xlsx to the worksheet named "Files".  If it exists already, clear the sheet then import the data.

    .EXAMPLE

        1..10 | Foreach-Object {
            New-Object -typename PSObject -Property @{
                Something = "Prop$_"
                Value = Get-Random
            }
        } |
            Select-Object Something, Value |
            Export-XLSX -Path C:\Random.xlsx -Force -Header Name, Val

        # Generate data
        # Send it to Export-XLSX
        # Give it new headers
        # Overwrite C:\random.xlsx if it exists

    .EXAMPLE

        # Create XLSX
        Get-ChildItem -file | Export-XLSX -Path C:\temp\multi.xlsx

        # Add a second worksheet to the xlsx
        Get-ChildItem -file | Export-XLSX -Path C:\temp\multi.xlsx -WorksheetName "Two"

    .EXAMPLE

        # Create XLSX
        Get-ChildItem -file | Export-XLSX -Path C:\temp\multi.xlsx

        # Add a second worksheet to the xlsx
        Get-ChildItem -file | Export-XLSX -Path C:\temp\multi.xlsx -WorksheetName "Two"

        # I don't like that second worksheet. Recreate it, deleting the existing worksheet if it exists.
        Get-ChildItem -file | Select -first 1 | Export-XLSX -Path C:\temp\multi.xlsx -WorksheetName "Two" -ReplaceSheet

    .EXAMPLE

        Get-ChildItem C:\ -file |
            Export-XLSX -Path C:\temp\files.xlsx -PivotRows Extension -PivotValues Length -ChartType Pie

        # Get files
        # Create an xlsx in C:\temp\files.xlsx
        # Pivot rows on 'Extension'
        # Pivot values on 'Length
        # Add a pie chart

        # This example gives you a pie chart breaking down storage by file extension

    .EXAMPLE

	    Get-Process | Export-XLSX -Path C:\temp\process.xlsx -Worksheet process -Table -TableStyle Medium1 -AutoFit

	    # Get all processes
	    # Create an xlsx
	    # Create a table with the Medium1 style and all cells autofit on the 'process' worksheet

    .EXAMPLE

    #
    # This example illustrates appending data

        1..10 | Foreach-Object {
            New-Object -typename PSObject -Property @{
                Something = "Prop$_"
                Value = Get-Random
            }
        } |
            Select-Object Something, Value |
            Export-XLSX -Path C:\Random.xlsx -Force

        # Generate data
        # Send it to Export-XLSX
        # Overwrite C:\random.xlsx if it exists

        1..5 | Foreach-Object {
            New-Object -typename PSObject -Property @{
                Something = "Prop$_"
                Value = Get-Random
            }
        } |
            Select-Object Something, Value |
            Export-XLSX -Path C:\Random.xlsx -Append

        # Generate data
        # Send it to Export-XLSX
        # Append to C:\random.xlsx

    .NOTES
        Thanks to Doug Finke for his example
        The pivot stuff is straight from Doug:
            https://github.com/dfinke/ImportExcel

        Thanks to Philip Thompson for an expansive set of examples on working with EPPlus in PowerShell:
            https://excelpslib.codeplex.com/

    .LINK
        https://github.com/RamblingCookieMonster/PSExcel

    .FUNCTIONALITY
        Excel
    #>
    [CmdletBinding(DefaultParameterSetName='Path')]
    param(
        [parameter( ParameterSetName='Path',
                    Position = 0,
                    Mandatory=$true )]
        [ValidateScript({
            $Parent = Split-Path $_ -Parent
            if( -not (Test-Path -Path $Parent -PathType Container) )
            {
                Throw "Specify a valid path.  Parent '$Parent' does not exist: $_"
            }
            $True
        })]
        [string]$Path,

        [parameter( ParameterSetName='Excel',
                    Position = 0,
                    Mandatory=$true )]
        [OfficeOpenXml.ExcelPackage]$Excel,

        [parameter( Position = 1,
                    Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromRemainingArguments=$false)]
        $InputObject,

        [string[]]$Header,

        [string]$WorksheetName = "Worksheet1",

        [string[]]$PivotRows,

        [string[]]$PivotColumns,

        [string[]]$PivotValues,
        
        [OfficeOpenXml.Table.PivotTable.DataFieldFunctions]$PivotFunction = [OfficeOpenXml.Table.PivotTable.DataFieldFunctions]::Count,

        [OfficeOpenXml.Drawing.Chart.eChartType]$ChartType,

        [switch]$Table,

        [OfficeOpenXml.Table.TableStyles]$TableStyle = [OfficeOpenXml.Table.TableStyles]"Medium2",

        [switch]$AutoFit,

        [switch]$Append,

        [switch]$Force,

		[switch]$ClearSheet,

        [switch]$ReplaceSheet,

        [switch]$Passthru
    )
    begin
    {
        if ( $PSBoundParameters.ContainsKey('Path'))
        {
            if ( Test-Path $Path ) 
            {
                if($Append)
                {
                    Write-Verbose "'$Path' exists. Appending data"
                }
                elseif($Force)
                {
                    Try
                    {
                        Remove-Item -Path $Path -Force -Confirm:$False
                    }
                    Catch
                    {
                        Throw "'$Path' exists and could not be removed: $_"
                    }
                }
                else
                {
                    Write-Verbose "'$Path' exists. Use -Force to overwrite. Attempting to add sheet to existing workbook"
                }
            }

            #Resolve relative paths... Thanks Oisin! http://stackoverflow.com/a/3040982/3067642
            $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
        }

        Write-Verbose "Export-XLSX '$($PSCmdlet.ParameterSetName)' PSBoundParameters = $($PSBoundParameters | Out-String)"

        $bound = $PSBoundParameters.keys -contains "InputObject"
        if(-not $bound)
        {
            [System.Collections.ArrayList]$AllData = @()
        }
    }
    process
    {
        #We write data by row, so need everything countable, not going to stream...
        if($bound)
        {
            $AllData = $InputObject
        }
        Else
        {
            foreach($Object in $InputObject)
            {
                [void]$AllData.add($Object)
            }
        }
    }
    end
    {
        #Deal with headers
            $ExistingHeader = @(

                # indexes might be an issue if we get array of strings, so select first
                ($AllData | Select -first 1).PSObject.Properties |
                    Select -ExpandProperty Name
            )

            $Columns = $ExistingHeader.count

            if($Header)
            {
                if($Header.count -ne $ExistingHeader.count)
                {
                    Throw "Found '$columns' columns, provided $($header.count) headers.  You must provide a header for every column."
                }
            }
            else
            {
                $Header = $ExistingHeader
            }

        #initialize stuff
            $RowIndex = 2
            Try
            {
                if( $PSBoundParameters.ContainsKey('Path'))
                {
                    $Excel = New-Object OfficeOpenXml.ExcelPackage($Path) -ErrorAction Stop
                }
                else
                {
                    $Path = $Excel.File.FullName
                }

                $Workbook = $Excel.Workbook
                if($ReplaceSheet)
                {
                    Try
                    {
                        Write-Verbose "Attempting to delete worksheet $WorksheetName"
                        $Workbook.Worksheets.Delete($WorksheetName)
                    }
                    Catch
                    {
                        if($_.Exception -notmatch 'Could not find worksheet to delete')
                        {
                            Write-Error "Error removing worksheet $WorksheetName"
                            Throw $_
                        }
                    }
                }

                # They asked for append but we don't have a worksheet.  Drop the append switch.
                if($Excel.Workbook.WorkSheets.count -eq 0 )
                {
                    $Append = $False
                }

                #If we have an excel or valid path, try to append or clearsheet as needed
                if (($Append -or $ClearSheet) -and ($PSBoundParameters.ContainsKey('Excel') -or (Test-Path $Path)) )
                {
                    $WorkSheet=$Excel.Workbook.Worksheets | Where-Object {$_.Name -like $WorkSheetName}
                    if($ClearSheet)
                    {
                        $WorkSheet.Cells[$WorkSheet.Dimension.Start.Row, $WorkSheet.Dimension.Start.Column, $WorkSheet.Dimension.End.Row, $WorkSheet.Dimension.End.Column].Clear()
                    }
                    if($Append)
                    {
                        $RealHeaderCount = $WorkSheet.Dimension.Columns
                        if($Header.count -ne $RealHeaderCount)
                        {
                            $Excel.Dispose()
                            Throw "Found $RealHeaderCount existing headers, provided data has $($Header.count)."
                        }
                        $RowIndex = 1 + $Worksheet.Dimension.Rows
                    }
                }
                else
                {
                    $WorkSheet = $Workbook.Worksheets.Add($WorkSheetName)
                }
            }
            Catch
            {
                Throw "Failed to initialize Excel, Workbook, or Worksheet. Try -ClearSheet switch if worksheet already exists:`n`n_"
            }

        #Set those headers if we aren't appending
            if(-not $Append)
            {
                for ($ColumnIndex = 1; $ColumnIndex -le $Header.count; $ColumnIndex++)
                {
                    $WorkSheet.SetValue(1, $ColumnIndex, $Header[$ColumnIndex - 1])
                }
            }

        #Write the data...
            foreach($RowData in $AllData)
            {
                Write-Verbose "Working on object:`n$($RowData | Out-String)"
                for ($ColumnIndex = 1; $ColumnIndex -le $Header.count; $ColumnIndex++)
                {
                    $Object = @($RowData.PSObject.Properties)[$ColumnIndex - 1]
                    $Value = $Object.Value
                    $WorkSheet.SetValue($RowIndex, $ColumnIndex, $Value)

                    Try
                    {
                        #Nulls will error, catch them
                        $ThisType = $Null
                        $ThisType = $Value.GetType().FullName
                    }
                    Catch
                    {
                        Write-Verbose "Applying no style to null in row $RowIndex, column $ColumnIndex"
                    }

                    #Idea from Philip Thompson, thank you Philip!
                    $StyleName = $Null
                    $ExistingStyles = @($WorkBook.Styles.NamedStyles | Select -ExpandProperty Name)
                    Switch -regex ($ThisType)
                    {
                        "double|decimal|single"
                        {
                            $StyleName = 'decimals'
                            $StyleFormat = "0.00"
                        }
                        "int\d\d$"
                        {
                            $StyleName = 'ints'
                            $StyleFormat = "0"
                        }
                        "datetime"
                        {
                            $StyleName = "dates"
                            $StyleFormat = "M/d/yyy h:mm"
                        }
                        "TimeSpan"
                        {
                            #Open to other ways to handle this
                            $WorkSheet.SetValue($RowIndex, $ColumnIndex, "$Value")
                        }
                        default
                        {
                            #No default yet...
                        }
                    }

                    if($StyleName)
                    {
                        if($ExistingStyles -notcontains $StyleName)
                        {
                            $StyleSheet = $WorkBook.Styles.CreateNamedStyle($StyleName)
                            $StyleSheet.Style.Numberformat.Format = $StyleFormat
                        }

                        $WorkSheet.Cells.Item($RowIndex, $ColumnIndex).Stylename = $StyleName
                    }

                }
                Write-Verbose "Wrote row $RowIndex"
                $RowIndex++
            }

            # Any pivot params specified?  add a pivot!
            if($PSBoundParameters.Keys -match 'Pivot')
            {
                $Params = @{}
                if($PivotRows)     {$Params.Add('PivotRows',$PivotRows)}
                if($PivotColumns)  {$Params.Add('PivotColumns',$PivotColumns)}
                if($PivotValues)   {$Params.Add('PivotValues',$PivotValues)}
                if($PivotFunction) {$Params.Add('PivotFunction',$PivotFunction)}
                if($ChartType)     {$Params.Add('ChartType',$ChartType)}
                $Excel = Add-PivotTable @Params -Excel $Excel -WorkSheetName $WorksheetName -Passthru -ErrorAction stop

            }

            # Create table
            elseif($Table)
            {
                $Excel = Add-Table -Excel $Excel -WorkSheetName $WorksheetName -TableStyle $TableStyle -Passthru
            }

            if($AutoFit)
            {
                $WorkSheet.Cells[$WorkSheet.Dimension.Address].AutoFitColumns()
            }

            # This is an export command. Save whether we have a path or ExcelPackage input...
            $Excel.SaveAs($Path)

            if($Passthru)
            {
                New-Excel -Path $Path
            }
    }
}
