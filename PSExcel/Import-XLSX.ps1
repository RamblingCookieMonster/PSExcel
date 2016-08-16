function Import-XLSX {
    <#
    .SYNOPSIS
        Import data from Excel
    .DESCRIPTION
        Import data from Excel
    .PARAMETER Path
        Path to an xlsx file to import
    .PARAMETER Sheet
        Index or name of Worksheet to import
    .PARAMETER Header
        Replacement headers.  Must match order and count of your data's properties.
    .PARAMETER RowStart
        First row to start reading from, typically the header. Default is 1
    .PARAMETER ColumnStart
        First column to start reading from. Default is 1
    .PARAMETER FirstRowIsData
        Indicates that the first row is data, not headers.  Must be used with -Header.
    .PARAMETER Text
        Extract cell text, rather than value.
        For example, if you have a cell with value 5:
            If the Number Format is '0', the text would be 5
            If the Number Format is 0.00, the text would be 5.00 
    .EXAMPLE
        Import-XLSX -Path "C:\Excel.xlsx"
        #Import data from C:\Excel.xlsx
    .EXAMPLE
        Import-XLSX -Path "C:\Excel.xlsx" -Header One, Two, Five
        # Import data from C:\Excel.xlsx
        # Replace headers with One, Two, Five
    .EXAMPLE
        Import-XLSX -Path "C:\Excel.xlsx" -Header One, Two, Five -FirstRowIsData -Sheet 2
        # Import data from C:\Excel.xlsx
        # Assume first row is data
        # Use headers One, Two, Five
        # Pull from sheet 2 (sheet 1 is default)
    .EXAMPLE
       #    A        B        C 
       # 1  Random text to mess with you!
       # 2  Header1  Header2  Header3
       # 3  data1    Data2    Data3
       # Your worksheet has data you don't care about in the first row or column
       # Use the ColumnStart or RowStart parameters to solve this.
       Import-XLSX -Path C:\RandomTextInRow1.xlsx -RowStart 2
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
    [cmdletbinding()]
    param(
        [parameter( Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        [validatescript({Test-Path $_})]
        [string[]]$Path,

        $Sheet = 1,

        [string[]]$Header,

        [switch]$FirstRowIsData,

        [ValidateSet('Text', 'Value')]
        [string]$Interpreter = 'Value',

        [UInt32]$RowStart = 1,

        [UInt32]$ColumnStart = 1,

        [UInt32]$RowCount,
        
        [UInt32]$ColumnCount,

        [switch]$IgnoreEmptyCells,

        [UInt32]$RowHeader = 1
    )
    Begin
    {
        function ColumnNumberToColumnLetter([uint64] $ColumnNumber)
        {
            while ($ColumnNumber -gt 0)
            {
                $Modulo = ($ColumnNumber - 1) % 26
                $ColumnName = [Char] (65 + $Modulo) + $ColumnName;
                $ColumnNumber = [uint64](($ColumnNumber - $Modulo) / 26);
            }

            $ColumnName
        }
    }
    
    Process
    {
        foreach($file in $path)
        {
            #Resolve relative paths... Thanks Oisin! http://stackoverflow.com/a/3040982/3067642
            $file = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($file)

            Write-Verbose "target excel file $($file)"
            
            Try
            {
                $xl = New-Object OfficeOpenXml.ExcelPackage $file
                $workbook  = $xl.Workbook
            }
            Catch
            {
                Write-Error "Failed to open '$file':`n$_"
                continue
            }

            Try
            {
                if( @($workbook.Worksheets).count -eq 0)
                {
                    Throw "No worksheets found"
                }
                else
                {
                    $worksheet = $workbook.Worksheets[$Sheet]
                    $dimension = $worksheet.Dimension

                    $Rows = $dimension.Rows
                    $Columns = $dimension.Columns

                    $RowEnd = if ($RowCount) {$RowCount + $RowStart} else {$Rows}

                    if ($ColumnCount)
                    {
                        $ColumnEnd =  $ColumnCount + $ColumnStart
                    }
                    else
                    {
                         $ColumnEnd = $Columns
                    }

                    $RowCount = $RowEnd - $RowStart + 1
                    $ColumnCount = $ColumnEnd - $ColumnStart + 1
                }

            }
            Catch
            {
                Write-Error "Failed to gather Worksheet '$Sheet' data for file '$file':`n$_"
                continue
            }
  
            # Define headears
            $Headers = @()

            foreach ($i in $ColumnStart..$ColumnEnd)
            {
                if ($Header -and -not [string]::IsNullOrEmpty($Header[$i - $ColumnStart]) -and $Header[$i - $ColumnStart] -notin $Headers)
                {
                    $Headers += $Header[$i - $ColumnStart]
                    continue
                }

                $Value = $worksheet.Cells.Item($RowHeader,$i).$Interpreter

                if ([string]::IsNullOrEmpty($Value) -or $FirstRowIsData)
                {
                    $Headers += "Column$(ColumnNumberToColumnLetter $i)"
                    continue
                }

                $i = 1
                $originalValue = $Value

                while ($Value -in $Headers)
                {
                    $Value = "$originalValue$i"
                    $i++ 
                }

                $Headers += $Value
            }

            Write-Verbose "Found $Rows rows, $Columns columns"
            Write-Verbose "Will read $RowCount rows, $ColumnCount columns, with headers:`n$($Header | Out-String)"

            $typeName = "Excel$(([System.IO.FileInfo]$File).BaseName)" 
            Update-TypeData -DefaultDisplayPropertySet $Headers -TypeName $typeName -Force

            if(-not $FirstRowIsData)
            {
                $RowStart++

                if ($RowStart -gt $RowEnd)
                {
                    return
                }
            }
            
            foreach ($RowId in $RowStart..$RowEnd)
            {
                $RowData = @{}
                $RowHeaders = @()

                foreach ($ColumnId in $ColumnStart..$ColumnEnd)
                {
                    $Name  = $Headers[$ColumnId - $ColumnStart]
                    
	                $Value = $worksheet.Cells.Item($RowId, $ColumnId).$Interpreter
                    Write-Debug "Row: $RowId, Column: $ColumnId, Name: $Name, Value = $Value"

                    #Handle dates, they're too common to overlook... Could use help, not sure if this is the best regex to use?
                    $Format = $worksheet.Cells.Item($RowId, $ColumnId).style.numberformat.format
                    if($Format -match '\w{1,4}/\w{1,2}/\w{1,4}( \w{1,2}:\w{1,2})?')
                    {
                        Try
                        {
                            $Value = [datetime]::FromOADate($Value)
                        }
                        Catch
                        {
                            Write-Verbose "Error converting '$Value' to datetime"
                        }
                    }

                    if ($IgnoreEmptyCells -and [string]::IsNullOrEmpty($Value))
                    {
                        Write-Verbose "Ignoring empty cell on row $RowId and column $ColumnId"
                    }
                    else
                    {
                        $RowData.Add($Name, $Value)
                        $RowHeaders += $Name
                    }
                }
                
                if (@($psObject.PSObject.Properties).Count -gt 0)
                {
                    $psObject = New-Object -TypeName PSObject -Property $RowData
                    $psObject.PSTypeNames.Insert(0, $typeName)
                    $psObject
                }
            }

            $xl.Dispose()
            $xl = $null
        }
    }
}
