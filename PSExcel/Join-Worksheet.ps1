Function Join-Worksheet {
    <#
    .SYNOPSIS
        Join two worksheets based on a common value

    .DESCRIPTION
        Join two worksheets based on a common value

        This wraps calls to Get-CellValue, Join-Object, and Export-XLSX.

        BETA NOTE:
            Minimal manual testing, no Pester tests
            Might add the option 

        NOTE:
            Each time you call this function, you need to save and re-create your Excel Object.
            If you attempt to modify the Excel object, save, modify, and save a second time, it will fail.
            See Save-Excel Passthru parameter for a workaround

            See Join-Object for more details on the join operation

    .PARAMETER Path
        Path to the file to write joined worksheet to.  We save changes to this.

    .PARAMETER Excel
        Excel package to write joined worksheet to.  We do not save this.

    .PARAMETER DestinationWorksheetName
        Name the worksheet you are adding joined data to

    .PARAMETER LeftWorksheet
        Left worksheet to join

    .PARAMETER RightWorksheet
        Right worksheet to join

    .PARAMETER LeftJoinColumn
        Column on left worksheet that we match up with RightJoinColumn on the right worksheet

    .PARAMETER RightJoinColumn
        Column on right worksheet that we match up with LeftJoinColumn on the left worksheet

    .PARAMETER LeftColumns
        One or more columns to keep from the left worksheet.  Default is to pull all left columns (*).

        Each property can:
            - Be a plain property name like "Name"
            - Contain wildcards like "*"
            - Be a hashtable like @{Name="Product Name";Expression={$_.Name}}.
                 Name is the output property name
                 Expression is the property value ($_ as the current object)
                
                 Alternatively, use the Suffix or Prefix parameter to avoid collisions
                 Each property using this hashtable syntax will be excluded from suffixes and prefixes

    .PARAMETER RightColumns
        One or more columns to keep from right worksheet.  Default is to pull all right columns (*).

        Each property can:
            - Be a plain property name like "Name"
            - Contain wildcards like "*"
            - Be a hashtable like @{Name="Product Name";Expression={$_.Name}}.
                 Name is the output property name
                 Expression is the property value ($_ as the current object)
                
                 Alternatively, use the Suffix or Prefix parameter to avoid collisions
                 Each property using this hashtable syntax will be excluded from suffixes and prefixes

    .PARAMETER Prefix
        If specified, prepend right column names with this prefix to avoid collisions

        Example:
            Column Name                     = 'Name'
            Suffix                          = 'j_'
            Resulting Joined Property Name  = 'j_Name'

    .PARAMETER Suffix
        If specified, append right column names with this suffix to avoid collisions

        Example:
            Column Name                     = 'Name'
            Suffix                          = '_j'
            Resulting Joined Property Name  = 'Name_j'

    .PARAMETER Type

        Type of join.  Default is AllInLeft.

        AllInLeft will have all elements from Left at least once in the output, and might appear more than once
          if the where clause is true for more than one element in right, Left elements with matches in Right are
          preceded by elements with no matches.
          SQL equivalent: outer left join (or simply left join)

        AllInRight is similar to AllInLeft.
        
        OnlyIfInBoth will cause all elements from Left to be placed in the output, only if there is at least one
          match in Right.
          SQL equivalent: inner join (or simply join)
          
        AllInBoth will have all entries in right and left in the output. Specifically, it will have all entries
          in right with at least one match in left, followed by all entries in Right with no matches in left, 
          followed by all entries in Left with no matches in Right.
          SQL equivalent: full join

    .PARAMETER AutoFit
        If specified, autofit everything

    .PARAMETER PivotRows
        If specified, add pivot table pivoting on these rows

    .PARAMETER PivotColumns
        If specified, add pivot table pivoting on these columns

    .PARAMETER PivotValues
        If specified, add pivot table pivoting on these values

    .PARAMETER ChartType
        If specified, add pivot chart of this type

    .PARAMETER Table
        If specified, add table to all cells

    .PARAMETER TableStyle
        If specified, add table style

    .PARAMETER Force
        If specified, and Path parameter is used, remove existing file if it is found

        If force is not specified and an existing XLSX is found, we try to add the worksheet to it

    .PARAMETER Passthru
        If specified, and Excel parameter is used, return Excel package object

    .EXAMPLE

        #Define some input data.

            $l = 1..5 | Foreach-Object {
                [pscustomobject]@{
                    Name = "jsmith$_"
                    Birthday = (Get-Date).adddays(-1)
                }
            }

            $r = 4..7 | Foreach-Object{
                [pscustomobject]@{
                    Department = "Department $_"
                    Name = "Department $_"
                    Manager = "jsmith$_"
                }
            }

        #Export it to a spreadsheet with specific worksheet names

            $l | export-xlsx -Path C:\temp\JoinTest.xlsx -WorksheetName Left
            $r | export-xlsx -Path C:\temp\JoinTest.xlsx -WorksheetName Right

        #Get the worksheets:
            $Excel = New-Excel -Path C:\temp\JoinTest.xlsx
            $LeftWorksheet = Get-Worksheet -Excel $Excel -Name 'Left'
            $RightWorksheet = Get-WorkSheet -Excel $Excel -Name 'Right'

        #We have the data - join it where Left.Name = Right.Manager
            Join-Worksheet -Path C:\temp\test.xlsx -LeftWorksheet $LeftWorksheet -RightWorksheet $RightWorksheet -LeftJoinColumn Name -RightJoinColumn Manager
            $Excel | Close-Excel

        #Verify the output:
            Import-XLSX -Path C:\temp\test.xlsx

            # Name         Birthday              Department   Manager
            # ----         --------              ----------   -------
            # jsmith1      4/15/2015 12:30:21 PM                     
            # jsmith2      4/15/2015 12:30:21 PM                     
            # jsmith3      4/15/2015 12:30:21 PM                     
            # Department 4 4/15/2015 12:30:21 PM Department 4 jsmith4
            # Department 5 4/15/2015 12:30:21 PM Department 5 jsmith5


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
        [parameter( ParametersetName = 'Path',
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

        [parameter( ParameterSetName = "Excel",
                    Position = 0,
                    Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$false)]
        [OfficeOpenXml.ExcelPackage]$Excel,

        [string]$DestinationWorksheetName = 'WorksheetJoin',
        
        [parameter( Mandatory=$true,
                    ValueFromPipeline=$false,
                    ValueFromPipelineByPropertyName=$false)]
        [OfficeOpenXml.ExcelWorksheet]$LeftWorksheet,
        
        [parameter( Mandatory=$true,
                    ValueFromPipeline=$false,
                    ValueFromPipelineByPropertyName=$false)]
        [OfficeOpenXml.ExcelWorksheet]$RightWorksheet,

        [Parameter(Mandatory = $true)]
        [string]$LeftJoinColumn,

        [Parameter(Mandatory = $true)]
        [string]$RightJoinColumn,

        [object[]]$LeftColumns,
        [object[]]$RightColumns,
        [string]$Prefix,
        [string]$Suffix,

        [validateset( 'AllInLeft', 'OnlyIfInBoth', 'AllInBoth', 'AllInRight')]
        [Parameter(Mandatory=$false)]
        [string]$Type = 'AllInLeft',

        [string[]]$Header,

        [switch]$Table,

        [OfficeOpenXml.Table.TableStyles]$TableStyle = [OfficeOpenXml.Table.TableStyles]"Medium2",

        [switch]$AutoFit,

        [switch]$Force,
        
        [switch]$Passthru
    )
    begin
    {
        #Resolve relative paths... Thanks Oisin! http://stackoverflow.com/a/3040982/3067642
        if($PSBoundParameters.ContainsKey('Path'))
        {
            $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
        }

        Try
        {
            $Left = Get-CellValue -WorkSheet $LeftWorksheet -ErrorAction stop
        }
        Catch
        {
            Throw "Error getting LeftWorksheet data: $_"
        }

        Try
        {
            $Right = Get-CellValue -WorkSheet $RightWorksheet -ErrorAction stop
        }
        Catch
        {
            Throw "Error getting RightWorksheet data: $_"
        }

        $MergeParams = @{
            Left = $Left
            Right = $Right
        }

        Switch($PSBoundParameters.Keys)
        {
            'LeftJoinColumn' { $MergeParams.Add('LeftJoinProperty',$PSBoundParameters['LeftJoinColumn'] ) }
            'RightJoinColumn' { $MergeParams.Add('RightJoinProperty',$PSBoundParameters['RightJoinColumn'] ) }
            'LeftColumns' { $MergeParams.Add('LeftProperties',$PSBoundParameters['LeftColumns'] ) }
            'RightColumns' { $MergeParams.Add('RightProperties',$PSBoundParameters['RightColumns'] ) }
            'Prefix' { $MergeParams.Add('Prefix',$PSBoundParameters['Prefix'] ) }
            'Suffix' { $MergeParams.Add('Suffix',$PSBoundParameters['Suffix'] ) }
            'Type' { $MergeParams.Add('Type',$PSBoundParameters['Type'] ) }
        }

        Try
        {
            $Merge = Join-Object @MergeParams -ErrorAction Stop
        }
        Catch
        {
            Write-Error $_
            Throw "Error merging data: $_"
        }
    }
    process
    {
        $ExportParams = @{ InputObject = $Merge }

        switch ($PSBoundParameters.Keys)
        {
            'Excel'      { $ExportParams.Add('Excel',$Excel) }
            'Path'       { $ExportParams.Add('Path',$Path) }
            'Header'     { $ExportParams.Add('Header',$Header) }
            'Table'      { $ExportParams.Add('Table',$Table) }
            'TableStyle' { $ExportParams.Add('TableStyle',$TableStyle) }
            'AutoFit'    { $ExportParams.Add('AutoFit',$AutoFit) }
            'Force'      { $ExportParams.Add('Force',$Force) }
        }

        Export-XLSX @ExportParams
        if($PSBoundParameters.ContainsKey('Excel') -and $Passthru)
        {
            $Excel
        }
    }
}
