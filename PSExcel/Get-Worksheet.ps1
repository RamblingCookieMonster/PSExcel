function Get-Worksheet {
    <#
    .SYNOPSIS
        Return an ExcelPackage Worksheet

    .DESCRIPTION
        Return an ExcelPackage Worksheet

    .PARAMETER Name
        If specified, return Worksheets named like this

    .PARAMETER Workbook
        Workbook to extract worksheets from

    .PARAMETER Excel
        ExcelPackage to extract worksheets from

    .EXAMPLE
        $Excel = New-Excel -Path "C:\Excel.xlsx"
        $WorkSheet = $Excel | Get-WorkSheet

        # Open C:\Excel.xlsx, view the worksheets in it

    .EXAMPLE
        $Workbook | Get-WorkSheet -Name "Worksheet2"

        # Get Worksheet with the name Worksheet2 from $WorkBook

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
    [cmdletbinding(DefaultParameterSetName = "Workbook")]
    param(
        [parameter(Mandatory=$false,
                    ValueFromPipeline=$false,
                    ValueFromPipelineByPropertyName=$false)]
        [string]$Name,

        [parameter( ParameterSetName = "Workbook",
                    Mandatory=$true,
                    ValueFromPipeline=$True,
                    ValueFromPipelineByPropertyName=$false)]
        [OfficeOpenXml.ExcelWorkbook]$Workbook,

        [parameter( ParameterSetName = "Excel",
                    Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$false)]
        [OfficeOpenXml.ExcelPackage]$Excel
    )
    Process
    {
        $Output = switch ($PSCmdlet.ParameterSetName)
        {
            "Workbook"
            {
                Write-Verbose "Processing Workbook"
                $Workbook.Worksheets
            }
            "Excel"
            {
                Write-Verbose "Processing ExcelPackage"
                $Excel.Workbook.Worksheets
            }
        }

        If($Name)
        {
            $FilteredOutput = $Output | Where-Object {$_.Name -like $Name}
            if($Name -notmatch '\*' -and -not $FilteredOutput)
            {
                Write-Error "$Name could not be found. Valid worksheets:`n$($Output | Select -ExpandProperty Name | Out-String)"
            }
            else
            {
                $FilteredOutput
            }
        }
        else
        {
            $Output
        }
    }
}