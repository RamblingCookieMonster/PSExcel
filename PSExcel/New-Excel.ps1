function New-Excel {
    <#
    .SYNOPSIS
        Create an OfficeOpenXml ExcelPackage to work with

    .DESCRIPTION
        Create an OfficeOpenXml ExcelPackage to work with

    .PARAMETER Path
        Path to an xlsx file to open
    
    .EXAMPLE
        $Excel = New-Excel -Path "C:\Excel.xlsx"
        $Excel.Workbook

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
    [OutputType([OfficeOpenXml.ExcelPackage])]
    [cmdletbinding()]
    param(
        [parameter( Mandatory=$false,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        [validatescript({
            $Parent = Split-Path $_ -Parent
            if( -not (Test-Path -Path $Parent -PathType Container) )
            {
                Throw "Specify a valid path.  Parent '$Parent' does not exist: $_"
            }
            $True
        })]
        [string]$Path
    )
    Process
    {

        if($path)
        {
            #Resolve relative paths... Thanks Oisin! http://stackoverflow.com/a/3040982/3067642
            $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
            Write-Verbose "Creating excel object with path '$path'"

            New-Object OfficeOpenXml.ExcelPackage $Path
        }
        else
        {
            Write-Verbose "Creating excel object with no specified path"
            New-Object OfficeOpenXml.ExcelPackage
        }
    }
}