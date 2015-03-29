function New-Excel {
    <#
    .SYNOPSIS
        Create an OfficeOpenXml ExcelPackage to work with

    .DESCRIPTION
        Create an OfficeOpenXml ExcelPackage to work with

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
        #Resolve relative paths... Thanks Oisin! http://stackoverflow.com/a/3040982/3067642
        $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)

        write-verbose "Creating excel object with path '$path'"

        if($path)
        {
            New-Object OfficeOpenXml.ExcelPackage $Path
        }
        else
        {
            New-Object OfficeOpenXml.ExcelPackage
        }
    }
}