function Save-Excel {
    <#
    .SYNOPSIS
        Save an OfficeOpenXml ExcelPackage

    .DESCRIPTION
        Save an OfficeOpenXml ExcelPackage

    .PARAMETER Excel
        An ExcelPackage object to close

    .PARAMETER Path
        If specified, save as this path

    .PARAMETER Close
        If specified, close after saving

    .PARAMETER Passthru
        If specified, re-create and return the Excel object

    .EXAMPLE
        Save-Excel -Excel $Excel

        #Save $Excel

    .EXAMPLE
        Save-Excel -Excel $Excel -Close

        #Save $Excel, close

    .EXAMPLE
        Save-Excel -Excel $Excel -Path C:\new.xlsx

        #Save $Excel as C:\new.xlsx

    .EXAMPLE
        $Excel = $Excel | Save-Excel -Passthru

        #Save $Excel, re-open it to continue working with it.

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
        [parameter( Mandatory=$true,
                    ValueFromPipeline=$true,
                    ValueFromPipelineByPropertyName=$true)]
        [OfficeOpenXml.ExcelPackage]$Excel,

        [parameter( Mandatory=$false,
                    ValueFromPipeline=$false,
                    ValueFromPipelineByPropertyName=$false)]
        [validatescript({
            $Parent = Split-Path $_ -Parent -ErrorAction SilentlyContinue
            if( -not (Test-Path -Path $Parent -PathType Container -ErrorAction SilentlyContinue) )
            {
                Throw "Specify a valid path.  Parent '$Parent' does not exist: $_"
            }
            $True
        })]        
        [string]$Path,

        [switch]$Close,

        [switch]$Passthru
    )
    Process
    {
        foreach($xl in $Excel)
        {            
            Try
            {
                if($Path)
                {
                    Try
                    {
                        #Resolve relative paths... Thanks Oisin! http://stackoverflow.com/a/3040982/3067642
                        $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
                    }
                    Catch
                    {
                        Write-Error "Could not resolve path for '$Path': $_"
                        continue
                    }
                    
                    write-verbose "Saving $($xl.File) as $($Path)"

                    $xl.saveas($Path)
                }
                else
                {
                    write-verbose "Saving $($xl.File)"

                    $xl.save()
                }

                if($Passthru)
                {
                    $OpenPath = $xl.File
                    $xl.Dispose()
                    $xl = $Null
                    New-Excel -Path $OpenPath
                }
            }
            Catch
            {
                Write-Error "Error saving file. $_"
                Continue
            }
            
            if($Close)
            {
                Try
                {
                    write-verbose "Closing $($xl.File)"

                    $xl.Dispose()
                    $xl = $null
                }
                Catch
                {
                    Write-Error $_
                    Continue
                }
            }
        }    
    }
}