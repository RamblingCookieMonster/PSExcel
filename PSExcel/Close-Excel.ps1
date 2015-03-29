function Close-Excel {
    <#
    .SYNOPSIS
        Close an OfficeOpenXml ExcelPackage

    .DESCRIPTION
        Close an OfficeOpenXml ExcelPackage

    .PARAMETER Excel
        An ExcelPackage object to close

    .PARAMETER Save
        Save the ExcelPackage before closing

    .PARAMETER Path
        If specified, Save the ExcelPackage as this path before closing

    .EXAMPLE
        Close-Excel -Excel $Excel -Save

        #Save and close $Excel

    .EXAMPLE
        Close-Excel -Excel $Excel

        #Close $Excel without saving

    .EXAMPLE
        Close-Excel -Excel $Excel -Path C:\new.xlsx

        #Save $Excel as C:\new.xlsx and close

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
        [OfficeOpenXml.ExcelPackage]$Excel,

        [Switch]$Save,

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
        [string]$Path
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
                elseif($Save)
                {
                    write-verbose "Saving $($xl.File)"

                    $xl.save()
                }
            }
            Catch
            {
                Write-Error "Error saving file.  Will not close this ExcelPackage: $_"
                Continue
            }
            
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