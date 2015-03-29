function Close-Excel {
    <#
    .SYNOPSIS
        Close an OfficeOpenXml ExcelPackage

    .DESCRIPTION
        Close an OfficeOpenXml ExcelPackage

    .EXAMPLE
        Close-Excel -Excel $Excel -Save

        #Save and close $Excel

    .EXAMPLE
        Close-Excel -Excel $Excel

        #Close $Excel without saving

    .NOTES
        Thanks to Doug Finke for his example:
            https://github.com/dfinke/ImportExcel/blob/master/ImportExcel.psm1

        Thanks to Philip Thompson for an expansive set of examples on working with EPPlus in PowerShell:
            https://excelpslib.codeplex.com/

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
        
        [switch]$SaveAs,

        [parameter( Mandatory=$false,
                    ValueFromPipeline=$false,
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
        foreach($xl in $Excel)
        {            
            Try
            {
                if($SaveAs)
                {

                    Try
                    {
                        #Resolve relative paths... Thanks Oisin! http://stackoverflow.com/a/3040982/3067642
                        $Path = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path)
                    }
                    Catch
                    {
                        Write-Error $_
                        continue
                    }
                    
                    write-verbose "Saving $($xl.File) as $($Path) and closing"

                    $xl.save()
                    $xl.Dispose()
                    $xl = $null
            
                }
                elseif($Save)
                {
                    write-verbose "Saving and closing $($xl.File)"

                    $xl.save()
                    $xl.Dispose()
                    $xl = $null

                }
                Else
                {
                    write-verbose "Closing $($xl.File)"

                    $xl.Dispose()
                    $xl = $null
                }
            }
            Catch
            {
                Write-Error $_
                Continue
            }
        }    
    }
}