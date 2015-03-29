function Save-Excel {
    <#
    .SYNOPSIS
        Save an OfficeOpenXml ExcelPackage

    .DESCRIPTION
        Save an OfficeOpenXml ExcelPackage

    .PARAMETER Excel
        An ExcelPackage object to save

    .PARAMETER Path
        If specified, save as this path

    .PARAMETER Close
        If specified, close after saving

    .EXAMPLE
        Save-Excel -Excel $Excel

        #Save $Excel

    .EXAMPLE
        Save-Excel -Excel $Excel -Close

        #Save $Excel, close

    .EXAMPLE
        Save-Excel -Excel $Excel -Path C:\new.xlsx

        #Save $Excel as C:\new.xlsx

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
                else
                {
                    write-verbose "Saving $($xl.File)"

                    $xl.save()
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