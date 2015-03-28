#handle PS2
    if(-not $PSScriptRoot)
    {
        $PSScriptRoot = Split-Path $MyInvocation.MyCommand.Path -Parent
    }

#Import assembly:
    $BinaryPath = Join-Path $PSScriptRoot 'lib\epplus.dll'
    if( -not ($Library = Add-Type -path $BinaryPath -PassThru -ErrorAction stop) )
    {
        Throw "Failed to load EPPlus binary from $BinaryPath"
    }

#Get public and private function definition files.
    $Public  = Get-ChildItem $PSScriptRoot\*.ps1 -ErrorAction SilentlyContinue
    #$Private = Get-ChildItem $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue 

#Dot source the files
    Foreach($import in @($Public))
    {
        Try
        {
            #PS2 compatibility
            if($import.fullname)
            {
                . $import.fullname
            }
        }
        Catch
        {
            Write-Error "Failed to import function $($import.fullname): $_"
        }
    }
    
#Create some aliases, export public functions
    Export-ModuleMember -Function $($Public | Select -ExpandProperty BaseName)