@{
    Psake = @{
        Target = 'C:\temp'
        Parameters = @{
                Force = $True
                Import = $True
        }
        AddToPath = $True
    }
    PSDeploy = @{
        Target = 'C:\temp'
    }
    Pester = @{
        Target = 'C:\temp'
    }
    BuildHelpers = @{
        Target = 'C:\temp'
    }
}