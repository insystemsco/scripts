# Display Demo Machine Information
Get-ComputerInfo | `
    Select-Object `
        @{N='Hostname';E={$env:COMPUTERNAME}}, `
        WindowsProductName, `
        WindowsCurrentVersion, `
        WindowsVersion, `
        WindowsBuildLabEx | `
            Format-Table `
                -AutoSize ;