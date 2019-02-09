<# 

.DESCRIPTION 
 Installs RSAT Tools for Windows 10 

#>
#Requires -Version 5.0 -RunAsAdministrator
[Cmdletbinding()]
param()

    $VerbosePreference = 'Continue'

    $x86 = 'https://download.microsoft.com/download/1/D/8/1D8B5022-5477-4B9A-8104-6A71FF9D98AB/WindowsTH-RSAT_WS2016-x86.msu'
    $x64 = 'https://download.microsoft.com/download/1/D/8/1D8B5022-5477-4B9A-8104-6A71FF9D98AB/WindowsTH-RSAT_WS2016-x64.msu'

    switch ($env:PROCESSOR_ARCHITECTURE)
    {
        'x86' {$version = $x86}
        'AMD64' {$version = $x64}
    }

    Write-Verbose -Message "OS Version is $env:PROCESSOR_ARCHITECTURE"
    Write-Verbose -Message "Now Downloading RSAT Tools installer"

    $Filename = $version.Split('/')[-1]
    Invoke-WebRequest -Uri $version -UseBasicParsing -OutFile "$env:TEMP\$Filename" 
    
    Write-Verbose -Message "Starting the Windows Update Service to install the RSAT Tools "
    
    Start-Process -FilePath wusa.exe -ArgumentList "$env:TEMP\$Filename /quiet" -Wait -Verbose
    
    Write-Verbose -Message "RSAT Tools are now be installed"
    
    Remove-Item "$env:TEMP\$Filename" -Verbose
    
    Write-Verbose -Message "Script Cleanup complete"
    
    Write-Verbose -Message "Remote Administration"


       