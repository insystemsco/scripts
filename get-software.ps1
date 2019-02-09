function Get-Software
{
    param
    (
        [Parameter(Position=0, Mandatory = $false, HelpMessage="DisplayName", ValueFromPipeline = $true)]
        $DisplayName='*',
        [Parameter(Position=0, Mandatory = $false, HelpMessage="Version", ValueFromPipeline = $true)]
        $Version='*',
        [Parameter(Position=0, Mandatory = $false, HelpMessage="Installation date", ValueFromPipeline = $true)]
        $InstallationDate='*',
        [Parameter(Position=0, Mandatory = $false, HelpMessage="How to uninstall", ValueFromPipeline = $true)]
        $UninstallationPath='*'
 
    )
     
    $UninstallationPathAllProfiles = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    $UninstallationPathCurrentPrfoile = "Registry::HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*"
    $UninstallationPathAllProfilesWOW64 = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    $UninstallationPathCurrentPrfoileWOW64 = "Registry::HKEY_CURRENT_USER\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
 
    $ResultArray = Get-ItemProperty -Path $UninstallationPathAllProfiles, $UninstallationPathCurrentPrfoile, $UninstallationPathAllProfilesWOW64, $UninstallationPathAllProfilesWOW64 | `
    Select-Object -Property DisplayVersion, DisplayName, InstallDate, @{Name="Size in MB";Expression={[Math]::Round($_.EstimatedSize / 1024)}} ,UninstallString
    $ResultArray | Where-Object {$_.DisplayName -ne $null -and $_.DisplayName -like $DisplayName -and $_.DisplayVersion -like $Version `
    -and $_.UninstallString -like $UninstallationPath -and $_.InstallDate -like $InstallationDate}
     
    return $ResultArray
}