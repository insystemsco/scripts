if ($env:PROCESSOR_ARCHITEW6432 -eq "AMD64") {

  $x64PS=join-path $PSHome.tolower().replace("syswow64","sysnative").replace("system32","sysnative") powershell.exe

  $cmd = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($myinvocation.MyCommand.Definition))

  $out = & "$x64PS" -NonInteractive -NoProfile -ExecutionPolicy Bypass -EncodedCommand $cmd

  $out

  exit $lastexitcode

}

 

# Write Script Below

Write-Output ("Running Context: " + $env:PROCESSOR_ARCHITECTURE)
<# 

.DESCRIPTION 
 This script enables many protection capabilities of Windows Defender Antivirus. These settings are not best practices or recommended settings for every organization, and should be used only when comparing Windows Defender AV or other 3rd party antimalware engines, not in production environments. 

#> 

Param()


<#  
.SYNOPSIS  
    This script sets Windows Defender AV to enable most features for the evaluation of protection capabilities in Windows 10 using the Windows Defender AV cmdlets, described at https://technet.microsoft.com/en-us/library/dn433280.aspx
#>

##  
# Start of Script  
##


# =================================================================================================
#                                              Functions
# =================================================================================================

# Verifies that the script is running as admin
function Check-IsElevated
{
    $id = [System.Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object System.Security.Principal.WindowsPrincipal($id)

    if ($p.IsInRole([System.Security.Principal.WindowsBuiltInRole]::Administrator))
    {
        Write-Output $true
    }            
    else
    {
        Write-Output $false
    }       
}

# Verifies that script is running on Windows 10 or greater
function Check-IsWindows10
{
    if ([System.Environment]::OSVersion.Version.Major -ge "10") 
    {
        Write-Output $true
    }
    else
    {
        Write-Output $false
    }
}

# Verifies that script is running on Windows 10 1709 or greater
function Check-IsWindows10-1709
{
    if ([System.Environment]::OSVersion.Version.Minor -ge "16299") 
    {
        Write-Output $true
    }
    else
    {
        Write-Output $false
    }
}

function SetRegistryKey([string]$key, [int]$value)
{
    #Editing Windows Defender settings AV via registry is NOT supported. This is a scripting workaround instead of using Group Policy or SCCM for Windows 10 version 1703
    $amRegistryPath = "HKLM:\Software\Policies\Microsoft\Microsoft Antimalware\MpEngine"
    $wdRegistryPath = "HKLM:\Software\Policies\Microsoft\Windows Defender\MpEngine"
    $regPathToUse = $wdRegistryPath #Default to WD path
    if (Test-Path $amRegistryPath)
    {
        $regPathToUse = $amRegistryPath
    }
    New-ItemProperty -Path $regPathToUse -Name $key -Value $value -PropertyType DWORD -Force | Out-Null
} 

# =================================================================================================
#                                              Main
# =================================================================================================
$scriptDate = Get-Date "6/13/2018"
$currentDate = Get-Date

if (!(Check-IsElevated))
{
    throw "Please run this script from an elevated PowerShell prompt"            
}

if (!(Check-IsWindows10))
{
    throw "Please run this script on Windows 10"            
}


Write-Host "This script helps configure Windows Defender Antivirus and Windows Defender Exploit Guard in order to evaluate its protection capabilities. `nFor more information see the Windows Defender protection evaluation guide (https://aka.ms/evaluatewdav)`nSome of these settings are set using unsupported methods, you should consult Windows Defender AV documentation for proper configuration methods at https://aka.ms/wdavdocs"
Write-Host "`nUpdating Windows Defender AV settings`n" -ForegroundColor Green 

"Enable real-time monitoring"
Set-MpPreference -DisableRealtimeMonitoring 0

"Enable cloud-deliveredprotection"
Set-MpPreference -MAPSReporting Advanced

"Enable sample submission"
Set-MpPreference -SubmitSamplesConsent Always

"Enable checking signatures before scanning"
Set-MpPreference -CheckForSignaturesBeforeRunningScan 1

"Enable behavior monitoring"
Set-MpPreference -DisableBehaviorMonitoring 0

"Enable IOAV protection"
Set-MpPreference -DisableIOAVProtection 0

"Enable script scanning"
Set-MpPreference -DisableScriptScanning 0

"Enable removable drive scanning"
Set-MpPreference -DisableRemovableDriveScanning 0

"Enable Block at first sight"
Set-MpPreference -DisableBlockAtFirstSeen 0

"Enable potentially unwanted apps"
Set-MpPreference -PUAProtection Enabled

"Schedule signature updates every 8 hours"
Set-MpPreference -SignatureUpdateInterval 8

"Enable archive scanning"
Set-MpPreference -DisableArchiveScanning 0

"Enable email scanning"
Set-MpPreference -DisableEmailScanning 0

if (!(Check-IsWindows10-1709))
{
    "Set cloud block level to 'High'"
    Set-MpPreference -CloudBlockLevel High

    "Set cloud block timeout to 1 minute"
    Set-MpPreference -CloudExtendedTimeout 50

    Write-Host "`nUpdating Windows Defender Exploit Guard settings`n" -ForegroundColor Green 

    Write-Host "Enabling Controlled Folder Access and setting to block mode"
    Set-MpPreference -EnableControlledFolderAccess Enabled 

    Write-Host "Enabling Network Protection and setting to block mode"
    Set-MpPreference -EnableNetworkProtection Enabled

    Write-Host "Enabling Exploit Guard ASR rules and setting to block mode. Some of these may block behavior that is acceptable in your organization, in this case please disable those specific rules. Learn more: https://docs.microsoft.com/en-us/windows/security/threat-protection/windows-defender-exploit-guard/attack-surface-reduction-exploit-guard"
    Add-MpPreference -AttackSurfaceReductionRules_Ids 75668C1F-73B5-4CF0-BB93-3ECF5CB7CC84 -AttackSurfaceReductionRules_Actions Enabled
    Add-MpPreference -AttackSurfaceReductionRules_Ids 3B576869-A4EC-4529-8536-B80A7769E899 -AttackSurfaceReductionRules_Actions Enabled
    Add-MpPreference -AttackSurfaceReductionRules_Ids D4F940AB-401B-4EfC-AADC-AD5F3C50688A -AttackSurfaceReductionRules_Actions Enabled
    Add-MpPreference -AttackSurfaceReductionRules_Ids D3E037E1-3EB8-44C8-A917-57927947596D -AttackSurfaceReductionRules_Actions Enabled
    Add-MpPreference -AttackSurfaceReductionRules_Ids 5BEB7EFE-FD9A-4556-801D-275E5FFC04CC -AttackSurfaceReductionRules_Actions Enabled
    Add-MpPreference -AttackSurfaceReductionRules_Ids BE9BA2D9-53EA-4CDC-84E5-9B1EEEE46550 -AttackSurfaceReductionRules_Actions Enabled
    Add-MpPreference -AttackSurfaceReductionRules_Ids 92E97FA1-2EDF-4476-BDD6-9DD0B4DDDC7B -AttackSurfaceReductionRules_Actions Enabled
    Add-MpPreference -AttackSurfaceReductionRules_Ids D1E49AAC-8F56-4280-B9BA-993A6D77406C -AttackSurfaceReductionRules_Actions Enabled
    Add-MpPreference -AttackSurfaceReductionRules_Ids B2B3F03D-6A65-4F7B-A9C7-1C7EF74A9BA4 -AttackSurfaceReductionRules_Actions Enabled
    Add-MpPreference -AttackSurfaceReductionRules_Ids C1DB55AB-C21A-4637-BB3F-A12568109D35 -AttackSurfaceReductionRules_Actions Enabled
    Add-MpPreference -AttackSurfaceReductionRules_Ids 01443614-CD74-433A-B99E-2ECDC07BFC25 -AttackSurfaceReductionRules_Actions Enabled

    if ($false -eq (Test-Path ProcessMitigation.xml))
    {
        Write-Host "Downloading Process Mitigation file from https://demo.wd.microsoft.com/Content/ProcessMitigation.xml"
        $url = 'https://demo.wd.microsoft.com/Content/ProcessMitigation.xml'
        Invoke-WebRequest $url -OutFile ProcessMitigation.xml
    }

    Write-Host "Enabling Exploit Protection"
    Set-ProcessMitigation -PolicyFilePath ProcessMitigation.xml

}

else
{
    ## Workaround for Windows 10 version 1703
    "Set cloud block level to 'High'"
    SetRegistryKey -key MpCloudBlockLevel -value 2

    "Set cloud block timeout to 1 minute"
    SetRegistryKey -key MpBafsExtendedTimeout -value 50
}
Write-Host "`nSettings update complete"  -ForegroundColor Green

Write-Host "`nOutput Windows Defender AV settings status"  -ForegroundColor Green
Get-MpPreference

if ($scriptDate.AddDays(180) -lt $currentDate)
{
    Write-Host "`nThis script is older than 180 days and there may be an updated version located here: https://aka.ms/wdavevalscript`n" -ForegroundColor yellow        
}


exit 0



#https://technet.microsoft.com/en-us/library/dn433280.aspx
#Set-MpPreference Options

#[-ExclusionPath <string[]>] 
#[-ExclusionExtension <string[]>] 
#[-ExclusionProcess <string[]>] 
#[-RealTimeScanDirection {Both | Incoming | Outcoming}] 
#[-QuarantinePurgeItemsAfterDelay <uint32>] 
#[-RemediationScheduleDay {Everyday | Sunday |  Monday | Tuesday | Wednesday | Thursday | Friday | Saturday | Never}] 
#[-RemediationScheduleTime <datetime>] 
#[-ReportingAdditionalActionTimeOut <uint32>] 
#[-ReportingCriticalFailureTimeOut <uint32>] 
#[-ReportingNonCriticalTimeOut <uint32>] 
#[-ScanAvgCPULoadFactor <byte>] 
#[-CheckForSignaturesBeforeRunningScan <bool>] 
#[-ScanPurgeItemsAfterDelay <uint32>] 
#[-ScanOnlyIfIdleEnabled <bool>] 
#[-ScanParameters {QuickScan | FullScan}] 
#[-ScanScheduleDay {Everyday | Sunday | Monday | Tuesday | Wednesday | Thursday | Friday | Saturday | Never}] 
#[-ScanScheduleQuickScanTime <datetime>] 
#[-ScanScheduleTime <datetime>] 
#[-SignatureFirstAuGracePeriod <uint32>] 
#[-SignatureAuGracePeriod <uint32>] 
#[-SignatureDefinitionUpdateFileSharesSources <string>] 
#[-SignatureDisableUpdateOnStartupWithoutEngine <bool>] 
#[-SignatureFallbackOrder <string>] 
#[-SignatureScheduleDay {Everyday | Sunday | Monday | Tuesday | Wednesday | Thursday | Friday | Saturday | Never}] 
#[-SignatureScheduleTime <datetime>] 
#[-SignatureUpdateCatchupInterval <uint32>] 
#[-SignatureUpdateInterval <uint32>] 
#[-MAPSReporting {Disabled | Basic | Advanced}] 
#[-SubmitSamplesConsent {None | Always | Never}] 
#[-DisableAutoExclusions <bool>] 
#[-DisablePrivacyMode <bool>] 
#[-RandomizeScheduleTaskTimes <bool>] 
#[-DisableBehaviorMonitoring <bool>] 
#[-DisableIntrusionPreventionSystem <bool>] 
#[-DisableIOAVProtection <bool>] 
#[-DisableRealtimeMonitoring <bool>] 
#[-DisableScriptScanning <bool>] 
#[-DisableArchiveScanning <bool>] 
#[-DisableCatchupFullScan <bool>] 
#[-DisableCatchupQuickScan <bool>] 
#[-DisableEmailScanning <bool>] 
#[-DisableRemovableDriveScanning <bool>] 
#[-DisableRestorePoint <bool>] 
#[-DisableScanningMappedNetworkDrivesForFullScan <bool>] 
#[-DisableScanningNetworkFiles <bool>] 
#[-UILockdown <bool>] 
#[-ThreatIDDefaultAction_Ids <long[]>] 
#[-ThreatIDDefaultAction_Actions {Clean | Quarantine | Remove | Allow | UserDefined | NoAction | Block}] 
#[-UnknownThreatDefaultAction {Clean | Quarantine | Remove | Allow | UserDefined | NoAction | Block}] 
#[-LowThreatDefaultAction {Clean | Quarantine | Remove | Allow | UserDefined | NoAction | Block}] 
#[-ModerateThreatDefaultAction {Clean | Quarantine | Remove | Allow | UserDefined | NoAction | Block}] 
#[-HighThreatDefaultAction {Clean | Quarantine | Remove | Allow | UserDefined | NoAction | Block}] 
#[-SevereThreatDefaultAction {Clean | Quarantine | Remove | Allow | UserDefined | NoAction | Block}] 
#[-Force] 
#[-DisableBlockAtFirstSeen <bool>] 
#[-PUAProtection {Disabled | Enabled | AuditMode}] 
#[-CimSession <CimSession[]>] 
#[-ThrottleLimit <int>] [-AsJob]  [<CommonParameters>]


exit 0
