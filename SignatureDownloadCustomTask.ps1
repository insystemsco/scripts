<# 

.DESCRIPTION
   It is crucial to keep antimalware definitions up-to-date to maintain optimal protection. Configuring VMs and VM hosts to stay up-to-date with antimalware definitions correctly can help
   avoid increased network bandwidth consumption and out of date AM definitions; misconfigurations can result in decreased usability of the VM and/or decreased protection from malware.

   This script simplifies the setting up of antimalware definitions for VMs and VM hosts. It allows VMs that don't have Internet connectivity or Windows Update (WU) connectivity to have
   up-to-date definitions.

   It does this by enabling the VM to pull from a UNC share, which is updated by the VM-host. This also results in overall network bandwidth savings as the antimalware definitions are
   downloaded only once for the VMs.

   The admin must select always run option or import the certificate and add the certificate to trusted publisher list.
   
   If you get an error trying to run 'create task', it is likely because the command line arguments tip over the max limit. Please place the script in a shortened path location and try again.

#> 


<#  
.SYNOPSIS  
    This script allows admins to create scheduled tasks on a Server host that would download the signatures regularly onto a specified share.
    The VMs can then be configured via a policy (not part of this script) to pick up the signatures from the share instead of directly downloading them from ADL or WU.
.NOTES  
    File Name  : SignatureDownloadCustomTask.ps1
    Author     : Windows Defender Team
    Requires   : PowerShell V3
    Note       : (1) Messages from this script will be logged to %windir%\Temp\DefenderSignatureDownloadTask.log.
                 (2) To download all the definitions, a customer must create 4 tasks:
                        X86 delta
                        X86 full
                        X64 delta
                        X64 full
                     We create the tasks separately so that admins can control the frequency of the delta versus the full downloads.
                     NIS sigs are generally small compared to the others, and hence are downloaded with all of these tasks.
                  (3) Ideally, the root destination folder should be the same for all the tasks with x86/x64 sub-folders so that it is easier to configure the UNC signature
                      source on the VMs.
                  (4) Requires admin rights to run.

.IMPORTANT-REMARKS
               : -scriptPath. Please make sure that this is a protected path. Otherwise a non-admin could replace the script file and have the task scheduler run it for them.
                              The task does however launch powershell with a flag to enforce running signed script only.
                              Also, the length of this cannot be greater than 150 characters.
               : -hoursInterval. If this is passed, the task will be created for hour-based intervals instead of day-based intervals.
.EXAMPLES  
    To create a task that downloads delta x86 signatures once every 2 days to a directory called D:\Share\Test, and assuming that this scripts resides in C:\Windows\Protected
        SignatureDownloadCustomTask.ps1 -action create -arch x86 -isDelta $true -destDir D:\Share\Test -scriptPath C:\Windows\Protected\SignatureDownloadCustomTask.ps1 -daysInterval 2
    To delete the above task
        SignatureDownloadCustomTask.ps1 -action delete -arch x86 -isDelta $true
    To run the task manually [JFYI, the scheduled task should take care of it anyways]
        SignatureDownloadCustomTask.ps1 -action run -arch x86 -isDelta $true -destDir D:\Share\Test
#>

param
(
    [parameter(Position=0, Mandatory=$true, HelpMessage="Action to do.")]
    [ValidateSet("create","delete","run")] 
    [string]$action,

    [parameter(Position=1, Mandatory=$true, HelpMessage="Architecture of the required signature package.")]
    [ValidateSet("x86","x64","arm")] 
    [string]$arch,

    [parameter(Position=2, Mandatory=$true, HelpMessage="False (0) - Task deals with full signature package, True (1) - Task deals with delta signature package")]
    [ValidateRange($False,$True)]
    [bool]$isDelta,

    [parameter(Mandatory=$false, HelpMessage="The destination directory where the sigs will be downloaded to.")]
    [string]$destDir,

    [parameter(Mandatory=$false, HelpMessage="The full path to where this script file resides.")]
    [string]$scriptPath,

    [parameter(Mandatory=$false, HelpMessage="The frequency desired for this task (in number of days).")]
    [ValidateRange(1, 365)]
    [int]$daysInterval = 1,

    [parameter(Mandatory=$false, HelpMessage="The frequency desired for this task if it has to be run multiple times per day (in number of hours).")]
    [ValidateRange(1, 23)]
    [int]$hoursInterval = 0
)

# Flushes the log file if it is bigger than 100 KB.
Function FlushIfTooBig-LogFile()
{
    [string]$path = Join-Path ($env:windir) "TEMP\DefenderSignatureDownloadTask.log"
    if (Test-Path $path)
    {
        $file = Get-Item $path    
        if ($file.Length -gt 100KB)
        {
            '' | Out-File $file
        }
    }
    else
    {
        New-Item $path -type file
        Write-Output 'Log file created.'
    }
}

# Appends a message to %windir%\Temp\DefenderSignatureDownloadTask.log with the time stamp.
Function Log-Message([string]$message)
{
    Write-Output $message
    [string]$path = Join-Path ($env:windir) "TEMP\DefenderSignatureDownloadTask.log"
    $date = Get-Date
    $date | Out-File $path -Append
    $message | Out-File $path -Append
    '----------End of message----------' | Out-File $path -Append
}

# Downloads file from a given URL.
Function Download-File([string]$url, [string]$targetFile)
{
    [System.Net.WebClient]$webClient = New-Object -TypeName System.Net.WebClient
    [System.Uri]$uri = New-Object -TypeName System.Uri -ArgumentList $url
    $webClient.DownloadFile($uri, $targetFile)
}

# Gets the specified registry value related to signatures.
Function Get-SignatureRegistryValue([string]$name) 
{
    [string]$path
    if ((Test-Path -Path 'HKLM:\SOFTWARE\Microsoft\Microsoft Antimalware') -And ([System.Environment]::OSVersion.Version.Major -lt 10))
    {
        $path = 'HKLM:\SOFTWARE\Microsoft\Microsoft Antimalware\Signature Updates'
    }
    else
    {
        $path = 'HKLM:\SOFTWARE\Microsoft\Windows Defender\Signature Updates'
    }

    $key = Get-Item -LiteralPath $path
    return $key.GetValue($name, $null)
}

# Gets the URL to download AM delta sigs. We use the hosts' AM engine and base sig version for the URL.
Function Get-AmDeltaSigUrl([string]$arch)
{
    [string]$engineVersionValue = Get-SignatureRegistryValue 'EngineVersion'
    [string]$avVersionValue = Get-SignatureRegistryValue 'AVSignatureBaseVersion'
    [string]$asVersionValue = Get-SignatureRegistryValue 'ASSignatureBaseVersion'
    [string]$deltaSigPath = [string]::Format("http://go.microsoft.com/fwlink/?LinkID=121721&clcid=0x409&arch={0}&eng={1}&avdelta={2}&asdelta={3}",
                                             $arch.Trim(),
                                             $engineVersionValue.Trim(),
                                             $avVersionValue.Trim(),
                                             $asVersionValue.Trim())
    return $deltaSigPath
}

# Gets the URL to download AM full sigs.
Function Get-AmFullSigUrl([string]$arch)
{
    [string]$fullSigPath = [string]::Format("http://go.microsoft.com/fwlink/?LinkID=121721&clcid=0x409&arch={0}", $arch.Trim())
    return $fullSigPath
}

# Gets the URL to download NIS sigs.
Function Get-NisSigUrl([string]$arch)
{
    [string]$nisSigPath = [string]::Format("http://go.microsoft.com/fwlink/?LinkID=260974&clcid=0x409&NRI=true&arch={0}", $arch.Trim())
    return $nisSigPath   
}

# Downloads AM and NIS sigs.
Function Run-Task([string]$arch, [bool]$isDelta, [string]$destDir)
{
    # Download AM sigs.
    [string]$amSigFileName
    [string]$url
    [string]$packageType
    if ($isDelta)
    {
        $packageType = 'Delta'
        $amSigFileName = 'mpam-d.exe'
        $url = Get-AmDeltaSigUrl $arch
    }
    else
    {
        $packageType = 'Full'
        $amSigFileName = 'mpam-fe.exe'
        $url = Get-AmFullSigUrl $arch
    }

    [string]$fullDestPath = Join-Path $destDir $amSigFileName
    Download-File $url $fullDestPath # Please note that Download-File can throw an error if ADL has no file to offer for the config we pass.
    [string]$message = 'Successfully ran task to download ' + $packageType + ' AM sigs of arch ' + $arch + ' to ' + $fullDestPath
    Log-Message $message

    # Download NIS sigs.
    [string]$nisSigFileName = 'nis_full.exe'
    [string]$fullNisDestPath = Join-Path $destDir $nisSigFileName
    [string]$nisUrl = Get-NisSigUrl $arch
    Download-File $nisUrl $fullNisDestPath
    [string]$nisMessage = 'Successfully ran task to download NIS sigs of arch ' + $arch + ' to ' + $fullNisDestPath
    Log-Message $nisMessage    
}

# Gets the name of the signature download scheduled task, based on arch and package type (delta/full).
Function Get-TaskName([string]$arch, [bool]$isDelta)
{
    [string]$packageType
    if ($isDelta)
    {
        $packageType = 'Delta'
    }
    else
    {
        $packageType = 'Full'
    }

    [string]$taskName = 'Windows Defender Custom Task for ' + $arch + $packageType + ' Signature Download'
    [string]$taskFullName = [string]::Format("\Microsoft\Windows\Windows Defender\{0}", $taskName)

    return $taskFullName
}

# Creates a scheduled task that would download the signatures regularly onto the specified share.
Function Create-Task([string]$scriptPath, [string]$arch, [bool]$isDelta, [string]$destDir, [int]$daysInterval, [int]$hoursInterval)
{
    [string]$schedule
    [int]$interval = 1
    if ($hoursInterval -eq 0)    # if this is not set, then use day-based interval for the scheduled task. Otherwise if it is set, use hour-based interval.
    {
        $schedule = 'DAILY'
        $interval = $daysInterval
    }
    else
    {
        $schedule = 'HOURLY'
        $interval = $hoursInterval
    }

    [string]$scriptCommand = '\\\"' + $scriptPath + '\\\"' + ' -action run' + ' -arch ' + $arch + ' -isDelta $' + $isDelta + ' -destDir ' + $destDir        
    [string]$schTasksPath = Join-Path ($env:windir) "system32\schtasks.exe"    # Needs to run downlevel as well; hence not using the Win8.1+ specific ScheduleTask* helpers.
    [string]$taskFullName = Get-TaskName $arch $isDelta
    $taskFullName = $taskFullName.Trim()

    [string]$fullPathToPowerShellExe = Join-Path $PsHome powershell.exe
    [string]$taskProgramArg = $fullPathToPowerShellExe + ' -Noprofile -ExecutionPolicy AllSigned -Command ' + '\"& ' + $scriptCommand + '\"'
    $schTasksProcess = Start-Process -FilePath "$schTasksPath" -ArgumentList "/create /tn ""$taskFullName"" /tr ""$taskProgramArg"" /ru ""NT AUTHORITY\SYSTEM"" /sc ""$schedule"" /MO ""$interval"" /NP /F /RL ""LIMITED""" -Wait -PassThru -NoNewWindow
    if ($schTasksProcess.ExitCode -ne 0)
    {
        throw "Error: Create task failed. Check Event log for more details."
    }

    [string]$message = 'Successfully created task ' + $taskFullName
    Log-Message $message
}

# Deletes specified scheduled task.
Function Delete-Task([string]$arch, [bool]$isDelta)
{
    [string]$schTasksPath = Join-Path ($env:windir) "system32\schtasks.exe"
    [string]$taskFullName = Get-TaskName $arch $isDelta
    $taskFullName = $taskFullName.Trim()

    $schTasksProcess = Start-Process -FilePath "$schTasksPath" -ArgumentList "/delete /tn ""$taskFullName"" /f" -Wait -PassThru -NoNewWindow
    if ($schTasksProcess.ExitCode -ne 0)
    {
        throw "Error: Create task failed. Check Event log for more details."
    }

    [string]$message = 'Successfully deleted task ' + $taskFullName
    Log-Message $message
}

# Main
try
{
    FlushIfTooBig-LogFile

    Log-Message 'Script started.'

    # Some additional parameter validation for specific actions.
    if (($action.ToLower() -eq 'create') -OR ($action.ToLower() -eq 'run'))
    {
        if (-NOT (Test-Path $destDir -PathType 'Container')) 
        { 
           throw "$($destDir) is not a valid folder" 
        }
    }

    if ($action.ToLower() -eq 'create')
    {
        if (-NOT (Test-Path $scriptPath -PathType 'Leaf')) 
        { 
           throw "$($scriptPath) is not a valid file" 
        }
    }

    # Execute specified action.
    if ($action.ToLower() -eq 'create')
    {
        Create-Task $scriptPath $arch $isDelta $destDir $daysInterval $hoursInterval
    }
    elseif ($action.ToLower() -eq 'delete')
    {
        Delete-Task $arch $isDelta
    }
    else # ($action.ToLower() -eq 'run')
    {
        Run-Task $arch $isDelta $destDir
    }

    Log-Message 'Script completed.'
}
catch [System.Exception]
{
    Log-Message $Error
}
# SIG # Begin signature block
# MIIaxQYJKoZIhvcNAQcCoIIatjCCGrICAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU+HAz8sbk9HTdBG5NfYq9/W23
# nGigghWCMIIEwzCCA6ugAwIBAgITMwAAAJvgdDfLPU2NLgAAAAAAmzANBgkqhkiG
# 9w0BAQUFADB3MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4G
# A1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEw
# HwYDVQQDExhNaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EwHhcNMTYwMzMwMTkyMTI5
# WhcNMTcwNjMwMTkyMTI5WjCBszELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjENMAsGA1UECxMETU9QUjEnMCUGA1UECxMebkNpcGhlciBEU0UgRVNO
# OjcyOEQtQzQ1Ri1GOUVCMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBT
# ZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8AMIIBCgKCAQEAjaPiz4GL18u/
# A6Jg9jtt4tQYsDcF1Y02nA5zzk1/ohCyfEN7LBhXvKynpoZ9eaG13jJm+Y78IM2r
# c3fPd51vYJxrePPFram9W0wrVapSgEFDQWaZpfAwaIa6DyFyH8N1P5J2wQDXmSyo
# WT/BYpFtCfbO0yK6LQCfZstT0cpWOlhMIbKFo5hljMeJSkVYe6tTQJ+MarIFxf4e
# 4v8Koaii28shjXyVMN4xF4oN6V/MQnDKpBUUboQPwsL9bAJMk7FMts627OK1zZoa
# EPVI5VcQd+qB3V+EQjJwRMnKvLD790g52GB1Sa2zv2h0LpQOHL7BcHJ0EA7M22tQ
# HzHqNPpsPQIDAQABo4IBCTCCAQUwHQYDVR0OBBYEFJaVsZ4TU7pYIUY04nzHOUps
# IPB3MB8GA1UdIwQYMBaAFCM0+NlSRnAK7UD7dvuzK7DDNbMPMFQGA1UdHwRNMEsw
# SaBHoEWGQ2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3Rz
# L01pY3Jvc29mdFRpbWVTdGFtcFBDQS5jcmwwWAYIKwYBBQUHAQEETDBKMEgGCCsG
# AQUFBzAChjxodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY3Jv
# c29mdFRpbWVTdGFtcFBDQS5jcnQwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZI
# hvcNAQEFBQADggEBACEds1PpO0aBofoqE+NaICS6dqU7tnfIkXIE1ur+0psiL5MI
# orBu7wKluVZe/WX2jRJ96ifeP6C4LjMy15ZaP8N0OckPqba62v4QaM+I/Y8g3rKx
# 1l0okye3wgekRyVlu1LVcU0paegLUMeMlZagXqw3OQLVXvNUKHlx2xfDQ/zNaiv5
# DzlARHwsaMjSgeiZIqsgVubk7ySGm2ZWTjvi7rhk9+WfynUK7nyWn1nhrKC31mm9
# QibS9aWHUgHsKX77BbTm2Jd8E4BxNV+TJufkX3SVcXwDjbUfdfWitmE97sRsiV5k
# BH8pS2zUSOpKSkzngm61Or9XJhHIeIDVgM0Ou2QwggTsMIID1KADAgECAhMzAAAB
# Cix5rtd5e6asAAEAAAEKMA0GCSqGSIb3DQEBBQUAMHkxCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xIzAhBgNVBAMTGk1pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBMB4XDTE1MDYwNDE3NDI0NVoXDTE2MDkwNDE3NDI0NVowgYMxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
# MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xDTALBgNVBAsTBE1PUFIx
# HjAcBgNVBAMTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjCCASIwDQYJKoZIhvcNAQEB
# BQADggEPADCCAQoCggEBAJL8bza74QO5KNZG0aJhuqVG+2MWPi75R9LH7O3HmbEm
# UXW92swPBhQRpGwZnsBfTVSJ5E1Q2I3NoWGldxOaHKftDXT3p1Z56Cj3U9KxemPg
# 9ZSXt+zZR/hsPfMliLO8CsUEp458hUh2HGFGqhnEemKLwcI1qvtYb8VjC5NJMIEb
# e99/fE+0R21feByvtveWE1LvudFNOeVz3khOPBSqlw05zItR4VzRO/COZ+owYKlN
# Wp1DvdsjusAP10sQnZxN8FGihKrknKc91qPvChhIqPqxTqWYDku/8BTzAMiwSNZb
# /jjXiREtBbpDAk8iAJYlrX01boRoqyAYOCj+HKIQsaUCAwEAAaOCAWAwggFcMBMG
# A1UdJQQMMAoGCCsGAQUFBwMDMB0GA1UdDgQWBBSJ/gox6ibN5m3HkZG5lIyiGGE3
# NDBRBgNVHREESjBIpEYwRDENMAsGA1UECxMETU9QUjEzMDEGA1UEBRMqMzE1OTUr
# MDQwNzkzNTAtMTZmYS00YzYwLWI2YmYtOWQyYjFjZDA1OTg0MB8GA1UdIwQYMBaA
# FMsR6MrStBZYAck3LjMWFrlMmgofMFYGA1UdHwRPME0wS6BJoEeGRWh0dHA6Ly9j
# cmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY0NvZFNpZ1BDQV8w
# OC0zMS0yMDEwLmNybDBaBggrBgEFBQcBAQROMEwwSgYIKwYBBQUHMAKGPmh0dHA6
# Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljQ29kU2lnUENBXzA4LTMx
# LTIwMTAuY3J0MA0GCSqGSIb3DQEBBQUAA4IBAQCmqFOR3zsB/mFdBlrrZvAM2PfZ
# hNMAUQ4Q0aTRFyjnjDM4K9hDxgOLdeszkvSp4mf9AtulHU5DRV0bSePgTxbwfo/w
# iBHKgq2k+6apX/WXYMh7xL98m2ntH4LB8c2OeEti9dcNHNdTEtaWUu81vRmOoECT
# oQqlLRacwkZ0COvb9NilSTZUEhFVA7N7FvtH/vto/MBFXOI/Enkzou+Cxd5AGQfu
# FcUKm1kFQanQl56BngNb/ErjGi4FrFBHL4z6edgeIPgF+ylrGBT6cgS3C6eaZOwR
# XU9FSY0pGi370LYJU180lOAWxLnqczXoV+/h6xbDGMcGszvPYYTitkSJlKOGMIIF
# vDCCA6SgAwIBAgIKYTMmGgAAAAAAMTANBgkqhkiG9w0BAQUFADBfMRMwEQYKCZIm
# iZPyLGQBGRYDY29tMRkwFwYKCZImiZPyLGQBGRYJbWljcm9zb2Z0MS0wKwYDVQQD
# EyRNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkwHhcNMTAwODMx
# MjIxOTMyWhcNMjAwODMxMjIyOTMyWjB5MQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMSMwIQYDVQQDExpNaWNyb3NvZnQgQ29kZSBTaWduaW5nIFBD
# QTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBALJyWVwZMGS/HZpgICBC
# mXZTbD4b1m/My/Hqa/6XFhDg3zp0gxq3L6Ay7P/ewkJOI9VyANs1VwqJyq4gSfTw
# aKxNS42lvXlLcZtHB9r9Jd+ddYjPqnNEf9eB2/O98jakyVxF3K+tPeAoaJcap6Vy
# c1bxF5Tk/TWUcqDWdl8ed0WDhTgW0HNbBbpnUo2lsmkv2hkL/pJ0KeJ2L1TdFDBZ
# +NKNYv3LyV9GMVC5JxPkQDDPcikQKCLHN049oDI9kM2hOAaFXE5WgigqBTK3S9dP
# Y+fSLWLxRT3nrAgA9kahntFbjCZT6HqqSvJGzzc8OJ60d1ylF56NyxGPVjzBrAlf
# A9MCAwEAAaOCAV4wggFaMA8GA1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFMsR6MrS
# tBZYAck3LjMWFrlMmgofMAsGA1UdDwQEAwIBhjASBgkrBgEEAYI3FQEEBQIDAQAB
# MCMGCSsGAQQBgjcVAgQWBBT90TFO0yaKleGYYDuoMW+mPLzYLTAZBgkrBgEEAYI3
# FAIEDB4KAFMAdQBiAEMAQTAfBgNVHSMEGDAWgBQOrIJgQFYnl+UlE/wq4QpTlVnk
# pDBQBgNVHR8ESTBHMEWgQ6BBhj9odHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtp
# L2NybC9wcm9kdWN0cy9taWNyb3NvZnRyb290Y2VydC5jcmwwVAYIKwYBBQUHAQEE
# SDBGMEQGCCsGAQUFBzAChjhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2Nl
# cnRzL01pY3Jvc29mdFJvb3RDZXJ0LmNydDANBgkqhkiG9w0BAQUFAAOCAgEAWTk+
# fyZGr+tvQLEytWrrDi9uqEn361917Uw7LddDrQv+y+ktMaMjzHxQmIAhXaw9L0y6
# oqhWnONwu7i0+Hm1SXL3PupBf8rhDBdpy6WcIC36C1DEVs0t40rSvHDnqA2iA6VW
# 4LiKS1fylUKc8fPv7uOGHzQ8uFaa8FMjhSqkghyT4pQHHfLiTviMocroE6WRTsgb
# 0o9ylSpxbZsa+BzwU9ZnzCL/XB3Nooy9J7J5Y1ZEolHN+emjWFbdmwJFRC9f9Nqu
# 1IIybvyklRPk62nnqaIsvsgrEA5ljpnb9aL6EiYJZTiU8XofSrvR4Vbo0HiWGFzJ
# NRZf3ZMdSY4tvq00RBzuEBUaAF3dNVshzpjHCe6FDoxPbQ4TTj18KUicctHzbMrB
# 7HCjV5JXfZSNoBtIA1r3z6NnCnSlNu0tLxfI5nI3EvRvsTxngvlSso0zFmUeDord
# EN5k9G/ORtTTF+l5xAS00/ss3x+KnqwK+xMnQK3k+eGpf0a7B2BHZWBATrBC7E7t
# s3Z52Ao0CW0cgDEf4g5U3eWh++VHEK1kmP9QFi58vwUheuKVQSdpw5OPlcmN2Jsh
# rg1cnPCiroZogwxqLbt2awAdlq3yFnv2FoMkuYjPaqhHMS+a3ONxPdcAfmJH0c6I
# ybgY+g5yjcGjPa8CQGr/aZuW4hCoELQ3UAjWwz0wggYHMIID76ADAgECAgphFmg0
# AAAAAAAcMA0GCSqGSIb3DQEBBQUAMF8xEzARBgoJkiaJk/IsZAEZFgNjb20xGTAX
# BgoJkiaJk/IsZAEZFgltaWNyb3NvZnQxLTArBgNVBAMTJE1pY3Jvc29mdCBSb290
# IENlcnRpZmljYXRlIEF1dGhvcml0eTAeFw0wNzA0MDMxMjUzMDlaFw0yMTA0MDMx
# MzAzMDlaMHcxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xITAf
# BgNVBAMTGE1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQTCCASIwDQYJKoZIhvcNAQEB
# BQADggEPADCCAQoCggEBAJ+hbLHf20iSKnxrLhnhveLjxZlRI1Ctzt0YTiQP7tGn
# 0UytdDAgEesH1VSVFUmUG0KSrphcMCbaAGvoe73siQcP9w4EmPCJzB/LMySHnfL0
# Zxws/HvniB3q506jocEjU8qN+kXPCdBer9CwQgSi+aZsk2fXKNxGU7CG0OUoRi4n
# rIZPVVIM5AMs+2qQkDBuh/NZMJ36ftaXs+ghl3740hPzCLdTbVK0RZCfSABKR2YR
# JylmqJfk0waBSqL5hKcRRxQJgp+E7VV4/gGaHVAIhQAQMEbtt94jRrvELVSfrx54
# QTF3zJvfO4OToWECtR0Nsfz3m7IBziJLVP/5BcPCIAsCAwEAAaOCAaswggGnMA8G
# A1UdEwEB/wQFMAMBAf8wHQYDVR0OBBYEFCM0+NlSRnAK7UD7dvuzK7DDNbMPMAsG
# A1UdDwQEAwIBhjAQBgkrBgEEAYI3FQEEAwIBADCBmAYDVR0jBIGQMIGNgBQOrIJg
# QFYnl+UlE/wq4QpTlVnkpKFjpGEwXzETMBEGCgmSJomT8ixkARkWA2NvbTEZMBcG
# CgmSJomT8ixkARkWCW1pY3Jvc29mdDEtMCsGA1UEAxMkTWljcm9zb2Z0IFJvb3Qg
# Q2VydGlmaWNhdGUgQXV0aG9yaXR5ghB5rRahSqClrUxzWPQHEy5lMFAGA1UdHwRJ
# MEcwRaBDoEGGP2h0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1
# Y3RzL21pY3Jvc29mdHJvb3RjZXJ0LmNybDBUBggrBgEFBQcBAQRIMEYwRAYIKwYB
# BQUHMAKGOGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljcm9z
# b2Z0Um9vdENlcnQuY3J0MBMGA1UdJQQMMAoGCCsGAQUFBwMIMA0GCSqGSIb3DQEB
# BQUAA4ICAQAQl4rDXANENt3ptK132855UU0BsS50cVttDBOrzr57j7gu1BKijG1i
# uFcCy04gE1CZ3XpA4le7r1iaHOEdAYasu3jyi9DsOwHu4r6PCgXIjUji8FMV3U+r
# kuTnjWrVgMHmlPIGL4UD6ZEqJCJw+/b85HiZLg33B+JwvBhOnY5rCnKVuKE5nGct
# xVEO6mJcPxaYiyA/4gcaMvnMMUp2MT0rcgvI6nA9/4UKE9/CCmGO8Ne4F+tOi3/F
# NSteo7/rvH0LQnvUU3Ih7jDKu3hlXFsBFwoUDtLaFJj1PLlmWLMtL+f5hYbMUVbo
# nXCUbKw5TNT2eb+qGHpiKe+imyk0BncaYsk9Hm0fgvALxyy7z0Oz5fnsfbXjpKh0
# NbhOxXEjEiZ2CzxSjHFaRkMUvLOzsE1nyJ9C/4B5IYCeFTBm6EISXhrIniIh0EPp
# K+m79EjMLNTYMoBMJipIJF9a6lbvpt6Znco6b72BJ3QGEe52Ib+bgsEnVLaxaj2J
# oXZhtG6hE6a/qkfwEm/9ijJssv7fUciMI8lmvZ0dhxJkAj0tr1mPuOQh5bWwymO0
# eFQF1EEuUKyUsKV4q7OglnUa2ZKHE3UiLzKoCG6gW4wlv6DvhMoh1useT8ma7kng
# 9wFlb4kLfchpyOZu6qeXzjEp/w7FW1zYTRuh2Povnj8uVRZryROj/TGCBK0wggSp
# AgEBMIGQMHkxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYD
# VQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xIzAh
# BgNVBAMTGk1pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBAhMzAAABCix5rtd5e6as
# AAEAAAEKMAkGBSsOAwIaBQCggcYwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQw
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFCmZ
# ueTpQcru4FxwkpKjZoMgfAFnMGYGCisGAQQBgjcCAQwxWDBWoDiANgBTAGkAZwBu
# AGEAdAB1AHIAZQBEAG8AdwBuAGwAbwBhAGQAQwB1AHMAdABvAG0AVABhAHMAa6Ea
# gBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20wDQYJKoZIhvcNAQEBBQAEggEAgvpF
# HSFy4b4DLy2ivILjoYbTWXNpbELdkKzGpwN3GzaRYYlwoF6WiAo12bD6Dn1XLvbA
# Y5FXMC/Y01tpK3ZBlLkFGsxx+luBqglqXlRzSEduD/CFbLNueHyGloN9Do5dkRyy
# 0dXYOQww6PuUZr/bz6qt4tDIUlon2zfCAAKTon8Jbs3/Xq/j3fyzP/tLgH13yBL2
# ZKQjn3vPHaoO0KnRd7sGX3nP2kr6Kti/wOIFkCzBjyY8uC7U2F9f4lDSjA//vlgy
# 8ROMivGVFopLNj4pya6vH3uK/wCtjtNjLHqWpHlY7leKJojQlhq9DLWryhDi8n4r
# tcHsNaLY5nHZ/Y1QO6GCAigwggIkBgkqhkiG9w0BCQYxggIVMIICEQIBATCBjjB3
# MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
# bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSEwHwYDVQQDExhN
# aWNyb3NvZnQgVGltZS1TdGFtcCBQQ0ECEzMAAACb4HQ3yz1NjS4AAAAAAJswCQYF
# Kw4DAhoFAKBdMBgGCSqGSIb3DQEJAzELBgkqhkiG9w0BBwEwHAYJKoZIhvcNAQkF
# MQ8XDTE2MDQxNDIzMzIwOVowIwYJKoZIhvcNAQkEMRYEFKhxQ4LIoYSeHJUik+PC
# MfUzJK1VMA0GCSqGSIb3DQEBBQUABIIBADRBunuLkBJCc0dRQGddecYsi4LUJFvl
# FIT5xLA8vbR0yK7/rtFMFddXO+CFM8HaMmz/t+4aP8CEsl5K2aaLIRlMT5SMaEA/
# 9RtG8E548CRd3MiRqtydtG1wJhjvhjsB2bCvhZZsKjYzol/PCtyz0fobd0f1HaGX
# V+5ichbjGCbQyE5dq0Ncd8dWM7L9Fo/lT4dtFCDOm/ujfLsk4u9UhAO5+bwtG//E
# VeoQggNlnyP/pU/2pZ1daRhGY9XrdDJpF8OgKPJXzSw3u/FMO3Kc5rPOKTh+o24M
# KI1eyOO/NaFF9liffCdZuyUsgeF6XUBTU0Htsga3QW/i7BZNAeKEO+Y=
# SIG # End signature block
