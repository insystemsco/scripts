<#

    What does this do?: 
                   This script creates a csv file with BitLocker Recovery Password/Key information for computers that have BitLocker enabled mount points.
                   - A time stamped "ComputerList" file is created when the script is run. The computer account file contains a list of computers the script will run against.
                   - A time stamped "NoResponse" file will be created if any of the computers from the list do not respond.
                   - A time stamped "Result" file will be created containing the hostname, tpm info, BitLockerID and RecoveryPassword/Key. All computers that respond to the 
                     query requests (that have BitLocker enabled) will be included.

    Prerequisites: 
                   - Script must be run on an account with Domain Administrative rights on a Domain Controller, a member server, or workstation with the RSAT tools installed.
                   - Script must be run in an elevated PowerShell instance.
                   - PowerShell Version 4.0 (at least) must be used.
                   - "Set-ExecutionPolicy remotesigned" must be enabled.

    How to use the script:
                   Simply run the script from an elevated PowerShell instance.
                   Ex: .\BitLocker_Pull_v1.0.ps1
   

#>
#
# Checks to see if the script is being run with administrator rights
if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator")) 
{
    Write-Warning "Administrative elevation has not been detected. `nPlease run the script as an Administrator`n "
    Start-Sleep 2
    Write-Warning "Script now exiting ...`n "
    Start-Sleep 2
    break
}
else {
    Write-Output "`n"    
    Write-Output "Administrative elevation detected, initiating script.`n"
    Start-Sleep 3
    Clear-Host
}
# Check PowerShell Version
If ($PSVersionTable.PSVersion.Major -lt 4) {
    Start-Sleep 2
    Write-Host ""
    Write-Host "This script requires PowerShell 4.0 or higher."
    Write-Host ""
    Start-Sleep 1
    Write-Host "Exiting Script ..."
    Start-Sleep 2
    Exit
}
# Function to start the bitlocker pull
function bitlockerpull {
    # Start Timer
    $timer = [System.Diagnostics.Stopwatch]::StartNew()
    # Sets the date format
    $Date=Get-Date -Format "MM-dd-yyyy_hh-mm-ss"
    # Location of where the "noresponse" and "result" files are created
    $noresponse = $logdrive+"\BitLockerPull_NoResponse_$Date.txt"
    $bitlockerinfo = $logdrive+"\BitLockerPull_Result_$Date.csv"
    # Location of where the "ComputerList" file is created
    $computerlist = $logdrive+"\BitLockerPull_ComputerList_$Date.txt"
    # 
    Get-ADComputer -Filter {OperatingSystem -like "Windows*"} -Properties Name | Select-Object -ExpandProperty Name > $computerlist
    $computers = Get-Content -Path $computerlist
    $eachcomp = @()
    ForEach ($computer in $computers) {
        if (Test-Connection $computer -Count 1 -Quiet){
            $mount = Invoke-Command $computer {Get-BitLockerVolume | Where-Object {$_.ProtectionStatus -eq "on"} | Select-Object MountPoint}
            $tpm = Invoke-Command $computer -ErrorAction SilentlyContinue {Get-BitLockerVolume | Select-Object -ExpandProperty KeyProtector | Where-Object {$_.KeyProtectorType -eq "Tpm"} | Select-Object KeyProtectorID}
            $keyprotectorid = Invoke-Command $computer {Get-BitLockerVolume | Select-Object -ExpandProperty KeyProtector | Where-Object {$_.KeyProtectorType -eq "RecoveryPassword"} | Select-Object KeyProtectorID}
            $recoverypassword = Invoke-Command $computer {Get-BitLockerVolume | Select-Object -ExpandProperty KeyProtector | Where-Object {$_.KeyProtectorType -eq "RecoveryPassword"} | Select-Object RecoveryPassword}
            #
            $object = New-Object psobject -Property @{
                ComputerName = (@($computer) | Out-String).Trim()
                TpmInfo = (@($tpm.KeyProtectorID) | Out-String).Trim()
                MountPoint = (@($mount.MountPoint) | Out-String).Trim()
                BitLockerID = (@($keyprotectorid.KeyProtectorID) | Out-String).Trim()
                RecoveryPassword = (@($recoverypassword.RecoveryPassword) | Out-String).Trim()
            }
            $eachcomp += $object
            #
        $eachcomp | Select-Object ComputerName,TpmInfo,MountPoint,BitLockerID,RecoveryPassword | Export-Csv $bitlockerinfo -NoTypeInformation
        } else {
        "Computer $computer is not responding. " >> $noresponse
            }
        }
        #
        $timer.stop()
        $timercomplete = $timer.Elapsed.TotalSeconds
        Write-Host "Script completed in: $timercomplete seconds.`n"
        Start-Sleep 2
        if (Test-Path $noresponse)
        {
            Write-Warning "Not all Windows computers responded. `n`nThis could be due to the Windows Remote Management Service (WinR) being disabled or the computer is turned off.`n`nReview the $noresponse file for details.`n"
            Start-Sleep 5
        }
        Write-Host "`nReview the $bitlockerinfo file for the Windows computers that responded.`n"
        Start-Sleep 2
    }
#
# Function to verify drive letter or path exists.
function logdriveexist {
    #$global:logdriveexist = Test-Path $logdrive
    Write-Host "Verifying the drive letter or path exists...`n"
    Start-Sleep 2
    if (Test-Path $logdrive){
        Start-Sleep 1
        Write-Host "Drive letter or path exists, starting script.`n"
        Start-Sleep 2
        Write-Host "Script started. It can take a while for the script to finish, please be patient...`n"
                # Call bitlockerpull function
                bitlockerpull
        } else {
        Write-Warning "The drive letter or folder path entered does not exist.`n`nVerify the driver letter or folder path exists and try again.`n"
        Start-Sleep 5
        Clear-Host
        Write-Host "`n--- BitLocker Pull v1.0 ---`n`n"
        # Call savelocation function
        savelocation
        }
    }    
# Function to prompt for a drive letter to store the files.
function savelocation {
$global:logdrive=Read-Host "Enter the drive letter or path that will house the 'BitLocker_Pull' files (Ex. C:, C:\BitLockerPull)"
Write-Host ""
# Call logdriveexist function
logdriveexist
}
# Start
Write-Host "`n--- BitLocker Pull v1.0 ---`n`n"
Start-Sleep 2
# Import Active Directory Module
Import-Module ActiveDirectory
# Import BitLocker Module
Import-Module BitLocker -DisableNameChecking
# Call savelocation function
savelocation
#
Write-Host ""