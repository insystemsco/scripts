<# 
This script checks if tpm is enabled and if so, it enables bitlocker
requires -Module ActiveDirectory
requires -runasadministrator
#>

Import-Module ActiveDirectory

# Credentials
$username = "domain\admin"
$passwordfile = 'C:\mysecurestring.txt'

if (!(Test-Path $Passwordfile))     
    {
    Read-Host "Enter password" | Out-File $Passwordfile
    Write-Output("$PasswordFile has been created")
    }

$password = Get-Content 'C:\mysecurestring.txt' | ConvertTo-SecureString -Force -AsPlainText
$credentials = new-object -typename System.Management.Automation.PScredential -argumentlist $username, $password
$BitlockerReport = 'C:\BitlockerReport.csv'
$OU = "DC=domain, DC=com"
$Computers = Get-ADComputer -Filter * -SearchScope Subtree -SearchBase $OU | select-object -expandproperty name

# Create TPM list if tpm is enabled it checks if bitlocker is enabled, if not it enables bitlocker
foreach ($Computer in $Computers) { 
    try {
        $tpmready = Invoke-Command -ComputerName $Computer -Credential $credentials -ScriptBlock {Get-Tpm | Select-Object -ExpandProperty Tpmready} -ErrorAction Stop
        $BLinfo = Invoke-Command -ComputerName $Computer -Credential $credentials -ScriptBlock {Get-Bitlockervolume -MountPoint 'C:'} -ErrorAction Stop
        $properties = @{Computer = $computer
                    Status = 'Connected'
                    TPM = $tpmready
                    Bitlocker = $BLinfo.ProtectionStatus}

            # If tpm is enabled and bitlocker is not enabled, enable bitlocker
            if ($tpmready -eq $true -and $BLinfo.ProtectionStatus -eq "Off"){
            # I've created a gpo that automatically backs up recovery keys to AD
            Invoke-Command -ComputerName $Computer -Credential $credentials -ScriptBlock {Add-BitLockerKeyProtector -MountPoint 'C:' -RecoveryPasswordProtector}
            Invoke-Command -ComputerName $Computer -Credential $credentials -ScriptBlock {Enable-BitLocker -MountPoint "C:" -EncryptionMethod Aes256 -TpmProtector}
            } 

        } catch {
            $properties = @{Computer = $computer
                        Status = 'Disconnected'
                        TPM = $null
                        Bitlocker = $null}

    } finally {
        $report = New-Object -TypeName PSObject -Property $properties
        Write-Output $report
        $report | export-csv -Append $BitlockerReport
    }
}

# Guide I used to backup recovery keys to AD
# http://jackstromberg.com/2015/02/tutorial-configuring-bitlocker-to-store-recovery-keys-in-active-directory/