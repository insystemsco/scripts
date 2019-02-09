<#
.SYNOPSIS
	Name: Remove-LocalAdminPermissions.ps1
	The purpose of this script is to selectively remove local administrative permissions.
	
.DESCRIPTION
	The purpose of this script is to remove local administrative permissions from all local user accounts except for "administrator", "Domain Admins", and any others specified in order to comply with Cyber Essentials Plus.
	
.PARAMETER LogOutput
	Logs the output to the default file path "C:\<hostname>.Remove-LocalAdminPermissions.txt".
	
.PARAMETER LogFile
	When used in combination with -LogOutput, logs the output to the custom specified file path.
	
.PARAMETER DisableDefaultAdmin
	Disables the local administrative user account "administrator".
	
.PARAMETER AdminWhitelist
	The provided list of user and/or group names will retain their administrative permissions. "Administrator" and "Domain Admins" are whitelisted by default.


.EXAMPLE
	Run with the default settings:
		Remove-LocalAdminPermissions
	
.EXAMPLE 
	Run with the default settings AND logging to the default path:
		Remove-LocalAdminPermissions -LogOutput
	
.EXAMPLE 
	Run with the default settings AND logging to a custom local path:
		Remove-LocalAdminPermissions -LogOutput -LogPath "C:\$env:computername.Remove-LocalAdminPermissions.txt"
	
.EXAMPLE 
	Run with the default settings AND logging to a custom network path:
		Remove-LocalAdminPermissions -LogOutput -LogPath "\\servername\filesharename\$env:computername.Remove-LocalAdminPermissions.txt"
	
.EXAMPLE 
	Run with the default settings AND the disabling of the local administrative user account "administrator":
		Remove-LocalAdminPermissions -DisableDefaultAdmin
	
.EXAMPLE 
	Run with the default settings AND an admin whitelist:
		Remove-LocalAdminPermissions -AdminWhitelist "ITAdmin", "Backup Admin"
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

Param(
	[Switch]$DisableDefaultAdmin,
	[Array]$AdminWhitelist,
	[Switch]$LogOutput,
	[String]$LogPath
)

#----------------------------------------------------------[Declarations]----------------------------------------------------------

$RunAsAdministrator = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator);

$LogPath_Default = "C:\$env:computername.Get-EndpointProtection.txt";

$LocalAdmins_ToRemain = "Administrator", "Domain Admins";
ForEach ($Admin in $AdminWhitelist){
	$LocalAdmins_ToRemain += $Admin;
}

#-----------------------------------------------------------[Functions]------------------------------------------------------------

Function List-ToRemainAdmins {
	Param()
	
	Begin {
		Write-Output "Getting list of to-remain admins...";
		Write-Output "";
	}
	
	Process {
		Try {
			ForEach ($Admin in $LocalAdmins_ToRemain){
				Write-Output "`t $Admin";
			}
		}
		
		Catch {
			Write-Output "";
			Write-Output "`t ...FAILURE. Something went wrong.";
			Break;
		}
	}
	
	End {
		If($?){ # only execute if the function was successful.
			
		}
	}
}

Function List-CurrentAdmins {
	Param()
	
	Begin {
		Write-Output "Getting list of current admins...";
		Write-Output "";
	}
	
	Process {
		Try {
			$LocalAdminGroup = [ADSI]("WinNT://$env:computername/Administrators,Group");
			# "$Script:" ensures that $LocalAdmins_Current_Paths can be used outside of this function;
			$Script:LocalAdmins_Current_Paths = $LocalAdminGroup.PSBase.Invoke("Members") | ForEach { $_.GetType().InvokeMember("ADSPath", "GetProperty", $Null, $_, $Null); };
			
			For ($i = 0; $i -NE $LocalAdmins_Current_Paths.Length; $i++){
				$LocalAdmin_Current_Username = $LocalAdmins_Current_Paths[$i].Split("/")[-1];
				
				Write-Output "`t $LocalAdmin_Current_Username";
			}
		}
		
		Catch {
			Write-Output "";
			Write-Output "`t ...FAILURE. Something went wrong.";
			Break;
		}
	}
	
	End {
		If($?){ # only execute if the function was successful.
			
		}
	}
}

Function Remove-Admins {
	Param()
	
	Begin {
		Write-Output "Removing local administrative permissions...";
		Write-Output "";
	}
	
	Process {
		Try {
			$LocalAdminGroup = [ADSI]("WinNT://$env:computername/Administrators,Group");
			
			For ($i = 0; $i -NE $LocalAdmins_Current_Paths.Length; $i++){
				$Retain = $False;
				
				$LocalAdmin_Current_Path = $LocalAdmins_Current_Paths[$i];
				$LocalAdmin_Current_Username = $LocalAdmin_Current_Path.Split("/")[-1];
				Write-Output "`t Analysing current admin '$LocalAdmin_Current_Username'...";
				
				If ($LocalAdmin_Current_Username -Eq "Administrator"){
					Write-Output "`t `t Checking whether DisableDefaultAdmin specified...";
					
					If ($DisableDefaultAdmin -Eq $True) {
						Write-Output "`t `t `t Found.";
						Write-Output "";
						Write-Output "`t `t `t Disabling...";
						
						$Error.Clear()
						Try {
							$DefaultAdmin = [ADSI]$LocalAdmin_Current_Path;
							$DefaultAdmin.UserFlags = 2;
							$DefaultAdmin.SetInfo();
						} Catch {
							Write-Output "`t `t `t FAILURE.";
						}
						If (!$Error){
							Write-Output "`t `t `t Success.";
						}
					} Else {
						Write-Output "`t `t `t NOT found.";
						Write-Output "";
						Write-Output "`t `t `t No changes will be made to status.";
					}
					
					Write-Output "";
				}
				
				Write-Output "`t `t Checking against list of to-remain admins...";

				
				For ($j = 0; $j -NE $LocalAdmins_ToRemain.Length; $j++){
					$LocalAdmin_ToRemain_Username = $LocalAdmins_ToRemain[$j];
					
					If ($LocalAdmin_Current_Username -NE $LocalAdmin_ToRemain_Username){
						$Retain = $False;
					} ElseIf ($LocalAdmin_Current_Username -Eq $LocalAdmin_ToRemain_Username) {
						$Retain = $True;
						
						Break;
					}
				}
				
				If ($Retain -Eq $False) {
					Write-Output "`t `t `t NOT found.";
					Write-Output "";
					Write-Output "`t `t `t Removing admin permissions...";
					
					$Error.Clear()
					Try {
						$LocalAdminGroup.Remove($LocalAdmin_Current_Path) | Out-Null;
					} Catch {
						Write-Output "`t `t `t FAILURE.";
					}
					If (!$Error){
						Write-Output "`t `t `t Success.";
					}
				} ElseIf ($Retain -Eq $True) {
					Write-Output "`t `t `t Found.";
					Write-Output "";
					Write-Output "`t `t `t No changes will be made to permissions / membership.";
				}
				
				Write-Output "";
			}
		}
		
		Catch {
			Write-Output "";
			Write-Output "`t ...FAILURE. Something went wrong.";
			Break;
		}
	}
	
	End {
		If($?){ # only execute if the function was successful.
			
		}
	}
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

If ($LogOutput -Eq $True) {
	If (-Not $LogPath) {
		$LogPath = $LogPath_Default;
	}
	Start-Transcript -Path $LogPath -Append | Out-Null;
}

Write-Output "This script requires administrative permissions. Checking...";
Write-Output "";
If ($RunAsAdministrator -Eq $False) {
	Write-Output "`t This script was NOT run as administrator. Exiting...";
	
	Break;
} ElseIf ($RunAsAdministrator -Eq $True) {
	Write-Output "`t This script WAS run as administrator. Proceeding...";
	Write-Output "";
	
	List-ToRemainAdmins;
	Write-Output "";
	List-CurrentAdmins;
	Write-Output "";
	Remove-Admins;
	
	Write-Output "Script complete. Exiting...";
}

If ($LogOutput -Eq $True) {
	Stop-Transcript | Out-Null;
}