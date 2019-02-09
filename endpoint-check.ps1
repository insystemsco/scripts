<#
.SYNOPSIS
	Name: List-EndpointProtection.ps1
	The purpose of this script is to determine what endpoint protection is used.
	
.DESCRIPTION
	The purpose of this script is to gather information on what endpoint protection is used in order to comply with Cyber Essentials Plus.
	
.PARAMETER LogOutput
	Logs the output to the default file path "C:\<hostname>.List-EndpointProtection.txt".
	
.PARAMETER LogFile
	When used in combination with -LogOutput, logs the output to the custom specified file path.

.EXAMPLE
	Run with the default settings:
		List-EndpointProtection
		
.EXAMPLE 
	Run with the default settings AND logging to the default path:
		List-EndpointProtection -LogOutput
	
.EXAMPLE 
	Run with the default settings AND logging to a custom local path:
		List-EndpointProtection -LogOutput -LogPath "C:\$env:computername.List-EndpointProtection.txt"
	
.EXAMPLE 
	Run with the default settings AND logging to a custom network path:
		List-EndpointProtection -LogOutput -LogPath "\\servername\filesharename\$env:computername.List-EndpointProtection.txt"
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

Param(
	[switch]$LogOutput,
	[string]$LogPath
)

#----------------------------------------------------------[Declarations]----------------------------------------------------------

$RunAsAdministrator = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator);

$LogPath_Default = "C:\$env:computername.List-EndpointProtection.txt";

#-----------------------------------------------------------[Functions]------------------------------------------------------------

Function List-EndpointProtection {
	Param()
	
	Begin {
		Write-Output "Searching for installed endpoint protection / security applications...";
		Write-Output "";
	}
	
	Process {
		Try {
			$EP_Active = $False;
			
			$EPApps = Get-WmiObject -Namespace "root\SecurityCenter2" -Class AntiVirusProduct;
			
			If ($EPApps -NE $Null){
				ForEach ($EPApp in $EPApps) {
					$EPApp_Name = $EPApp.displayName;
					Write-Output "`t Found endpoint protection app '$EPApp_Name'. Analysing...";
					Write-Output "`t `t Detecting status...";
					
					$EPApp_Status = $EPApp.productState;
					Switch ($EPApp_Status) {
						"262144" {$EPApp_Status_Definitions = "Up-to-date"; $EPApp_Status_RealTimeProtection = "Disabled";}
						"262160" {$EPApp_Status_Definitions = "Out-of-date"; $EPApp_Status_RealTimeProtection = "Disabled";}
						"266240" {$EPApp_Status_Definitions = "Up-to-date"; $EPApp_Status_RealTimeProtection = "Enabled";}
						"266256" {$EPApp_Status_Definitions = "Out-of-date"; $EPApp_Status_RealTimeProtection = "Enabled";}
						"393216" {$EPApp_Status_Definitions = "Up-to-date"; $EPApp_Status_RealTimeProtection = "Disabled";}
						"393232" {$EPApp_Status_Definitions = "Out-of-date"; $EPApp_Status_RealTimeProtection = "Disabled";}
						"393472" {$EPApp_Status_Definitions = "Up-to-date"; $EPApp_Status_RealTimeProtection = "Disabled";}
						"393488" {$EPApp_Status_Definitions = "Out-of-date"; $EPApp_Status_RealTimeProtection = "Disabled";}
						"397312" {$EPApp_Status_Definitions = "Up-to-date"; $EPApp_Status_RealTimeProtection = "Enabled";}
						"397328" {$EPApp_Status_Definitions = "Out-of-date"; $EPApp_Status_RealTimeProtection = "Enabled";}
						"397568" {$EPApp_Status_Definitions = "Up-to-date"; $EPApp_Status_RealTimeProtection = "Enabled";}
						"397584" {$EPApp_Status_Definitions = "Out-of-date"; $EPApp_Status_RealTimeProtection = "Enabled";}
						Default {$EPApp_Status_Definitions = "Unknown"; $EPApp_Status_RealTimeProtection = "Unknown";}
					}
					
					Write-Output "`t `t `t Definitions: $EPApp_Status_Definitions.";
					Write-Output "`t `t `t Real-time protection: $EPApp_Status_RealTimeProtection.";
					
					If ($EP_Active -Eq $False){
						If (($EPApp_Status_Definitions -Eq "Up-to-date") -And ($EPApp_Status_RealTimeProtection -Eq "Enabled")){
							$EP_Active = $True;
						}
					}
					
					Write-Output "";
				}
				
				If ($EP_Active -Eq $True){
					Write-Output "This computer IS protected by at least one endpoint protection app that is up-to-date AND enabled.";
				} Else {
					Write-Output "This computer IS NOT protected by at least one up-to-date and active endpoint protection app.";
				}
			} ElseIf ($EPApps -Eq $Null) {
				Write-Output "NO endpoint protection apps found!";
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

List-EndpointProtection;

If ($LogOutput -Eq $True) {
	Stop-Transcript | Out-Null;