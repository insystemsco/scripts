<#
.SYNOPSIS
	Name: Secure-WindowsServices.ps1
	Description: The purpose of this script is to secure any Windows services with insecure permissions.

.PARAMETER LogOutput
	Logs the output to the default file path "C:\<hostname>_Secure-WindowsServices.txt".
	
.PARAMETER LogFile
	When used in combination with -LogOutput, logs the output to the custom specified file path.

.EXAMPLE
	Run with the default settings:
		Secure-WindowsServices
		
.EXAMPLE 
	Run with the default settings AND logging to the default path:
		Secure-WindowsServices -LogOutput
	
.EXAMPLE 
	Run with the default settings AND logging to a custom local path:
		Secure-WindowsServices -LogOutput -LogPath "C:\$env:computername_Secure-WindowsServices.txt"
	
.EXAMPLE 
	Run with the default settings AND logging to a custom network path:
		Secure-WindowsServices -LogOutput -LogPath "\\servername\filesharename\$env:computername_Secure-WindowsServices.txt"
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------

Param(
	[switch]$LogOutput,
	[string]$LogPath
)

#----------------------------------------------------------[Declarations]----------------------------------------------------------

$RunAsAdministrator = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator);

$LogPath_Default = "C:\$env:computername`_Secure-WindowsServices.txt";

#-----------------------------------------------------------[Functions]------------------------------------------------------------

Function Secure-WindowsServices {
	Param()
	
	Begin {
		Write-Output "Securing all Windows services...";
	}
	
	Process {
		Try {
			If ($AlreadyRun -Eq $Null){
				$AlreadyRun = $False;
			} Else {
				$AlreadyRun = $True;
			}
			
			If ($AlreadyRun -Eq $False){
				[System.Collections.ArrayList]$FilesChecked = @(); # This is critical to ensuring that the array isn't a fixed size so that items can be added;
				[System.Collections.ArrayList]$FoldersChecked = @(); # This is critical to ensuring that the array isn't a fixed size so that items can be added;
			}
			
			Write-Output "";
			
			Write-Output "`t Searching for Windows services...";
			
			$WindowsServices = Get-WmiObject Win32_Service | Select Name, DisplayName, PathName | Sort-Object DisplayName;
			$WindowsServices_Total = $WindowsServices.Length;
			
			Write-Output "`t`t $WindowsServices_Total Windows services found.";
			
			Write-Output "";
			
			For ($i = 0; $i -LT $WindowsServices_Total; $i++) {
				$Count = $i + 1;
				
				$WindowsService_DisplayName = $WindowsServices[$i].DisplayName;
				$WindowsService_Path = $WindowsServices[$i].PathName;
				$WindowsService_File_Path = ($WindowsService_Path -Replace '(.+exe).*', '$1').Trim('"');
				$WindowsService_Folder_Path = Split-Path -Parent $WindowsService_File_Path;
				
				Write-Output "`t Windows service ""$WindowsService_DisplayName"" ($Count of $WindowsServices_Total)...";
				
				If ($FoldersChecked -Contains $WindowsService_Folder_Path){
					Write-Output "`t`t Folder ""$WindowsService_Folder_Path"": Security has already been ensured.";
					Write-Output "";
				} Else {
					$FoldersChecked += $WindowsService_Folder_Path;
					
					Write-Output "`t`t Folder ""$WindowsService_Folder_Path"": Security has not yet been ensured.";
					
					Correct-InsecurePermissions -Path $WindowsService_Folder_Path;
				}
				
				If ($FilesChecked -Contains $WindowsService_File_Path){
					Write-Output "`t`t File ""$WindowsService_File_Path"": Security has already been ensured.";
					Write-Output "";
				} Else {
					$FilesChecked += $WindowsService_File_Path;
					
					Write-Output "`t`t File ""$WindowsService_File_Path"": Security has not yet been ensured.";
					
					Correct-InsecurePermissions -Path $WindowsService_File_Path;
				}
			}
		}
		
		Catch {
			Write-Output "...FAILURE securing all Windows services.";
			$_.Exception.Message;
			$_.Exception.ItemName;
			Break;
		}
	}
	
	End {
		If($?){
			Write-Output "...Success securing all Windows services.";
		}
	}
}

Function Correct-InsecurePermissions {
	Param(
		[Parameter(Mandatory=$true)][String]$Path
	)
	
	Begin {
		
	}
	
	Process {
		Try {
			$ACL = Get-ACL $Path;
			$ACL_Access = $ACL | Select -Expand Access;
			
			$InsecurePermissionsFound = $False;
			
			ForEach ($ACE_Current in $ACL_Access) {
				$SecurityPrincipal = $ACE_Current.IdentityReference;
				$Permissions = $ACE_Current.FileSystemRights.ToString() -Split ", ";
				$Inheritance = $ACE_Current.IsInherited;
				
				ForEach ($Permission in $Permissions){
					If ((($Permission -Eq "FullControl") -Or ($Permission -Eq "Modify") -Or ($Permission -Eq "Write")) -And (($SecurityPrincipal -Eq "Everyone") -Or ($SecurityPrincipal -Eq "NT AUTHORITY\Authenticated Users") -Or ($SecurityPrincipal -Eq "BUILTIN\Users") -Or ($SecurityPrincipal -Eq "$Env:USERDOMAIN\Domain Users"))) {
						$InsecurePermissionsFound = $True;
						
						Write-Output "`t`t`t [WARNING] Insecure permissions found: ""$Permission"" granted to ""$SecurityPrincipal"".";
						
						If ($Inheritance -Eq $True){
							$Error.Clear();
							Try {
								$ACL.SetAccessRuleProtection($True,$True);
								Set-Acl -Path $Path -AclObject $ACL;
							} Catch {
								Write-Output "`t`t`t`t [FAILURE] Could not convert permissions from inherited to explicit.";
							}
							If (!$error){
								Write-Output "`t`t`t`t [SUCCESS] Converted permissions from inherited to explicit.";
							}
							
							# Once permission inheritance has been disabled, the permissions need to be re-acquired in order to remove ACEs
							$ACL = Get-ACL $Path;
						} Else {
							Write-Output "`t`t`t`t [NOTIFICATION] Permissions not inherited.";
						}
						
						Write-Output "";
						
						$Error.Clear();
						Try {
							$ACE_New = New-Object System.Security.AccessControl.FileSystemAccessRule($SecurityPrincipal, $Permission, , , "Allow");
							$ACL.RemoveAccessRuleAll($ACE_New);
							Set-Acl -Path $Path -AclObject $ACL;
						} Catch {
							Write-Output "`t`t`t`t [FAILURE] Insecure permissions could not be removed.";
						}
						If (!$error){
							Write-Output "`t`t`t`t [SUCCESS] Removed insecure permissions.";
						}
						
						Write-Output "";
					}
				}
			}
			
			If ($InsecurePermissionsFound -Eq $False) {
				Write-Output "`t`t`t [NOTIFICATION] No insecure permissions found.";
				Write-Output "";
			}
		}
		
		Catch {
			Write-Output "`t`t`t ...FAILURE.";
			$_.Exception.Message;
			$_.Exception.ItemName;
			Break;
		}
	}
	
	End {
		If($?){
			
		}
	}
}

#-----------------------------------------------------------[Execution]------------------------------------------------------------

If ($LogOutput -Eq $True) {
	If (-Not $LogPath) {
		$LogPath = $LogPath_Default;
	}
	Start-Transcript -Path $LogPath -Append | Out-Null;
	
	Write-Output "Logging output to file ""$LogPath""...";
	
	Write-Output "";
}

Write-Output "Administrative permissions required. Checking...";
If ($RunAsAdministrator -Eq $False) {
	Write-Output "`t This script was not run as administrator. Exiting...";
	
	Break;
} ElseIf ($RunAsAdministrator -Eq $True) {
	Write-Output "`t This script was run as administrator. Proceeding...";
	
	Write-Output "";
	Write-Output "----------------------------------------------------------------";
	Write-Output "";
	
	Secure-WindowsServices;
}

Write-Output "";
Write-Output "----------------------------------------------------------------";
Write-Output "";

Write-Output "Script complete. Exiting...";

If ($LogOutput -Eq $True) {
	Stop-Transcript | Out-Null;