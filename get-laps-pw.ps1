<#---------------------------------------------------------------------------

 REQUIREMENTS
-----------------------------------------------------------------------------
 Must have RSAT tools installed


 USAGE / SYNTAX
-----------------------------------------------------------------------------
 Run from Command Line:
      *  Open CMD with ADM creds
      *  Browse to get-laps-pw.ps1 location
      *  powershell.exe -file get-laps-pw.ps1
      
      Optional with computer argument
      *  powershell.exe -file get-laps-pw.ps1 <computername>
      *  ie...powershell.exe -file get-laps-pw.ps1 abc-d000345
      
      Optional with computer and domain arguments
      *  powershell.exe -file get-laps-pw.ps1 <computername> <domain>
      *  ie...powershell.exe -file get-laps-pw.ps1 abc-d000345 corp.twcable.com
      *  ie...powershell.exe -file get-laps-pw.ps1 abc-d000345 corp.chartercom.com

 Open with powershell:
      *  Open powershell with ADM creds
      *  browse to get-laps-pw.ps1 location
      *  .\get-laps-pw.ps1
      
      Optional with computer and domain arguments
      *  .\get-laps-pw.ps1 <computername>  ---  .\get-laps-pw.ps1 abc-d000345
      *  .\get-laps-pw.ps1 <computername> <domain>  --- .\get-laps-pw.ps1 abc-d000345 corp.twcable.com

 https://github.com/JFFail/LAPS-Password
 Requires -Modules ActiveDirectory
---------------------------------------------------------------------------#>

#Optional parameter to specify the computer name and domain as a parameter.
Param
(
	[Parameter(Mandatory=$False)]
		[string]$compName,
        [string]$i_domain
)

clear-host

#check if RSAT tools are installed, exit if not found
#If ((Get-WmiObject -class win32_optionalfeature | Where-Object { $_.Name -eq 'RemoteServerAdministrationTools'}) -ne $null) {Exit 0} else {If ((Get-Module -Name ActiveDirectory -ListAvailable) -ne $null) {Exit 0} else {Exit 1}}
if(get-module -list activedirectory) {
    # RSAT tools detected, import module
    Import-Module ActiveDirectory
    }
    else {
    # RSAT tools not found, notify and exit
    write-host "-------------------------------------------------" -ForegroundColor Yellow
    write-host ""
    write-host "Remote Server Administration Tools not installed!" -ForegroundColor Red
    write-host ""
    write-host "RSAT required to query Active Directory." -ForegroundColor Red
    write-host ""
    write-host "-------------------------------------------------" -ForegroundColor Yellow
    Exit 1
    }

#Get the machine name from the user if necessary.
#Didn't want to set Mandatory to $true because I wanted the flavor text.
if($compName -eq "") {
	write-host ""
    Write-Host "Enter the shortname of the computer; don't specify domain information!" -ForegroundColor Cyan
	$compName = Read-Host "Computer "
}


#Make sure the name is in the correct format.
$check = $compName.Split(".")

if($check.Count -ne 1) {
	write-host ""
    Write-Host "That isn't a valid name format! Try again..." -ForegroundColor Red
	$needName = $true
	
	while($needName) {
		write-host ""
        Write-Host "Enter the shortname of the computer; don't specify domain information!" -ForegroundColor Cyan
        $compName = Read-Host "Computer "
		$check = $compName.Split(".")
		
		#Re-do this if the name has dots.
		if($check.Count -eq 1) {
			$needName = $false
		} else {
			write-host ""
            Write-Host "Try entering a valid name!" -ForegroundColor Red
		}
	}
}

#Get the domain name from the user if necessary.
#Didn't want to set Mandatory to $true because I wanted the flavor text.
if($i_domain -eq "") {
	write-host ""
    Write-Host "Enter the domain name to search." -ForegroundColor Cyan
	write-host "corp.twcable.com / corp.chartercom.com / corp.local" -ForegroundColor Gray
    write-host "Leave blank for corp.chartercom.com" -ForegroundColor Gray
    $i_domain = Read-Host "Domain "
}

<#
#Make sure the object exists in AD.
#Running the script only works within a domain. Helps to prevent duplicate
#	name issues since my domains were migrated from 2008 where it didn't complain
#	when duplicate names existed between domains in a single forest.
#	Also means I don't need to import AdmPwd.PS.
$currentDomain = (Get-WmiObject -Class Win32_ComputerSystem).Domain
$currentDomain = $currentDomain.Split(".")
$domainValue = ""
$counter = 0
#>

<#
#Parse the "domain.something.tld" into "dc=domain,dc=something,dc=tld"
foreach($part in $currentDomain) {
	if($counter -lt ($currentDomain.Count - 1)) {
		$domainValue += "DC=" + $part + ","
	} else {
		$domainValue += "DC=" + $part
	}
	
	$counter++
}
#>

if($i_domain -ne "") {
    $domainvalue = $i_domain
    } else {
        $domainvalue = "corp.chartercom.com"
    }

#Pull back the AD object itself while getting the password property.
# $compObject = Get-ADComputer -Filter {name -eq $compName} -SearchBase $domainValue -SearchScope Subtree -Properties ms-Mcs-AdmPwd
$compObject = Get-ADComputer -Filter {name -eq $compName} -Server $domainValue -SearchScope Subtree -Properties ms-Mcs-AdmPwd

#Notify if the object doesn't exist in AD, the password is blank, or the user can't read it.
if($compObject -eq $null) {
	write-host ""
    Write-Host "That computer doesn't exist in AD! Please verify the name/domain and try again!" -ForegroundColor Red
} elseif($compObject."ms-Mcs-AdmPwd" -eq $null) {
	write-host ""
    Write-Host "Either the object has no password value or you don't have rights to access it!" -ForegroundColor Yellow
    write-host ""
} else {
	#No double-quotes will make PowerShell think you're trying to specify a parameter...
	write-host ""
    write-host "Password retrieved " -NoNewline -ForegroundColor Cyan
    write-host $compName"" -NoNewline -ForegroundColor gray
    write-host ": " -nonewline -ForegroundColor Cyan
    #write-host ": " -NoNewline -ForegroundColor Cyan
    Write-Host $compObject."ms-Mcs-AdmPwd" #-ForegroundColor Green
    write-host ""
    write-host "note....All EUCELEV password lookups are recorded by the DC and forwarded to splunk"
    write-host ""
}