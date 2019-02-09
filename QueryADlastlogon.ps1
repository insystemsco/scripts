<#
	Title 	-	QueryADlastlogon.ps1
	Purpose -	Looks in AD to see when a computer object last logged on with
				domain
	Notes	-	This needs the Quest Active Directory module
#>

# set domain
$dom2chk = "mydomain.com"
$queryadlastlogon = "C:\queryadlastlogon.csv"

Get-QADComputer -service $dom2chk -ip lastlogontimestamp -sizelimit 0 | select name, lastLogonTimestamp | Export-Csv $queryadlastlogon