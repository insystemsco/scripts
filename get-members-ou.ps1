$OUs = "Your", "OUs", "here"
Foreach ($ou in $OUs){
Get-ADGroupMember -Identity $OU -Recursive | Get-AdUser -Properties * | Select-Object SamAccountName, GivenName, SurName, Mail, @{N='LastLogon' ; E={[DateTime]::FromFileTime($_.LastLogon)}}, Company
}
$List
$List | Export-Csv -Path 'C:\Adgroups.csv' -NoTypeInformation -Encoding UTF8
