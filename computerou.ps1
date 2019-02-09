#load list of computers
$computerList = Get-Content "c:\scripts\List.txt"

$results = @()

foreach($computerName in $computerList) {
    #add result of command to results array
    $results += Get-ADComputer $computerName -Properties Name, DistinguishedName | Select Name, DistinguishedName
}

#results to CSV file
$results | Export-Csv "c:\scripts\computerOUs.txt" -NoTypeInformation