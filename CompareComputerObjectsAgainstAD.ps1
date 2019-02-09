<# 
.SYNOPSIS
This script compares a list of servers from a text file against Active Directory computer objects and outputs the reults to a CSV file.

.DESCRIPTION
This script compares a list of servers from a text file against Active Directory computer objects and outputs the reults to a CSV file.
Provide the input (server names) one per line under ServerNames.txt file. Output is stored in Output.csv file.


.IMPORTANT
This script should be executed on a Domain Controller.


.NOTES
Author: Pratik Pudage (Pratik.Pudage@hotmail.com)

#>

Import-Module ActiveDirectory
$Servers = Get-Content .\ServerNames.txt
$Results = ForEach ($server in $Servers) 
{   Try {
        Get-ADComputer $server -ErrorAction Stop
        $Result = $true
    }
    Catch {
        $Result = $False
    }
    [PSCustomObject]@{
        ComputerName = $server
        Found = $Result
    }
} 

$Results |Select ComputerName,Found | Export-Csv .\Output.csv -NoTypeInformation