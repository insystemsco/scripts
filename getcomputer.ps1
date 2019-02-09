<# 
.SYNOPSIS 
This script will find the OU of the current computer or a specified computer or computers. 
 
.DESCRIPTION 
This script will use the System.DirectoryServices functionality to retrieve the OU for 
the current computer or a list of computers. This will return a list of objects with 
a Name and OU property. If the ValueOnly switch is used, then only the OU of the objects 
is returned. 
 
.PARAMETER ComputerName 
The name(s) of the computer or computers to retrieve the OUs for. This can be specified as a string 
or array. Value from the pipeline is accepted. 
 
.PARAMETER ThisComputer 
This parameter does not need to be specified. It is used to differentiate between parameter sets. 
 
.PARAMETER ValueOnly 
This switch parameter specifies that only the OU of the computer(s) should be returned. 
 
.EXAMPLE 
.\GetComputerOU.ps1 -ValueOnly 
 
This will return the OU of the current computer. 
 
.EXAMPLE 
$ComputerList = @("Computer1", "Computer2", "Computer3") 
$ComputerList | .\GetComputer.ps1
 
This will return an array of object (Name and OU) of the three computer names specified. 
#> 
[CmdletBinding()] 
#requires -version 3 
param( 
    [parameter(ParameterSetName = "ComputerName", Mandatory = $true, ValueFromPipeline = $true, Position = 0)] 
    $ComputerName, 
    [parameter(ParameterSetName = "ThisComputer")] 
    [switch]$ThisComputer, 
    [switch]$ValueOnly 
) 
 
begin 
{ 
    $rootDse = New-Object System.DirectoryServices.DirectoryEntry("LDAP://RootDSE") 
    $Domain = $rootDse.DefaultNamingContext 
    $root = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$Domain") 
} 
 
process 
{ 
    if ($PSCmdlet.ParameterSetName -ne "ComputerName") 
    { 
        $ComputerName = $env:COMPUTERNAME 
    } 
 
    $searcher = New-Object System.DirectoryServices.DirectorySearcher($root) 
    $searcher.Filter = "(&(objectClass=computer)(name=$ComputerName))" 
    [System.DirectoryServices.SearchResult]$result = $searcher.FindOne() 
    if (!$?) 
    { 
        return 
    } 
    $dn = $result.Properties["distinguishedName"] 
    $ouResult = $dn.Substring($ComputerName.Length + 4) 
    if ($ValueOnly) 
    { 
        $ouResult 
    } else { 
        New-Object PSObject -Property @{"Name" = $ComputerName; "OU" = $ouResult} 
    } 
}