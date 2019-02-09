#Requires -Module PSWindowsUpdate


<#
	.SYNOPSIS
		A brief description of the Update-Windows.ps1 file.
	
	.DESCRIPTION
		Install Windows Updates
	
	.PARAMETER AcceptAll
		A description of the AcceptAll parameter.
	
	.PARAMETER Verbose
		A description of the Verbose parameter.
	
	.NOTES
		Additional information about the file.
#>
param
(
	[switch]$AcceptAll,
	[switch]$Verbose
)

Get-WUServiceManager | ForEach-Object { Install-WindowsUpdate -ServiceID $_.ServiceID -AcceptAll:$AcceptAll -Verbose:$Verbose }
