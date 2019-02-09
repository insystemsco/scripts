if ($env:PROCESSOR_ARCHITEW6432 -eq "AMD64") {

  $x64PS=join-path $PSHome.tolower().replace("syswow64","sysnative").replace("system32","sysnative") powershell.exe

  $cmd = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($myinvocation.MyCommand.Definition))

  $out = & "$x64PS" -NonInteractive -NoProfile -ExecutionPolicy Bypass -EncodedCommand $cmd

  $out

  exit $lastexitcode

}

 

# Write Script Below

Write-Output ("Running Context: " + $env:PROCESSOR_ARCHITECTURE)

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
