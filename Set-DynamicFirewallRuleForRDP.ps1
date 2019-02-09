<#
.DESCRIPTION
Creates a Windows Firewall rule that blocks the IP addresses of all the network clients
that have connected to RDP (not necessarily authenticated - just established a TCP connection)
within the last 24 hours. This will effectively "ban" those IP addresses from making RDP password
guesses for 24 hours. Modify the 'exclusions file' with IP addresses that you want to always
have RDP access. E.g., adding "192.168." to the exclusions file will exempt the entire 192.168.0.0/16
network. Adding "192.168.1." will exclude the entire 192.168.1.0/24 network. Exclusions are entered
one per line in the text file. If you do not enter any exclusions at all, and then run this script,
a firewall rule will be created to block every IP address that has completed a TCP handshake with the RDP port in the 
past 24 hours. Including you. Which is probably not what you want. So you should really enter some exclusions.
This script requires that the auditing of Windows Filtering Platform (WFP)
be set to log successful connections. (Event ID 5156 in the Security log.) This can be configured via Advanced Audit Policy. It also
requires that the Windows Firewall be turned on.

.EXAMPLE
.\Set-DynamicFirewallRuleForRDP.ps1
You could put this in a scheduled task. Run it every 24 hours as SYSTEM. I do.
#>
#Requires -Version 4.0

Set-StrictMode -Version Latest

# Edit these variables to your liking.

[Int]$RDPPort = 3389
# If this script is running as SYSTEM, the LOCALAPPDATA environment variable probably resolves to
# C:\Windows\System32\config\systemprofile\AppData\Local
# Otherwise it will be in the profile directory of the user who launched the script, like
# C:\Users\Joe\AppData\Local
[String]$ExclusionsFilePath = (Join-Path $Env:LOCALAPPDATA 'Set-DynamicFirewallRuleForRDPExclusions.txt')
[String]$FirewallRuleName = 'Dynamic Firewall Rule for RDP'
[String]$LogFilePath = (Join-Path $Env:LOCALAPPDATA 'Set-DynamicFirewallRuleForRDP.log')
[Bool]$CreateFWRuleEnabled = $True # Set this to false for testing purposes, the firewall rule will be created but left disabled.
[Bool]$ForceNoExclusions = $False # Set this to true if you are REALLY sure you don't want to exclude ANY IP addresses.

# Stop editing here 


Add-Type -TypeDefinition @'
    public enum LOGLEVEL
    {
        INFO,
        WARNING,
        ERROR
    }
'@

# This function simplifies writing a message simultaneously to the console and to a file.
# Why don't I just use Start-Transcript, you ask? Well shut up, that's why!
Function Trace
{
    Param([Parameter(Mandatory=$True)][LOGLEVEL]$Level,
          [Parameter(Mandatory=$True)][String]$Message,
          [Parameter(Mandatory=$True)][String]$LogPath)

    [String]$TraceMessage = [String]::Empty
                
    Switch ($Level)
    {
        ERROR   { $TraceMessage = "[ERROR] $Message"; Write-Error $TraceMessage }
        WARNING { $TraceMessage = "[WARN]  $Message"; Write-Warning $TraceMessage }
        INFO    { $TraceMessage = "[INFO]  $Message"; Write-Output $TraceMessage }
        Default { Write-Error 'Debug assertion in Trace!' }
    }
    
    Out-File -Append -FilePath $LogPath -Encoding Unicode -InputObject $TraceMessage
}

Trace -Level INFO -LogPath $LogFilePath -Message "Set-DynamicFirewallRuleForRDP beginning at $(Get-Date)."


If (Test-Path $ExclusionsFilePath -PathType Leaf)
{
    Trace -Level INFO -LogPath $LogFilePath -Message "Exclusions file $ExclusionsFilePath found."
}
Else
{
    Trace -Level INFO -LogPath $LogFilePath -Message "Exclusions file $ExclusionsFilePath not found - creating it."
    New-Item -Path $ExclusionsFilePath -Type File
}

$ExclusionsFileContents = Get-Content $ExclusionsFilePath

If ($ExclusionsFileContents -EQ $Null)
{
    If ($ForceNoExclusions -EQ $False)
    {
        Trace -Level WARNING -LogPath $LogFilePath -Message "0 exclusions found. Everyone who has connected to RDP on this server in the last 24 hours will be blocked."
        Trace -Level WARNING -LogPath $LogFilePath -Message "Are you sure that's what you want?"
        $Answer = Read-Host '[Y|N]'

        If (-Not($Answer.ToLower().Trim().StartsWith('y')))
        {
            Return
        }
    }
}
Else
{
    Trace -Level INFO -LogPath $LogFilePath -Message "$($ExclusionsFileContents.Count) exclusions found."
}

Try
{
    # This requires that auditing of Windows Filtering Platform (WFP) be set to log successful connections. (Advanced Audit Policy.)
    $AllWFPEvents = Get-WinEvent -LogName Security -FilterXPath "*[System[EventID=5156]]" -ErrorAction Stop | Where TimeCreated -GT (Get-Date).AddDays(-1)
}
Catch
{
    If ($_.FullyQualifiedErrorId -Like "*NoMatchingEventsFound*")
    {
        Trace -Level WARNING -LogPath $LogFilePath -Message "No events with ID 5156 were found in the Security event log. Nothing else to do."
    }
    Else
    {
        Trace -Level ERROR -LogPath $LogFilePath -Message "$($_.Exception.Message)"
    }

    Return
}

$RDPConnections = @()

Foreach ($WFPEvent In $AllWFPEvents)
{    
    Foreach ($Line In $WFPEvent.Message -Split [Environment]::NewLine)
    {
        If ($Line.Trim() -Like "Direction:*")
        {
            If ($Line.Split(':')[1].Trim() -NE 'Inbound')
            {
                Continue
            }
        }

        If ($Line.Trim() -Like "Source Port:*")
        {
            If ($Line.Split(':')[1].Trim() -EQ "$RDPPort")
            {
                $RDPConnections += $WFPEvent
            }
        }
    }
}

Trace -Level INFO -LogPath $LogFilePath -Message "$($RDPConnections.Count) connections to port $RDPPort in the last 24 hours."

If ($RDPConnections.Count -EQ 0)
{
    Trace -Level INFO -LogPath $LogFilePath -Message "No connections to port $RDPPort detected in the last 24 hours. Exiting."
    Return
}

$DestinationIPs = @()

Foreach ($RDPConnection In $RDPConnections)
{
    Foreach ($Line In $RDPConnection.Message -Split [Environment]::NewLine)
    {
        If ($Line.Trim() -Like "Destination Address:*")
        {
            $DestinationIPs += $Line.Split(':')[1].Trim()
        }
    }
}
 
Trace -Level INFO -LogPath $LogFilePath -Message "A summary of connections to port $RDPPort is as follows:"
Trace -Level INFO -LogPath $LogFilePath -Message (($DestinationIPs | Group-Object | Select Count, Name | Sort Count -Descending | Format-Table -AutoSize) | Out-String)

$FirewallRule = (Get-NetFirewallRule | Where DisplayName -EQ $FirewallRuleName)

If ($FirewallRule -NE $Null)
{
    Trace -Level INFO -LogPath $LogFilePath -Message 'Removing previous firewall rule.'
    $FirewallRule | Remove-NetFirewallRule    
}

$DestinationIPsExclusionsRemoved = @()

Foreach ($IP In ($DestinationIPs | Sort | Get-Unique))
{
    [Bool]$ShouldExclude = $False
    Foreach ($Exclusion In $ExclusionsFileContents)
    {
        If ($IP -Like "$Exclusion*")
        {
            $ShouldExclude = $True            
        }
    }

    If ($ShouldExclude)
    {
        Trace -Level INFO -LogPath $LogFilePath -Message "Excluding IP $IP"
    }
    Else
    {
        $DestinationIPsExclusionsRemoved += $IP
    }
}

If ($DestinationIPsExclusionsRemoved.Count -EQ 0)
{
    Trace -Level WARNING -LogPath $LogFilePath -Message "All RDP connections in the past 24 hours were excluded as part of the exclusions list file. Exiting."
    Return
}

Trace -Level INFO -LogPath $LogFilePath -Message "Adding $($DestinationIPsExclusionsRemoved.Count) unique IPs to the firewall rule."

Try
{
    If ($CreateFWRuleEnabled)
    {
        New-NetFirewallRule -DisplayName $FirewallRuleName -Direction Inbound -Description 'This firewall rule is managed dynamically by the Set-DynamicFirewallRuleForRDP Powershell script.' -PolicyStore 'PersistentStore' -EdgeTraversalPolicy Block -Action Block -Enabled True -Protocol 'TCP' -LocalPort $RDPPort -RemoteAddress $DestinationIPsExclusionsRemoved -ErrorAction Stop
    }
    Else
    {
        New-NetFirewallRule -DisplayName $FirewallRuleName -Direction Inbound -Description 'This firewall rule is managed dynamically by the Set-DynamicFirewallRuleForRDP Powershell script.' -PolicyStore 'PersistentStore' -EdgeTraversalPolicy Block -Action Block -Enabled False -Protocol 'TCP' -LocalPort $RDPPort -RemoteAddress $DestinationIPsExclusionsRemoved -ErrorAction Stop
    }
    Trace -Level INFO -LogPath $LogFilePath -Message "A firewall rule named $FirewallRuleName has been created to block the IPs listed above."
}
Catch
{
    Trace -Level ERROR -LogPath $LogFilePath -Message "Failed to create firewall rule. $($_.Exception.Message)"
}

Trace -Level INFO -LogPath $LogFilePath -Message "Set-DynamicFirewallRuleForRDP finished at $(Get-Date).`r`n`r`n"