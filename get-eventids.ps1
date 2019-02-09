#Server list
$Servers = Get-Content "c:\scripts\servers.txt"
 
#Query remote machines
Invoke-Command $Servers {
 
    $Filter = @{
           ProviderName = 'AD FS'
           ID = 238,246,247,305,306,353
           StartTime =  [datetime]::Today.AddDays(-5)
           EndTime = [datetime]::Today
    }
    Get-WinEvent -FilterHashtable $Filter
 
} | Select-Object MachineName,TimeCreated,ID,Message | Out-GridView -Title "Results" 