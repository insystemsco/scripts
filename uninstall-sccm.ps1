#Browsing file
Add-Type -AssemblyName System.Windows.Forms
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog
$FileBrowser.filter = "Txt (*.txt)| *.txt"
[void]$FileBrowser.ShowDialog()
  
#Getting servers from txt file
Try
{
    $FilePath = $FileBrowser.FileName
    $Servers = Get-Content -Path $FilePath -ErrorAction Stop
}
Catch
{
    $_.Exception.Message
    Break
}
 
#Break if file is empty 
If( $Servers.length -eq 0 ) 
{ 
    Write-Warning "Servers list is empty"
    Pause 
    Break
}
 
#Uninstall command 
$scriptBlock = { C:\windows\ccmsetup\ccmsetup.exe /uninstall }
 
#Looping each server
Foreach($Server in $Servers)
{
    $Server = $Server.Trim()
    Write-Host "Processing $Server"
  
    #Check if server exist
    $FQDN = ([System.Net.Dns]::GetHostByName(("$Server")))
  
    If(!$FQDN)
    {
        Write-Warning "$Server does not exist"
    }
    Else
    {
        #Launch ccmsetup.exe /uninstall
        Invoke-Command -Scriptblock $scriptBlock -ComputerName $Server
    }
}