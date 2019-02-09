function NET_Check {
# .NET 4.6 or higher
If ((Get-ItemProperty -Path 'HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction SilentlyContinue).Version -ge '4.6.1')
{
$version = ((Get-ItemProperty -Path 'HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction SilentlyContinue).Version)
write-host ".NET $version is installed - Meets MD-STAFF Requirements"
}
else
{
$version = ((Get-ItemProperty -Path 'HKLM:\Software\Microsoft\NET Framework Setup\NDP\v4\Full' -ErrorAction SilentlyContinue).Version)
write-host ".NET $version is installed and DOES NOT meet MD-STAFF Requirements"
#Update .NET Framework
Function Update-NetFramework
{
New-Item "C:\NetFramework" -type directory -Force
$path = "C:\NetFramework"
$webclient = New-Object System.Net.WebClient
$OS = (Get-CimInstance Win32_OperatingSystem)
write-host $OS.version
If ($OS.Version -ge "10.*"){
write-host "Windows Server 2016 detected, installing .NET 4.7.1"
$url = 'https://download.microsoft.com/download/9/E/6/9E63300C-0941-4B45-A0EC-0008F96DD480/NDP471-KB4033342-x86-x64-AllOS-ENU.exe'
}
else
{
write-host "Windows Server 2012 detected, installing .NET 4.6.2"
$url = 'https://download.microsoft.com/download/F/9/4/F942F07D-F26F-4F30-B4E3-EBD54FABA377/NDP462-KB3151800-x86-x64-AllOS-ENU.exe'
}

$filename = [System.IO.Path]::GetFileName($url)
$file = "$path\$filename"
Try {
$webclient.DownloadFile($url,$file)
}
catch
{
write-host "Could not download .NET Framework from Microsoft Downloads - Check Internet Connectivity"
}

#Start-Process $file -ArgumentList '/q' -Wait
}
write-host "Updating .NET Framework"
Update-NetFramework
}
    }
