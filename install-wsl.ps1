# Display Demo Machine Information
Get-ComputerInfo | `
    Select-Object `
        @{N='Hostname';E={$env:COMPUTERNAME}}, `
        WindowsProductName, `
        WindowsCurrentVersion, `
        WindowsVersion, `
        WindowsBuildLabEx | `
            Format-Table `
                -AutoSize ;
				
# Validate if Linux Subsystem feature state has been enabled on Windows 10
Get-WindowsOptionalFeature `
    -FeatureName Microsoft-Windows-Subsystem-Linux `
    -Online ;

# Enable the Microsoft Windows Subsystem Linux Feature on Windows 10
#  and reboot the Windows 10
Enable-WindowsOptionalFeature `
    -FeatureName Microsoft-Windows-Subsystem-Linux `
    -Online `
    -NoRestart:$False ;
	
# After enabling Linux Subsystem feature and rebooted the Windows 10,
#  validate if Linux Subsystem feature state has been enabled on Windows 10
Get-WindowsOptionalFeature `
    -FeatureName Microsoft-Windows-Subsystem-Linux `
    -Online ;
	
# To immediately start using Bash from the current default Ubuntu distro
lxrun /install /y

# Now, let us try switching into Bash
bash
	
# Let us validate the version in Bash
lsb_release -a

# Before doing anything in Bash, also ensure your Bahs is updated and upgraded
sudo apt-get update
sudo apt-get upgrade
	
# Now that we are in Bash, let us 
#  try generating a screenfetch
sudo apt-get install lsb-release scrot
wget -O screenfetch 'https://raw.github.com/KittyKatt/screenFetch/master/screenfetch-dev'
chmod +x screenfetch
./screenfetch

# Now, Launch Windows Store and get your preferred Linux distro
Start-Process `
    -FilePath "ms-windows-store://collection/?CollectionId=LinuxDistros" ;