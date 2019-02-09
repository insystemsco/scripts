# define variables
$appSetup = & "\\mdt1\DeploymentShare`$\Applications\Microsoft Office 2016\setup.exe" # use "setup /admin" to create a custom.msp and save to updates folder 
$appsetup = $appSetup -replace "`"",""
$appName = 'Office 2016'
$localLog = 'log.txt'
$timeFormat = 'yyyy/MM/dd hh:mm:ss tt'


# Office 2013
$checkOffice2013_32 = 'c:\Program Files (x86)\Microsoft Office\Office15\WINWORD.EXE'
$scrubOffice2013 = "\\mdt1\DeploymentShare$\Applications\Microsoft Office 2016scrub2013.vbs" # scrubber from Microsoft

#Office 2016
$checkOffice2016_32 = 'c:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE'

# create event log source
new-eventlog -Logname Application -source $appName -ErrorAction SilentlyContinue
$logstamp = (get-date).toString($timeFormat) ; $logstamp + " Created event log source." | out-file -filepath $localLog -Append

Clear-Host 

##### Check if Office 2013 installed #####

# Office 32 bit Checker
If (Test-Path -Path $checkOffice2013_32){
#Uninstall Office 2013
Write-Host "Removing Office 2013..."
Start-Process cscript $scrubOffice2013 ALL /Quiet /NoCancel
$logstamp = (get-date).toString($timeFormat) ; $logstamp + " Removed Office 2013 32 bit" | out-file -filepath $localLog -Append
 }Else{
   "32 bit Office does not exist"
}

##### Check if Office 2016 is already installed #####

# Office 32 bit Checker
If (Test-Path -Path $checkOffice2016_32){
#Uninstall Office 2013
Write-Host "Office 2016 32bit already installed..."
$logstamp = (get-date).toString($timeFormat) ; $logstamp + " Office 2016 32Bit already installed" | out-file -filepath $localLog -Append
{
# write event log
Clear-Host
Write-Host "Writing event log..."
$startTime = Get-date
$startLog = $appName + ' FAILED SETUP ' 
Write-Eventlog -Logname Application -Message $startLog -Source $appName -id 777 -entrytype Information -Category 0
$logstamp = (get-date).toString($timeFormat) ; $logstamp + " Failed setup!" | out-file -filepath $localLog -Append
}
Exit
 }Else{
   "32 bit Office does not exist"
}

# install application
Clear-Host 
Write-Host "Launching Setup..."
Start-Process -Wait -FilePath $appsetup -ArgumentList "/s"
$logstamp = (get-date).toString($timeFormat) ; $logstamp + " Launched setup" | out-file -filepath $localLog -Append
$logstamp = (get-date).toString($timeFormat) ; $logstamp + " Exit code: " + $LastExitCode | out-file -filepath $localLog -Append

# write event log
Clear-Host
Write-Host "Writing event log..."
$startTime = Get-date
$startLog = $appName + ' COMPLETED SUCCESSFULLY ' + $startTime
Write-Eventlog -Logname Application -Message $startLog -Source $appName -id 777 -entrytype Information -Category 0
$logstamp = (get-date).toString($timeFormat) ; $logstamp + " Installed successfully!" | out-file -filepath $localLog -Append
 
# exiting
Clear-Host
Write-Host "Installed successfully! Exiting now..."
Start-Sleep -s 4
$logstamp = (get-date).toString($timeFormat) ; $logstamp + " Exiting..." | out-file -filepath $localLog -Append