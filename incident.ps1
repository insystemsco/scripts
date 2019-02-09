#SOME WMI EVENTS NEED ADMIN PRIVELEGES TO BE REGISTERED !!!


 <################################  INITIALIZATION  #####################################>


 #If you're not running from powershell_ise these assemblies need to be loaded explicitly
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.NotifyIcon")

#Getting current script path
#powershell 2 compliant
$scriptPath = split-path -parent $MyInvocation.MyCommand.Definition

$balloonTipInterval = 10000

#Hash tables that link notifications and there corresponding timer objects
$notification_Timer = @{}
$timer_Notification = @{}

#Messages for the paraoid security guy 🙂
$ParanoidMsgs =        @("Are we under attack?",
                         "They're hacking the sh*t out of you!",
                         "They're probably exfiltrating by now!",
                         "You are getting owned!",
                         "Do NOT keep calm !",
                         "Jeeeeeez !",
                         "Tight security in here...",
                         "DUDE !!!" ,
                         "You might as well give them your password...",
                         "Your security <= 0",
                         "Kiss your data goodbye!"
                        )


<########################################  FUNCTIONS FOR NOTIFICATIONS ###########################################

<#
createNotification
    Creates windows notifications that pop up in the system tray.
    When the icon of the notifications is clicked once the message re-appears.
    When the icon of the notifications is double-clicked the notification is removed

#>
function createNotification($text){

    $objNotifyIcon = New-Object System.Windows.Forms.NotifyIcon
    $objNotifyIcon.Icon = "$scriptPath\psg.ico"
    $objNotifyIcon.BalloonTipIcon = "info"
    $objNotifyIcon.BalloonTipText = $text
    $objNotifyIcon.BalloonTipTitle = "Paranoid Security Guy"
    $objNotifyIcon.Visible = $True
    $objNotifyIcon.ShowBalloonTip($balloonTipInterval)

    #Timer is used to be able to differentiate between single and double mouse clicks
    #We use the timer to create a delay for the message to re-appear
    #Without the delay from the timer the system only registers single clicks
    $timer = New-Object System.Timers.Timer
    $timer.Interval = 100
    $timer.AutoReset = $False
    $notification_Timer.add($timer.GetHashCode(),$objNotifyIcon)
    $timer_Notification.add($objNotifyIcon.GetHashCode(), $timer)


    Register-ObjectEvent -InputObject $objNotifyIcon -EventName MouseClick -Action {
        $t = $timer_Notification.item($Sender.getHashCode())
        $t.Start()
    }

    Register-ObjectEvent -InputObject $objNotifyIcon -EventName DoubleClick -Action {
        #remove the notification  + cleanup hashtables
        $t = $timer_Notification.item($Sender.getHashCode())
        $notification_Timer.remove($t.getHashCode())
        $timer_Notification.remove($Sender.getHashCode())
        $Sender.Dispose()
    }

    Register-ObjectEvent -InputObject $timer -EventName Elapsed -Action {
        $notif = $notification_Timer.item($Sender.getHashCode())
        $notif.ShowBalloonTip($balloonTipInterval)

    }
}


<#
showNotification
    Display's the notification text and add's some paranoid interpretations
#>
function global:showNotification($text){
    $random = Get-Random($ParanoidMsgs.Length)
    createNotification((get-date).ToString() +" " + $text + "`r`n" + $ParanoidMsgs[$random])
}


<###################################  WMI - EVENTS  ############################################



<#
 #
 # Monitoring of services
 # All services who stop cause a popup
 #
 #>


$wmiStoppedService = @{
     Query ="SELECT * FROM __InstanceModificationEvent WITHIN 2 WHERE TargetInstance Isa 'Win32_Service' AND TargetInstance.State = 'Stopped'"

     Action = {
        showNotification($Event.SourceEventArgs.NewEvent.TargetInstance.Name + " service was stopped.")
    }
    SourceIdentifier = "Service.Stopped"
}
$Null = Register-WMIEvent @wmiStoppedService
"Monitoring for stopped services."



<#
 #
 Monitoring of network logons
 Notifications about logons of type 3 get
 #get-help
 #>


$wmiNetworkLogon = @{
     Query ="select * from __InstanceCreationEvent where TargetInstance ISA 'Win32_NTLogEvent' and TargetInstance.eventcode = 4624 "

     Action = {
        $msg = $Event.SourceEventArgs.NewEvent.TargetInstance.message
        $logtype3 = ($msg.Split( [environment]::NewLine) | Select-String("Logon Type:") -SimpleMatch) -like "*3*"

        if ($logtype3) {

            $source =  ($msg.Split( [environment]::NewLine) | Select-String("Source Network Address") -SimpleMatch)

            showNotification("A network logon was performed." + $source)
        }

    }
    SourceIdentifier = "network.Logon"
}

$Null = Register-WMIEvent @wmiNetworkLogon
"Monitoring for network logons."




<#

 Monitoring of processes
 Processes who start from %temp% folder or $recycle.bin get flagged.

#>


$wmiProcessStarted = @{
    #we don't use Win32_ProcessStartTrace , sometimes it seems to miss really short living processes
    Query ="select * from __InstanceCreationEvent within 1 where TargetInstance ISA 'Win32_process' "

    Action = {
        $process = Get-Process -Id $Event.SourceEventArgs.NewEvent.TargetInstance.processid | Select-Object id,path
        if ($process.path -and ( (Get-Item -LiteralPath $process.path).FullName -like (Get-Item -LiteralPath $env:temp).FullName + '*' `
                                                            -or $process.path -like "*\`$Recycle.Bin\*") )
        {
            showNotification("Process with pid " + $process.id + " runs from       " + $process.path)
        }

    }
    SourceIdentifier = "Process.Started"
}

$Null = Register-WMIEvent @wmiProcessStarted
"Monitoring for processes starting from $env:temp and \`$Recycle.Bin."




<#
 #
 Monitoring for changes in runkeys of the active user
 #
 #>


#We monitor the hive of the current user so we search his SID
$userSid = ((Get-WmiObject -Class Win32_UserProfile -Namespace "root\cimv2"  | select sid,localpath) | Where-Object localpath -like "*$env:USERNAME*" ).sid
$wmiRegistryKeys = @{
    Query ="SELECT * FROM RegistryKeyChangeEvent WHERE Hive='HKEY_USERS' AND (KeyPath='$userSid\\Software\\Microsoft\\Windows\\CurrentVersion\\RunOnce' `
                                                                         OR  KeyPath='$userSid\\Software\\Microsoft\\Windows\\CurrentVersion\\Run' )"

    Action = { showNotification("Current user's run keys were modified.") }

    SourceIdentifier = "RegistryKey.Changed"
}

$Null = Register-WMIEvent @wmiRegistryKeys
"Monitoring for registry changes in current user's run key and runonce key."



<#

 Monitoring of System32 folder
 We do not monitor recursivly, to avoid an overload of events
 We monitor for creation,deletion,renaming

#>

$fsWatcher = New-Object System.IO.FileSystemWatcher
$fsWatcher.Path = $env:windir+"\system32"
$fsWatcher.IncludeSubdirectories = $false
$fsWatcher.EnableRaisingEvents = $true

"Monitoring for file creation in " + $env:windir +"\system32."
$System32Created = Register-ObjectEvent $fsWatcher "Created" -Action {
   showNotification("File created in system32 folder.           " +$($eventArgs.FullPath))
}

"Monitoring for file deletion in " + $env:windir +"\system32."
$System32Deleted = Register-ObjectEvent $fsWatcher "Deleted" -Action {
   showNotification("File deleted in system32 folder.           " +$($eventArgs.FullPath))
}

"Monitoring for file renaming in " + $env:windir +"\system32."
$System32Renamed = Register-ObjectEvent $fsWatcher "Renamed" -Action {
   showNotification("File renamed in system32 folder.           "+$($eventArgs.FullPath))
}



#########################  Main  ###########"##############

"
##############################################################################################
If You see errors during initialization make sure you have admin privileges.
Paranoid Security Guy is ready!
Hit ctrl-c to exit "


try{

    #We stay in an endless loop to keep the session open
    while($True){
        Wait-Event
    }
}
catch{

}
finally{
    write-host "Thanks for using The Paraoid Security Guy!"

    unregister-event -sourceIdentifier "Service.Stopped"   -ErrorAction Stop
    unregister-event -sourceIdentifier "Network.Logon" -ErrorAction Stop
    unregister-event -sourceIdentifier "Process.Started"   -ErrorAction Stop
    unregister-event -sourceIdentifier "RegistryKey.Changed"  -ErrorAction Stop
    Unregister-Event $System32Created.id
    Unregister-Event $System32Deleted.id
    Unregister-Event $System32Renamed.id

    write-host "All bindings are unloaded."

}
