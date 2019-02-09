Import-Module ActiveDirectory
Get-ADComputer –Filter * | Select –Expand Name | Tee-Object uphostst.txt

"C:\Program Files (x86)\Nmap\nmap.exe" -oG _upHosts -sP -n 192.100.1.101-200
for /F "tokens=1-5" %%A IN ('type _upHosts ^| find "Status: Up"') DO echo %%B >> upHosts.txt
del _upHosts

wmic /node:@hostlist.txt process call create "powershell enable-psremoting -force"

#CheckPsRemoting.ps1
#
#

[CmdletBinding()]
#One mandatory parameter, the hostfile we’re going to test
#Each line has one hostname or ip address
#Warning no error checking , we assume the file to be correct
Param(
   [ Parameter(Mandatory=$True,Position=1)]
   [string]$hostFile
)

# IF our output file exists we remove it
$outputFile = "Ps-Remoting-Enabled-Hosts.txt"
If (Test-Path $outputFile){
    Remove-Item $outputFile
}

Write-Host "Checking for Ps-Remoting enabled hosts : "
# For all the entries in hostlist .
get-content $hostFile | ForEach-Object{
    $machine = $_
    try{
        Write-Host $machine
        #Remote execution on $machine , -ErrorAction Stop used for neatly handling errors in try catch block
        Invoke-Command -ComputerName $machine -ScriptBlock{Write-Host " - PS Remoting Enabled" } -ErrorAction Stop
        Add-Content $outputFile $machine
    }
    catch{
        Write-Host "PS Remoting Failed "
    }
}
Write-Host "Test Finished."


.\PS_IOC.ps1 -hostFile .\hosts.txt -fileIOC_File .\fileioc.txt -regIOC_File .\regiocs.txt


<#
PS_IOC.ps1

Arguments :
    $hostFile     =>  File with each line containing one hostname or ip adress.
    $fileIOC_File =>  File with each line containing file or directory to search for , wildcards are accepted.

                      Example:

                      c:\temp\EvilFile.exe
                      c:\temp
                      c:\Users\*\ntuser.dat


    $regIOC_File  =>  File with each line containing registry key (wildcards are accepted) and optionally a value (no wildcards allowed here) to search for in that key.
                      Keys and values need to be seperated by a tab.
                      You need to give the powershell notation for the registry keys

                      Example:
                      HKLM:SOFTWARE\7-Zip   Path

                       --> checks whether value Path exists for a key HKLM:SOFTWARE\7-Zip on the system

                      HKLM:SOFTWARE\*\IBM

                       --> example of wildcard use , NO value specified

                      HKCU:Software\Microsoft\Windows\CurrentVersion\Run\   EvilPersistenceValue

                       --> Checks if EvilPersistenceValue appears in Software\Microsoft\Windows\CurrentVersion\Run\
                           in the user hive of the user executing the script !!
                           Next blog post will include version that scans for all users on the machine
#>

[CmdletBinding()]
Param(
  [ Parameter(Mandatory=$False,Position=1)]
   [string]$fileIOC_File,

  [ Parameter(Mandatory=$False,Position=2)]
   [string]$regIOC_File,

   [ Parameter(Mandatory=$False,Position=3)]
   [string]$hostFile

 )


####################Main##############

#load iocfunctions
. .\IocFunctions.ps1

if( -not $fileIOC_File -and -not $regIOC_File){
    "Warning: No IOCs were passed to script. Exiting."
    Exit
}

#Check if outputfile exists, if so delete
$outputFile = "ioc-scan-results.txt"
If (Test-Path $outputFile){
    Remove-Item $outputFile
}

# List all the IOCs were going to check
$fileIocArray = @()
if( $fileIOC_File){
    $fileIocArray = Get-Content $fileIOC_File
    "Scanning for File IOCs: " | Tee-Object $outputFile -Append
        foreach ($i in $fileIocArray){
              $i | Tee-Object $outputFile -Append
        }
}

$RegIocArray = @()
if( $regIOC_File){
    $regIocArray = Get-Content $regIOC_File
    "Scanning for Registry IOCs: " | Tee-Object $outputFile -Append
        foreach ($i in $regIocArray){
              $i | Tee-Object $outputFile -Append
        }
}

#no hostfile ? -> check only local system
if (-not $hostFile){
    write-host "NO HOSTFILE, checking on localhost only."
    $hosts = @('localhost')
 }
 else{
    $hosts = (get-content $hostFile)
 }


#Building the remote session
$psSessions = New-PSSession -ComputerName $hosts
#add the functions to our session
Invoke-Command  -Session $psSessions -FilePath .\IocFunctions.ps1
#remote invocation of checkFileIocs as a Job
$Job = Invoke-Command  -Session $psSessions -ScriptBlock ${function:checkFileIOCs} -ArgumentList (,$fileIocArray) -AsJob
Wait-Job $Job
#remote invocation of checkRegIocs as a Job
$Job = Invoke-Command  -Session $psSessions -ScriptBlock ${function:checkRegIOCs} -ArgumentList (,$RegIocArray) -AsJob
Wait-Job $Job

#After waiting for execution we collect output from the jobs
foreach($job in Get-Job){
    Receive-Job -Job $job | Tee-Object $outputFile -Append
}
#Close the session (important)
Remove-PSSession $psSessions

<#
Checks file system for existence of files / directories .
Hidden file's will be included

$File_IOC_List =>  Array with each entry containing file or directory to search for , wildcards are accepted.

                      Example of entries:

                      c:\temp\EvilFile.exe
                      c:\temp
                      c:\Users\*\ntuser.dat
#>

function checkFileIOCs([string[]] $File_IOC_List ){

    "`r`n<<<<<<<<<Checking filesystem " +  $env:COMPUTERNAME + " user : " + $env:USERDOMAIN +'\' + $env:USERNAME + ">>>>>>>>>"

    foreach ($ioc in $File_IOC_List) {
       try{
             "`r`n*********SCANNING " + $env:COMPUTERNAME + " FOR " + $ioc.ToUpper() + "    ***********"
             Get-ChildItem $ioc -force  -ErrorAction Stop
             "`r`nHIT FOUND !`r`n"
        }
        catch{
            #do nothing , just to hide errors from the screen
        }
    }
}


<#
 Reg_IOC_List =>      Array with each entry containing registry key (wildcards are accepted in the keyname) and optionally a value (no wildcards allowed here) to search for in that key.

                      Keys and values need to be seperated by a TAB character.
                      You need to give the powershell notation for the registry keys

                      Example:
                      HKLM:SOFTWARE\7-Zip   Path

                       --> checks whether value Path exists for a key HKLM:SOFTWARE\7-Zip on the system

                      HKLM:SOFTWARE\*\IBM

                       --> example of wildcard use , NO value specified

                      HKCU:Software\Microsoft\Windows\CurrentVersion\Run\   EvilPersistenceValue

                       --> Checks if EvilPersistenceValue appears in Software\Microsoft\Windows\CurrentVersion\Run\
                           in the user hive of the user executing the script !!
                           Next blog post will include version that scans users registry of ALL the users on the remote machine.
#>
function checkRegIOCs([string[]] $Reg_IOC_List ){

    "`r`n<<<<<<<<<Checking registry " +  $env:COMPUTERNAME + " as user : " + $env:USERDOMAIN +'\' + $env:USERNAME + ">>>>>>>>>"

     foreach ($ioc in $Reg_IOC_List) {
       try{
             #split Using tab as delimmiter
             $regKey ,$regValue= $ioc.split("`t",2)

             #check if key ends on "\" else add it , otherwise lookups may not be correct
             if (-not $regKey.EndsWith("\")){
                 $regKey+= '\'
             }

             "`r`n*********SCANNING " + $env:COMPUTERNAME + " FOR -  Reg Key: " + $regKey.ToUpper() + "  - Value : "+ $regValue +" **********"
             if($regValue -eq $null){
                #If no Reg value specified (only a key)
                ($foundValue = Get-ChildItem $regKey -force  -ErrorAction Stop)
                if ( $foundValue -eq $null) {
                    # If it's a key with no more subkeys we need to use Get-ItemProperty to list any values
                    ($foundValue = Get-ItemProperty -ErrorAction stop $regKey)
                    #no values found ?
                    if ( $foundValue -eq $null) {
                        Write-Error("No Hit") -ErrorAction stop
                    }
                }

             }
             else{
                 #Reg Key + value specified
                 ($foundValue = Get-ItemProperty -ErrorAction stop $regKey   | Select-Object -ExpandProperty $regValue -ErrorAction stop)
                 if ( $foundValue -eq $null) {Write-Error("No Hit") -ErrorAction stop}
             }
             "`r`nHIT FOUND !"
        }
        catch{
            #do nothing , just to hide errors from the screen

        }
    }
}
