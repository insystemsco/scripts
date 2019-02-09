<#
    .SYNOPSIS
    Gets the Perfmon counters from either a local system or a specified target system.
    
    .DESCRIPTION
    Script will track the counters over a specified amount of time (default of 1 minute) and record the most significant value for each counter.

    .PARAMETER TargetComputer
    Leave unchanged to run against the local computer. Specify a hostname otherwise. 

    .PARAMETER MonitorDuration
    The value (in minutes) you want to monitor the computer. Default value is 1

    .PARAMETER SamplingRate
    How many seconds between each poll. Default value is 2

    .PARAMETER logPath
    The folder where the generated log file will reside.  The default is my Uploads folder. 

    .PARAMETER TargetProcess
    I just put this in here in case you want to only check against a single process. If you don't want to get process data for every single process in our library (which you can do by leaving this blank), this is the flag to use

    .EXAMPLE
    .\LocalCounterScript.ps1
    Runs the script against the local computer and uses all the default settings.  

    .EXAMPLE
    .\LocalCounterScript.ps1 -TargetComputer TEST-BOX-1 -MonitorDuration 5 -SamplingRate 5
    Captures performance information from computer "TEST-BOX-1" once per five seconds for five minutes

    .EXAMPLE
    .\LocalCounterScript.ps1 -TargetComputer TEST-BOX-2 -TargetProcess TaniumClient -logPath C:\Temp
    Captures performance information from TEST-BOX-2, but instead of getting process data from our massive list of them, only tracks the TaniumCllient one (and the generic system ones). 
    The output report has been redirected to the local temp folder in this example as well

    .EXAMPLE
    cat .\TestComputers.txt | % {
        if (Test-Connection -Count 1 -ComputerName $_ -quiet)
        {
            Start-Job {\\MyShare\LocalCounterScript.ps1 -TargetComputer $args[0] -monitorDuration 5 -samplingRate 3} -ArgumentList $_
        }
    }
    Runs through a list of computers and runs this script against each of them. Runs them as a job so you don't have to wait a million hours for each one to run before moving onto the next one. 
    Lets you run longer duration scans against multiple machines at once, quickly and easily. 

#>

param (
        [Parameter()]
        [string]$TargetComputer=$ENV:COMPUTERNAME,
        [Parameter()]
        [double]$monitorDuration = 1,
        [Parameter()]
        [int]$samplingRate = 2,
        [Parameter()]
        [string]$logPath = "\\testServer\uploads\synack\counterscript\",
        [Parameter()]
        [string]$targetProcess = "BLANK"  # I put this in here if you want to only track a specific process (i.e collecting Tanium Report data so you only want to monitor the TaniumClient process). It also collects the generic system ones
)

# Variable Declaration
$masterCountersList = @() #this is the array that will hold the GenericCounters combined with the process specific counters so I can run them all at once
$startTime = (Get-Date -format "yyyy-MM-d_hhmmss")
$GenericCounters =
"\\$TargetComputer\PhysicalDisk(*)\% Idle Time",
"\\$TargetComputer\Memory\% committed bytes in use",
"\\$TargetComputer\Memory\Available MBytes",
"\\$TargetComputer\Memory\Free System Page Table Entries",
"\\$TargetComputer\Memory\Pool Paged Bytes",
"\\$TargetComputer\Memory\Pool Nonpaged Bytes",
"\\$TargetComputer\Memory\Pages/sec",
"\\$TargetComputer\Processor(_total)\% processor time",
"\\$TargetComputer\Processor(_total)\% user time" # these are all the counters that aren't process specific

if ($targetProcess -eq "BLANK") {
    $ProcessList = 
    "CcmExec",
    "ACCM_MSGBUS",
    "ACCM_WATCH",
    "AuditManagerService",
    "fcag",
    "fcags",
    "fcagswd",
    "FireSvc",
    "HipMgmt",
    "macmnsvc",
    "macompatsvc",
    "masvc",
    "mcshield",
    "mctray",
    "mfecanary",
    "mfeesp",
    "mfefire",
    "mfehcs",
    "mfemactl",
    "mfemms",
    "mfetp",
    "mfevtps",
    "RSDPP",
    "scsrvc",
    "UpdaterUI",
    "ac.activclient.gui.scagent",
    "acevents",
    "concentr",
    "dvservice",
    "dvtrayapp",
    "receiver",
    "redirector",
    "selfserviceplugin",
    "wfcrun32",
    "cmdagent",
    "taniumclient",
    "taniumclient#1",
    "taniumclient#2",
    "nomadbranch" # these are the processes I want to get the Handle Count, Thread Count, and Private Byte information from.
}
else {
    $ProcessList = $targetProcess
    if ($targetProcess -eq "TaniumClient")
    {
        $ProcessList = 
        "taniumclient",
        "taniumclient#1",
        "taniumclient#2"
    }
    
}

# This turns the process names into full counter paths so we don't have to enter three of them per each process we add later down the line
function Return-CounterArray ($processName)
{
      $counters = @()
      $counters += "\\$TargetComputer\Process($processName)\Handle Count"
      $counters += "\\$TargetComputer\Process($processName)\Thread Count"
      $counters += "\\$TargetComputer\Process($processName)\Private Bytes"
      $counters += "\\$TargetComputer\Process($processName)\% Processor Time"
      return $counters
}

# This takes the generic counters and adds them to the master list along with all the processs based ones. I feel like this could probably be rolled up into the Return-CounterArray one, but it looks pretty like this
function Generate-Counters (){
    $allCounters = @()
    $allCounters += $GenericCounters
    $ProcessList | % {$allCounters += (Return-CounterArray $_)}
    return $allCounters
}

# Actually start the script now
$ReportArray = @() # this holds all the generate report objects for later export
$TempArray = @() # this holds the objects while we figure out the highest value
Write-host -ForegroundColor Green "$(get-date -format hh:mm:ss) - Starting report on $TargetComputer. Please be patient as the process begins."
Write-Host -ForegroundColor Green "$(get-date -format hh:mm:ss) - Generating master list based on $($GenericCounters.Count) System counters and $($ProcessList.count) Processes."
$masterCountersList = (Generate-Counters)
Write-Host -ForegroundColor Green "$(get-date -format hh:mm:ss) - Master list created with $($masterCountersList.count) items."
$maxSamples = [Math]::Round(($monitorDuration*60/$samplingRate), 0)	 #multiplies your monitor duration minutes by 60 and divides by your sampling interval. Rounds to 0 decimal places because Integers
Write-Host -ForegroundColor Green "$(get-date -format hh:mm:ss) - Will take $maxSamples samples over the course of $monitorDuration minutes."
$rawCounterDump = @()

# This actually goes and gets the counter information. Woot. 
$rawCounterDump = Get-Counter -Counter $masterCountersList -SampleInterval $samplingRate -MaxSamples $maxSamples -ComputerName $TargetComputer -ErrorAction SilentlyContinue

# This will export everything to BLG files so you can review them in Perfmon later if you'd like (gives a pretty line graph!) 
if ($logPath[-1] -ne "\") {$logPath += "\"}
$endTime = (Get-Date -format "yyyy-MM-d_hhmmss")
$blgDump = $logPath+"$TargetComputer-$endTime-RawData.blg"
Write-Host -ForegroundColor Green "$(get-date -format hh:mm:ss) - Dumping raw Perfmon data to $blgDump."

# now that the raw data has already been exported, this chunk turns that raw data into an array for further processing.
$rawCounterDump | Export-Counter -Path $blgDump
$rawCounterDump.countersamples | % {
    $path = $_.Path
    $obj = new-object psobject -property @{
        ComputerName = $TargetComputer
        Counter = $path.Replace("\\$($TargetComputer.ToLower())","")
        Item = $_.InstanceName
        Value = [Math]::Round($_.CookedValue, 2)	
        DateTime = (Get-Date -format "yyyy-MM-d hh:mm:ss")
    }
    $TempArray += $obj
}
Write-Host -ForegroundColor Green "$(get-date -format hh:mm:ss) - $($TempArray.count) total samples collected."

# This bit takes all the entries in TempArray, gets the unique counter names, finds all entries for that counter name, looks for the highest (or lowest where it matters) value, and then adds only the matching entry to the "highest value" report
$UniqueCounters = ($TempArray | select -Property Counter -Unique).counter
Write-Host -ForegroundColor Green "$(get-date -format hh:mm:ss) - $($UniqueCounters.count) unique counters discovered"
foreach ($c in $UniqueCounters)
{
    $targetEntries = $TempArray | ? {$_.Counter -eq $c}
    if ($c -eq "\PhysicalDisk(*)\% Idle Time" -or $c -eq "\Memory\Available MBytes" -or $c -eq "\Memory\Pool Nonpaged Bytes") {$highValue = ($targetEntries | Measure-Object -Property Value -Minimum).Minimum}
    else {$highValue = ($targetEntries | Measure-Object -Property Value -Maximum).Maximum}
    $selectedEntry = $TempArray | ? {$_.Counter -eq $c -and  $_.Value -eq $highValue}
    if ($selectedEntry.count -gt 1) {$selectedEntry = $selectedEntry[0]}
    $ReportArray += $selectedEntry
}


# Generates a file name based on what you asked the script to do, and dumps it to a CSV for manager-ization later. 
if ($targetProcess -eq "BLANK") {$outLog = $logPath+"$TargetComputer-$startTime-to-$endTime-Results.csv"}
else {$outLog = $logPath+"$TargetComputer-$TargetProcess-$startTime-to-$endTime-Results.csv"}
Write-Host -ForegroundColor Green "$(get-date -format hh:mm:ss) - Writing report to $outLog."
$ReportArray | Export-Csv -Path $outLog -NoClobber -NoTypeInformation -Force
Write-Host -ForegroundColor Green "$(get-date -format hh:mm:ss) - Complete.`n"
