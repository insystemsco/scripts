function Get-ADSystem {
          
    [CmdletBinding()]
    [OutputType([Array])] 
    param
    (
        [Parameter(Position=0, Mandatory = $true, HelpMessage="Provide server names", ValueFromPipeline = $true)]
        $Server
    )
 
    $SystemArray = @()
 
        $Server = $Server.trim()
        $Object = '' | Select ServerName, BootUpTime, UpTime, "Physical RAM", "C: Free Space", "Memory Usage", "CPU usage"
                         
        $Object.ServerName = $Server
 
        # Get OS details using WMI query
        $os = Get-WmiObject win32_operatingsystem -ComputerName $Server -ErrorAction SilentlyContinue | Select-Object LastBootUpTime,LocalDateTime
                         
        If($os)
        {
            # Get bootup time and local date time  
            $LastBootUpTime = [Management.ManagementDateTimeConverter]::ToDateTime(($os).LastBootUpTime)
            $LocalDateTime = [Management.ManagementDateTimeConverter]::ToDateTime(($os).LocalDateTime)
 
            # Calculate uptime - this is automatically a timespan
            $up = $LocalDateTime - $LastBootUpTime
            $uptime = "$($up.Days) days, $($up.Hours)h, $($up.Minutes)mins"
 
            $Object.BootUpTime = $LastBootUpTime
            $Object.UpTime = $uptime
        }
        Else
        {
            $Object.BootUpTime = "(null)"
                $Object.UpTime = "(null)"
        }
 
        # Checking RAM, memory and cpu usage and C: drive free space
        $PhysicalRAM = (Get-WMIObject -class Win32_PhysicalMemory -ComputerName $server | Measure-Object -Property capacity -Sum | % {[Math]::Round(($_.sum / 1GB),2)})
                         
        If($PhysicalRAM)
        {
            $PhysicalRAM = ("$PhysicalRAM" + " GB")
            $Object."Physical RAM"= $PhysicalRAM
        }
        Else
        {
            $Object.UpTime = "(null)"
        }
    
        $Mem = (Get-WmiObject -Class win32_operatingsystem -ComputerName $Server  | Select-Object @{Name = "MemoryUsage"; Expression = { “{0:N2}” -f ((($_.TotalVisibleMemorySize - $_.FreePhysicalMemory)*100)/ $_.TotalVisibleMemorySize)}}).MemoryUsage
                        
        If($Mem)
        {
            $Mem = ("$Mem" + " %")
            $Object."Memory Usage"= $Mem
        }
        Else
        {
            $Object."Memory Usage" = "(null)"
        }
 
        $Cpu =  (Get-WmiObject win32_processor -ComputerName $Server  |  Measure-Object -property LoadPercentage -Average | Select Average).Average 
                         
        If($PhysicalRAM)
        {
            $Cpu = ("$Cpu" + " %")
            $Object."CPU usage"= $Cpu
        }
        Else
        {
            $Object."CPU Usage" = "(null)"
        }
 
        $FreeSpace =  (Get-WmiObject win32_logicaldisk -ComputerName $Server -ErrorAction SilentlyContinue  | Where-Object {$_.deviceID -eq "C:"} | select @{n="FreeSpace";e={[math]::Round($_.FreeSpace/1GB,2)}}).freespace 
                         
        If($FreeSpace)
        {
            $FreeSpace = ("$FreeSpace" + " GB")
            $Object."C: Free Space"= $FreeSpace
        }
        Else
        {
            $Object."C: Free Space" = "(null)"
        }
 
        $SystemArray += $Object
  
        $SystemArray
} 