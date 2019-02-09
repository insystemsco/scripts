#=======================================================================
#Servers
$Servers = Get-Content -Path "D:\scripts\input.txt"
 
#Location with procdump exe files
$ProcDumpPath = "D:\procdump_script\tmp"
 
#Process name
$ProcessName = "Microsoft.IdentityServer.ServiceHost"
 
#=======================================================================
#Looping each server
Foreach($Server in $Servers){
    $Server = $Server.Trim()
    Write-Host "Processing $Server" -ForegroundColor Green
     
    #Testing path 
    $DestinationPath = "\\$Server\d$\"
    $TestPath = Test-Path $DestinationPath
  
    If(!$TestPath){
        Write-Warning "Failed to access: $DestinationPath"
    }
    Else{
        Write-Host "Copying Procdump files to $DestinationPath"
        $CopyError = $Null
        $Srv = $Null
         
        Try{
            #Copying procdump files to server
            Copy-Item -Recurse $ProcDumpPath -Destination $DestinationPath -Force -ErrorVariable CopyError -ErrorAction Stop
 
            #Creating folder for dumps and getting process id
            If(!$CopyError){
                $Fold = New-Item -Path \\$Server\d$\tmp\dumps -Type directory -Force -ErrorAction Stop
                $Srv = Get-Process -ComputerName $Server -Name $ProcessName -ErrorAction Stop | Select-Object ProcessName,Id
            }
        }
        Catch{
            $_.Exception.Message
            Continue
        }
         
        If($Srv){
            $SrvDump = $Null
            $SrvDumpPath = $Null
 
            Write-Host 'Creating dump file "d:\tmp\dumps\"'
            #Creating procdump and save in "d:\tmp\dumps\"
            Try{
                $SrvDump = Invoke-Command $Server -ErrorAction Stop -ScriptBlock{param($Srv)  
  
                       cmd.exe /c "d:\tmp\procdump.exe"  -ma $Srv.id -accepteula "d:\tmp\dumps\"
  
                } -ArgumentList $Srv
  
                $SrvDumpPath = $SrvDump | Where-Object {$_ -match "initiated"} 
                $SrvDumpPath -Replace  '(\[\d\d:\d\d:\d\d\]\ Dump \d initiated: )'
            }
            Catch{
                $_.Exception.Message
                Continue
            }
        }
    }
}#Foreach loop end
  
Read-Host "`nPress any key to exit..."
exit