:: TR3 - rewrite to be effective on newer windows versions. And use
::	Script start...
@echo off
::	Change Directory to the script's one, since Stupid windows will execute at "\windows\system32" when chosen to run as administrator...
cd /d "%~d0%~p0"

::Get Date and Time
For /F "tokens=1,2,3,4 delims=/ " %%A in ('Date /t') do @( 
Set FullDate="%%D-%%C-%%B"
)

For /F "tokens=1,2,3 delims=: " %%A in ('time /t') do @( 
Set FullTime="%%A-%%B"
)

set dirname="%computername%_%FullDate%_%FullTime%"
mkdir "%dirname%"
cd "%dirname%"
set var=%computername%.txt


type ..\commands.txt >> %var%

for /f "tokens=*" %%a in (..\commands.txt) do @(
echo. >> %var%
echo. >> %var%
echo ======================================== >> %var%
echo ======================================== >> %var%
echo %%a >> %var%
echo ======================================== >> %var%
echo ======================================== >> %var%
cmd /c "%%a" >> %var%
)

::extract tools
"c:\program files (x86)"\tanium\tanium clienttools\StdUtils\7za.exe e "c:\program files (x86)\tanium\tanium clienttools\tr3\deps.zip"  -o"\%dirname%\" > %dirname%\deps.log

cd ..
cd Tools
autorunsc.exe -a -f -c >> ..\%dirname%\autoruns.csv
handle.exe -a -u >> ..\%dirname%\handles.txt
HiJackThis.EXE /silentautolog
move hijackthis.log ..\%dirname%\
listdlls.exe >> ..\%dirname%\dlls.txt
psinfo -s -h -d >> ..\%dirname%\psinfo.txt
pslist -t >> ..\%dirname%\pslist.txt
psgetsid >> ..\%dirname%\psgetsid.txt
psloggedon.exe >> ..\%dirname%\psloggedon.txt
Tcpvcon.exe -a -c -n >> ..\%dirname%\tcpvcon.txt
handle.exe >> ..\%dirname%\handle.txt

::Get system info
@echo off
echo +++++Live Response Collection+++++ >> %computername%.txt
echo +++++Date and Time+++++ >> %computername%.txt
date /t >> %computername%.txt
time /t >> %computername%.txt
echo +++++System Information: Time Zone, Installed Software, OS Version, Uptime, File System+++++ >> %computername%.txt
systeminfo >> %computername%.txt
echo +++++User Accounts+++++ >> %computername%.txt
net user >> %computername%.txt
echo +++++Groups+++++ >> %computername%.txt
net localgroup >> %computername%.txt
echo +++++Network Interfaces+++++ >> %computername%.txt
ipconfig /all >> %computername%.txt
echo +++++Routing Table+++++ >> %computername%.txt
route print >> %computername%.txt
echo +++++ARP Table+++++ >> %computername%.txt
arp -a >> %computername%.txt
echo +++++DNS Cache+++++ >> %computername%.txt
ipconfig /displaydns >> %computername%.txt
echo +++++Network Connections+++++ >> %computername%.txt
netstat -abn >> %computername%.txt
