@echo off
Title ******Disk Clean Up******
:: full junk clean up and restarts computer
echo --------------------------------------
echo !!!!Clean up in progress!!!!
echo --------------------------------------
:: kills chrome to enable cache file clean up
taskkill /F /IM "chrome.exe"
echo --------------------------------------
echo Beginning clean up
IF EXIST "C:\Users\" (
    for /D %%x in ("C:\Users\*") do (
		del /q /s /f "%%x\AppData\Local\Google\Chrome\User Data\Default\cache\*.*"
	    del /f /s /q "%%x\AppData\Local\Temp\*.*"
        del /f /s /q "%%x\AppData\Local\Microsoft\Windows\Temporary Internet Files\*.*"
        del /f /s /q "C:\Windows\Prefetch\*.*"
        del /f /s /q "C:\Windows\Temp\*.*"
		del /f /s /q "C:\WINDOWS\pchealth\ERRORREP\UserDumps\*.*"
    )
)
echo --------------------------------------
:: run disk clean up manager to clear system temp files and anything missed above.
echo Running Disk Cleanup
cd %systemroot%\
cleanmgr.exe /verylowdisk /d c
echo --------------------------------------
echo -------- !!Tasks Complete!!  --------
:end
exit