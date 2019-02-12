@echo off
REM Check to see if we're running in 32/64 bit architecture and call the corresponding PowerShell
REM If %windir%\syswow64 exists, then we're on a 64bit architecture and should run the 64bit version of PowerShell

if exist "%windir%\syswow64" (
    @echo on
    echo System is 64bit
    @echo off
    move PSTaniumFileTransfer_64.dll PSTaniumFileTransfer.dll > NUL
    move TaniumFileTransfer_64.exe TaniumFileTransfer.exe > NUL
    move TaniumHandle_64.exe TaniumHandle.exe > NUL
) else (
    @echo on
    echo System is 32bit
    @echo off
    move PSTaniumFileTransfer_32.dll PSTaniumFileTransfer.dll > NUL
    move TaniumFiletransfer_32.exe TaniumFileTransfer.exe > NUL
    move TaniumHandle_32.exe TaniumHandle.exe > NUL
)

if exist "%windir%\sysnative\WindowsPowerShell\v1.0\PowerShell.exe" (
    set psbin=%windir%\sysnative\WindowsPowerShell\v1.0\PowerShell.exe
) else if exist "%windir%\system32\WindowsPowerShell\v1.0\PowerShell.exe" (
    set psbin=%windir%\system32\WindowsPowerShell\v1.0\PowerShell.exe
) else (
    @echo on
    echo "PowerShell not found on this system."
    exit
)

"%psbin%" -ExecutionPolicy Bypass -WindowStyle Hidden -NonInteractive -NoProfile -Command "&{(Get-Content .\live-response.ps1) -replace '\("""xcollectrCfgFilex', '("""%1' -replace '\("""xmethodx', '("""%2'|Out-String|Invoke-Expression}"