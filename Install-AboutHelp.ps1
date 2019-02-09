<# 

.DESCRIPTION 
 Fixes the about_help topics missing in WMF5+ installs 

#> 
Param()

Invoke-WebRequest -Uri https://github.com/kilasuit/Install-AboutHelp/raw/master/about_help.zip -OutFile $env:TEMP\About_help.zip

Expand-Archive $env:TEMP\About_help.zip C:\Windows\System32\WindowsPowerShell\v1.0\en\ -Force

Remove-Item $env:TEMP\About_help.zip
