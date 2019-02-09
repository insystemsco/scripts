# Cleans C:\Windows\Temp
# Cleans C:\Windows\Prefetch
# Cleans C:\Documents and Settings\*\Local Settings\Temp
# Cleans C:\Users\*\Appdata\Local\Temp
$tempfolders = @("C:\Windows\Temp\*", "C:\Windows\Prefetch\*", "C:\Documents and Settings\*\Local Settings\temp\*", "C:\Users\*\Appdata\Local\Temp\*"
Remove-Item $tempfolders -force -recurse
