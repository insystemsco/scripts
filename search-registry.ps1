<#
    Description: 
        This script will accept a parameter of a registry key and return all child values with last modified time of the specified registry key and/or sub keys.
    
    Usage:
        Parameter must be in the format of (HKLM:\ or HKCU:\)
		You can choose to match regex patterns and return only registry keys/value/data that match a specific regex pattern. Default will get everything.
#>

param (
	[Parameter(Mandatory = $true)]
	$RegistryKeyQuery,
	$MatchPattern
)

#Checks Match Pattern to have data, if it doesn't, everything under the specified key will be returned.'
if (($MatchPattern -eq $null) -or ($MatchPattern -eq ""))
{
	$MatchPattern = ".*"
}

#Below JSON Conversion from (https://gist.github.com/mdnmdn/6936714)
function Escape-JSONString($str)
{
	if ($str -eq $null) { return "" }
	$str = $str.ToString().Replace('"', '\"').Replace('\', '\\').Replace("`n", '\n').Replace("`r", '\r').Replace("`t", '\t')
	return $str;
}

#Below JSON Conversion from (https://gist.github.com/mdnmdn/6936714)
function ConvertTo-JSON-V2($maxDepth = 4, $forceArray = $false)
{
	begin
	{
		$data = @()
	}
	process
	{
		$data += $_
	}
	
	end
	{
		
		if ($data.length -eq 1 -and $forceArray -eq $false)
		{
			$value = $data[0]
		}
		else
		{
			$value = $data
		}
		
		if ($value -eq $null)
		{
			return "null"
		}
		
		
		
		$dataType = $value.GetType().Name
		
		switch -regex ($dataType)
		{
			'String'  {
				return "`"{0}`"" -f (Escape-JSONString $value)
			}
			'(System\.)?DateTime'  { return "`"{0:yyyy-MM-dd}T{0:HH:mm:ss}`"" -f $value }
			'Int32|Double' { return "$value" }
			'Boolean' { return "$value".ToLower() }
			'(System\.)?Object\[\]' {
				# array
				
				if ($maxDepth -le 0) { return "`"$value`"" }
				
				$jsonResult = ''
				foreach ($elem in $value)
				{
					#if ($elem -eq $null) {continue}
					if ($jsonResult.Length -gt 0) { $jsonResult += ', ' }
					$jsonResult += ($elem | ConvertTo-JSON-V2 -maxDepth ($maxDepth - 1))
				}
				return "[" + $jsonResult + "]"
			}
			'(System\.)?Hashtable' {
				# hashtable
				$jsonResult = ''
				foreach ($key in $value.Keys)
				{
					if ($jsonResult.Length -gt 0) { $jsonResult += ', ' }
					$jsonResult +=
					@"
	"{0}": {1}
"@ -f $key, ($value[$key] | ConvertTo-JSON-V2 -maxDepth ($maxDepth - 1))
				}
				return "{" + $jsonResult + "}"
			}
			default
			{
				#object
				if ($maxDepth -le 0) { return "`"{0}`"" -f (Escape-JSONString $value) }
				
				return "{" +
				(($value | Get-Member -MemberType *property | % {
							@"
	"{0}": {1}
"@ -f $_.Name, ($value.($_.Name) | ConvertTo-JSON-V2 -maxDepth ($maxDepth - 1))
							
						}) -join ', ') + "}"
			}
		}
	}
}


#Function is from TechNet Rohn Edwards (https://gallery.technet.microsoft.com/scriptcenter/Get-Last-Write-Time-and-06dcf3fb)
function Add-RegKeyMember
{
<#
.SYNOPSIS
Adds note properties containing the last modified time and class name of a 
registry key.

.DESCRIPTION
The Add-RegKeyMember function uses the unmanged RegQueryInfoKey Win32 function
to get a key's last modified time and class name. It can take a RegistryKey 
object (which Get-Item and Get-ChildItem output) or a path to a registry key.

.EXAMPLE
PS> Get-Item HKLM:\SOFTWARE | Add-RegKeyMember | Select Name, LastWriteTime

Show the name and last write time of HKLM:\SOFTWARE

.EXAMPLE
PS> Add-RegKeyMember HKLM:\SOFTWARE | Select Name, LastWriteTime

Show the name and last write time of HKLM:\SOFTWARE

.EXAMPLE
PS> Get-ChildItem HKLM:\SOFTWARE | Add-RegKeyMember | Select Name, LastWriteTime

Show the name and last write time of HKLM:\SOFTWARE's child keys

.EXAMPLE
PS> Get-ChildItem HKLM:\SYSTEM\CurrentControlSet\Control\Lsa | Add-RegKeyMember | where classname | select name, classname

Show the name and class name of child keys under Lsa that have a class name defined.

.EXAMPLE
PS> Get-ChildItem HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall | Add-RegKeyMember | where lastwritetime -gt (Get-Date).AddDays(-30) | 
>> select PSChildName, @{ N="DisplayName"; E={gp $_.PSPath | select -exp DisplayName }}, @{ N="Version"; E={gp $_.PSPath | select -exp DisplayVersion }}, lastwritetime |
>> sort lastwritetime

Show applications that have had their registry key updated in the last 30 days (sorted by the last time the key was updated).
NOTE: On a 64-bit machine, you will get different results depending on whether or not the command was executed from a 32-bit
      or 64-bit PowerShell prompt.

#>
	
	[CmdletBinding()]
	param (
		[Parameter(Mandatory = $true, ParameterSetName = "ByKey", Position = 0, ValueFromPipeline = $true)]
		[ValidateScript({ $_ -is [Microsoft.Win32.RegistryKey] })]
		# Registry key object returned from Get-ChildItem or Get-Item. Instead of requiring the type to
		# be [Microsoft.Win32.RegistryKey], validation has been moved into a [ValidateScript] parameter
		# attribute. In PSv2, PS type data seems to get stripped from the object if the [RegistryKey]
		# type is an attribute of the parameter.
		$RegistryKey,
		[Parameter(Mandatory = $true, ParameterSetName = "ByPath", Position = 0)]
		# Path to a registry key
		[string]$Path
	)
	
	begin
	{
		# Define the namespace (string array creates nested namespace):
		$Namespace = "CustomNamespace", "SubNamespace"
		
		# Make sure type is loaded (this will only get loaded on first run):
		Add-Type @"
            using System; 
            using System.Text;
            using System.Runtime.InteropServices; 

            $($Namespace | ForEach-Object {
				"namespace $_ {"
			})

                public class advapi32 {
                    [DllImport("advapi32.dll", CharSet = CharSet.Auto)]
                    public static extern Int32 RegQueryInfoKey(
                        IntPtr hKey,
                        StringBuilder lpClass,
                        [In, Out] ref UInt32 lpcbClass,
                        UInt32 lpReserved,
                        out UInt32 lpcSubKeys,
                        out UInt32 lpcbMaxSubKeyLen,
                        out UInt32 lpcbMaxClassLen,
                        out UInt32 lpcValues,
                        out UInt32 lpcbMaxValueNameLen,
                        out UInt32 lpcbMaxValueLen,
                        out UInt32 lpcbSecurityDescriptor,
                        out Int64 lpftLastWriteTime
                    );

                    [DllImport("advapi32.dll", CharSet = CharSet.Auto)]
                    public static extern Int32 RegOpenKeyEx(
                        IntPtr hKey,
                        string lpSubKey,
                        Int32 ulOptions,
                        Int32 samDesired,
                        out IntPtr phkResult
                    );

                    [DllImport("advapi32.dll", CharSet = CharSet.Auto)]
                    public static extern Int32 RegCloseKey(
                        IntPtr hKey
                    );
                }
            $($Namespace | ForEach-Object { "}" })
"@
		
		# Get a shortcut to the type:    
		$RegTools = ("{0}.advapi32" -f ($Namespace -join ".")) -as [type]
	}
	
	process
	{
		switch ($PSCmdlet.ParameterSetName)
		{
			"ByKey" {
				# Already have the key, no more work to be done :)
			}
			
			"ByPath" {
				# We need a RegistryKey object (Get-Item should return that)
				$Item = Get-Item -Path $Path -ErrorAction Stop
				
				# Make sure this is of type [Microsoft.Win32.RegistryKey]
				if ($Item -isnot [Microsoft.Win32.RegistryKey])
				{
					throw "'$Path' is not a path to a registry key!"
				}
				$RegistryKey = $Item
			}
		}
		
		# Initialize variables that will be populated:
		$ClassLength = 255 # Buffer size (class name is rarely used, and when it is, I've never seen 
		# it more than 8 characters. Buffer can be increased here, though. 
		$ClassName = New-Object System.Text.StringBuilder $ClassLength # Will hold the class name
		$LastWriteTime = $null
		
		# Get a handle to our key via RegOpenKeyEx (PSv3 and higher could use the .Handle property off of registry key):
		$KeyHandle = New-Object IntPtr
		
		if ($RegistryKey.Name -notmatch "^(?<hive>[^\\]+)\\(?<subkey>.+)$")
		{
			Write-Error ("'{0}' not a valid registry path!")
			return
		}
		
		$HiveName = $matches.hive -replace "(^HKEY_|_|:$)", "" # Get hive in a format that [RegistryHive] enum can handle
		$SubKey = $matches.subkey
		
		# Get hive. $HiveName should contain a valid MS.Win32.RegistryHive enum, but it will be in all caps. It seems that
		# [enum]::IsDefined is case sensitive, so that won't work. There's an awesome static method [enum]::TryParse, but it
		# appears that it was introduced in .NET 4. So, I'm just wrapping it in a try {} block:
		try
		{
			$Hive = [Microsoft.Win32.RegistryHive]$HiveName
		}
		catch
		{
			Write-Error ("Unknown hive: {0} (Registry path: {1})" -f $HiveName, $RegistryKey.Name)
			return # Exit function or we'll get an error in RegOpenKeyEx call
		}
		
		Write-Verbose ("Attempting to get handle to '{0}' using RegOpenKeyEx" -f $RegistryKey.Name)
		switch ($RegTools::RegOpenKeyEx(
				$Hive.value__,
				$SubKey,
				0, # Reserved; should always be 0
				[System.Security.AccessControl.RegistryRights]::ReadKey,
				[ref]$KeyHandle
			))
		{
			0 {
				# Success
				# Nothing required for now
				Write-Verbose "  -> Success!"
			}
			
			default
			{
				# Unknown error!
				Write-Error ("Error opening handle to key '{0}': {1}" -f $RegistryKey.Name, $_)
			}
		}
		
		switch ($RegTools::RegQueryInfoKey(
				$KeyHandle,
				$ClassName,
				[ref]$ClassLength,
				$null, # Reserved
				[ref]$null, # SubKeyCount
				[ref]$null, # MaxSubKeyNameLength
				[ref]$null, # MaxClassLength
				[ref]$null, # ValueCount
				[ref]$null, # MaxValueNameLength 
				[ref]$null, # MaxValueValueLength 
				[ref]$null, # SecurityDescriptorSize
				[ref]$LastWriteTime
			))
		{
			
			0 {
				# Success
				$LastWriteTime = [datetime]::FromFileTime($LastWriteTime)
				
				# Add properties to object and output them to pipeline
				$RegistryKey |
				Add-Member -MemberType NoteProperty -Name LastWriteTime -Value $LastWriteTime -Force -PassThru |
				Add-Member -MemberType NoteProperty -Name ClassName -Value $ClassName.ToString() -Force -PassThru
			}
			
			122  {
				# ERROR_INSUFFICIENT_BUFFER (0x7a)
				throw "Class name buffer too small"
				# function could be recalled with a larger buffer, but for
				# now, just exit
			}
			
			default
			{
				throw "Unknown error encountered (error code $_)"
			}
		}
		
		# Closing key:
		Write-Verbose ("Closing handle to '{0}' using RegCloseKey" -f $RegistryKey.Name)
		switch ($RegTools::RegCloseKey($KeyHandle))
		{
			0 {
				# Success, no action required
				Write-Verbose "  -> Success!"
			}
			default
			{
				Write-Error ("Error closing handle to key '{0}': {1}" -f $RegistryKey.Name, $_)
			}
		}
	}
}

#Checking for large registry query (HKLM:\Software) 

if ($RegistryKeyQuery -eq "HKLM:\Software")
{
	Write-Warning "You have chosen to get all subkeys under HKLM:\Software`n`nThis will take a long time to complete.`n`nAre you sure you want to proceed?"
	
	$Proceed = Read-Host "Please type ""Y"" for Yes or ""N"" for No"
	
}
else
{
	$Proceed = "Y"
}

if ($Proceed -eq "Y")
{
	$Continue = 0
	
	#Checking for HKU:\ and creating Drive
	if ($RegistryKeyQuery -match "HKU:\.*")
	{
		Remove-PSDrive HKU -Force -ErrorAction SilentlyContinue
		New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS
	}
	
	#Checks for Registry Key matching proper format.
	if ($RegistryKeyQuery -notmatch ".*:\\.*")
	{
		Write-warning "Registry Path is incorrect.`nPlease use this format when specifying the registry key`n`nHKLM:\Path\To\Key"
	}
	else
	{
		#Tests Registry Key Path to ensure valid key
		if (!(Test-Path "$RegistryKeyQuery"))
		{
			Write-Warning "Registry key ($RegistryKeyQuery) is invalid. Please specify an existing registry key"
		}
		else
		{
			Write-Host "$RegistryKeyQuery is valid`n`nBegin Registry Search"
			$Continue = 1
		}
	}
	
	#If path is good, continue
	
	if ($Continue -eq 1)
	{
		
		
		Write-Host "Begin recursive search for $RegistryKeyQuery" -ForegroundColor Green
		
		#Sets Headers for CSV (if CSV is your export option)
		
		$CSVHeaders = "Key Name,Key Value,Key Data,Last Write Time,Key Value Type"
		
		$CSVData = @()
		
		#Gets data from Root key and All SubKeys
		$RootKey = Get-Item "$RegistryKeyQuery"
		$SubKeys = Get-ChildItem "$RegistryKeyQuery" -Recurse
		
		#Gets last write time of the Root Key
		
		
		#Checks for number of values in root registry key
		if ($RootKey.Property.count -gt 1)
		{
			#Goes through each value and creates PSObject with information regarding the value
			foreach ($Value in $RootKey.Property)
			{
				if (($RootKey.Name -match $MatchPattern) -or ($Value -match $MatchPattern) -or (($($RootKey.GetValue($Value)) -join "|") -match $MatchPattern))
				{
					$LastWriteTime = $RootKey | Add-RegKeyMember | Select Name, LastWriteTime
					$ObjInput = New-Object System.Object
					$ObjInput | Add-Member -MemberType NoteProperty -Name "Key Name" -Value $($RootKey.Name)
					$ObjInput | Add-Member -MemberType NoteProperty -Name "Key Value" -Value $($Value)
					#Handles REG_MULTI_SZ            
					if ($RootKey.GetValue($Value).count -gt 1)
					{
						$ObjInput | Add-Member -MemberType NoteProperty -Name "Key Value Data" -Value ($($RootKey.GetValue($Value)) -join "|") -Force
					}
					else
					{
						$ObjInput | Add-Member -MemberType NoteProperty -Name "Key Value Data" -Value $($RootKey.GetValue($Value)) -Force
					}
					$ObjInput | Add-Member -MemberType NoteProperty -Name "Key Last Write Time" -Value $($LastWriteTime.LastWriteTime)
					$ObjInput | Add-Member -MemberType NoteProperty -Name "Key Value Type" -Value ($RootKey.GetValueKind($Value)).ToString() -ErrorAction SilentlyContinue
					$CSVData += $ObjInput
					$ObjInput
				}
			}
		}
		else
		{
			if (($RootKey.Name -match $MatchPattern) -or ($RootKey.Property -match $MatchPattern) -or (($($RootKey.GetValue($RootKey.Property)) -join "|") -match $MatchPattern))
			{
				$LastWriteTime = $RootKey | Add-RegKeyMember | Select Name, LastWriteTime
				#there is 1 or less values, display said value
				$ObjInput = New-Object System.Object
				$ObjInput | Add-Member -MemberType NoteProperty -Name "Key Name" -Value $($RootKey.Name)
				
				
				#Checks for Values existing in registry key
				if ($RootKey.Property -ge 1)
				{
					$ObjInput | Add-Member -MemberType NoteProperty -Name "Key Value" -Value $($RootKey.Property)
					#Handles REG_MULTI_SZ
					if (($RootKey.GetValue($RootKey.Property)).count -gt 1)
					{
						$ObjInput | Add-Member -MemberType NoteProperty -Name "Key Value Data" -Value ($($RootKey.GetValue($RootKey.Property)) -join "|") -Force
					}
					else
					{
						$ObjInput | Add-Member -MemberType NoteProperty -Name "Key Value Data" -Value $($RootKey.GetValue($RootKey.Property)) -Force
					}
					$ObjInput | Add-Member -MemberType NoteProperty -Name "Key Value Type" -Value ($RootKey.GetValueKind($RootKey.Property)).ToString()
				}
				else
				{
					$ObjInput | Add-Member -MemberType NoteProperty -Name "Key Value" -Value "null"
					$ObjInput | Add-Member -MemberType NoteProperty -Name "Key Value Data" -Value "null"
				}
				$ObjInput | Add-Member -MemberType NoteProperty -Name "Key Last Write Time" -Value $($LastWriteTime.LastWriteTime)
				
				$CSVData += $ObjInput
				$ObjInput
			}
		}
		
		#Parses through sub keys of root key
		foreach ($Key in $SubKeys)
		{
			
			#Gets Last write Time of Registry KEY (Not value)
			
			
			#Checks for number of values in sub registry key
			if ($Key.Property.count -gt 1)
			{
				
				#Goes through each value and creates PSObject with information regarding the value
				foreach ($Value in $Key.Property)
				{
					if (($Key.Name -match $MatchPattern) -or ($Value -match $MatchPattern) -or (($($Key.GetValue($Value)) -join "|") -match $MatchPattern))
					{
						$LastWriteTime = $Key | Add-RegKeyMember | Select Name, LastWriteTime
						$ObjInput = New-Object System.Object
						$ObjInput | Add-Member -MemberType NoteProperty -Name "Key Name" -Value $($Key.Name)
						$ObjInput | Add-Member -MemberType NoteProperty -Name "Key Value" -Value $($Value)
						#Handles REG_MULTI_SZ            
						if ($Key.GetValue($Value).count -gt 1)
						{
							$ObjInput | Add-Member -MemberType NoteProperty -Name "Key Value Data" -Value ($($Key.GetValue($Value)) -join "|") -Force
						}
						else
						{
							$ObjInput | Add-Member -MemberType NoteProperty -Name "Key Value Data" -Value $($Key.GetValue($Value)) -Force
						}
						$ObjInput | Add-Member -MemberType NoteProperty -Name "Key Last Write Time" -Value $($LastWriteTime.LastWriteTime)
						$ObjInput | Add-Member -MemberType NoteProperty -Name "Key Value Type" -Value ($(try { ($Key.GetValueKind($Value)).ToString() }
								catch [exception]{ "Default value - String" }))
						
						$CSVData += $ObjInput
						$ObjInput
					}
				}
			}
			else
			{
				if (($Key.Name -match $MatchPattern) -or ($Key.Property -match $MatchPattern) -or (($($Key.GetValue($Key.Property)) -join "|") -match $MatchPattern))
				{
					$LastWriteTime = $Key | Add-RegKeyMember | Select Name, LastWriteTime
					#there is 1 or less values, display said value
					
					$ObjInput = New-Object System.Object
					$ObjInput | Add-Member -MemberType NoteProperty -Name "Key Name" -Value $($Key.Name)
					
					#Checks for Values existing in registry key
					if ($Key.Property -ge 1)
					{
						$ObjInput | Add-Member -MemberType NoteProperty -Name "Key Value" -Value $($Key.Property)
						if (($Key.GetValue($Key.Property)).count -gt 1)
						{
							$ObjInput | Add-Member -MemberType NoteProperty -Name "Key Value Data" -Value ($($Key.GetValue($Key.Property)) -join "|") -Force
						}
						else
						{
							$ObjInput | Add-Member -MemberType NoteProperty -Name "Key Value Data" -Value $($Key.GetValue($Key.Property)) -Force
						}
						$ObjInput | Add-Member -MemberType NoteProperty -Name "Key Value Type" -Value ($Key.GetValueKind($Key.Property)).ToString()
					}
					else
					{
						$ObjInput | Add-Member -MemberType NoteProperty -Name "Key Value" -Value "null"
						$ObjInput | Add-Member -MemberType NoteProperty -Name "Key Value Data" -Value "null"
					}
					$ObjInput | Add-Member -MemberType NoteProperty -Name "Key Last Write Time" -Value $($LastWriteTime.LastWriteTime)
					
					$CSVData += $ObjInput
					$ObjInput
				}
			}
		}
		
		if (($CSVData | Out-String) -eq "")
		{
			Write-Warning "Nothing found in registry that matches $MatchPattern"
		}
		else
		{
			
			Write-Host "Exporting to JSON`n`nView of Export:"
			
			
			try
			{
				#Displays Data that will be exported in PSObject Format
				#$CSVData
				
				#Checks for C:\temp existing and creates folder if it doesn't exist.
				if (!(test-path "C:\temp"))
				{
					Write-Host "C:\temp does not exist, creating directory"
					
					New-Item -Name temp -Path C:\ -ItemType Directory -Force
					
				}
				
				#Exports PSObject to JSON 
				$CSVData | ConvertTo-JSON-V2 | Out-File "C:\temp\registry-Export.json"
				
				#$CSVData | Export-Csv -NoTypeInformation -Path "c:\temp\Regsitry-Export.csv" -Force
				Write-Host "Data Successfully Exported!`n`nFile is located in C:\temp\Registry-Export.json" -ForegroundColor Green
			}
			catch [exception]
			{
				Write-Error -Message "Failed to Export Data`n`nERROR: $_"
			}
		}
	}
	else
	{
		#End Script without exiting powershell
	}
}
else
{
	Write-Warning "$Proceed was typed... Ending Script.`n`nTo Proceed, Please type 'Y'"
}