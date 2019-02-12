<#
LiveResponse

Example command line:
PowerShell -NoProfile -ExecutionPolicy Bypass -Noninteractive -WindowStyle Hidden .\live-response.ps1 -methodParam smb -collectrCfgFileParam Standard_Collection
#>

# Parameters below are required for Pester testing.
# Command Line arguemnts can't be used in production because we're passing
# this script to IEX as a string, so we have to do a find and replace. 
param (
  [Parameter(Mandatory=$False,Position=1)][string]$collectrCfgFileParam='Extended_Collection',
  [Parameter(Mandatory=$false,Position=2)][string]$methodParam='SMB'
)

# Stop when cmdlets fail, use try/catch/finally to handle
$ErrorActionPreference = "Stop"

# Code below sets parameters to command line parameters for pester tests
# For production runs, variables like $x...x will be replaced at run time
$collectrCfgFile = [System.Uri]::UnescapeDataString("xcollectrCfgFilex")
if ($collectrCfgFile -match 'xcollectrCfgFilex') {
  $collectrCfgFile = $collectrCfgFileParam
}
$collectrCfgFile = $collectrCfgFile + '.json'

$method = [System.Uri]::UnescapeDataString("xmethodx")
if ($method -match 'xmethodx') {
  $method = $methodParam
}
$method = $method + '.json'

function LoadPowerForensics {
  $PFTanium = $LRScriptRoot + "\powerforensics.dll"

  if ( -Not (Test-Path -Path $PFTanium)) {
    $Message = "PowerForensics module not found."
    Write-Log -Message $Message -Level 'Fatal'
    MoveLog @params
    throw $Message
  }
  try {
    # PowerForensics is a dll based module, so Import-Module should work
    Import-Module $PFTanium 
  } catch {
    $Message = "PowerForensics failed to import. Error: $_"
    Write-Log -Message $Message -Level 'Fatal'
    MoveLog @params
    throw $Message
  }
}

function MoveLog {
  param(
    $dataDest
  )
  Write-Log -Message ('Moving {0} to {1}' -f $LogFile,$dataDest)
  Write-Log -Message ('Live Response processing complete.')
  $oldPosition = $LogStream.BaseStream.Position
  $LogStream.BaseStream.Position = 0
  try {
    $DestLogFile = $(Get-Item $Logfile).Name
    $TFT.SendStream($LogStream.BaseStream,"${dataDest}/${DestLogFile}",(Get-Date),$LogStream.BaseStream.Length)
    $LogStream.BaseStream.Position = $oldPosition
  } catch {
    $LogStream.BaseStream.Position = $oldPosition
    $Message = ('Failed to move {0} to {1}. Error: {2}' -f $LogFile, $(${dataDest} + '/' +${DestLogFile}), $_)
    Write-Log -Message $Message
    $Message # This will write to the action log
  }
}

function RunLiveResponse {
  # Commenting out the bits below because they can have unintended consequences
  # In general setting thread priority is a bad idea as reducing thread priority 
  # can introduce latencies elsewhere in the system as the system may have to wait
  # for the lower priority thread to release resources
  # "Setting Thread Priority.........."
  # [System.Threading.Thread]::CurrentThread.Priority = 'Lowest'

  # Parse the config into collectors and files

  # set up output directory

  try {
    $collectrCfgJson = ConvertFrom-Json -Path $collectrCfgFile
  } catch {
    $Message = "Failed to parse $collectrCfgFile. Error: $_"
    Write-Log -Message $Message -Level 'Fatal'
    MoveLog @params 
    throw $Message
  }
  Write-Log -Message "$collectrCfgFile parsing complete." -Level 'Info'
  Import-GlobalOptions $collectrCfgJson.Options
  try {
    $scriptCollectors = $collectrCfgJson.scripts | Sort-Object -Property { [int32]$_.order }
  } catch {
    $Message = "Failed to parse script collectors. Error: $_"
    Write-Log -Message $Message -Level 'Fatal'
    MoveLog @params
    throw $Message
  }
  Write-Log -Message "Script collector parsing complete." -Level 'Info'

  try {
    $moduleCollectors = $collectrCfgJson.modules | Sort-Object -Property { [int32]$_.order }
  } catch {
    $Message = "Failed to parse module collectors. Error: $_"
    Write-Log -Message $Message -Level 'Fatal'
    MoveLog @params
    throw $Message
  }
  Write-Log -Message "Module collector parsing complete." -Level 'Info'

  try {
    $fileCollectors = $collectrCfgJson.files | Sort-Object -Property { [int32]$_.order }
  } catch {
    $Message = "Failed to parse file collectors. Error: $_"
    Write-Log -Message $Message -Level 'Fatal'
    MoveLog @params
    throw $Message
  }
  Write-Log -Message "File collector parsing complete." -Level 'Info'

  if ($moduleCollectors) {
    Write-Log -Message "Running Module Collectors.........." -Level 'Info'
    Invoke-ModuleCollectors -moduleCollectors $moduleCollectors @params
  } else {
    Write-Log -Message "No Module Collectors configured" -Level 'Info'
  }

  if ($scriptCollectors) {
    Write-Log -Message "Running Script Collectors.........." -Level 'Info'
    Invoke-ScriptCollectors -ScriptCollectors $scriptCollectors @params
  } else {
    Write-Log -Message "No Script Collectors configured" -Level 'Info'
  }
  
  if ($fileCollectors) {
    Write-Log -Message "Running File Collectors.........." -Level 'Info'
    Invoke-FileCollectors -fileCollectors $fileCollectors -collectionDate $datetime @params
  } else {
    Write-Log -Message "No File Collectors configured" -Level 'Info'
  }
  
  MoveLog @params
} # End RunLiveResponse

function GetTaniumDir {
  # Return Tanium Client Path from the Registry or $null
  $Error.Clear()
  $TaniumDir = $null
  if ($TaniumDir -eq $null) { if ((Test-Path 'HKLM:\Software\Tanium\Tanium Client')) {$TaniumDir = Get-ItemProperty 'HKLM:\Software\Tanium\Tanium Client' }}
  if ($TaniumDir -eq $null) { if ((Test-Path 'HKLM:\Software\Wow6432Node\Tanium\Tanium Client')) {$TaniumDir = Get-ItemProperty 'HKLM:\Software\Wow6432Node\Tanium\Tanium Client' }}
  if ($TaniumDir -eq $null) { if ((Test-Path 'HKLM:\Software\McAfee\Real Time')) {$TaniumDir = Get-ItemProperty 'HKLM:\Software\McAfee\Real Time' }}
  if ($TaniumDir -eq $null) { if ((Test-Path 'HKLM:\Software\Wow6432Node\McAfee\Real Time')) {$TaniumDir = Get-ItemProperty 'HKLM:\Software\Wow6432Node\McAfee\Real Time' }}
  if ($TaniumDir -eq $null -and $Error) {
    $Message = "Failed to find Tanium Client directory."
    Write-Log -Message $Message -Level 'Fatal'
    Write-Log -Message $Error[0].Message -Level 'Fatal'
    MoveLog
    throw "$Message Error: $_"
  }
  $Error.Clear()
  $TaniumDir | Select-Object -ExpandProperty Path
} # End GetTaniumDir

function Add-HashTableEntry {
  param (
    $HashTable,
    $Name,
    $Value
  )
  if ($HashTable.ContainsKey($Name)) {
    $HashTable.$Name = $Value
  } else {
    $HashTable.Add($Name, $Value)
  }
}

function Import-GlobalOptions {
  # Sets up the global options from the config file
  param(
    $Options
  )
  foreach ($item in $Options.PSObject.Properties) {
      Add-HashTableEntry -HashTable $script:options -Name $item.Name -Value $item.Value
  }
  Write-Log -Level 'Debug' "Global Config set: $(FormatHashTable $script:options)"
}

function GetArgs {
  # Makes hash table used for calling some of the various functions in this module.
  # Intended to make merging options from various scopes easy.
  param (
    $Collector,
    $Options = $script:options,
    $Destination
  )
  $ht = @{}
  #Merge global and local settings, eventually can be easily expanded to 'section'-level as well.
  foreach ($setting in $Options.GetEnumerator() + $Collector.PSObject.Properties) {
    $name = $setting.Name -split '_' -join ''
    Add-HashTableEntry -HashTable $ht -Name $name -Value $setting.Value
  }
  if ($ht.'Copy' -eq $True -and -not [String]::IsNullOrEmpty($Destination)) {
    Add-HashTableEntry -HashTable $ht -Name 'dest' -Value $Destination
  }
  $ht.Remove('Copy')
  #Convert snake_case config file values into PascalCase

  Write-Log -Level 'Debug' "Args: $(FormatHashTable $ht)"
  $ht
}

function FormatHashTable {
  param (
    $ht
  )
  ($ht.GetEnumerator() | Foreach-Object {"`r`n$($_.Name) = $($_.Value)"})
}
<#
function LoadSQLite {
  if ( -Not (Test-Path .\sqlite.psd1)) {
    throw "SQLite module not found. Verify sqlite.psd1 is in the package. Error: $_"
  }
  try {
    Import-Module \sqlite.psd1
  } catch {
    throw "SQLite failed to import. Error: $_"
  }
}
#>

function Get-FileDetails {
  # Rewrite of Get-FileDetails
  param (
    [String]$Path,
    [Boolean]$DiskInfo,
    [Array]$Hashes,
    [String]$Dest="",
    [Boolean]$Raw = $false,
    [Boolean]$RawFallback = $false,
    [HashTable]$Cache = $null
  )

  # If a cache is provided and we have a cached result, then return that. 
  # Note that our key is a combination of the Path and Destination.
  $cacheKey = "$Path-$Dest".ToLower()
  if ($Cache -ne $null -and $Cache.ContainsKey($cacheKey)) {
    Write-Log -Level "Debug" "Returning cached file details for key '$cacheKey'"
    return $Cache[$cacheKey]
  }

  Write-Log "Processing $Path" -Level 'Info'
  $hashesToProcess = @()
  $record = $null

  #Default file record. Will be re-used for output hashtable.
  $defaultFile = @{
    "FilePath" = $Path
    "SIModifiedTime" = 'FALSE'
    "SIAccessedTime" = 'FALSE'
    "SIChangedTime" = 'FALSE'
    "SIBornTime" = 'FALSE'
    "FNModifiedTime" = 'FALSE'
    "FNAccessedTime" = 'FALSE'
    "FNChangedTime" = 'FALSE'
    "FNBornTime" = 'FALSE'
    "HashMD5" = "FALSE"
    "HashSHA1" = "FALSE"
    "HashSHA256" = "FALSE"
    "Copied" = "FALSE"
  }
  if ($Raw) {
    try {
      Write-Log "Getting forensic file record for $Path" -Level 'Debug'
      $record = Get-ForensicFileRecord -Path $Path
    } catch {
      Write-Log -Message "Failed to get forensic file record for $Path. Error: $_" -Level 'Warn'
    }
    try {
      $recordString = $record.ToString()
    } catch {
      Write-Log -Message "Failed to get forensic file record string for $Path. Error: $_" -Level 'Warn'
      $recordString = $null
    }
    if ([String]::IsNullOrEmpty($recordString) -or $recordString.StartsWith("[0]") -or $recordString.StartsWith("[Directory]")) {
      # File is 0 bytes, a Directory or we had an issue getting its size - Disable hashing.
      $hashes = $null
      $dest = ""
    } else {
      try {
        Write-Log "Getting raw bytes for $Path" -Level 'Debug'
        $stream = $record.GetStream("")
        $recSize = $stream.Length
        $modTime = Convert-UTCtoLocal($record.ModifiedTime)
      } catch {
        Write-Log -Message "Failed to get file raw bytes for $record Error: $_" -Level 'Warn'
      }
    }
    $fullname = $record.FullName
  } else {
    try {
      $Error.Clear()
      Write-Log "Getting item for $Path" -Level 'Debug'
      $item = Get-Item -Path $Path -ErrorAction 'SilentlyContinue'
      if ($null -eq $item) {
        throw $Error
      }
    } catch {
      # Configuration wasn't set for raw some don't bother with raw fallback, but let the user know they
      # may need to use raw for acquisition of this file.
      Write-Log -Message "Failed to get item for $Path. Acquisition may require raw mode. Error: $_" -Level 'Warn'
    }
    if ($item) {
      if (($item.Attributes -band [System.IO.FileAttributes]::Directory) -eq [System.IO.FileAttributes]::Directory) {
        # Item is a directory, so there's no data to be sent.
        $hashes = $null
        $dest = ""
      } else {
        try {
          Write-Log "Opening file $Path" -Level 'Debug'
          $File = Get-Item -Path $Path
          $stream = $File.OpenRead()
          $stream.Position = 0
          $fullname = $item.FullName
          $modTime = $item.LastWriteTime
          $recsize = $item.Length
        } catch {
          if ($RawFallback) {
            # Fallback to raw acquisition
            Write-Log "Failed to get file bytes for $Path - switching to raw acquisition" -Level 'Warn'
            $item = $null
            try {
              Write-Log "Getting forensic file record for $Path" -Level 'Debug'
              $record = Get-ForensicFileRecord -Path $Path
            } catch {
              Write-Log -Message "Failed to get forensic file record for $Path. Error: $_" -Level 'Warn'
            }
            try {
              $recordString = $record.ToString()
            } catch {
              Write-Log -Message "Failed to get forensic file record string for $Path. Error: $_" -Level 'Warn'
              $recordString = $null
            }
            if ([String]::IsNullOrEmpty($recordString) -or $recordString.StartsWith("[0]") -or $recordString.StartsWith("[Directory]")) {
              # File is 0 bytes or we had an issue getting its size - Disable hashing.
              $hashes = $null
              $dest = ""
            } else {
              try {
                Write-Log "Getting raw bytes for $Path" -Level 'Debug'
                $stream = $record.GetStream("")
              } catch {
                Write-Log -Message "Failed to get file raw bytes for $record Error: $_" -Level 'Warn'
              }
            }
            $fullname = $record.FullName
            $modTime  = Convert-UTCtoLocal($record.ModifiedTime)
          } else {
            Write-Log "Failed to get file bytes for $Path - $_" -Level 'Warn'
          }
        }
      }
    }
  }

  if (-not $FileHashes.ContainsKey($Path)) {
    #If we don't have this path in $FileHashes, add it with a default.
    $FileHashes.Add($Path, $defaultFile)
    if ($record) {
      # If we have a record, we need to set the file details because this is a new file.
      $FileHashes.$Path."SIModifiedTime" = $record.ModifiedTime
      $FileHashes.$Path."SIAccessedTime" = $record.AccessedTime
      $FileHashes.$Path."SIChangedTime" = $record.ChangedTime
      $FileHashes.$Path."SIBornTime" = $record.BornTime
      $FileHashes.$Path."FNModifiedTime" = $record.FNModifiedTime
      $FileHashes.$Path."FNAccessedTime" = $record.FNAccessedTime
      $FileHashes.$Path."FNChangedTime" = $record.FNChangedTime
      $FileHashes.$Path."FNBornTime" = $record.FNBornTime
    } elseif ($item) {
      $FileHashes.$Path."SIModifiedTime" = $item.LastWriteTime
      $FileHashes.$Path."SIAccessedTime" = $item.LastAccessTime
      $FileHashes.$Path."SIChangedTime" = 'Not available without raw'
      $FileHashes.$Path."SIBornTime" = $item.CreationTime
      $FileHashes.$Path."FNModifiedTime" = 'Not available without raw'
      $FileHashes.$Path."FNAccessedTime" = 'Not available without raw'
      $FileHashes.$Path."FNChangedTime" = 'Not available without raw'
      $FileHashes.$Path."FNBornTime" = 'Not available without raw'
    } else {
      # If we don't have a record or an item, we can't do anything else, so set everything to failed.
      $FileHashes.$Path."SIModifiedTime" = 'Failed'
      $FileHashes.$Path."SIAccessedTime" = 'Failed'
      $FileHashes.$Path."SIChangedTime" = 'Failed'
      $FileHashes.$Path."SIBornTime" = 'Failed'
      $FileHashes.$Path."FNModifiedTime" = 'Failed'
      $FileHashes.$Path."FNAccessedTime" = 'Failed'
      $FileHashes.$Path."FNChangedTime" = 'Failed'
      $FileHashes.$Path."FNBornTime" = 'Failed'
      $FileHashes.$Path."HashMD5" = "Failed"
      $FileHashes.$Path."HashSHA1" = "Failed"
      $FileHashes.$Path."HashSHA256" = "Failed"
      $FileHashes.$Path."Copied" = "Failed"
    }
  }

  if ($hashes -contains 'md5' -and $FileHashes.$Path.'HashMD5' -eq 'FALSE') {
    $hashesToProcess += 'MD5'
  }
  if ($hashes -contains 'sha1' -and $FileHashes.$Path.'HashSHA1' -eq 'FALSE') {
    $hashesToProcess += 'SHA1'
  }
  if ($hashes -contains 'sha256' -and $FileHashes.$Path.'HashSHA256' -eq 'FALSE') {
    $hashesToProcess += 'SHA256'
  }
  if (($record -or $item) -and ($stream.length -gt 0) -and ($hashesToProcess.Count -gt 0)) {
    foreach ($hash in $hashesToProcess) {
      try {
        Write-Log -Message "Generating $hash hash for $Path" -Level 'Debug'
        $hashObject = [Security.Cryptography.HashAlgorithm]::Create($hash)
        $stream.Position = 0
        $fileHash = ([string]$(foreach ($piece in $hashObject.ComputeHash($stream)) {("{0:x2}" -f $piece)}) -Replace ' ','')
      } catch {
        Write-Log -Message "Failed to generate $hash hash for $Path. Error: $_" -Level 'Warn'
        $fileHash = 'Failed'
      }
      $FileHashes.$Path."Hash$hash" = $fileHash
    }
  }
  if (-not ($stream.length -gt 0)) {
    #No rawbytes, so can't copy the file
    $FileHashes.$Path.'Copied' = 'Failed'
  } elseif ( -not ([String]::IsNullOrEmpty($dest))) {
    $sanitizedFilepath = ($fullname -replace "^\\\\\?\\","" -replace ":","" -replace "\\","_")
    $destination = ("${Dest}/file/" + $sanitizedFilepath)
    try {
      Write-Log -Message "Writing file bytes to $destination" -Level 'Debug'
      $null = $stream.seek(0,0)
      $TFT.SendStream($stream, $destination, $modTime, -1)
      $stream.dispose()
      $copied = $sanitizedFilepath
    } catch {
      Write-Log -Message "Failed to write file bytes to destination file for $fullname. Error: $_" -Level 'Warn'
      $copied = 'Failed'
    }
    $FileHashes.$Path.'Copied' = $copied
  }
  if ($diskinfo) {
    $defaultFile.'SIModifiedTime' = $FileHashes.$Path.'SIModifiedTime'
    $defaultFile.'SIAccessedTime' = $FileHashes.$Path.'SIAccessedTime'
    $defaultFile.'SIChangedTime' = $FileHashes.$Path.'SIChangedTime'
    $defaultFile.'SIBornTime' = $FileHashes.$Path.'SIBornTime'
    $defaultFile.'FNModifiedTime' = $FileHashes.$Path.'FNModifiedTime'
    $defaultFile.'FNAccessedTime' = $FileHashes.$Path.'FNAccessedTime'
    $defaultFile.'FNChangedTime' = $FileHashes.$Path.'FNChangedTime'
    $defaultFile.'FNBornTime' = $FileHashes.$Path.'FNBornTime'
  }
  $defaultFile.'HashMD5' = $FileHashes.$Path.'HashMD5'
  $defaultFile.'HashSHA1' = $FileHashes.$Path.'HashSHA1'
  $defaultFile.'HashSHA256' = $FileHashes.$Path.'HashSHA256'
  $defaultFile.'Copied' = $FileHashes.$Path.'Copied'
  if ($Cache -ne $null) {
    $Cache[$cacheKey] = $defaultFile
  }
  return $defaultFile
}
function Write-FileMetaData {
  param (
    [Parameter(Mandatory=$true)]$file,
    [Parameter(Mandatory=$true)][System.IO.StreamWriter]$writer
  )

  if ($file.FilePath -ne $null) {
    $line = ""
    $line = $line + """$($file.FilePath)""" + '|'
    $line = $line + """$($file.SIModifiedTime)""" + '|'
    $line = $line + """$($file.SIAccessedTime)""" + '|'
    $line = $line + """$($file.SIChangedTime)""" + '|'
    $line = $line + """$($file.SIBornTime)""" + '|'
    $line = $line + """$($file.FNModifiedTime)""" + '|'
    $line = $line + """$($file.FNAccessedTime)""" + '|'
    $line = $line + """$($file.FNChangedTime)""" + '|'
    $line = $line + """$($file.FNBornTime)""" + '|'
    $line = $line + """$($file.HashMD5)""" + '|'
    $line = $line + """$($file.HashSHA1)""" + '|'
    $line = $line + """$($file.HashSHA256)""" + '|'
    $line = $line + """$($file.Copied)"""
    $writer.WriteLine($line)
  }
} # End Write-FileMetaData

# Run all script collectors provided by -scriptCollectors argument
function Invoke-ScriptCollectors {
  param (
    [Parameter(Mandatory=$true)]$ScriptCollectors,
    [Parameter(Mandatory=$true)]$dataDest
  )
  $names = @{}
  foreach ($collector in $ScriptCollectors | Where-Object {$_.Enabled -eq $true}) {
    try {
      $names.Add($collector.Name,$collector.Order)
    } catch {
      Write-Log "Non-unique script collector name: $($collector.Name)" -Level 'Error'
      continue
    }

    $outFile = "${dataDest}/collector/$($collector.name)-results.txt"

    Try {
      $process = ('{0}' -f (Get-Command -Name PowerShell).Definition)
      $script  = ('{0}' -f (Join-Path $LRScriptRoot ($collector.filename -replace '\.\\')))
      $argList = ('-ExecutionPolicy','bypass','-NoProfile','-NonInteractive','-file',"$script")
      if ($collector.safe_args) {
        $argList = $argList + $collector.safe_args
      }
      Send-ProcessOutput -TaniumFileTransfer $TFT -FilePath $Process -ArgumentList $argList -RemoteFilePath $outfile -SHA256 $True -MD5 $True -CloseStandardInput $True
    } catch {
      Write-Log -Level 'Error' -Message "Error running $script - $_"
    }
  }
} # End Invoke-ScriptCollectors

$REGISTERED_MODULE_COLLECTORS = @{}
# Registers a module collector
function Register-ModuleCollector {
  param (
    [Parameter(Mandatory=$true)][string]$Name,
    [Parameter(Mandatory=$true)][System.Management.Automation.CommandInfo]$Cmd
  )

  if ($REGISTERED_MODULE_COLLECTORS.ContainsKey($Name)) {
    throw "Module Collector with key '$Name' already registered."
  }

  $REGISTERED_MODULE_COLLECTORS.Add($Name, @{
    'Name' = $Name
    'Cmd' = $Cmd
  })
}

# Run all module collectors provided by -moduleCollectors argument
function Invoke-ModuleCollectors {
  param (
    [Parameter(Mandatory=$true)]$moduleCollectors,
    [Parameter(Mandatory=$true)]$dataDest
  )

  foreach ($module in $moduleCollectors | Where-Object {$_.Enabled}) {
    $collectorInfo = $REGISTERED_MODULE_COLLECTORS.Item($module.Name)
    if (!$collectorInfo) {
      Write-Log -Level 'Error' -Message "Module Collector with name '$($module.Name)' does not exist"
      continue
    }

    $stream = New-Object -TypeName System.IO.MemoryStream
    $writer = New-Object -TypeName System.IO.StreamWriter($stream, [System.Text.Encoding]::UTF8)
  
    $outFile = "${dataDest}/collector/$($collectorInfo.name)-results.txt"
    $args = GetArgs -Collector $module -Destination "${dataDest}/collector/$($collectorInfo.Name)"
    $args.add("Writer",$writer)

    Write-Log -Message "Running Module: $($collectorInfo.Name)" -Level 'Info'
    try {
      & $collectorInfo.Cmd @args
    } catch {
      Write-Log -Level 'Error' -Message "Error executing module $($collectorInfo.Name) - $_"
    }

    $writer.Flush()
    $null = $stream.seek(0,0)
    $modTime = Get-Date
    try {
      $TFT.SendStream($stream, $outFile, $modTime, -1)
    } catch {
      Write-Log -Level 'Error' -Message "Error transfering report for module $($collectorInfo.Name) - $_"
    } finally {
      $stream.Dispose()
      $writer.Dispose()
    }
  }
} # End Invoke-ModuleCollectors

function Get-APIFiles {
  param (
    [String]$Path,
    [String]$Regex,
    [int]$Depth,
    [int]$MaxNumFiles
  )
  Write-Log -Level 'Info' "Beginning API file collection of '$Path', Regex: '$Regex', Depth: '$Depth', MaxNumFiles: '$MaxNumFiles'"

  $queue = New-Object System.Collections.Queue
  foreach ($p in $(Expand-EnvVariables -Path $Path)) {
    Write-Log -Level 'Info' "Will crawl '$p'"
    $queue.Enqueue($p)
  }
  # check directory depth for the first path in the queue, assuming that all have the same depth,
  # as variables typically don't intoduce additional directory levels 
  $samplePath = $queue.Peek()
  $depthOffSet = $samplePath.Split([System.IO.Path]::DirectorySeparatorChar).Length - 1
  $currentDepth = 0
  $numFiles = 0
  Write-Log -Level 'Info' "Collecting maximum of $MaxNumFiles files"
  while ($queue.Count -gt 0) {
    $currentPath = $queue.Dequeue()
    $currentDepth = $currentPath.Split([System.IO.Path]::DirectorySeparatorChar).Length - 1 - $depthOffSet
    $relativePath = $currentPath -replace ([Regex]::Escape($Path))
    $currentDepth = $relativePath.Split([System.IO.Path]::DirectorySeparatorChar).Length - 1
    if (-not ([System.IO.Directory]::Exists($currentPath))) {
      Write-Log -Level 'Warn' "Failed to find $currentPath"
      continue
    }
    try {
      $files = [System.IO.Directory]::GetFiles($currentPath)
      foreach ($file in $files) {
        if ($file.Split('\')[-1] -match $Regex) {
          if ($numFiles -lt $MaxNumFiles -or $MaxNumFiles -eq -1) {
            Write-Output $file
          }
          $numFiles++
        }
      }
    } catch {
      Write-Log -Level 'Warn' "Error getting files from $currentPath - $_"
    }
    if ($currentDepth -lt $Depth -or $Depth -eq -1) {
      try {
        $directories = [System.IO.Directory]::GetDirectories($currentPath)
        foreach ($directory in $directories) {
          #Don't follow reparse points
          if (([System.IO.FileAttributes]::ReparsePoint) -contains ([System.IO.DirectoryInfo]$directory).Attributes) {
            Write-Log -Level 'Info' "Not enqueueing $directory because it's a reparse point."
          } else {
            $queue.Enqueue($directory)
          }
        }
      } catch {
        Write-Log -Level 'Warn' "Error getting directories from $currentPath - $_"
      }
    }
  }
  Write-Log -Level 'Info' "Found $numFiles files in $Path"
} # End Get-APIFiles

function Get-RawFiles {
  param (
    [String]$Path,
    [String]$Regex,
    [Int]$Depth,
    [Int]$MaxNumFiles
  )
  Write-Log -Level 'Info' "Beginning Raw file collection of '$Path', Regex: '$Regex', Depth: '$Depth', MaxNumFiles: '$MaxNumFiles'"
  
  $queue = New-Object System.Collections.Queue
  foreach ($p in $(Expand-EnvVariables -Path $Path)) {
    Write-Log -Level 'Info' "Will crawl '$p'"
    $queue.Enqueue($p)
  }
  $currentDepth = 0
  $numFiles = 0
  Write-Log -Level 'Info' "Collecting maximum of $MaxNumFiles files"
  while ($queue.Count -gt 0) {
    $currentPath = $queue.Dequeue()
    $relativePath = $currentPath -replace ([Regex]::Escape($Path))
    $currentDepth = $relativePath.Split([System.IO.Path]::DirectorySeparatorChar).Length - 1
    try {
      $items = Get-ForensicChildItem -Path $currentPath
    } catch {
      Write-Log "Error getting child items from $dir - $_" -Level 'Warn'
      continue
    }

    foreach ($item in $items) {
      # Until Get-ForensicChildItem correctly determines directories vs files, do this:
      try {
        $record = Get-ForensicFileRecord -VolumeName $SysVolume -Index $item.RecordNumber
      } catch {
        Write-Log "Error getting record from $item - $_" -Level 'Warn'
        continue
      }
      if (-not $record.Directory) {
        if ($item.FileName -match $Regex) {
          if ($numFiles -lt $MaxNumFiles -or $MaxNumFiles -eq -1) {
            Write-Output $item.FullName
          }
          $numFiles++
        }
      } elseif ($currentDepth -lt $Depth -or $Depth -eq -1) {
        $queue.Enqueue($item.FullName)
      }
    }
  }
  Write-Log -Level 'Info' "Found $numFiles files in $Path"
} # End Get-RawFiles

function Get-Files {
  param (
    [String]$Path,
    [String]$Regex,
    [Int]$Depth,
    [Boolean]$Raw,
    [Int]$MaxNumFiles
  )
  if ($Raw) {
    Get-RawFiles -Path $Path -Regex $Regex -Depth $Depth -MaxNumFiles $MaxNumFiles
  } else {
    Get-APIFiles -Path $Path -Regex $Regex -Depth $Depth -MaxNumFiles $MaxNumFiles
  }
}

function Invoke-FileCollectors {
  param (
    $FileCollectors,
    $dataDest,
    $CollectionDate
  )
  $stream = New-Object -TypeName System.IO.MemoryStream
  $writer = New-Object -TypeName System.IO.StreamWriter($stream, [System.Text.Encoding]::UTF8)

  $fileDetailsCache = @{}

  $writer.WriteLine("""FullPath""|""SIModifiedTime""|""SIAccessedTime""|""SIChangedTime""|""SIBornTime""|""FNModifiedTime""|""FNAccessedTime""|""FNChangedTime""|""FNBornTime""|""MD5""|""SHA1""|""SHA256""|""Copied""")
  # Add-Content ("$dataDest\files\_" + $CollectionDate + "_collection_summary.txt") """FullPath""|""SIModifiedTime""|""SIAccessedTime""|""SIChangedTime""|""SIBornTime""|""FNModifiedTime""|""FNAccessedTime""|""FNChangedTime""|""FNBornTime""|""MD5""|""SHA1""|""SHA256""|""Copied"""
  # $fileDetailOutput += $line
  foreach ($collector in $FileCollectors | Where-Object {$_.Enabled -eq $true}) {
    $args = GetArgs -Collector $collector -Destination $dataDest
    Get-Files @args | ForEach-Object {
      $path = $_
      $args = GetArgs -Collector $collector -Destination $dataDest
      Add-HashTableEntry -HashTable $args -Name 'Path' -Value $path
      if ($collector.Raw) {
        $args.'Raw' = $true
        #Determine whether path is file or folder. Should be improved to not double up on Get-ForensicFileRecord calls.
        if ((Get-ForensicFileRecord $path).Directory) {
          $args.Hashes = $false
        } elseif ($collector.Copy) {
          $args.Add('Dest', $dataDest)
        }
      } else {
        if ([System.IO.Directory]::Exists($path)) {
          $args.Hashes = $False
        } elseif ($collector.Copy) {
          $args.Add('Dest', $dataDest)
        }
      }
      $args["Cache"] = $fileDetailsCache
      $fileDetails = Get-FileDetails @args
      Write-FileMetaData -File $fileDetails -Writer $writer
    }
  }
  $outfile = ("${dataDest}/" + $CollectionDate + "_file_collection_summary.txt")
  $writer.Flush()
  $null = $stream.seek(0,0)
  $modTime = Get-Date
  $TFT.SendStream($stream, $outfile, $modTime, -1)
  $stream.dispose()
}

function Expand-EnvVariables {
  param (
    [Parameter(Mandatory=$true)][string]$path
  )

  switch -wildcard ($path) {
    '*%appdata%*' { $subPath,$envVariable = '\appdata\roaming','%appdata%' }
    '*%homepath%*' { $subPath,$envVariable = '','%homepath%' }
    '*%localappdata%*' { $subPath,$envVariable = '\appdata\local','%localappdata%' }
    '*%psmodulepath%*' { $subPath,$envVariable = '\documents\windowspowershell\modules','%psmodulepath%' }
    '*%temp%*' { $subPath,$envVariable = '\appdata\local\temp','%temp%' }
    '*%tmp%*' { $subPath,$envVariable = '\appdata\local\temp','%tmp%' }
    '*%userprofile%*' { $subPath,$envVariable = '','%userprofile%' }
    default { $subPath,$envVariable = '','' }
  }

  if ($envVariable -eq '') { Expand-SystemEnvVariables $path }
  else {
    $paths = Get-ProfileImagePaths -subPath $subPath -filterBuiltIn $True
    foreach ($p in $paths) {
      if ($envVariable -eq '%homepath%') { Expand-SystemEnvVariables $($path -Replace $envVariable, $p -Replace $env:systemdrive, '') }
      else { Expand-SystemEnvVariables $($path -Replace $envVariable, $p) }
    }
  }
}

function Get-ProfileImagePaths {
  param (
    [Parameter(Mandatory=$false)][string]$subPath,
    [Parameter(Mandatory=$false)][bool]$filterBuiltIn
  )
  $profiles = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList"
  foreach ($p in $($profiles | Select-Object pspath)) {
    $result = $p | Get-ItemProperty -Name ProfileImagePath | Select-Object -ExpandProperty ProfileImagePath
    if ($filterBuiltIn) {
      if (-Not ($result -Match "(Windows\\system32\\config\\systemprofile)|(Windows\\ServiceProfiles\\LocalService)|(Windows\\ServiceProfiles\\NetworkService)")) {
        $result = $result + $subPath
        $result
      }
    } else {
      $result = $result + $subPath
      $result
    }
  }
}

function Expand-SystemEnvVariables {
  param (
    [Parameter(Mandatory=$true)][string]$path
  )
  $path -match '(%.*%)' | Out-Null
  if ($Matches -ne $null){
    $envPath = [Environment]::GetEnvironmentVariable($($Matches[1] -Replace "%",""))
    $path -replace "(%.*%)", "$envPath"
  } else {
    $path
  }
}

function Get-FullPathColumn {
  param (
    [Parameter(Mandatory=$true,Position=1)][string]$header
  )
  $columns = $header -split "\|"
  $position = 0
  foreach ($c in $columns) {
    if ($c -eq 'FullPath') {
      return $position
    }
    $position++
  }
}

function Test-PSVersion {
  if ($null -eq $PSVersionTable) {
    return $false
  } else {
    return $true
  }
}

function Get-OSVersion {
  [version](Get-WmiObject Win32_OperatingSystem).Version
}
function Test-IsGeWin8 {
  if ((Get-OSVersion) -gt [version]"6.2") {
    return $true
  } else {
    return $false
  }
}

function Test-IsLtWin10 {
  if ((Get-OSVersion) -lt [version]"10.0") {
    return $true
  } else {
    return $false
  }
}

function Test-IsDeviceGuardHVCIRunning {
  if ( (Get-OSVersion).Major -lt 10 )
  {
    # If we're not on Major version 10 (Win10/2016), then HVCI will not be running.
    return $False
  }

  try {
    $DeviceGuardInfo = Get-WmiObject -ClassName "Win32_DeviceGuard" -Namespace "root\Microsoft\Windows\DeviceGuard"
  } catch {
    Write-Log -Message "Unable to get WMI object Win32_DeviceGuard. Assuming HVCI is running. Error: $_" -Level 'Warn'
    return $True
  }

  # 2 = HVCI Running
  return $DeviceGuardInfo.SecurityServicesRunning -Contains 2
}

function Test-IsWinpmemSafe {
  if ( (Get-OSVersion).Major -lt 10 )
  {
    # If we're not on Major version 10 (Win10/2016), we're safe
    return $True
  }

  if (!(Test-IsDeviceGuardHVCIRunning)) {
    # Device Guard HVCI not running, so we should be OK.
    Write-Log -Message 'DeviceGuard HVCI is not running.' -Level 'Info'
    return $True
  }
  Write-Log -Message 'DeviceGuard HVCI is running.' -Level 'Info'

  try {
    $VersionDetail = Get-ItemProperty 'HKLM:Software\Microsoft\Windows NT\CurrentVersion' | Select-Object ReleaseId, CurrentBuild, UBR

    if ($VersionDetail.ReleaseId -eq $null) {
      Write-Log -Message "Failed to retrieve 'ReleaseId' from registry. Assuming unsafe for Winpmem." -Level 'Warn'
      return $False
    }

    if ($VersionDetail.CurrentBuild -eq $null) {
      Write-Log -Message "Failed to retrieve 'CurrentBuild' from registry. Assuming unsafe for Winpmem." -Level 'Warn'
      return $False
    }

    if ($VersionDetail.UBR -eq $null) {
      Write-Log -Message "Failed to retrieve 'UBR' from registry. Assuming unsafe for Winpmem." -Level 'Warn'
      return $False
    }

    $ReleaseId = [int]$VersionDetail.ReleaseId
    $CurrentBuild = [int]$VersionDetail.CurrentBuild
    $UBR = [int]$VersionDetail.UBR
  } catch {
    # If we can't get specific version information, assume it's not safe
    Write-Log -Message "Unable to retrieve Windows 10 build revision from the registry. Assuming unsafe for Winpmem. Error: $_" -Level 'Warn'
    return $False
  }

  Write-Log -Message "Windows 10 Version: ReleaseId '$ReleaseId', CurrentBuild: '$CurrentBuild', UBR: '$UBR'" -Level 'Info'

  if ($ReleaseId -ge 1809 ) {
    # Assume newer versions are unaffected by BSOD fixed by Sep 2018 rollups.
    return $True
  }

  # We're on Win10, which version and what UBR?
  switch ( $ReleaseId ) {
    1803 {
      if ( $UBR -ge 320 ) {
        return $True
      }
      break
    }
    1709 {
      # Two different 'CurrentBuild' values for this version
      if ( $CurrentBuild -eq 16299 -and $UBR -ge 699 ) {
        return $True
      }
      break
    }
    1703 {
      if ( $UBR -ge 1356 ) {
        return $True
      }
      break
    }
    1607 {
      if ( $UBR -ge 2515 ) {
        return $True
      }
      break
    }
  }
  
  # Default to false
  return $False
}

function Get-Memory {
  param (
    [Boolean]$DiskInfo,
    [Array]$Hashes,
    [String]$Dest=$null,
    [Boolean]$Raw,
    [Boolean]$RawFallback,
    [System.IO.StreamWriter]$writer
  )

  if (Test-IsWinpmemSafe)
  {
    $winpmembin = "$($LRScriptRoot)\winpmem.gb414603.exe"

    if ($(Test-Path $winpmembin -PathType Leaf) -eq $true) {
      try {
        $argument = ('--format','raw','--volume_format','raw','--output','-')
        $outfile = "${Dest}/memory.raw"

        $MD5    = $Hashes -Contains 'MD5'
        $SHA1   = $Hashes -Contains 'SHA1'
        $SHA256 = $Hashes -Contains 'SHA256'

        Send-ProcessOutput -TaniumFileTransfer $TFT -FilePath $winpmembin -ArgumentList $argument -RemoteFilePath $outfile -MD5 $MD5 -SHA1 $SHA1 -SHA256 $SHA256 -CloseStandardInput $False
        $writer.WriteLine('System memory copied to {0}.' -f $outfile)
      } catch {
        $Message = ('Unable to image memory with winpmen. Error: {0}' -f $_)
        Write-Log -Message $Message -Level 'Warn'
      }
    } else {
      $Message = ('Winpmem binary not found.')
      Write-Log -Message $Message -Level 'Warn'
    }
  } 
  else 
  {
    $Message = ('System is running a version of Device Guard that is incompatible with Winpmem.')
    Write-Log -Message $Message -Level 'Warn'
  }
}

function Get-DriverDetails {
  param (
    [Boolean]$DiskInfo,
    [Array]$Hashes,
    [String]$Dest=$null,
    [Boolean]$Raw,
    [Boolean]$RawFallback,
    [System.IO.StreamWriter]$writer
  )
  $driverFileDetailsCache = @{}
  try {
    $items = Get-WmiObject -Class win32_SystemDriver  # Get-wmiobject for all drivers loaded on the system
  }
  catch {
    $Message = "Unable to get WMI object win32_SystemDriver. Error: $_"
    Write-Log -Message $Message -Level 'Warn'
    $writer.WriteLine($Message)
    $items = $null
  }

  if ($items -ne $null) {
    $writer.WriteLine("""FullPath""|""Status""|""Name""|""State""|""Started""|""ServiceType""|""StartMode""|""SIModifiedTime""|""SIAccessedTime""|""SIChangedTime""|""SIBornTime""|""FNModifiedTime""|""FNAccessedTime""|""FNChangedTime""|""FNBornTime""|""MD5""|""SHA1""|""SHA256""|""Copied""")
    foreach ($item in $items) {
      if ($item -ne $null) {
        if (-not [String]::IsNullOrEmpty($item.PathName)) {
          $item.PathName = ($item.PathName).replace("\??\", "")
          $args = @{'Path' = $item.PathName} + $PSBoundParameters
          $args["Cache"] = $driverFileDetailsCache
          $fileDetails = Get-FileDetails @args
        } else {
          $item.PathName = "Path not found"
        }
        $returnStr = '"{0}"|"{1}"|"{2}"|"{3}"|"{4}"|"{5}"|"{6}"|' -f $item.PathName,
        $item.Status, $item.Name, $item.State, $item.Started, $item.ServiceType, $item.StartMode
        $returnStr += '"{0}"|"{1}"|"{2}"|"{3}"|"{4}"|"{5}"|"{6}"|"{7}"|"{8}"|"{9}"|"{10}"|"{11}"' -f $fileDetails.'SIModifiedTime',
        $fileDetails.'SIAccessedTime', $fileDetails.'SIChangedTime', $fileDetails.'SIBornTime', $fileDetails.'FNModifiedTime',
        $fileDetails.'FNAccessedTime', $fileDetails.'FNChangedTime', $fileDetails.'FNBornTime', $fileDetails.'HashMD5',
        $fileDetails.'HashSHA1', $fileDetails.'HashSHA256', $fileDetails.'Copied'
        $writer.WriteLine($returnStr)
      }
    }
  }
}

function Get-ModuleDetails {
  param (
    [Boolean]$DiskInfo,
    [Array]$Hashes,
    [String]$Dest=$null,
    [Boolean]$Raw,
    [Boolean]$RawFallback,
    [System.IO.StreamWriter]$writer
  )
  try {
    $items = Get-WmiObject -Class CIM_ProcessExecutable  # Get-wmiobject for all modules loaded on the system
  }
  catch {
    $Message = "Unable to get WMI object CIM_ProcessExecutable. Error: $_"
    Write-Log -Message $Message -Level 'Warn'
    $writer.WriteLine($Message)
    $items = $null
  }

  $moduleFileDetailsCache = @{}

  $processID = ""
  $processPath = ""

  if ($items -ne $null) {
    $writer.WriteLine("""ParentProcess""|""ParentProcessId""|""FullPath""|""BaseAddress""|""SIModifiedTime""|""SIAccessedTime""|""SIChangedTime""|""SIBornTime""|""FNModifiedTime""|""FNAccessedTime""|""FNChangedTime""|""FNBornTime""|""MD5""|""SHA1""|""SHA256""|""Copied""")
    foreach ($item in $items) {

      if ($item -ne $null) {
        [regex]$regHandle = "(.Handle="")([0-9]){1,9}"
        [regex]$regModule = "(.Name="")([^""]){0,300}"
        $item.Dependent = ($regHandle.Matches($item.Dependent) | ForEach-Object {$_.Value}).Replace(".Handle=""", "")
        $item.Antecedent = ($regModule.Matches($item.Antecedent) | ForEach-Object {$_.Value}).Replace(".Name=""", "").Replace("\\","\")

        if ($item.Dependent -ne $processID) {
          $processID = $item.Dependent
          $processPath = $item.Antecedent
        }

        $returnStr = '"{0}"|"{1}"|"{2}"|"{3}"|' -f $processPath, $item.Dependent, $item.Antecedent, $item.BaseAddress
        $args = @{'Path' = $item.Antecedent} + $PSBoundParameters
        $args["Cache"] = $moduleFileDetailsCache
        $fileDetails = Get-FileDetails @args
        $returnStr += '"{0}"|"{1}"|"{2}"|"{3}"|"{4}"|"{5}"|"{6}"|"{7}"|"{8}"|"{9}"|"{10}"|"{11}"' -f $fileDetails.'SIModifiedTime',
        $fileDetails.'SIAccessedTime', $fileDetails.'SIChangedTime', $fileDetails.'SIBornTime', $fileDetails.'FNModifiedTime',
        $fileDetails.'FNAccessedTime', $fileDetails.'FNChangedTime', $fileDetails.'FNBornTime', $fileDetails.'HashMD5',
        $fileDetails.'HashSHA1', $fileDetails.'HashSHA256', $fileDetails.'Copied'
        $writer.WriteLine($returnStr)
      }
    }
  }
}

function Get-ProcessDetails {
  param (
    [Boolean]$DiskInfo,
    [Array]$Hashes,
    [String]$Dest=$null,
    [Boolean]$Raw,
    [Boolean]$RawFallback,
    [System.IO.StreamWriter]$writer
  )
  try {
    $items = Get-WmiObject -Class Win32_Process  # Get-wmiobject for all processes on the system
  }
  catch {
    $Message = "Unable to get WMI object Win32_Process. Error: $_"
    $writer.WriteLine($Message)
    Write-Log -Message $Message -Level 'Warn'
    $items = $null
  }
  $processFileDetailsCache = @{}

  $protectedProcesses = 'audiodg.exe','smss.exe', 'csrss.exe', 'wininit.exe', 'services.exe', 'svchost.exe', 'MsMpEng.exe', 'NisSrv.exe'

  if ($items -ne $null) {
    $writer.WriteLine("""Caption""|""FullPath""|""ParentProcessId""|""ProcessId""|""Domain""|""User""|""CreationDate""|""CommandLine""|""SIModifiedTime""|""SIAccessedTime""|""SIChangedTime""|""SIBornTime""|""FNModifiedTime""|""FNAccessedTime""|""FNChangedTime""|""FNBornTime""|""MD5""|""SHA1""|""SHA256""|""Copied""")
    foreach ($item in $items) {
      $caption = $item.Caption
      $fullPath = ""
      $parentProcessID = $item.ParentProcessId
      $processID = $item.ProcessId
      $domain = ""
      $user = ""
      $creation = $item.CreationDate
      $command = $item.CommandLine
      if (($item.ProcessId -ne $null) -and ($item.ProcessId -ne "")) {
        if (($item.ProcessId -eq '4') -and ($item.Caption.ToLower() -eq "system")) { # If System process use ntoskrnl.exe path
          $fullPath = "$env:SystemRoot\System32\ntoskrnl.exe"
          $domain = "N/A"
          $user = "N/A"
        } elseif (($item.ParentProcessId -eq '4') -and ($item.Caption.ToLower() -eq "memory compression")) { # If Memory Compressor
          $fullPath = "N/A"
          $domain = "N/A"
          $user = "N/A"
        } elseif ($item.ExecutablePath -eq $null) { # Correct protected process paths
          try {
            $domain = $item.getOwner().domain
            $user = $item.getOwner().user
          } catch {
            $domain = "Process No Longer Exists"
            $user = "Process No Longer Exists"
          }
          if ($protectedProcesses -contains $item.Caption) {
            if (($item.Caption -eq 'NisSrv.exe') -or ($item.Caption -eq 'MsMpEng.exe')) {
              $fullPath = "$env:ProgramFiles\Windows Defender\$($item.Caption)"
            } else {
              $fullPath = "$env:SystemRoot\System32\$($item.Caption)"
            }
          } else {
            $fullPath = "N/A"
          }
        } else {
          try {
            $domain = $item.getOwner().domain
            $user = $item.getOwner().user
          } catch {
            $domain = "Process No Longer Exists"
            $user = "Process No Longer Exists"
          }
          $fullPath = $item.ExecutablePath
        }
        $returnStr = '"{0}"|"{1}"|"{2}"|"{3}"|"{4}"|"{5}"|"{6}"|"{7}"|' -f $caption,
        $fullPath, $parentProcessId, $processId, $domain, $user, $creation, $command
        $args = @{'Path' = $fullPath} + $PSBoundParameters
        $args["Cache"] = $processFileDetailsCache
        $fileDetails = Get-FileDetails @args
        $returnStr += '"{0}"|"{1}"|"{2}"|"{3}"|"{4}"|"{5}"|"{6}"|"{7}"|"{8}"|"{9}"|"{10}"|"{11}"' -f $fileDetails.'SIModifiedTime',
        $fileDetails.'SIAccessedTime', $fileDetails.'SIChangedTime', $fileDetails.'SIBornTime', $fileDetails.'FNModifiedTime',
        $fileDetails.'FNAccessedTime', $fileDetails.'FNChangedTime', $fileDetails.'FNBornTime', $fileDetails.'HashMD5',
        $fileDetails.'HashSHA1', $fileDetails.'HashSHA256', $fileDetails.'Copied'
        $writer.WriteLine($returnStr)
      }
    }
  }
}

function Convert-UTCtoLocal {
  param(
  [parameter(Mandatory=$true)]
  [DateTime] $UTCTime
  )

  ($UTCTime).ToLocalTime()
}

function Get-AddrPort {
  Param(
    [Parameter(Mandatory=$True,Position=0)]
    [String]$AddrPort
  )
  # Write-Verbose "Entering $($MyInvocation.MyCommand)"
  # Write-Verbose "Processing $AddrPort"
  if ($AddrPort -match '[0-9a-f]*:[0-9a-f]*:[0-9a-f%]*\]:[0-9]+') {
    $Addr, $Port = $AddrPort -split "]:"
    $Addr += "]"
  } else {
    $Addr, $Port = $AddrPort -split ":"
  }
  $Addr, $Port
  # Write-Verbose "Exiting $($MyInvocation.MyCommand)"
}

function Get-NetworkConnectionDetails {
  param (
    [System.IO.StreamWriter]$writer
  )
  $netstatScriptBlock = { & $env:windir\system32\netstat.exe -naob }
  $results = @()
  foreach($line in $(& $netstatScriptBlock)) {
    if ($line.length -gt 1 -and $line -notmatch "Active |Proto ") {
      $line = $line.trim()
      if ($line.StartsWith("TCP")) {
        $Protocol, $LocalAddress, $ForeignAddress, $State, $ConPId = ($line -split '\s{2,}')
        $Component = $Process = $False
      } elseif ($line.StartsWith("UDP")) {
        $State = "STATELESS"
        $Protocol, $LocalAddress, $ForeignAddress, $ConPid = ($line -split '\s{2,}')
        $Component = $Process = $False
      } elseif ($line -match "^\[[-_a-zA-Z0-9.]+\.(exe|com|ps1)\]$") {
        $Process = $line
        if ($Component -eq $False) {
          # No Component given
          $Component = $Process
        }
      } elseif ($line -match "Can not obtain ownership information") {
        $Process = $Component = $line
      } else {
        # We have the $Component
        $Component = $line
      }
      if ($State -match "TIME_WAIT") {
        $Component = "Not provided"
        $Process = "Not provided"
      }
      if ($Component -and $Process) {
        $LocalAddress, $LocalPort = Get-AddrPort($LocalAddress)
        $ForeignAddress, $ForeignPort = Get-AddrPort($ForeignAddress)

        $o = "" | Select-Object Protocol, LocalAddress, LocalPort, ForeignAddress, ForeignPort, State, ConPId, Component, Process
        $o.Protocol, $o.LocalAddress, $o.LocalPort, $o.ForeignAddress, $o.ForeignPort, $o.State, $o.ConPId, $o.Component, $o.Process = `
          $Protocol, $LocalAddress, $LocalPort, $ForeignAddress, $ForeignPort, $State, $ConPid, $Component, $Process
        $results +=  $o
      }
    }
  }
  $writer.WriteLine($(($results | ConvertTo-CSV -Delimiter '|' -NoTypeInformation)|Out-String))
}

function Get-HandleDetails {
  param(
    [System.IO.StreamWriter]$writer
  )
  $handles = & ($env:TANIUMDIR + "\Tools\IR\TaniumHandle.exe")
  $results = @()
  foreach ($line in $handles) {
    if ($line -match '<unable to open process>|^[-]+$') {
      continue
    }

    if (-not $line.ToString().StartsWith(' ')) {
      $line = $line -split '\s'
      $process = $line[0]
      $procid = $line[2]
    } else {
      $line = $line.trim()
      $line = $line -replace '\s+', ' '
      $handleid = ($line -split ':')[0]
      $handletype = ($line -split ' ')[1]
      if ($line -match 'File \( \)') {
        $handle = (($line -split '\)')[1]).Trim()
      } else {
        try {
          $temp = ($line -split ' ')
          $handle = (($line -split ' ')[2..$temp.length] -join " "|out-string).trim()
        } catch {
          $handle = ''
        }
      }
    }
    $properties = @{
      Process = $process
      PID = $procid
      Type = $handletype
      HandleId = $handleid
      Handle = $handle
    }
    # Output object
    $results += New-Object -TypeName PSObject -Property $properties
  }
  $writer.WriteLine($(($results | Select-Object Process,PID,Type,HandleId,Handle | ConvertTo-CSV -Delimiter '|' -NoTypeInformation)|Out-String))
}

function Get-PFAmCache {
  param(
    [System.IO.StreamWriter]$writer
  )
  if (-not (Test-IsGeWin8)) {
    $Message = "Error: Amcache only available on Windows 8 or greater."
    Write-Log -Level 'Warn' -Message $Message
    $writer.WriteLine($Message)
    return
  }
  try {
    $amcache = Get-ForensicAmcache -VolumeName $SysVolume -ErrorAction SilentlyContinue
    if ($amcache.length -lt 1) {
      $Message = "No amcache files found."
      Write-Log -Level 'Info' -Message $Message
      $writer.WriteLine($Message)
    } else {
      $writer.WriteLine($(($amcache | ConvertTo-CSV -Delimiter '|' -NoTypeInformation)|Out-String))
    }
  } catch {
    $Message = ('Exception in Collect-PFAMCache. {0}' -f $_)
    $writer.WriteLine($Message)
    Write-Log -Message $Message -Level 'Warn'
  }
}

function Get-PFShellLink {
  param(
    [System.IO.StreamWriter]$writer
  )  
  try {
    $shelllinks = Get-ForensicShellLink -VolumeName $SysVolume
    if ($shelllinks.length -lt 1) {
        $Message = "No .lnk files found."
        Write-Log -Level 'Warn' -Message $Message
        $writer.WriteLine($Message)
    } else {
      $writer.WriteLine($(($shelllinks | ConvertTo-CSV -Delimiter '|' -NoTypeInformation)|Out-String))
    }
  } catch {
    $Message = ('Exception in Collect-PFShellLink. {0}' -f $_)
    Write-Log -Level 'Warn' -Message $Message
    Write-Log -Message $Message -Level 'Warn'
  }
}

function Get-PFScheduledJob {
  param(
    [System.IO.StreamWriter]$writer
  )
  try {
    $scheduledjobs = Get-ForensicScheduledJob -VolumeName $SysVolume
    if ($scheduledjobs.length -lt 1) {
        $Message = "No .job files found."
        Write-Log -Level 'Info' -Message $Message
        $writer.WriteLine($Message)
    } else {
      $writer.WriteLine($(($scheduledjobs | ConvertTo-CSV -Delimiter '|' -NoTypeInformation)|Out-String))
    }
  } catch {
    $Message = ('Exception in Collect-PFScheduledJob. {0}' -f $_)
    $writer.WriteLine($Message)
    Write-Log -Message $Message -Level 'Warn'
  }
}

function Get-PFRecentFileCache {
  param(
    [System.IO.StreamWriter]$writer
  )
  if (-not (Test-IsGeWin8)) {
    $Message = "Error: RecentFileCache only available on Windows 8 or greater."
    Write-Log -Level 'Warn' -Message $Message
    $writer.WriteLine($Message)
    return
  }
  try {
    $recentfilecache = Get-ForensicRecentFileCache
    if ($recentfilecache.length -lt 1) {
        $writer.WriteLine("No recent files found in cache.")
    } else {
      $writer.WriteLine($(($recentfilecache | ConvertTo-CSV -Delimiter '|' -NoTypeInformation)|Out-String))
    }
  } catch {
    $Message = ('Exception in Collect-PFRecentFileCache. {0}' -f $_)
    $writer.WriteLine($Message)
    Write-Log $Message -Level 'Warn'
  }
}

function Get-PFUserAssist {
  param(
    [System.IO.StreamWriter]$writer
  )

  $writer.WriteLine("""User""|""ImagePath""|""RunCount""|""FocusTime""|""LastExecutionTimeUtc""")
  foreach ($profile in Get-ChildItem 'HKLM:Software\Microsoft\Windows NT\CurrentVersion\ProfileList') {
    try {
      $writer.WriteLine($((Get-ForensicUserAssist -HivePath ($profile.GetValue('ProfileImagePath') + '\ntuser.dat') | ConvertTo-CSV -Delimiter '|' -NoTypeInformation | Select-Object -Skip 1) | Out-String))
    } catch {
      $writer.WriteLine("""$($profile.pschildname)""|""$($_.Exception.Message)""|||")
    }
  }
}

function Get-PFShimCache {
  param(
    [System.IO.StreamWriter]$writer
  )
  try {
    $shimcache = Get-ForensicShimcache -VolumeName $SysVolume
    if ($shimcache.length -lt 1) {
        $writer.WriteLine("No Shimcache files found.")
    } else {
      $writer.WriteLine($(($shimcache | ConvertTo-CSV -Delimiter '|' -NoTypeInformation)|Out-String))
    }
  } catch {
    $Message = ('Exception in Collect-PFSchimCache. {0}' -f $_)
    $writer.WriteLine($Message)
    Write-Log -Message $Message -Level 'Warn'
  }
}

function Get-PFPrefetch {
  param (
    [System.IO.StreamWriter]$writer
  )  
  try {
    $prefetch = Get-ForensicPrefetch -VolumeName $SysVolume
    if ($null -eq $prefetch) {
        throw
    } else {
      $writer.WriteLine($(($prefetch | Select-Object version, name, path, pathhash, dependencycount, prefetchaccesstime, devicecount, runcount, @{Name='dependencyfiles';Expression={$_.DependencyFiles -join ';'}} | ConvertTo-CSV -Delimiter '|' -NoTypeInformation)|out-string))
    }
  } catch {
    Write-Log -Message ('Exception in Collect-PFPrefetch. {0}' -f $_) -Level 'Warn'
    $writer.WriteLine("No prefetch files found. Prefetch may be disabled on this host.")
  }
}

function Get-PFMasterBootRecord {
  param(
    [System.IO.StreamWriter]$writer
  )
  try {

    # PF uses the Win32 device namespace for MBR object
    $DRIVE = "\\.\PHYSICALDRIVE"
    # The first 440 bytes of the MBR contain the boot code to be hashed
    # First 3 bytes of that are "jump to boot program"
    # Next 43 bytes are disk parameters, we just want the "boot program code"
    # https://raw.githubusercontent.com/Invoke-IR/ForensicPosters/master/Posters/_MBR.png
    $MBRCODELENGTH = 394

    $outObject = "" | Select-Object OS,Hash,Hex

    #Enumerate bootable drive
    Try {
      $Error.Clear()
      $DriveNum = Get-WmiObject win32_diskpartition | Where-Object {$_.BootPartition} | Foreach-Object {(($_.Caption).Split(",")[0]).Split("#")[1]}
      If ( $Error ) {$DriveNum = "0"}
      $Error.Clear()
      If ( -Not $DriveNum ) { $DriveNum = "0" }
    } Catch{
      $DriveNum = "0" #if fail to enumerate drive number then assume drive 0
    }

    $DRIVE = $DRIVE + $DriveNum

    Try  {
      $MBR = (Get-ForensicMasterBootRecord -path $Drive)
      if (($MBR.PartitionTable | Select-Object -ExpandProperty SystemId) -match  "GPT") {
        $Message = "Disk is GPT"
        Write-Log -message $Message -Level 'Warn'
        $writer.WriteLine($Message)
        return
      }
      $MBRCodeSection = $MBR.CodeSection | Select-Object -Skip 46
    } Catch {
      $Message = "Failed to enumerate boot record for drive: $DRIVE. $_"
      Write-Log -message $Message -Level 'Warn'
      $writer.WriteLine($Message)
      return
    }
    # This check confirms the correct length of data was found
    If ( $MBRCodeSection.Length -ne $MBRCODELENGTH) {
      $Message = "Incorrect MBR code length found"
      Write-Log -message $Message -Level 'Warn'
      $writer.WriteLine($Message)
      return
    }

    #Instantiate MD5Object for hashing the MBRCodeSection
    Try {
        $MD5Object = [System.Security.Cryptography.HashAlgorithm]::Create("MD5")
    } Catch {
      $Message = "Failed to generate MD5 hash generator object"
      Write-Log -message $Message -Level 'Warn'
      $writer.WriteLine($Message)
      return
    }

    Try {
        $Error.Clear()
        $OperatingSystem = (Get-WmiObject -class Win32_OperatingSystem).Caption
        If ( $Error ) { $OperatingSystem = "Unknown" }
        $Error.Clear()
    } Catch {
        $OperatingSystem = "Unknown"
    }

    #Assign values to output object
    $outObject.OS = $OperatingSystem.Trim()
    $outObject.hash = (-Join ($MD5Object.ComputeHash($MBRCodeSection) | Foreach-Object {"{0:x2}" -f $_})).ToUpper()
    $outObject.hex =  (-Join ($MBRCodeSection | Foreach-Object {"{0:x2}" -f $_})).ToUpper()

    $writer.WriteLine($(($outObject | Select-Object os, hash, hex | ConvertTo-CSV -Delimiter '|' -NoTypeInformation)|Out-String))
  } catch {
    $Message = ('Exception in Collect-PFMasterBootRecord. {0}' -f $_)
    Write-Log -Message $Message -Level 'Warn'
    $writer.WriteLine($Message)
  }
}

function Get-Autoruns {
  param (
    [System.IO.StreamWriter]$writer
  )

  # Fallback to Autoruns in the Tools IR directory, if it exists
  $IRToolsAutoruns = (GetTaniumDir) + ('\Tools\IR\Autoruns\Autorunsc.exe')

  if (Test-Path -Path ($LRScriptRoot + '\Autoruns.zip')) {
    $arunsArchive = ls ($LRScriptRoot + '\Autoruns.zip')
    try {
      Extract-ZipArchive -archivePath $arunsArchive
      $AutorunscPath = $LRScriptRoot + "\Autorunsc.exe"
      if (-Not (Test-Path $AutorunscPath))
      {
        throw ('Autorunsc.exe not found at {0}. Check the Autoruns zip file.' -f $AutorunscPath)    
      }
    } catch {
      Write-Log -Level 'Warn' -Message $_
      $writer.WriteLine($_)
      return
    }
  }
  elseif ( Test-Path -Path $IRToolsAutoruns )
  {
    $AutorunscPath = $IRToolsAutoruns
    Write-Log -Level 'Info' -Message 'Using Autoruns from IR Tools.'
  }
  else 
  {
    $Message = ('Autoruns.zip not found. Skipping Autoruns collection.')
    Write-Log -Level 'Warn' -Message 'Autoruns.zip not found. Skipping Autoruns collection.'
    $writer.WriteLine($Message)
    return
  }

  $Arguments = @("/accepteula","-a","*","-c","-h","-s","-t","'*'","-nobanner")
  $arrAseps = @()

  # Build an array of ASEPs from the Autorunsc created CSV
  try {
    $arrAseps = (& $AutorunscPath @Arguments | ConvertFrom-Csv)
    $writer.WriteLine($(($arrAseps | ConvertTo-Csv -NoTypeInformation -Delimiter '|')|Out-String))
  } catch {
    $Message = ('Exception in Collect-Autoruns. {0}' -f $_) 
    Write-Log -Message $Message -Level 'Warn'
    $writer.WriteLine($Message)
  }
}

function ConvertTo-DelimiterSeparatedValues {
  <#
    This function is like ConverTo-CSV but with
    support for multi-character delimiters. The
    function will return noteproperty names as
    a header row.
  #>
  param(
    [Parameter(Mandatory=$True,ValueFromPipeLine=$True,Position=0)]
      [pscustomobject[]]$arrObject,
    [Parameter(Mandatory=$False,Position=1)]
          [String]$strDelimiter=":|"
  )
  # Create a header row from the names of NoteProperties
  $header = @()
  $header += ($arrObject | Get-Member -Type NoteProperty | Select-Object -ExpandProperty Name)

  # return the delimited header
  $header -join $strDelimiter
  # return delimited rows of data
  (
    $arrObject | ForEach-Object {
      $arrObject_ = $_              # Name the automatic variable, we need it in the inner loop
      ( $header | ForEach-Object {
        $arrObject_.$_
      } ) -join $strDelimiter
    }
  )
}

function Write-Log {
  [CmdletBinding()]
  Param(
    [String]
    [ValidateNotNullOrEmpty()]
    [Parameter(Mandatory = $True, Position = 1)]
    $Message,
    [ValidateSet('Error','Fatal','Warn','Info','Debug')]
    [String]
    $Level = 'Info'
  )
  try {
    $Level = $Level.ToUpper()
    $msg = "{0:yyyyMMddHHmm}: {1}: {2}" -f (Get-Date), $Level, $Message
    # write
    $LogStream.WriteLine($msg)
    $LogStream.Flush()
    $color = 'White'
    switch($Level) {
      'Debug' {$color = 'DarkYellow'}
      'Info'  {$color = 'Cyan'}
      'Warn'  {$color = 'Yellow'}
      'Error' {$color = 'Red'}
      'Fatal' {$color = 'DarkMagenta'}
    }
    # Commenting out the next line as it causes everything to log to the Action log
    #Write-Host -Object $msg -ForegroundColor $color
  }
  catch {
    throw "Error writing to log stream. $_"
  } 
}

function Is-LRRunning {
  param(
    $thisInvocation
  )
  # Looks for an instance of LR already running
  
  # Gather all PowerShell commandlines
  $parameters = @{Class = "Win32_Process"; Filter = "name='PowerShell.exe'"}
  $properties = ('ProcessId','CommandLine')
  try {
    $PSProcs = Get-CimInstance @parameters | Select-Object $properties
  } catch {
    try {
      # Fallback for older systems
      $PSProcs = Get-WmiObject @parameters | Select-Object $properties
    } catch {
      'Get-CimInstance and Get-WmiObject both failed. Not running LR to be safe.'
      return $True
    }
  }

  foreach($Proc in $PSProcs) {
    if ($Proc.ProcessId -ne $PID -and $Proc.CommandLine -match [regex]::escape($thisInvocation)) {
      return $True
    }
  }
  $False
}

function Remove-SensitiveFiles {
  $safeFilePattern = ".+\.zip|.+\.ps1$|.+\.psm1$|.+\.dll$|.+\.exe$|.+\.bat$|.+\.log$|^Custom_Collection\.json$|^Extended_Collection\.json$|^Memory_Collection\.json$|^Standard_Collection\.json$"
  
  foreach($file in (Get-ChildItem -Path $LRScriptRoot | Where-Object { $_.Name -notmatch $safeFilePattern })) {
    try {
      Remove-Item -Path $file.FullName
      Write-Log -Level 'Info' -Message $('Removed {0}' -f $file.Name)
    } catch {
      Write-Log -Level 'Warn' -Message $('Failed to remove {0} from {1}.' -f $file.Name, $LRScriptRoot)
    }
  }
}

function Extract-ZipArchive {
  param(
    [System.IO.FileInfo]$archivePath
  )

  # Extracts the $archivePath or throws a message
  try {
    $shell = New-Object -ComObject shell.application
    $archive = $shell.Namespace($archivePath.fullname)
    foreach ($item in $archive.items()) {
      $shell.Namespace($LRScriptRoot).copyhere($item)
    }
  } catch {
    $Message = ('Exception extracting {0} to {1}. Error {2}' -f $archivePath, $LRScriptRoot, $_)
    throw $Message
  }
}

function New-LoggerStream {
  param(
    [Parameter(Mandatory=$true,Position=1)][string]$Path
  )
  $mode = [System.IO.FileMode]::CreateNew
  $access = [System.IO.FileAccess]::ReadWrite
  $sharing = [System.IO.FileShare]::Read

  # create
  $fs = New-Object System.IO.FileStream($Path,$mode,$access,$sharing)
  $sw = New-Object System.IO.StreamWriter($fs)
  Write-Output $sw
}

Register-ModuleCollector -Name 'Memory' -Cmd (Get-Command Get-Memory)
Register-ModuleCollector -Name 'DriverDetails' -Cmd (Get-Command Get-DriverDetails)
Register-ModuleCollector -Name 'ProcessDetails' -Cmd (Get-Command Get-ProcessDetails)
Register-ModuleCollector -Name 'ModuleDetails' -Cmd (Get-Command Get-ModuleDetails)
Register-ModuleCollector -Name 'NetworkConnectionDetails' -Cmd (Get-Command Get-NetworkConnectionDetails)
Register-ModuleCollector -Name 'HandleDetails' -Cmd (Get-Command Get-HandleDetails)
Register-ModuleCollector -Name 'PFAmCache' -Cmd (Get-Command Get-PFAmCache)
Register-ModuleCollector -Name 'PFShellLink' -Cmd (Get-Command Get-PFShellLink)
Register-ModuleCollector -Name 'PFScheduledJob' -Cmd (Get-Command Get-PFScheduledJob)
Register-ModuleCollector -Name 'PFRecentFileCache' -Cmd (Get-Command Get-PFRecentFileCache)
Register-ModuleCollector -Name 'PFUserAssist' -Cmd (Get-Command Get-PFUserAssist)
Register-ModuleCollector -Name 'PFShimCache' -Cmd (Get-Command Get-PFShimCache)
Register-ModuleCollector -Name 'PFPrefetch' -Cmd (Get-Command Get-PFPrefetch)
Register-ModuleCollector -Name 'PFMasterBootRecord' -Cmd (Get-Command Get-PFMasterBootRecord)
Register-ModuleCollector -Name 'Autoruns' -Cmd (Get-Command Get-Autoruns)

# Force the console input and output encoding to be UTF-8, no BOM.
[Console]::InputEncoding = new-object System.Text.UTF8Encoding $false
[Console]::OutputEncoding = new-object System.Text.UTF8Encoding $false

# Set up a global for logging
if (-Not(Get-Variable -Name LogFile -EA 'SilentlyContinue')) {
  $LogFile = (Get-Location | Select-Object -ExpandProperty Path) + "\"+ (Get-Date -format yyyyMMddHHmm) + "_LR.log"
} # End global logging prep
$LogStream = New-LoggerStream -Path $LogFile

# We can't rely on the automatic $LRScriptRoot variable in PSv3+
# because of the way we run this content via:
# Get-Content <script> | Out-String | IEX
# which makes $LRScriptRoot == "Get-Content <script>"
# but we need this code for some cases, like when running Pester 
# tests or running live-response.ps1 from the cli for dev & debug.
$LRScriptRoot = ($pwd).tostring()

# Setting an environment variable for the Tanium Client directory as
# some collection configuration options may depend on it.
Set-Item -Path env:TANIUMDIR -Value (GetTaniumDir)

if ($null -eq $PSVersionTable) {
  "Live Response requires PSv2 or later."
  Remove-SensitiveFiles
  exit
} elseif ($PSVersionTable.PSVersion.Major -lt 3) {
  # This code is needed because of this issue:
  # http://www.leeholmes.com/blog/2008/07/30/workaround-the-os-handles-position-is-not-what-filestream-expected/
  
  $parentProcessId = (Get-WmiObject -class win32_process -Filter "ProcessId='$pid'").parentProcessId
  $parentProcessName = Get-Process -id $parentProcessId | Select-Object -ExpandProperty name
  if ($parentProcessName -eq 'cmd') {
    $bindingFlags = [Reflection.BindingFlags] "Instance,NonPublic,GetField"
    $objectRef = $host.GetType().GetField("externalHostRef", $bindingFlags).GetValue($host)
    $bindingFlags = [Reflection.BindingFlags] "Instance,NonPublic,GetProperty"
    $consoleHost = $objectRef.GetType().GetProperty("Value", $bindingFlags).GetValue($objectRef, @())
    [void] $consoleHost.GetType().GetProperty("IsStandardOutputRedirected", $bindingFlags).GetValue($consoleHost, @())
    $bindingFlags = [Reflection.BindingFlags] "Instance,NonPublic,GetField"
    $field = $consoleHost.GetType().GetField("standardOutputWriter", $bindingFlags)
    $field.SetValue($consoleHost, [Console]::Out)
    $field2 = $consoleHost.GetType().GetField("standardErrorWriter", $bindingFlags)
    $field2.SetValue($consoleHost, [Console]::Out)
  }
}

# Get the first part of our command line since 
# we don't want to run twice even with different arguments
$thisInvocation = ($MyInvocation.Line -split "-replace")[0]
if (Is-LRRunning -thisInvocation $thisInvocation) {
  'An instance of Live Response is already running.'
  Remove-SensitiveFiles
  Exit
}

# global scoped
$datetime = Get-Date -format yyyyMMddHHmm
if ((Get-OSVersion) -gt [version]"5.2") {
  $FQDN = (Get-WmiObject win32_ComputerSystem).DNSHostName + "." + (Get-WmiObject win32_ComputerSystem).Domain
} else {
  $FQDN = $env:ComputerName + "." + $env:UserDNSDomain
}
$dataDest = "$datetime-$FQDN"
$params = @{
  'dataDest' = $dataDest
}

Write-Log -Message "Loading Modules.........." -Level 'Info'

# PowerForensics is a dll module, not a script module so we can load it via function call
LoadPowerForensics
$Message = "PowerForensics loaded..."
Write-Log -Message $Message -Level 'Info'

try {
  Get-Content ($LRScriptRoot + '\json.psm1') | Out-String | Invoke-Expression
} catch {
  $Message = "Failed loading json.psm1. Verify that it is in the package. Error: $_"
  Write-Log -Message $Message -Level 'Fatal'
  Remove-SensitiveFiles
  MoveLog @params
  throw $Message
}
$Message = "JSON Parser loaded..."
Write-Log -Message $Message -Level 'Info'

# Global hashtable for file hashes
$FileHashes = @{}
# Set some other globals we may need
try {
  $SysVolume = (Get-WmiObject -Class Win32_OperatingSystem).SystemDrive
  $Windir    = (Get-WmiObject -Class Win32_OperatingSystem).WindowsDirectory
} catch {
  $Message = "Failed to get system volume and/or Windows directory. Error: $_"
  Write-Log -Message $Message -Level 'Fatal'
  Remove-SensitiveFiles
  MoveLog
  throw $Message
}

$script:options = @{}

Write-Log -Message ('PSTaniumFileTransfer version: {0}' -f (Get-Item ($LRScriptRoot + "\PSTaniumFileTransfer.dll")).VersionInfo.ProductVersion) -Level 'Info'
Import-Module ($LRScriptRoot + "\PSTaniumFileTransfer.dll")
$TFT = New-TaniumFileTransfer -Path ($LRScriptRoot + "\taniumfiletransfer.exe") -ConfigFile ($LRScriptRoot + "\$method")
try {
  $outfile = "$dataDest/LRConnectionTest"
  $stream = New-Object -TypeName System.IO.MemoryStream
  $writer = New-Object -TypeName System.IO.StreamWriter($stream,[System.Text.Encoding]::UTF8)
  $writer.WriteLine('Testing destination.')
  $writer.Flush()
  $null = $stream.seek(0,0)
  $TFT.SendStream($stream,$outfile,(Get-Date),-1)
  $TFTGood = $True
} catch {
  Remove-SensitiveFiles
  $Message = ('TFT may have encountered an error. Review and include {0} when reporting this issue.' -f ($LRScriptRoot + '.log'))
  Write-Log -Message $Message -Level 'Fatal'
  $Message
} finally {
  $stream.Dispose()
  $writer.Dispose()
}

if ($TFTGood) {
  try {
    RunLiveResponse

  } catch {
    ('Caught exception running Live Response. Error: {0}' -f $_)
  } finally {

    try {
      $TFT.Stop()
    } catch {
      ('Failed to stop TFT. Error: {0}' -f $_)
    }

    Remove-SensitiveFiles
  }
}
#------------ INCLUDES after this line. Do not edit past this point -----
