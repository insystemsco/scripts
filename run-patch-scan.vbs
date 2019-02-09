'Tanium File Version:2.2.2.0011

Option Explicit

'requires access to 64-bit areas of registry for uninstall checking
'use 64-bit WUA scanner on 64-bit OS
x64Fix

' allow override
RunOverride

' Global classes
Dim tLog
Set tLog = New TaniumContentLog
tLog.Log "----------------Beginning Patch Scan----------------"

Dim tRandom
Set tRandom = New TaniumRandomSeed ' Performs Randomize

Dim tContentReg
Set tContentReg = New TaniumContentRegistry
' All functions share this same object
' all activity is in the PatchManagement key
' and are all string values
tContentReg.RegValueType = "REG_SZ"
tContentReg.ClientSubKey = "PatchManagement"
On Error Resume Next
If tContentReg.ErrorState Then
	tLog.log "Severe Patch Management Registry Error: " & tContentReg.ErrorMessage
	tLog.Log "Quitting"	
	WScript.Quit
End If
On Error Goto 0

' This is updated by Build Tools, must match Has Patch Tools
Dim strPatchToolsVersion : strPatchToolsVersion = "2.2.2.0011" ' match header

EnsureRunsOneCopy

' Require greater than specific version of Windows UpdateAgent
Dim strMinWUAVersion : strMinWUAVersion = "6.1.0022.4"
If WUAVersionTooLow(strMinWUAVersion) Then
	tLog.Log "Error: Windows Update Agent needs to be at least " & strMinWUAVersion _
		& " - please upgrade. Cannot continue."
	WScript.Quit
End If


'Argument handling
Dim ArgsParser
Set ArgsParser = New TaniumNamedArgsParser
ParseArgs ArgsParser

If ArgsParser.GetArg("DoNotSaveOptions").ArgValue Then
	tLog.Log "One time argument usage, do not save command line options to registry"
Else
	' Put set values into registry
	MakeSticky ArgsParser, tContentReg
End If

' Create a config - combination of default values and passed in arguments - for use in this script
Dim dictPConfig
Set dictPConfig = CreateObject("Scripting.Dictionary")
dictPConfig.CompareMode = vbTextCompare
' Load default values
LoadDefaultConfig ArgsParser,dictPConfig
' Read from Registry (parsed values are here now if it is 'sticky')
LoadRegConfig tContentReg, dictPConfig
' Load parsed values - in case it not 'sticky'
LoadParsedConfig ArgsParser,dictPConfig
EchoConfig dictPConfig
RandomSleep TryFromDict(dictPConfig,"RandomWaitTimeInSeconds",0)

'Globals needed throughout script
Dim wuaService,wuaNeedsStop,wuaNeedsDisabled,intLocale
Dim dtmStartTime,dtmTempTime,intTotalRunSeconds
intTotalRunSeconds = "Error Calculating"
dtmStartTime = Now()

' for lookup of any error codes
Dim dictErrorCodes : Set dictErrorCodes = CreateErrorCodesDict

' can set global locale code via content
' the LocaleID string value
intLocale = GetTaniumLocale()
SetLocale(intLocale) ' sets locale options (date, commas and decimals, etc ...)

' check for / run 'pre' files
' put any preflight vbscript files in a directory called run-patch-scan\pre
RunFilesInDir("pre")

CheckWindowsUpdate() ' Ensure service started and store previous state

' Build update history via the WU API
' This is the kind of data that populates showing updates via the
' checkbox in Add/Remove programs. We use this data to track the
' updates which have been installed - in the update history file
Dim dictLiveHistory
FillAPIHistoryDictWithTimings dictLiveHistory

tLog.Log "Running patch scan"

Dim dictMSIPatchNames : Set dictMSIPatchNames = CreateObject("Scripting.Dictionary")
' look at all installed products (MSI based) and their patches, and list all names here. Should be global, done one time only
BuildMSIPatchNamesDict dictMSIPatchNames

Dim dictCabs,strCabPath
' Tanium Patch Scan supports custom cab file scanning.
' An Example would be the Windows XP extended support cabs available
' on connect.microsoft.com
FillCustomCabsDict dictCabs
' The default cab is automatically added
For Each strCabPath In dictCabs.Keys
	' takes the path to a cab and also the file name of the cab file
	RunPatchScan strCabPath,dictCabs.Item(strCabPath)
Next

tLog.Log "Finished running patch scan, sleeping for 4 seconds"
WScript.Sleep(4000)
StopWindowsUpdate() ' restore service's previous state

' put files in a directory called run-patch-scan\post
' candidates would be actions to be done after a patch scan
RunFilesInDir("post")

' Write total time to registry
intTotalRunSeconds = DateDiff("s",dtmStartTime,Now())
tLog.log "Patch scan complete in " & intTotalRunSeconds & " seconds"
tContentReg.ErrorClear
tContentReg.ValueName = "LastScanDurationS"
tContentReg.Data = intTotalRunSeconds
tContentReg.Write
If tContentReg.ErrorState Then
	tLog.Log "Error: " & tContentReg.ErrorMessage
End If

WScript.Quit

Function StopWindowsUpdate()
	Dim oShell,objWMIService,colServices,objService,WuaService
	Dim strService,strServiceStatus,strServiceMode
	
	strService = "wuauserv"
	If wuaNeedsStop Or wuaNeedsDisabled Then 
		tLog.log "Stopping Windows Update service"

		Set oShell = WScript.CreateObject ("WScript.Shell")
		oShell.run "net stop wuauserv /y", 0, True
		Set oShell = Nothing
	End If

	If wuaNeedsDisabled Then

		Set objWMIService = GetObject("winmgmts:" &  "{impersonationLevel=impersonate}!\\.\root\cimv2")  
		Set colServices = objWMIService.ExecQuery ("select State, StartMode from win32_Service where Name='"&strService&"'")    
		
		For Each objService in colServices
			strServiceStatus = objService.State
			strServiceMode = objService.StartMode
			Set WuaService = objService
		Next
	
		If IsEmpty(strServiceStatus) Then
			tLog.log "Scan Error: Cannot find Windows Update (wuauserv)"
			Exit Function
		End If
		
		tLog.log "Return code: " & WuaService.ChangeStartMode("Disabled")
	End If
End Function

Function CheckWindowsUpdate()
	'Check to see if Windows Update Service needs to be enabled and/or stopped at end
	Dim objWMIService,objService
	Dim colComputer, objComputer, strService, colServices
	Dim strServiceStatus,strServiceMode
	
	strService = "wuauserv"
	
	Set objWMIService = GetObject("winmgmts:" &  "{impersonationLevel=impersonate}!\\.\root\cimv2")  
	Set colServices = objWMIService.ExecQuery ("select State, StartMode from win32_Service where Name='"&strService&"'")    
	
	
	For Each objService in colServices
		strServiceStatus = objService.State
		strServiceMode = objService.StartMode
		Set WuaService = objService
	Next
	
	
	If IsEmpty(strServiceStatus) Then
		tLog.log "Scan Error: Cannot find Windows Update (wuauserv)"
		WScript.Quit
	End If
	
	If strServiceStatus = "Stopped" Then
		tLog.log "Windows Update is stopped, will stop after Patch Scan Complete"
		wuaNeedsStop = true
	End If
	
	If strServiceMode = "Disabled" Then
		tLog.log "Attempting to change 'Windows Update' start mode to 'Manual'"
		tLog.log "Return code: " & WuaService.ChangeStartMode("Manual")
		wuaNeedsDisabled = True
	End If

End Function

Function GetAllCabs
	tLog.log "Looking for extra patch scan cab files"
	' Returns a dictionary of patch scan cabinet files
	Dim dictCabs,objFSO,objCabFolder,objFile,strToolsDir
	Dim strExtraCabDir,strCabPath
	Set dictCabs = CreateObject("Scripting.Dictionary")
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strToolsDir = GetTaniumDir("Tools")
	strExtraCabDir = strToolsDir&"ExtraPatchCabs"
		
	If objFSO.FolderExists(strExtraCabDir) Then
		tLog.log "Extra Cab Folder exists, looking for additional cab files"
		Set objCabFolder = objFSO.GetFolder(strExtraCabDir)
		For Each objFile In objCabFolder.Files
			If LCase(Right(objFile.Name,4)) = ".cab" Then
				If Not dictCabs.Exists(objFile.Path) Then
					dictCabs.Add objFile.Path,objFile.Name
					tLog.log "Found extra cab file " & objFile.Path
				End If
			End If
		Next
	End If
	
	' Now add the default distributed wsusscn2.cab
	strCabPath = strToolsDir & "wsusscn2.cab"
	If objFSO.FileExists(strCabPath) And Not dictCabs.Exists(strCabPath) Then
		dictCabs.Add strCabPath,"wsusscn2.cab"
	End If
	Set GetAllCabs = dictCabs
End Function 'GetAllCabs

Function RunPatchScan(strCabPath,strCabName)
	Dim objFSO,objTextFile,oAgentInfo,ProductVersion
	Dim strSep, strClientDirPath,strToolsDir,strScanDir,strResultsPath
	Dim strResultsReadablePath
	Dim strCVEFilePath,objCVEDict,strEntry,dictHistory
	Dim arrHistoryLine,strHistoryLine
	Dim intHistoryFileMode,bBadHistoryLines,bGoodHistoryLine,dtmUTCTime
	Dim dictAlreadySupersededHistoryEntries,dictAlreadyFirstNeededHistoryEntries
	Dim bBadHistoryVersion,strDesiredHistoryVersion,strUseScanSource
	
	On Error Resume Next ' disable when necessary, but should remain on!
	
	strSep = "|"
	'title|severity|bulletins|date|download|filename|status|updateId|size|KB|CVEs
	dtmUTCTime = DateAdd("n",-GetTZBias,Now())
	dtmUTCTime = FormatDateTime(dtmUTCTime,vbShortDate)&" "&FormatDateTime(dtmUTCTime,vbShortTime)
	tLog.log "Time in UTC is " & dtmUTCTime
	
	strClientDirPath = GetTaniumDir("")
	strToolsDir = GetTaniumDir("Tools")
	strScanDir = GetTaniumDir("Tools\Scans")
	
	strUseScanSource = LCase(TryFromDict(dictPConfig,"UseScanSource","cab"))
	
	If LCase(strCabname) = "wsusscn2.cab" Then
		strResultsPath = strScanDir & "patchresults.txt"
		strResultsReadablePath = strScanDir & "patchresultsreadable.txt"
	Else
		strResultsPath = strScanDir & "patchresults-"&strCabName&".txt"
		strResultsReadablePath = strScanDir & "patchresultsreadable-"&strCabName&".txt"
	End If
	tLog.log "Scanning against " & strCabName & " and storing results at " & strResultsReadablePath
    If objFSO.FileExists(strResultsErrorPath) Then
        fso.DeleteFile(strResultsErrorPath)
    End If
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objTextFile = objFSO.OpenTextFile(strResultsPath, 2, true)
	
	' Read in CVE data
	strCVEFilePath = GetTaniumDir("Tools")&"MS-CVEs.dat"
	
	If TryFromDict(dictPConfig,"ShowTimings",False) Then
		tLog.log "CVE Data File Read In: Started"
		dtmTempTime = Timer()
	End If

	Set objCVEDict = GetCVEDict(strCVEFilePath)

	If TryFromDict(dictPConfig,"ShowTimings",False) Then 
		tLog.log "CVE Data File Read In: End - Took " & Abs(Timer() - dtmTempTime) & " seconds"
	End If
	
	If Not objFSO.FileExists(strCabPath) And strUseScanSource = "cab" Then
		tLog.log strCabPath & " not deployed"
		objTextFile.WriteLine("Scan Error: "&strCabName&" not deployed")		
		objTextFile.Close
		
		Exit Function
	End If
	
	Set oAgentInfo = CreateObject("Microsoft.Update.AgentInfo")
	
	If IsEmpty(oAgentInfo) Then
		tLog.log "Error creating Update Agent object"
		objTextFile.WriteLine("Scan Error: Cannot create Update Agent object")
		objTextFile.Close
		
		StopWindowsUpdate()
		Exit Function
	End If
	
	' -------- Update History File ----------' 
	If TryFromDict(dictPConfig,"ShowTimings",False) Then 
		tLog.log "History File Read In: Started"
		dtmTempTime = Timer()
	End If

	' Create / write to update history text file
	Dim objHistoryTextFile
	Dim strHistoryTextFilePath,strHistoryTextReadableFilePath
	Dim strHistoryTextFileVersionLine	
	
	strHistoryTextFilePath = strScanDir&"patchhistory.txt"
	strHistoryTextReadableFilePath = strScanDir&"patchhistoryreadable.txt"

	' Create or clear the history file if it doesn't exist or has wrong version
	' sensors should assume this line exists and do a .skipline before doing real work

	bBadHistoryVersion = False
	strDesiredHistoryVersion = "6.1.314.2372" ' this build updates history output to account for modifications to error code output
	
	If Not objFSO.FileExists(strHistoryTextFilePath) Then
		objFSO.CreateTextFile strHistoryTextFilePath,True
		Set objHistoryTextFile = objFSO.OpenTextFile(strHistoryTextFilePath,2,True) ' write version header
		objHistoryTextFile.WriteLine("'Tanium Windows Patch History Version:"&strDesiredHistoryVersion)
		objHistoryTextFile.Close
		' reopen for read-in
		Set objHistoryTextFile = objFSO.OpenTextFile(strHistoryTextFilePath,1,True)
	Else
		strHistoryTextFileVersionLine = ""
		Set objHistoryTextFile = objFSO.OpenTextFile(strHistoryTextFilePath,1,True)
		If Not objHistoryTextFile.AtEndOfStream Then strHistoryTextFileVersionLine = objHistoryTextFile.ReadLine
		If Not InStr(strHistoryTextFileVersionLine,"'Tanium Windows Patch History Version:"&strDesiredHistoryVersion) = 1 Then
			tLog.log "History file has bad / no version string"
			bBadHistoryVersion = True
		End If
	End If
	
	' Read existing history snapshot into a dictionary object
	Set dictHistory = CreateObject("Scripting.Dictionary")
	Set dictAlreadySupersededHistoryEntries = CreateObject("Scripting.Dictionary")
	Set dictAlreadyFirstNeededHistoryEntries = CreateObject("Scripting.Dictionary")

	' the history file is read and appended as opposed to being overwritten each time.
	' If we change the format of the file, we are probably changing column count. 
	' The Version number in the header should change as well.
	' If the line entry does not match the column count, do not append and instead
	' overwrite the file later.
	' only unique lines are read into the dictionary object, and only unique lines
	' are written back.
	
	If Not bBadHistoryVersion Then
		intDesiredHistoryColumnCount = 14
		bBadHistoryLines = False
		While objHistoryTextFile.AtEndOfStream = False
			bGoodHistoryLine = True
			strHistoryLine = objHistoryTextFile.ReadLine
			If InStr(strHistoryLine,"'Tanium Windows Patch History Version:") = 1 Then 
				' skip processing
			Else
				arrHistoryLine = Split(strHistoryLine,"|")
				If IsArray(arrHistoryLine) Then
					If UBound(arrHistoryLine) <> (intDesiredHistoryColumnCount-1) Then
						tLog.log "bad history file line detected:" & strHistoryLine
						bGoodHistoryLine = False
						bBadHistoryLines = True
					End If
					If (Not dictHistory.Exists(strHistoryLine)) And bGoodHistoryLine Then
						dictHistory.Add strHistoryLine,1
						' note if it's a superseded line
						If Trim(LCase(arrHistoryLine(1))) = "superseded" Then
							If Not dictAlreadySupersededHistoryEntries.Exists(arrHistoryLine(0)) Then
								' add to dictionary of superseded entries.  Adding by GUID to save time in lookup
								dictAlreadySupersededHistoryEntries.Add arrHistoryLine(0),strHistoryLine
							End If
						End If
						' note if it's a first needed line
						If Trim(LCase(arrHistoryLine(1))) = "firstneeded" Then
							If Not dictAlreadyFirstNeededHistoryEntries.Exists(arrHistoryLine(0)) Then
								' add to dictionary of superseded entries.  Adding by GUID to save time in lookup
								dictAlreadyFirstNeededHistoryEntries.Add arrHistoryLine(0),strHistoryLine
							End If
						End If					
					End If
					If Not bGoodHistoryLine Then
						tLog.log "skipped " & strHistoryLine
					End If
				End If
			End If
		Wend
	End If
	
	Dim bClearHistoryOnBadLine
	If dictPConfig.Exists("ClearHistoryOnBadLine") Then
		bClearHistoryOnBadLine = dictPConfig.Item("ClearHistoryOnBadLine")
	Else
		bClearHistoryOnBadLine = False
	End If
	
	' note that because we append, we will always retain any bad line
	' sensors should use simliar code here (checking column count) to ignore bad lines
	If ( bClearHistoryOnBadLine And bBadHistoryLines ) Or bBadHistoryVersion Then ' overwrite file
		intHistoryFileMode = 2 ' overwrite
		tLog.log "Will overwrite history results"		
	Else
		intHistoryFileMode = 8 ' append
	End If
	
	' Will need to re-open for either writing or appending now that it's in dictionary
	objHistoryTextFile.Close

	If TryFromDict(dictPConfig,"ShowTimings",False) Then 
		tLog.log "History File Read In: End - Took " & Abs(Timer() - dtmTempTime) & " seconds"
	End If

	Set objHistoryTextFile = objFSO.OpenTextFile(strHistoryTextFilePath,intHistoryFileMode,True)

	' write new version header before writing any data
	If bBadHistoryVersion Then objHistoryTextFile.WriteLine("'Tanium Windows Patch History Version:"&strDesiredHistoryVersion)

	''###SEARCH FOR PACKAGES###''
	Dim UpdateSession,UpdateServiceManager,UpdateService,UpdateSearcher
	Dim SearchResult,Updates,update
	Dim numInstalled,numNotInstalled,strInstalled,i,strUpdateID
	Dim bulletins,bulletin,bundledUpdate,contents,FileToCopy
	Dim strBulletins,severity,strDownloadSize,strCVEIDs,strResultsCVEIDs
	Dim urls,q,creationDate,mo,yr,j,k,words,filenames,strTemp,strKBArticle
	Dim strKBArticles,strFirstNeededDate
	Dim strSupersededDate,dtmSupersededDate
	Dim intDesiredHistoryColumnCount,strHistoryOutputLine,dtmOSInstallDate
	Dim bPartialCVEs,bCVEFileCreated,strSupersededUpdateID,strServicePackID
	Dim strDictHistoryLine,arrDictHistoryLine,dictSupersededIDs,strGUID
	Dim dictValidServicePacks,strCategory, dictUpdatesSupersededByValidServicePacks,bValidUpdate,dictServicePacks
	Dim strOSMajorVersion,bVistaPlus,sngOSMajorVersion,dictSupersededByOnlyServicePacks
	Dim strSupersededItemString,dictUpdateTitles,strScanError,strTrulyUniqueID,strRebootBehavior,strUpdateImpact
	Dim strRevisionNumber
	Dim arrUpdateImpact : arrUpdateImpact = Array("Normal","Minor","Exclusive Handling") 
	Dim arrUpdateRebootBehavior :  arrUpdateRebootBehavior = Array("Never","Always","Maybe")
	Dim bScanSourceOverrideToCab,bShowSupersededUpdates,bDisableMicrosoftUpdate
	
	If Not strCabName = "wsusscn2.cab" Then
		' it's a cab file, guarantee we're not scanning online
		bScanSourceOverrideToCab = True
		If Not strUseScanSource = "cab" Then 
			tLog.log "Directed to use non-cab scan source, but " _
				& "must scan locally against atypical cab file " & strCabName
		End If
	Else
		' Allow scan online
		bScanSourceOverrideToCab = False
	End If
	
	Set UpdateSession = CreateObject("Microsoft.Update.Session")
	
	UpdateSession.UserLocale = intLocale ' Changes output language
	' Log that it's Tanium (as seen in WindowsUpdate.log)
	UpdateSession.ClientApplicationID = "Tanium Patch Scan " & strPatchToolsVersion
	UpdateSession.WebProxy.AutoDetect = True ' Set proxy to use IE autodetect settings
	
	If dictPConfig.Exists("UseScanSource") Then
		strUseScanSource = dictPConfig.Item("UseScanSource")
	Else
		tLog.Log "No ScanSource specified (Is there a default arg value?) - choosing cab"
		strUseScanSource = "cab"
	End If
	
	Set UpdateServiceManager = CreateObject("Microsoft.Update.ServiceManager")

	bDisableMicrosoftUpdate = TryFromDict(dictPConfig,"DisableMicrosoftUpdate",False)
	If Not strUseScanSource = "cab" And Not bScanSourceOverrideToCab Then
		If Not bDisableMicrosoftUpdate Then
			' add microsoft update if we're using an online scan source
			On Error Resume Next
			Set UpdateService = UpdateServiceManager.AddService2("7971f918-a847-4430-9279-4a52d1efe18d",7,"")
			If Err.Number <> 0 Then
				tLog.log "Could not set update service to Microsoft Update, Error was " & Err.Number
			End If
			On Error Goto 0
			Set UpdateSearcher = UpdateSession.CreateUpdateSearcher()
			UpdateSearcher.ServiceID = UpdateService.ServiceID ' set microsoft update
		Else
			tLog.log "Would have scanned against Microsoft Update, but skipping"
			Set UpdateSearcher = UpdateSession.CreateUpdateSearcher()
		End If
	Else
		' for cab, add cab file
		Set UpdateSearcher = UpdateSession.CreateUpdateSearcher()
		If Not objFSO.FileExists(strCabPath) Then
			tLog.log "Scan Error: Cannot locate offline cab file at "&strCabPath&" when scan source is set to cab"
			objTextFile.WriteLine("Scan Error: Cannot locate offline cab file at "&strCabPath&" when scan source is set to cab")
			objTextFile.Close
			StopWindowsUpdate()
			Exit Function
		Else ' cab is there
			Set UpdateService = UpdateServiceManager.AddScanPackageService("Offline Sync Service", strCabPath)
			If IsEmpty(UpdateService) Then
				tLog.log "Error creating Offline Sync Service object (Windows Update may be disabled or bad cab file), aborting scan"
                objTextFile.WriteLine("Scan Error: Error creating Offline Sync Service object (Windows Update may be disabled or bad cab file)")
                objTextFile.Close
				StopWindowsUpdate()
				Exit Function
			End If
		End If
	End If
	
	bShowSupersededUpdates = TryFromDict(dictPConfig,"ConsiderSupersededUpdates",True)
	
	If Not bShowSupersededUpdates Then tLog.log "Will not consider any superseded updates"
	UpdateSearcher.IncludePotentiallySupersededUpdates = bShowSupersededUpdates
	
	If bScanSourceOverrideToCab Then ' force cab settings
		UpdateSearcher.ServerSelection = 3
		UpdateSearcher.ServiceID = UpdateService.ServiceID
	Else
		' how to scan (/UseScanSource argument)
		Select Case strUseScanSource
			Case "systemdefault"
				UpdateSearcher.ServerSelection = 0
			Case "wsus"
				UpdateSearcher.ServerSelection = 1
			Case "internet"
				UpdateSearcher.ServerSelection = 2
			Case "optimal" ' this is online with Microsoft Update
				If bDisableMicrosoftUpdate Then
					UpdateSearcher.ServerSelection = 2 'Windows Update only
				Else
					UpdateSearcher.ServerSelection = 3 ' now requires a serviceID if it's Microsoft Update					
					UpdateSearcher.ServiceID = UpdateService.ServiceID
				End If
			Case "cab"
				UpdateSearcher.ServerSelection = 3
				UpdateSearcher.ServiceID = UpdateService.ServiceID
			Case Else ' unknown option, defaults to optimal
				tLog.log "Unknown Scan Source reference, choosing local cab file"
				UpdateSearcher.ServerSelection = 3
				UpdateSearcher.ServiceID = UpdateService.ServiceID
				strUseScanSource = "cab"
		End Select
		
		' sleep if scan source was a service based, non-cab source
		' skip sleep if non-default cab file specified
		If Not strUseScanSource = "cab" Then
			tLog.log "Scan source is online / service-based, sleeping for a pre-determined amount of time"
			RandomSleep(TryFromDict(dictPConfig,"OnlineScanRandomWaitTimeInSeconds",0))
		End If
	End If
	
	' Get OS version (needed to determine uninstallability)
	' Windows XP doesn't use update.isuninstallable
	strOSMajorVersion = GetOSMajorVersion
	If Not IsNumeric(strOSMajorVersion) Then
		' the OS major version is indeterminate, assume vista+
		' so the API will check for uninstallability
		bVistaPlus = True
	Else
		sngOSMajorVersion = CSng(strOSMajorVersion)
		If sngOSMajorVersion >= 6.0 Then 
			bVistaPlus = True
		Else
			bVistaPlus = False
		End If
	End If
	
	If TryFromDict(dictPConfig,"ShowTimings",False) Then 
		tLog.log "Update Search: Started"
		dtmTempTime = Timer()
	End If
	
	'Set SearchResult = UpdateSearcher.Search("Type='Software'")
	Set SearchResult = UpdateSearcher.Search("Type='Software' and IsHidden=1 or IsHidden=0")
	If Err.Number <> 0 Then ' Try scanning with cab file as failover
		strScanError = Err.Number
		If dictErrorCodes.Exists("0x"&CStr(Hex(strScanError))) Then 
			strScanError = dictErrorCodes.Item("0x"&CStr(Hex(strScanError)))
		End If
		
		tLog.log "Cannot complete patch scan via scan source " & strUseScanSource & ", Scan Error: " & strScanError
		tLog.log "Retrying with offline cab file as backup scan source"
		'Retry scan with offline cab file as last resort
		If Not objFSO.FileExists(strCabPath) Then
			tLog.log "Scan Error: Cannot locate offline cab file at "&strCabPath&" when scan source is set to cab"
			objTextFile.WriteLine("Scan Error: Cannot locate offline cab file at "&strCabPath&" when scan source is set to cab")
			objTextFile.Close
			StopWindowsUpdate()
			Exit Function
		End If
		Set UpdateService = UpdateServiceManager.AddScanPackageService("Offline Sync Service", strCabPath)
		' Reset update objects
		' Set UpdateServiceManager = CreateObject("Microsoft.Update.ServiceManager") 'redundant
		Set UpdateSearcher = UpdateSession.CreateUpdateSearcher()
		' UpdateSearcher.ServiceID = UpdateService.ServiceID 'redundant
		UpdateSearcher.IncludePotentiallySupersededUpdates = bShowSupersededUpdates	
		If IsEmpty(UpdateService) Then
			tLog.log "Error creating Offline Sync Service object (Windows Update may be disabled)"
			objTextFile.WriteLine("Scan Error: Cannot create Offline Sync Service object (Windows Update may be disabled)")
			objTextFile.Close
			
			StopWindowsUpdate()
			Exit Function
		End If
		UpdateSearcher.ServerSelection = 3
		UpdateSearcher.ServiceID = UpdateService.ServiceID	' only set when cab	
		Err.Clear
		Set SearchResult = UpdateSearcher.Search("Type='Software' and IsHidden=1 or IsHidden=0")
		If Err.Number <> 0 Then
			strScanError = Err.Number
			If dictErrorCodes.Exists("0x"&CStr(Hex(strScanError))) Then 
				strScanError = dictErrorCodes.Item("0x"&CStr(Hex(strScanError)))
			End If
			tLog.log "Cannot complete patch scan with offline cab file as backup scan source, Scan Error: " & strScanError
			objTextFile.WriteLine("Scan Error: " & strScanError)
			objTextFile.Close

			StopWindowsUpdate()
			Exit Function
		Else ' no error scanning with cab as backup
			tLog.log "Scan failed with " & strUseScanSource & " as scan source, but was successful with offline cab file as backup"
			' WritePatchManagementValueToRegistry strUseScanSource,"CabScanUsedAsBackup" ' note that a failover occurred
			tContentReg.ValueName = "CabScanUsedAsBackup"
			tContentReg.Data = strUseScanSource
			On Error Resume Next
			tContentReg.Write
			If Err.Number <> 0 Then
				tLog.Log "Error: Could not store whether the cab file was used as a backup: "&tContentReg.ErrorMessage
				Err.Clear
				tContentReg.ErrorClear
			End If
			On Error Goto 0
		End If
	Else ' clear if cabscanusedasbackup is set as there are no errors with chosen scan source
		tContentReg.ValueName = "CabScanUsedAsBackup"
		On Error Resume Next
		tContentReg.DeleteVal
		If Err.Number <> 0 Then
			' We try the deletion whether it exists or not, so this may throw an error each time
			tContentReg.ErrorClear
			Err.Clear
		End If
		On Error Goto 0
	End If
	
	Set Updates = SearchResult.Updates

	dtmOSInstallDate = GetOSInstallDate

	' map all updates to their titles by ID (useful for debugging)
	Set dictUpdateTitles = CreateObject("Scripting.Dictionary")
	For i = 0 To SearchResult.Updates.Count-1
		Set update = SearchResult.Updates.Item(i)
		If Not dictUpdateTitles.Exists(update.Identity.UpdateID) Then
			' tLog.log "Adding " & update.Identity.UpdateID & " " & update.Title
			dictUpdateTitles.Add update.Identity.UpdateID, update.Title
		End If
	Next

	' Get a list of all update IDs and what they are superseded by
	Set dictSupersededIDs = CreateObject("Scripting.Dictionary")
	strSupersededUpdateID = ""
	For i = 0 to searchResult.Updates.Count-1
		Set update = SearchResult.Updates.Item(i)
		For Each strSupersededUpdateID In update.SupersededUpdateIDs
			' tLog.log "Superseded update ID is " & strSupersededUpdateID
			If Not dictSupersededIDs.Exists(strSupersededUpdateID) Then
				' This superseded ID is seen for the first time, simply add the details of the ID that supersedes it.
				' for each superseded update, track the title of the update that superseded it and its publish date
				' tLog.log "First add of " & strSupersededUpdateID & " : " & update.Identity.UpdateID&strSep&update.Title&strSep&update.LastDeploymentChangeTime
				dictSupersededIDs.Add strSupersededUpdateID,update.Identity.UpdateID&strSep&update.Title&strSep&update.LastDeploymentChangeTime
			Else
				' superseded ID is already seen - read in the already existing superseder information and second delimiter
				' to the item section, adding the new superseder information on the end.
				strSupersededItemString = dictSupersededIDs.Item(strSupersededUpdateID)
				' tLog.log  "already know about " & strSupersededItemString
				' tLog.log "concat adding to " & strSupersededUpdateID & " : " & strSupersededItemString&"|-|"&update.Identity.UpdateID&strSep&update.Title&strSep&update.LastDeploymentChangeTime
				dictSupersededIDs.Item(strSupersededUpdateID) = strSupersededItemString&"|-|"&update.Identity.UpdateID&strSep&update.Title&strSep&update.LastDeploymentChangeTime
			End If
		Next
	Next

	' Useful debug output via /PrintSupersedenceInfo command line argument
	' This will print detail on which updates are superseded by others
	If TryFromDict(dictPConfig,"PrintSupersedenceInfo",False) Then PrintSupersedenceInfo dictSupersededIDs, dictUpdateTitles

	strSupersededUpdateID = ""
	Set dictValidServicePacks = CreateObject("Scripting.Dictionary")
	Set dictServicePacks = CreateObject("Scripting.Dictionary")
	Set dictUpdatesSupersededByValidServicePacks = CreateObject("Scripting.Dictionary")
	' next get a list of service packs that are not superseded
	' and get a list of update IDs that those valid service packs supersede
	For i = 0 to searchResult.Updates.Count-1
		Set update = SearchResult.Updates.Item(i)
		For Each strCategory In update.Categories
			If LCase(strCategory) = "service packs" Then
				If Not dictServicePacks.Exists(update.Identity.UpdateID) Then dictServicePacks.Add update.Identity.UpdateID,update.Title
				If Not dictSupersededIDs.Exists(update.Identity.UpdateID) Then
					If Not dictValidServicePacks.Exists(update.Identity.UpdateID) Then
						dictValidServicePacks.Add update.Identity.UpdateID,update.Title
						' tLog.log update.Title & " Is a non-superseded service pack"
						For Each strSupersededUpdateID In update.SupersededUpdateIDs
							' record the IDs that the service pack supersedes
							If Not dictUpdatesSupersededByValidServicePacks.Exists(strSupersededUpdateID) Then
								' tLog.log "service Pack " & update.Title & " supersedes " & strSupersededUpdateID
								dictUpdatesSupersededByValidServicePacks.Add strSupersededUpdateID, update.Title
							End If
						Next
					End If
				End If
			End If
		Next
	Next


' loop through supersedence map dictionary and write a dictionary which holds updates that are superseded
' by any update which is not a service pack
	Dim strSupersededID,bSupersededByNonServicePack,arrSupersedersList,strSupersederInfo,strSupersederID
	Dim dictSupersededByAnyNonServicePack
	Set dictSupersededByAnyNonServicePack = CreateObject("Scripting.Dictionary")
	For Each strSupersededID In dictSupersededIDs.Keys
		bSupersededByNonServicePack = False ' false until we find any non-SP update
		arrSupersedersList = Split(dictSupersededIDs.Item(strSupersededID),"|-|")
		For Each strSupersederInfo In arrSupersedersList
			strSupersederID = Split(strSupersederInfo,"|")(0)
			'tLog.log "Checking superseder ID " & strSupersederID & " based on " & strSupersederInfo
			If Not dictServicePacks.Exists(strSupersederID) Then 
				bSupersededByNonServicePack = True
				'tLog.log "will add " & strSupersededID & " to dictionary, superseded by a non-service pack " & dictUpdateTitles.Item(strSupersededID)
			End If
		Next
		If bSupersededByNonServicePack Then
			If Not dictSupersededByAnyNonServicePack.Exists(strSupersededID) Then
				dictSupersededByAnyNonServicePack.Add strSupersededID, 1				
			End If
		End If
	Next
		
	' loop through the supersedence map dictionary and write another dictionary which holds updates that are superseded
	' only by updates which are service packs
	Set dictSupersededByOnlyServicePacks = CreateObject("Scripting.Dictionary")
	Dim bSupersededOnlyByServicePacks
	For Each strSupersededID In dictSupersededIDs.Keys
		bSupersededOnlyByServicePacks = True
		arrSupersedersList = Split(dictSupersededIDs.Item(strSupersededID),"|-|")
		For Each strSupersederInfo In arrSupersedersList
			strSupersederID = Split(strSupersederInfo,"|")(0)
			'tLog.log "Checking superseder ID " & strSupersederID & " based on " & strSupersederInfo
			If Not dictServicePacks.Exists(strSupersederID) Then 
				bSupersededOnlyByServicePacks = False
				'tLog.log "will  add " & strSupersededID & " to dictionary, superseded by a service pack " & dictServicePacks.Item(strSupersededID)
			End If
		Next
		If bSupersededOnlyByServicePacks Then
			If Not dictSupersededByOnlyServicePacks.Exists(strSupersededID) Then
				dictSupersededByOnlyServicePacks.Add strSupersededID, 1				
			End If
		End If
	Next

	If TryFromDict(dictPConfig,"ShowTimings",False) Then
		tLog.log "Update Search: End - Took " & Abs(Timer() - dtmTempTime) & " seconds"
	End If
	
	If searchResult.Updates.Count = 0 Then
	    objTextFile.WriteLine("Scan: No patches needed or installed")
		objTextFile.Close
		
		StopWindowsUpdate()
	    Exit Function
	End If
	
	numInstalled = 0
	numNotInstalled = 0
	
	bCVEFileCreated = False

	Dim hasher
	Set hasher = New MD5er

	' Hashing requires calculating the filenames and URLs
	' so store hash and other data and tie to result index
	Dim dictUpdateSearchResultIDtoDetails
	Set dictUpdateSearchResultIDtoDetails = CreateObject("Scripting.Dictionary")
	
	For i = 0 to searchResult.Updates.Count-1
		''get guid for update, save as updateID
	    Set update = searchResult.Updates.Item(I)
		strUpdateId = update.Identity.UpdateID

	'If dictSupersededIds.Exists(update.Identity.UpdateID) Then
	'	If update.IsInstalled Then
	'		tLog.log update.Title & "("&update.Identity.UpdateID & ") IS installed and appears to be in the superseded IDs dictionary"
	'	Else
	'		tLog.log update.Title & "("&update.Identity.UpdateID & ") is NOT installed and appears to be in the superseded IDs dictionary"
	'	End If
	'End If
		
		bValidUpdate = True
		' if an update is superseded only by a non-superseded service pack, it is valid
		' if it is superseded by any other update, it is not valid and we will not report on it as a needed update

		strKBArticles = " "
		For Each strKBArticle In update.KBArticleIDs
			strKBArticles = strKBArticles & "KB"&strKBArticle & " "
		Next
		strKBArticles = Trim(strKBArticles)

		bValidUpdate = IsValidUpdate(update,dictSupersededIDs,dictSupersededByAnyNonServicePack,dictSupersededByOnlyServicePacks,dictServicePacks,strKBArticles)
				
		If update.IsInstalled Then
			numInstalled = numInstalled + 1
			If bVistaPlus Then
				If sngOSMajorVersion = 6.0 Then
					' Uninstallation for Vista and Server 2008 requires that we check
					' the API and also the packages area of the registry
					If update.IsUninstallable And IsUninstallable(strKBArticles,update.Title) Then
						strInstalled = "Already Installed and Uninstallable"
					Else
						strInstalled = "Already Installed"
					End If
				Else
					' Uninstallation for Windows 7 and beyond just uses WUSA
					' so there is no need to check the registry to determine. The API
					' is authoratative
					If update.IsUninstallable Then
						strInstalled = "Already Installed and Uninstallable"
					Else
						strInstalled = "Already Installed"
					End If
				End If
			Else
				If IsUninstallable(strKBArticles,update.Title) Then 'non-vista, don't ask API
					strInstalled = "Already Installed and Uninstallable"
				Else
					strInstalled = "Already Installed"
				End If
			End If			
		Else
			' only count updates which are not installed that are valid (supersedence checking)
			If bValidUpdate Then
				numNotInstalled = numNotInstalled + 1
				tLog.log update.Title & " is not installed"
			End If
			strInstalled = "Not Installed"
		End If
		
	    urls = ""
		q = ""
	    creationDate = update.LastDeploymentChangeTime
	    
	    mo = Month(creationDate)
	    yr = Year(creationDate)
	    If Len(mo) = 1 Then
	    	mo = "0" & mo
	    End If
	    
	    creationDate = yr & "-" & mo
		' SCUP updates are never bundled
	    For j = 0 To update.DownloadContents.Count-1
    		Set contents = bundledUpdate.DownloadContents.Item(j)
			words = Split(contents.DownloadUrl, "/")
    		If urls = "" Then
    			urls = contents.DownloadUrl
    			filenames = words(UBound(words))
			Else
	    		urls = urls & "," & contents.DownloadUrl
	    		filenames = filenames & "," & words(UBound(words))
			End If
    	Next
		If Len(urls) = 0 Then ' not a SCUP update, traditional method
			' Typically, updates are in a bundle, loop through all of them
		    For j = 0 To update.BundledUpdates.Count-1
		    	Set bundledUpdate = update.BundledUpdates.Item(J)
		    		    	
		    	For k = 0 To bundledUpdate.DownloadContents.Count-1
		    		Set contents = bundledUpdate.DownloadContents.Item(K)
					words = Split(contents.DownloadUrl, "/")
		    		If urls = "" Then
		    			urls = contents.DownloadUrl
		    			filenames = words(UBound(words))
					Else
			    		urls = urls & "," & contents.DownloadUrl
			    		filenames = filenames & "," & words(UBound(words))
					End If 
		    	Next
		    Next
		 End If
		'look through bulletin IDs
		Set bulletins = update.SecurityBulletinIDs
		strBulletins = " "
		strCVEIDs = " "
		
		For Each bulletin In bulletins
			strBulletins = strBulletins & bulletin & " "
			strCVEIDs = strCVEIDs & objCVEDict.Item(UCase(bulletin)) & " "
		Next
	    strBulletins = trim(strBulletins)
	    strCVEIDs = Trim(strCVEIDs)
	    ' keep separate CVEIds value for results file
	    strResultsCVEIDs = strCVEIDs
				
	    severity = ""
	    If (IsNull(update.MsrcSeverity) or update.MsrcSeverity = "")  then 
	    	severity = "None"
	    else
			severity = update.MsrcSeverity
	    end If

	    strRebootBehavior = ""
	    If (IsNull(update.InstallationBehavior.RebootBehavior) or update.InstallationBehavior.RebootBehavior = "")  then 
	    	strRebootBehavior = "Unknown"
	    Else
			strRebootBehavior = arrUpdateRebootBehavior(update.InstallationBehavior.RebootBehavior)
	    end If

	    strRevisionNumber = ""
	    If (IsNull(update.Identity.RevisionNumber) or update.Identity.RevisionNumber = "")  then 
	    	strRevisionNumber = "Unknown"
	    else
			strRevisionNumber = update.Identity.RevisionNumber
	    end If
	    	
	    strUpdateImpact = ""
	    If (IsNull(update.InstallationBehavior.Impact) or update.InstallationBehavior.Impact = "")  then 
	    	strUpdateImpact = "Unknown"
	    else
			strUpdateImpact = arrUpdateImpact(update.InstallationBehavior.Impact)
	    end If

		strDownloadSize = GetPrettyFileSize(update.MaxDownloadSize)
		
	    ' build an ID value which is more unique than the GUID
	    ' a GUID can be shared between two updates which are different binaries
	    ' example: silverlight update for end users and silverlight update
	    ' for developers share the same GUID value
	    
		strTrulyUniqueID = ""
		If ( Not IsNull(update.Identity.UpdateID) ) Then
			strTrulyUniqueID = hasher.GetMD5(update.Identity.UpdateID&urls&filenames&strDownloadSize)
		Else
			strTrulyUniqueID = "None"
		End If
		
		If Trim(severity) = "" Then severity = "None"
		If Trim(strBulletins) = "" Then strBulletins = "None"
		If Trim(creationDate) = "" Then creationDate = "None"
		If Trim(urls) = "" Then urls = "None"
		If Trim(filenames) = "" Then filenames = "None"
		If Trim(strInstalled) = "" Then strInstalled = "None"
		If Trim(strUpdateId) = "" Then strUpdateId = "None"
		If Trim(strCveIDs) = "" Then strCVEIDs = "None"
		If Trim(strResultsCVEIDs) = "" Then strResultsCVEIDs = "None"
		If strKBArticles = "" Then strKBArticles = "None"	
		
		Dim strUpdateLongDetails
		strUpdateLongDetails = update.Title & strSep & _
			   	severity & strSep & _
			   	strBulletins & strSep & _
			   	creationDate & strSep & _
			   	urls  & strSep & _
			   	filenames & strSep & _
			   	strInstalled & strSep & _
			   	strUpdateId & strSep & _
			   	strDownloadSize & strSep & _
			   	strKBArticles & strSep & _
			   	strResultsCVEIDs & strSep & _
			   	strTrulyUniqueID & strSep & _
			   	strRebootBehavior & strSep & _
			   	strRevisionNumber & strSep & _				   				   		   	
			   	strUpdateImpact
		If Not dictUpdateSearchResultIDtoDetails.Exists(I) Then
			dictUpdateSearchResultIDtoDetails.Add I, strUpdateLongDetails
		End If
		
		' do not write superseded updates to patchresults
		' unless the update is installed and also not superseded
		' this prevents the deployment of (most) superseded updates but allows
		' installed updates to be reported on even if they are superseded
		If bValidUpdate Or ( update.IsInstalled And Not dictSupersededIDs.Exists(strUpdateID) ) Then
		   	objTextFile.WriteLine(UnicodeToAscii(strUpdateLongDetails))
			   	' ResultsCveIDs can be None or the actual ID list depending on the date parameter passed in
			   	' strCVEIDs Is always the actual result
		Else
			' tLog.log "Skipped superseded update " & update.Title			   	
		End If	

		If IsDate(dtmOSInstallDate) Then
		'strFirstNeededDate must be either OS install Date Or update publish Time
		' whichever is later.		
			If dtmOSInstallDate > update.LastDeploymentChangeTime Then
				strFirstNeededDate = dtmOSInstallDate
			Else
				strFirstNeededDate = update.LastDeploymentChangeTime
			End If

		' strSupersededDate is the OS install date or the date at which the current update
		' was superseded by another
			If dictSupersededIDs.Exists(strUpdateID) Then
				' an update may be superseded by more than one update, must pick the earliest date
				Dim dtmEarliestSupersededDate
				arrSupersedersList = Split(dictSupersededIDs.Item(strUpdateID),"|-|")
				dtmEarliestSupersededDate = Date()
				For Each strSupersederInfo In arrSupersedersList
					strSupersededDate = Split(strSupersederInfo,strSep)(2)
					On Error Resume Next
					dtmSupersededDate = CDate(strSupersededDate)
					If Err.Number <> 0 Then
						tLog.log "Error converting superseded date for an entry"
					End If
					On Error Goto 0
					If dtmSupersededDate < dtmEarliestSupersededDate Then
						dtmEarliestSupersededDate = dtmSupersededDate
					End If
					' earliest date
				Next
					On Error Resume Next
					dtmSupersededDate = CDate(dtmEarliestSupersededDate)
					If Err.Number <> 0 Then
						tLog.log "Error converting superseded date for " & update.Title &": - " & Err.Description
					End If
					On Error Goto 0
					
					
				If dtmOSInstallDate > dtmSupersededDate Then
					strSupersededDate = dtmOSInstallDate
				Else
					strSupersededDate = dtmSupersededDate
				End If
			End If
	
		End If
	
		' Write history for update if the GUID exists in the patchresults
		' History Columns are (times are UTC)
		' GUID|Operation|Result|Error|Install Date|Publish Date|Title|FirstNeededDate|Severity|MS ID|KB|CVE ID|Size
		
		strHistoryOutputLine = update.LastDeploymentChangeTime&strSep&update.Title&strSep _
					&strFirstNeededDate&strSep&severity&strSep&strBulletins&strSep&strKBArticles&strSep&strCVEIDs _
					&strSep&strDownloadSize&strSep&strTrulyUniqueID
					
		Dim arrHistoryEventsList,strHistoryEvent,strEventGUID
		If dictLiveHistory.Exists(strUpdateID) Then
		' only write updates which are known to the cab file at write time				
			arrHistoryEventsList = Split(dictLiveHistory.Item(strUpdateID),"|-|")
			For Each strHistoryEvent In arrHistoryEventsList
				strEventGUID = Split(strHistoryEvent,"|")(0)
				If bBadHistoryLines Then
					'tLog.log "bad history lines - writing " & strHistoryEvent&strSep&strHistoryOutputLine
					objHistoryTextFile.WriteLine(UnicodeToAscii(strHistoryEvent&strSep&strHistoryOutputLine))
				Else
					' only write completely unique lines, preserving all history			
					If Not dictHistory.Exists(strHistoryEvent&strSep&strHistoryOutputLine) Then 
						objHistoryTextFile.WriteLine(UnicodeToAscii(strHistoryEvent&strSep&strHistoryOutputLine))
						'tLog.log "history output line " & strHistoryEvent&strSep&strHistoryOutputLine
					Else
						'tLog.log "history output line already existed"
					End If
				End If
			Next
		End If
		

'####### Write first needed and supersedence information to history for updates which are in the cab file
		
		' if it's not already noted in the history file, we must write that we have seen it for the first time
		' and it has not alreayd been written.  Only if it's a valid update (appropriately superseded)
		If bValidUpdate And Not dictAlreadyFirstNeededHistoryEntries.Exists(strUpdateID) Then
			'tLog.log "writing " & strUpdateID&strSep&"FirstNeeded"&strSep _
			'	&"FirstNeeded"&strSep&"FirstNeeded"&strSep&strFirstNeededDate&strSep&strHistoryOutputLine
			objHistoryTextFile.WriteLine(UnicodeToAscii(strUpdateID&strSep&"FirstNeeded"&strSep _
				&"FirstNeeded"&strSep&"FirstNeeded"&strSep&strFirstNeededDate&strSep&strHistoryOutputLine))
		End If

		' same for superseded updates.
		' if the update is superseded in any way (by a valid service pack or not)
		' and if the update is not already written in the history file
		' note when it was first superseded
		If Not dictAlreadySupersededHistoryEntries.Exists(strUpdateID) And dictSupersededIDs.Exists(strUpdateID) Then
			'tLog.log "writing " & strUpdateID&strSep&"Superseded"&strSep _
			'	&"Superseded"&strSep&"Superseded"&strSep&strSupersededDate&strSep&strHistoryOutputLine
			objHistoryTextFile.WriteLine(UnicodeToAscii(strUpdateID&strSep&"Superseded"&strSep _
				&"Superseded"&strSep&"Superseded"&strSep&strSupersededDate&strSep&strHistoryOutputLine))
		End If
	Next
					
	tLog.log "Number of updates not installed: " & numNotInstalled
	tLog.log "Number of updates installed: " & numInstalled
	
	objTextFile.Close
	objHistoryTextFile.Close
	
	set FiletoCopy = objFSO.GetFile(strResultsPath)
	FiletoCopy.Copy(strResultsReadablePath)
		
	set FiletoCopy = objFSO.GetFile(strHistoryTextFilePath)
	FiletoCopy.Copy(strHistoryTextReadableFilePath)
	
	If bPartialCVEs Then ' close the cve list file and make readable copy
		objPatchResultsAllCVETextFile.Close
		set FiletoCopy = objFSO.GetFile(strCVEPatchResultsFile)
		FiletoCopy.Copy(strCVEPatchResultsReadableFile)
	End If
End Function 'RunPatchScan

Function GetPrettyFileSize(strSize)
	Dim dblSize
	dblSize = CDbl(strSize)
	
	If dblSize > 1024*1024*1024 Then ''Should be GB
		strSize = CStr(Round(dblSize / 1024 / 1024 / 1024, 1)) & " GB"	
	ElseIf dblsize > 1024*1024 Then  ''Should be MB
		strSize = CStr(Round(dblSize / 1024 / 1024, 1)) & " MB"
	ElseIf dblSize > 1024 Then  ''Should be kB
		strSize = CStr(Round(dblSize / 1024)) & " KB"
	Else
		strSize = CStr(dblSize) & " B"	
	End If
	strSize = Replace(strSize,",",".")
	GetPrettyFileSize = strSize
End Function 'GetPrettyFileSize

Sub BuildKBToUninstallableDict(intColIndex, ByRef dictTitles, ByRef dictKBToUninstallable)
' build a dictionary of KB article IDs and True or False indicating their uninstallability

	' loop through installed products and patches and check for those which match an update title
	

End Sub 'BuildPatchGUIDToProductGUIDDictByColumnIndex

Sub BuildMSIPatchNamesDict(ByRef dictMSIPatchNames)

	Const MSIINSTALLCONTEXT_ALL = 7 

	Dim oMSI,iContext,objProducts,objProd
	Dim dictProducts,dictPatches
	Dim allProducts,product,strDisplayName,objPatches
	Dim objPatch
	
	Set dictProducts = CreateObject("Scripting.Dictionary")
	Set dictPatches = CreateObject("Scripting.Dictionary")
	
	Set oMsi = CreateObject("WindowsInstaller.Installer")
	iContext = MSIINSTALLCONTEXT_ALL
	
	Set allProducts = oMSI.ProductsEx("","",4)
	
	For Each product In allProducts
		Set objPatches = oMSI.PatchesEx(product.ProductCode,"",4,1)
		For Each objPatch In objPatches
			strDisplayName = objPatch.PatchProperty("DisplayName")
			If Not dictMSIPatchNames.Exists(strDisplayName) Then
				dictMSIPatchNames.Add strDisplayName,objPatch.PatchCode
			End If
		Next
	Next

End Sub 'BuildMSIPatchNamesDict

Function IsUninstallable(strKB,strTitle)
' Determines if an update can be uninstalled by Tanium

	Const HKLM = &h80000002	
	
	Dim objReg,strKeyPath,strUninstallStringFromReg
	Dim strUninstallKeyPath,strUninstallCommand,strOSMajorVersion,sngOSMajorVersion,strPackagesKeyPath
	Dim arrPackagesKeys,strPackagesKey,bUninstallableByTitle,strMSIPatchDisplayName
		
	' First determine if the update is an MSI based patch (Office update, for example)
	For Each strMSIPatchDisplayName In dictMSIPatchNames.Keys
		'tLog.log "checking MSI patch display name: "&strMSIPatchDisplayName&" against patch name: "&strTitle
		' the update's title is assumed to be shorter than the MSI patch update title
		If InStr(LCase(strMSIPatchDisplayName),LCase(strTitle)) = 1 And Len(strMSIPatchDisplayName) > 2 Then
			'tLog.log "windows update " & strTitle & " - KB article " & strKB _
			'	& " - appears to be the same as MSI patch update title " & strMSIPatchDisplayName _
			'	& ". Ensure update is marked as uninstallable"
			IsUninstallable = True
		ElseIf InStr(LCase(strMSIPatchDisplayName),LCase(strTitle)) And Len(strMSIPatchDisplayName) > 2 Then
			'tLog.log "!!!!!"&strTitle & " appeared somewhere strange in " & strMSIPatchDisplayName
		End If
	Next
	
	strOSMajorVersion = GetOSMajorVersion
	
	strUninstallStringFromReg = ""
	strUninstallKeyPath = "Software\Microsoft\Windows\CurrentVersion\Uninstall\"&strKB
	strPackagesKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages"
	Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")	

	If Not IsNumeric(strOSMajorVersion) Then
		'tLog.log "Uninstall function cannot determine OS version"
		IsUninstallable = False
		Exit Function
	Else
		sngOSMajorVersion = CSng(strOSMajorVersion)
	End If
	
	If sngOSMajorVersion < 5.0 And sngOSMajorVersion >= 4.0 Then
		'tLog.log "Windows NT 4 uninstall not supported" ' WSH won't execute this script!
		IsUninstallable = False
	End If
	
	If sngOSMajorVersion >= 5 And sngOSMajorVersion < 6.0 Then
		'tLog.log "pre-vista OS"
		' Pre-vista OS
		' Pull uninstall key for pre-vista OS's
		If RegKeyExists(objReg,HKLM,strUninstallKeyPath) Then
			On Error Resume Next
			objReg.GetStringValue HKLM,strUninstallKeyPath,"UninstallString",strUninstallStringFromReg
			If Err.Number <> 0 Or strUninstallStringFromReg = "" Or IsNull(strUninstallStringFromReg) Then
				'tLog.log "Unexpected Error retrieving uninstall registry key"
				On Error Goto 0
				IsUninstallable = False				
				Exit Function
			Else
				IsUninstallable = True
				Exit Function
			End If
			On Error Goto 0
		Else
			'tLog.log "Cannot obtain uninstall string from registry"
			IsUninstallable = False		
			Exit Function
		End If
	End If 'Pre-Vista
	
	If sngOSMajorVersion >= 6.0 Then
		'Windows Vista or Server 2008, Windows 7 and Server 2008 R2, Windows 8 or 2012
		If RegKeyExists(objReg,HKLM,strPackagesKeyPath) Then
		' Looking for a key that looks like
		' Package_for_KB960803~31bf3856ad364e35~amd64~~6.0.1.0
			objReg.EnumKey HKLM,strPackagesKeyPath,arrPackagesKeys
			If IsArray(arrPackagesKeys) Then
				For Each strPackagesKey In arrPackagesKeys
					If InStr(strPackagesKey,"Package_for_"&strKB) = 1 Then 'starts with
						IsUninstallable = True ' assume one hit means it can be uninstalled				
					End If
				Next
			Else
				'tLog.log "Unexpected Error enumerating Packages keys, cannot continue"
				IsUninstallable = False				
				Exit Function
			End If
		Else
			'tLog.log "Unexpected Error locating Packages key, cannot continue"
			IsUninstallable = False
			Exit Function
		End If
	End If ' Post-XP OS's

End Function 'IsUninstallable

Function GetOSMajorVersion
' Returns the OS Major Version
' A different way to examine the OS instead of by name

	Dim objWMIService,colItems,objItem
	Dim strVersion,arrVersion
	
	strVersion = "Unknown"

	Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
	Set colItems = GetObject("WinMgmts:root/cimv2").ExecQuery("select Version from win32_operatingsystem")    
	For Each objItem In colItems
		strVersion = objItem.Version ' like 6.2.9200
		arrVersion = Split(strVersion,".")
		If UBound(arrVersion) >= 1 Then
			strVersion = arrVersion(0)&"."&arrVersion(1)
		End If
	Next
	GetOSMajorVersion = strVersion
End Function 'GetOSMajorVersion

Function GetTaniumRegistryPath
'GetTaniumRegistryPath works in x64 or x32
'looks for a valid Path value

	Dim objShell
	Dim keyNativePath, keyWoWPath, strPath, strFoundTaniumRegistryPath
	  
    Set objShell = CreateObject("WScript.Shell")
    
	keyNativePath = "Software\Tanium\Tanium Client"
	keyWoWPath = "Software\Wow6432Node\Tanium\Tanium Client"
    
    ' first check the Software key (valid for 32-bit machines, or 64-bit machines in 32-bit mode)
    On Error Resume Next
    strPath = objShell.RegRead("HKLM\"&keyNativePath&"\Path")
    On Error Goto 0
	strFoundTaniumRegistryPath = keyNativePath
 
  	If strPath = "" Then
  		' Could not find 32-bit mode path, checking Wow6432Node
  		On Error Resume Next
  		strPath = objShell.RegRead("HKLM\"&keyWoWPath&"\Path")
  		On Error Goto 0
		strFoundTaniumRegistryPath = keyWoWPath
  	End If
  	
  	If Not strPath = "" Then
  		GetTaniumRegistryPath = strFoundTaniumRegistryPath
  	Else
  		GetTaniumRegistryPath = False
  		tLog.log "Error: Cannot locate Tanium Registry Path"
  	End If
End Function 'GetTaniumRegistryPath

Function GetTaniumDir(strSubDir)
'GetTaniumDir with GeneratePath, works in x64 or x32
'looks for a valid Path value
	
	Dim objShell
	Dim keyNativePath, keyWoWPath, strPath
	  
    Set objShell = CreateObject("WScript.Shell")
    
	keyNativePath = "HKLM\Software\Tanium\Tanium Client"
	keyWoWPath = "HKLM\Software\Wow6432Node\Tanium\Tanium Client"
    
    ' first check the Software key (valid for 32-bit machines, or 64-bit machines in 32-bit mode)
    On Error Resume Next
    strPath = objShell.RegRead(keyNativePath&"\Path")
    On Error Goto 0
 
  	If strPath = "" Then
  		' Could not find 32-bit mode path, checking Wow6432Node
  		On Error Resume Next
  		strPath = objShell.RegRead(keyWoWPath&"\Path")
  		On Error Goto 0
  	End If
  	
  	If Not strPath = "" Then
		If strSubDir <> "" Then
			strSubDir = "\" & strSubDir
		End If	
	
		Dim fso
		Set fso = WScript.CreateObject("Scripting.Filesystemobject")
		If fso.FolderExists(strPath) Then
			If Not fso.FolderExists(strPath & strSubDir) Then
				''Need to loop through strSubDir and create all sub directories
				GeneratePath strPath & strSubDir, fso
			End If
			GetTaniumDir = strPath & strSubDir & "\"
		Else
			' Specified Path doesn't exist on the filesystem
			tLog.log "Error: " & strPath & " does not exist on the filesystem"
			GetTaniumDir = False
		End If
	Else
		tLog.log "Error: Cannot find Tanium Client path in Registry"
		GetTaniumDir = False
	End If
End Function 'GetTaniumDir

Function GeneratePath(pFolderPath, fso)
	GeneratePath = False

	If Not fso.FolderExists(pFolderPath) Then
		If GeneratePath(fso.GetParentFolderName(pFolderPath), fso) Then
			GeneratePath = True
			Call fso.CreateFolder(pFolderPath)
		End If
	Else
		GeneratePath = True
	End If
End Function 'GeneratePath

Function RegKeyExists(objRegistry, sHive, sRegKey)
	Dim aValueNames, aValueTypes
	If objRegistry.EnumValues(sHive, sRegKey, aValueNames, aValueTypes) = 0 Then
		RegKeyExists = True
	Else
		RegKeyExists = False
	End If
End Function 'RegKeyExists

Function GetTaniumLocale
'' This function will retrieve the locale value
' previously set which governs Tanium content that
' is locale sensitive.

	Dim objWshShell
	Dim intLocaleID
	
	Set objWshShell = CreateObject("WScript.Shell")
	On Error Resume Next
	intLocaleID = objWshShell.RegRead("HKLM\Software\Tanium\Tanium Client\LocaleID")
	If Err.Number <> 0 Then
		intLocaleID = objWshShell.RegRead("HKLM\Software\Wow6432Node\Tanium\Tanium Client\LocaleID")
	End If
	On Error Goto 0
	If intLocaleID = "" Then
		GetTaniumLocale = 1033 ' default to us/English
	Else
		GetTaniumLocale = intLocaleID
	End If

	' Cleanup
	Set objWshShell = Nothing

End Function 'GetTaniumLocale

Function GetInstallHistory(intLocale)
' scans history file and current state and 
' returns dictionary object with history data in it
' History data is the patch

	Set dictLiveHistory = CreateObject("Scripting.Dictionary") ' Global variable

	Dim objSession, objSearcher, colHistory
	Dim arrResultCodes,arrOperations
	Dim intHistoryCount, objEntry, strResult
	Dim strOperation,strError,strDate,strGUID,strSep
	Dim strTitle,strEventSep,strEventEntry
	
	strSep = "|" ' delimiter for use within single events
	strEventSep = "|-|" ' delimiter to indicate a new event for the GUID
	
	arrResultCodes = Array( "Not Started", "In Progress", "Succeeded", _
	                        "Succeeded With Errors", "Failed", "Aborted" )
	arrOperations  = Array( "", "Installation", "Uninstallation" )
	
	Set objSession = CreateObject("Microsoft.Update.Session")
	' Use locale settings
	objSession.UserLocale = intLocale
	' note - Title is not localized.
	
	Set objSearcher = objSession.CreateUpdateSearcher

	intHistoryCount = objSearcher.GetTotalHistoryCount
	
	'// Get WU history data:
	
	If intHistoryCount > 0 Then
		Set colHistory = objSearcher.QueryHistory(0,intHistoryCount+6)
		
		For Each objEntry In colHistory
			' dictionary will look like
			' GUID,Operation|Result|ErrorCode|Date
			On Error Resume Next
			strOperation = arrOperations(objEntry.Operation)

			strResult = arrResultCodes(objEntry.ResultCode)
			If Err.Number <> 0 Then
				Err.Clear
				strResult = "Unknown"
			End If
			
			' Error (HResult) is always a number
			strError = CLng(objEntry.HResult)
			If Err.Number <> 0 Then
				Err.Clear
				strError = "Unknown"
			End If
			
			' If no error, show "None"
			If strError = "0" Then strError = "None"

			' Try to pull a friendly error name

			If dictErrorCodes.Exists("0x"&CStr(Hex(strError))) Then 
				strError = dictErrorCodes.Item("0x"&CStr(Hex(strError)))
			End If
			
			strDate = objEntry.Date 
			If Err.Number <> 0 Then
				Err.Clear
				strDate = "Unknown"
			End If
			strGUID = objEntry.UpdateIdentity.UpdateID

			If Err.Number <> 0 Then
				Err.Clear
				strGUID = "Unknown"
			End If
			strTitle = objEntry.Title
			If Err.Number <> 0 Then
				Err.Clear
				strTitle = "Unknown"
			End If
			
								
			On Error Goto 0
			
			If Not ( dictLiveHistory.Exists(strGUID) Or strGUID = "Unknown" ) Then
				' tLog.log "Adding "strOperation&strSep&strResult&strSep _
				'	&strError&strSep&strDate& " for "&strGUID
				dictLiveHistory.Add strGUID,strGUID&strSep&strOperation&strSep _
					&strResult&strSep&strError&strSep&strDate
			Else
				' if the GUID entry is already there, read in the already existing event information and add second delimiter
				' to the item section, adding the new event information on the end.
				strEventEntry = dictLiveHistory.Item(strGUID)
				' tLog.log "concat adding to " & strGUID & " : " & strEventEntry&"|-|"&strGUID&strSep&strOperation&strSep&strResult&strSep&strError&strSep&strDate
				dictLiveHistory.Item(strGUID) = strEventEntry&strEventSep&strGUID&strSep&strOperation&strSep&strResult&strSep&strError&strSep&strDate	
			End If
		Next
	End If

	'Cleanup
	Set colHistory = Nothing
	Set objSession = Nothing
	Set objSearcher = Nothing

	Set GetInstallHistory = dictLiveHistory
	Set dictLiveHistory = Nothing

End Function 'GetInstallHistory

Function GetCVEDict(strCVEFilePath)
' produces a dictionary entry with KB article as key and CVE data as entry
	
	Const FORREADING = 1
	
	Dim objFSO,dictCVEs,objCVETextFile
	
	Dim strLine,arrBulletinLine,arrCVELine,strBulletinID,arrLine,bBadLine
	Dim strCVE
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set dictCVEs = CreateObject("Scripting.Dictionary")
	If Not objFSO.FileExists(strCVEFilePath) Then
		' Return empty dictionary
		Set GetCVEDict = dictCVEs
		Exit Function
	Else
		Set objCVETextFile = objFSO.OpenTextFile(strCVEFilePath,FORREADING,True)
		Do Until objCVETextFile.AtEndOfStream = True
			bBadLine = False 'Continue statement for loop control not in vbscript
			strLine = Trim(objCVETextFile.ReadLine)
			If Not InStr(strLine,"MS") = 1 Then
				If Not InStr(strLine,"Q") = 1 Then
					' There is always unmappable Q169461,CVE-1999-0275 lines
					tLog.log "Warning: Bad CVE Line: " & strLine
				End If
				bBadLine = True
			End If
			arrLine = Split(strLine,",",2) ' two element array, MS14-035|CVE1,CVE2,CVE3
			If Not UBound(arrLine) > 0 Then
				tLog.log "Warning: Bad CVE Line: " & strLine
				bBadLine = True
			End If
			strBulletinID = arrLine(0)
			strCVE = arrLine(1)
			strCVE = Replace(strCVE,","," ") ' Patch scan output is space delimited
			If ( Not bBadLine ) Or dictCVEs.Exists(strBulletinID) Then
				' tLog.log "Adding " & strBulletinID&","&strCVE
				dictCVEs.Add Trim(strBulletinID),Trim(strCVE)
			End If
		Loop
		objCVETextFile.Close()
	End If
	
	Set GetCVEDict = dictCVEs
		
End Function 'GetCVEDict

Function GetOSInstallDate
' retrieves OS install date for use in update needed text file

	Dim dtmConvertedDate,objWMIService,colOperatingSystems
	Dim objOperatingSystem,dtmInstallDate
	Set dtmConvertedDate = CreateObject("WbemScripting.SWbemDateTime")
	
	Set objWMIService = GetObject("winmgmts:" _
	    & "{impersonationLevel=impersonate}!\\.\root\cimv2")
	
	Set colOperatingSystems = objWMIService.ExecQuery _
	    ("Select InstallDate from Win32_OperatingSystem")
	
	For Each objOperatingSystem in colOperatingSystems
		
		dtmConvertedDate.Value = objOperatingSystem.InstallDate
	    dtmInstallDate = dtmConvertedDate.GetVarDate
	    GetOSInstallDate = dtmInstallDate
	Next
	'Cleanup
	Set dtmConvertedDate = Nothing
	Set colOperatingSystems = Nothing
	Set objWMIService = Nothing
End Function 'GetOSInstallDate

Function RunOverride
' This funciton will look for a file of the same name in a subdirectory
' called Override.  If it exists, it will run that instead, passing all arguments
' to it.

	Dim objFSO,objArgs,objShell,objExec
	Dim strFileDir,strFileName,strOriginalArgs,strArg,strLaunchCommand
	
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strFileDir = Replace(WScript.ScriptFullName,WScript.ScriptName,"")
	strFileName = WScript.ScriptName
	
	
	If objFSO.FileExists(strFileDir&"override\"&strFileName) Then
		tLog.log "Relaunching"
		strOriginalArgs = ""
		Set objArgs = WScript.Arguments
		
		For Each strArg in objArgs
		    strOriginalArgs = strOriginalArgs & " " & strArg
		Next
		' after we're done, we have an unnecessary space in front of strOriginalArgs
		strOriginalArgs = LTrim(strOriginalArgs)
	
		strLaunchCommand = Chr(34) & strFileDir&"override\"&strFileName & Chr(34) & " " & strOriginalArgs
		' tLog.log "Script full path is: " & WScript.ScriptFullName
		
		Set objShell = CreateObject("WScript.Shell")
		Set objExec = objShell.Exec(Chr(34)&WScript.FullName&Chr(34) & " " & strLaunchCommand)
		
		' skipping the two lines and space after that look like
		' Microsoft (R) Windows Script Host Version
		' Copyright (C) Microsoft Corporation
		'
		objExec.StdOut.SkipLine
		objExec.StdOut.SkipLine
		objExec.StdOut.SkipLine
	
		' catch the stdout of the relaunched script
		tLog.log objExec.StdOut.ReadAll()
		
		' prevent endless loop
		WScript.Quit
		' Remember to call this function only at the very top, before x64fix
		
		' Cleanup
		Set objArgs = Nothing
		Set objExec = Nothing
		Set objShell = Nothing
	End If
	
End Function 'RunOverride

Function RunFilesInDir(strSubDirArg)
' This function will run all vbs files in a directory
' in alphabetical order
' the directory must be called <script name-vbs>\strSubDirArg
' so for run-patch-scan.vbs it must be run-patch-scan\strSubDirArg

	Dim objFSO,objShell,objFolder,objExec
	Dim objFile,strFileDir,strSubDir,intResult
	Dim strFileName,strExtension,strFolderName,strTargetExtension
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	
	strFileDir = Replace(WScript.ScriptFullName,WScript.ScriptName,"")

	strExtension = objFSO.GetExtensionName(WScript.ScriptFullName)
	strFolderName = Replace(WScript.ScriptName,"."&strExtension,"")
	strSubDir = strFileDir&strFolderName&"\"&strSubDirArg
	
	If objFSO.FolderExists(strSubDir) Then ' Run each file in the directory
		tLog.log "Found subdirectory " & strSubDirArg
		Set objFolder = objFSO.GetFolder(strSubDir)
		Set objShell = CreateObject("WScript.Shell")
		For Each objFile In objFolder.Files
			strTargetExtension = Right(objFile.Name,3)
			If strTargetExtension = "vbs" Then
				tLog.log "Running " & objFile.Path
				Set objExec = objShell.Exec(Chr(34)&WScript.FullName&Chr(34) & "//T:1800 " & Chr(34)&objFile.Path&Chr(34))
			
				' skipping the two lines and space after that look like
				' Microsoft (R) Windows Script Host Version
				' Copyright (C) Microsoft Corporation
				'
				objExec.StdOut.SkipLine
				objExec.StdOut.SkipLine
				objExec.StdOut.SkipLine
			
				' catch the stdout of the relaunched script
				tLog.log objExec.StdOut.ReadAll()
			    Do While objExec.Status = 0
					WScript.Sleep 100
				Loop
				intResult = objExec.ExitCode
				If intResult <> 0 Then
					tLog.log "Non-Zero exit code for file " & objFile.Path & ", Quitting"
					WScript.Quit(-1)
				End If
			End If 'VBS only
		Next
	End If

	
	'Cleanup
	Set objFSO = Nothing
	Set objShell = Nothing
	Set objExec = Nothing
	Set objFolder = Nothing
	
End Function 'RunFilesInDir

Function CommandCount(strExecutable, strCommandLineMatch)
' This function will return a count of the number of exectuable / command line
' instances running.  if the executable
' passed in is running with a command line that matches part of what
' the CommandLineMatch parameter, it will be added to the count.  If the count is greater
' than one, we can assume this process is already running, so don't run the scan.

	Const HKLM = &h80000002
	
	Dim objWMIService,colItems
	Dim objItem,strCmd,intRunningCount

	intRunningCount = 0
	On Error Resume Next
	
	SetLocale(1033) ' Uses Date Math which requires us/english to work correctly
	
	Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
	
	Set colItems = objWMIService.ExecQuery("Select CommandLine from Win32_Process where Name = '"&strExecutable&"'",,48)
	For Each objItem in colItems
		strCmd = objItem.CommandLine
		If InStr(strCmd,strCommandLineMatch) > 0 Then
			intRunningCount = intRunningCount + 1
		End If
	Next
	On Error Goto 0

	CommandCount = intRunningCount

End Function 'CommandCount

Function ForceRemoveMicrosoftUpdate
	Dim updateServiceManager,service,strServiceID,objSession
	Set objSession = CreateObject("Microsoft.Update.Session")
	' Microsoft Update Service ID
	strServiceID = "7971f918-a847-4430-9279-4a52d1efe18d"
	objSession.WebProxy.AutoDetect = True

	tLog.log "==== Removing Microsoft Update ===="
	Set UpdateServiceManager = CreateObject("Microsoft.Update.ServiceManager")
	On Error Resume Next
	UpdateServiceManager.SetOption "AllowWarningUI", False
	UpdateServiceManager.UnregisterServiceWithAU(strServiceID)
	UpdateServiceManager.RemoveService(strServiceID)
	On Error Goto 0
	bResult = False
	' validate removal
	For each service in UpdateServiceManager.Services
		if Trim(Lcase(service.ServiceID)) = strServiceID Then
			tLog.log "Error: Removal Failed!"
			bResult = False
		End If
	Next
	
	ForceRemoveMicrosoftUpdate = bResult
End Function 'ForceRemoveMicrosoftUpdate

Function GetTZBias
' This function returns the number of minutes
' (positive or negative) to add to current time to get UTC
' considers daylight savings

	Dim objLocalTimeZone, intTZBiasInMinutes


	For Each objLocalTimeZone in GetObject("winmgmts:").InstancesOf("Win32_ComputerSystem")
		intTZBiasInMinutes = objLocalTimeZone.CurrentTimeZone
	Next

	GetTZBias = intTZBiasInMinutes
		
End Function 'GetTZBias


Sub EnsureRunsOneCopy

	' Do not run this more than one time on any host
	' This is useful if the job is done via start /B for any reason (like random wait time)
	' or to prevent any other situation where multiple scans could run at once
	Dim intCommandCount,intCommandCountMax
	intCommandCount = CommandCount("cscript.exe","run-patch-scan.vbs")
	
	' There will always be one copy of this script running
	' where we want to stop is if there are two running
	' which would be the one doing the work and then another checking to see
	' if it scan start
	' must take into account the double launch with x64Fix when run in 32-bit mode
	' on a 64-bit system
	
	If Is64 Then 
		intCommandCountMax = 3
	Else
		intCommandCountMax = 2
	End If
	
	If intCommandCount < intCommandCountMax Then
		tLog.log "Patch scan not running, continuing"
	Else
		tLog.log "Patch scan currently running, won't scan - Quitting"
		WScript.Quit
	End If
	
End Sub 'EnsureRunsOneCopy


Function WUAVersionTooLow(strNeededVersion)
	' Return True or false if version is too low. Pass in desired version like
	' "6.1.0022.4"
	Dim i, objAgentInfo, intMajorVersion
	Dim arrNeededVersion
	Dim strVersion, arrVersion, intVersionPiece, bOldVersion
	
	WUAVersionTooLow = True ' Assume bad until proven otherwise
	'adjust as required version changes
	arrNeededVersion = Split(strNeededVersion,".")
	If UBound(arrNeededVersion) < 3 Then
		WScript.Echo "Version passed to WUA version check is malformed"
		Exit Function
	End If

	On Error Resume Next 
	Set objAgentInfo = CreateObject("Microsoft.Update.AgentInfo")	
	strVersion = objAgentInfo.GetInfo("ProductVersionString") 
	If Err.Number <> 0 Then
		WScript.Echo "Could not reliably determine Windows Update Agent version"
		Exit Function
	End If
	On Error Goto 0
	arrVersion = Split(strVersion,".")
	' loop through each part
	' if any individual part is less than its corresponding required part
	bOldVersion = False
	For i = 0 To UBound(arrVersion)
		If CInt(arrVersion(i)) > CInt(arrNeededVersion(i)) Then
			bOldVersion = False
			Exit For ' No further checking necessary, it's newer
		ElseIf CInt(arrVersion(i)) < CInt(arrNeededVersion(i)) Then
			bOldVersion = True
			Exit For ' No further checking necessary, it's out of date
		End If
		'For Loop will only continue if the first set of numbers were equal
	Next
	If bOldVersion Then
		Exit Function 'still false
	End If
	
	WUAVersionTooLow = False
	
End Function 'WUAVersionTooLow


Sub FillAPIHistoryDictWithTimings(dictLiveHistory)
	dtmTempTime = Timer()
	If TryFromDict(dictPConfig,"ShowTimings",False) Then tLog.log "Update History Gather from WU API: Started"
	Set dictLiveHistory = GetInstallHistory(intLocale)
	If TryFromDict(dictPConfig,"ShowTimings",False) Then tLog.log "Update History Gather from WU API: End - Took " & Abs(Timer() - dtmTempTime) & " seconds"
End Sub 'FillAPIHistoryDict


Sub FillCustomCabsDict(ByRef dictCabs)
	Dim strCustomCabSupportVal,strToolsPath
	tContentReg.ErrorClear
	On Error Resume Next
	tContentReg.ValueName = "CustomCabSupport"
	strCustomCabSupportVal = LCase(tContentReg.Read)
	If Err.Number <> 0 Then
		WScript.Echo "Custom Cab Support is not enabled"
		Err.Clear
		tContentReg.ErrorClear
	End If
	On Error Goto 0
	If strCustomCabSupportVal = "true" Or strCustomCabSupportVal = "yes" Then
		tLog.log "Custom Cab Scan support enabled in registry"
		Set dictCabs = GetAllCabs
	Else
		Set dictCabs = CreateObject("Scripting.Dictionary")
		' scan only the default cab
		strToolsPath = GetTaniumDir("Tools")
		dictCabs.Add strToolsPath&"wsusscn2.cab","wsusscn2.cab"
	End If
End Sub 'FillCustomCabsDict


Function SupersededOnlyByValidServicePack(update,dictUpdatesSupersededByValidServicePacks,dictSupersededIDs)
' this function returns true if the update is not superseded
' unless it's superseded by a valid service pack
' The best answer here is True which means we will process the update
' This function is no longer used in the main branch, replaced by IsValidUpdate

	SupersededOnlyByValidServicePack = True ' assume updates are valid

	' first check if the update itself is superseded.  If so, do not process the update.
	If dictSupersededIDs.Exists(update.Identity.UpdateID) Then
		' tLog.log update.Title & " is superseded, will not process"
		SupersededOnlyByValidServicePack = False
	End If
	
	' now see if it is in the list of updates superseded by service packs.
	' if it is superseded by a service pack, we will process the update.
	If dictUpdatesSupersededByValidServicePacks.Exists(update.Identity.UpdateID) Then
		SupersededOnlyByValidServicePack = True
	End If

End Function 'SupersededOnlyByValidServicePack

Function IsValidUpdate(update,ByRef dictSupersededIDs,ByRef dictSupersededByAnyNonServicePack,ByRef dictSupersededByOnlyServicePacks,ByRef dictServicePacks,strKBArticles )

'Determines if the update should be displayed
' If not superseded, it's valid
' If the update was superseded within the last superseded updates published days old threshold (/ShowSupersededUpdatesPublishedDaysOld:45), it's valid
' If the update is in the list of updates in the NeverSupersedeList registry key (comma separated KBXXXX,KBXXXX), it's valid
' If superseded by only service packs, it's valid
' if superseded by any non-sp update, it's not valid

	Dim strID,bValid,intSupersededUpdatesDaysOldThreshold
	strID = update.Identity.UpdateID

	bValid = False 'assume update is not shown
	
	' tLog.log "Validity Checking for " & update.Title
	' First, if the update is not superseded at all, it is valid with no further checks necessary
	If Not dictSupersededIDs.Exists(strID) Then
		'tLog.log update.Title & " - Update not superseded, will process"
		IsValidUpdate = True
		Exit Function
	End If

	intSupersededUpdatesDaysOldThreshold = TryFromDict(dictPConfig,"ShowSupersededUpdatesPublishedDaysOld",0)
	
	' If the update was superseded within the last superseded updates published days old threshold (/ShowSupersededUpdatesPublishedDaysOld:45), it's valid
	If intSupersededUpdatesDaysOldThreshold <> 0 Then
		Dim strPublishDate,intDaysOld
		strPublishDate = update.LastDeploymentChangeTime
		intDaysOld = DateDiff("d",strPublishDate,Now())
		' tLog.log "last publish date is " & strPublishDate & " which was published " & intDaysOld & " days ago, threshold value of " & intSupersededUpdatesDaysOldThreshold
		
		If CInt(intDaysOld) < CInt(intSupersededUpdatesDaysOldThreshold) Then
			IsValidUpdate = True
			' tLog.log update.title & " - last publish date is " & strPublishDate & " which was " & intDaysOld & " days ago - less than threshold value of " & intSupersededUpdatesDaysOldThreshold
			Exit Function
		End If
	End If
	
	Dim strNeverSupersedeList
	strNeverSupersedeList = TryFromDict(dictPConfig,"NeverSupersedeList","")
	
	' if any of the update KB articles exist in the list of KB article IDs to never supersede, the update is valid
	Dim strKBToNeverSupersede,strThisUpdatesKBArticle
	For Each strThisUpdatesKBArticle In Split(strKBArticles,",")
		' tLog.log "Checking " & UCase(strThisUpdatesKBArticle) & " "
		For Each strKBToNeverSupersede In Split(strNeverSupersedeList,",")
			' tLog.log "Against " & UCase(strKBToNeverSupersede)
			If UCase(strThisUpdatesKBArticle) = UCase(strKBToNeverSupersede) Then
				tLog.log "Update " & strThisUpdatesKBArticle & " is in the never supersede list, and is always valid"
				IsValidUpdate = True
				Exit Function
			End If
		Next
	Next

	'' if the update is superseded by only service packs, show
	If dictSupersededByOnlyServicePacks.Exists(strID) Then
		'tLog.log update.Title & " -  Update is superseded only by service packs, will process"
		IsValidUpdate = True
		Exit Function
	End If
	
	'' if the update is superseded by any non-sp update, do not show
	If dictSupersededByAnyNonServicePack.Exists(strID) Then
		'tLog.log update.Title & " -  Update is superseded by at least one update which are non-service packs, will process"
		IsValidUpdate = False
		Exit Function
	End If
	
	'tLog.log update.Title & " did not match validity rules, and is not valid"
	
	IsValidUpdate = bValid

End Function 'IsValidUpdate

Function RandomSleep(intSleepTimeSeconds)
' sleeps for a random period of time, intSleepTime is in seconds
	Dim intWaitTime
	If intSleepTimeSeconds = 0 Then Exit Function
	intWaitTime = CLng(intSleepTimeSeconds) * 1000 ' convert to milliseconds
	' wait random interval between 0 and the max
	' assign random value to wait time max value
	intWaitTime = Int( ( intWaitTime + 1 ) * Rnd )
	tLog.log "Sleeping for " & intWaitTime & " milliseconds"
	WScript.Sleep(intWaitTime)
	tLog.log "Done sleeping, continuing ..."
End Function 'RandomSleep

Sub PrintSupersedenceInfo(ByRef dictSupersededIDs,ByRef dictUpdateTitles)
' prints supersedence map to screen
	
	Dim strSupersededID,arrSupersedersList,strSupersederInfo,strSupersederID
	tLog.log "-------------Supersedence Info-----------------"
	For Each strSupersededID In dictSupersededIDs.Keys
		tLog.log strSupersededID & " - " & dictUpdateTitles(strSupersededID) & " is superseded"
		arrSupersedersList = Split(dictSupersededIDs.Item(strSupersededID),"|-|")
		For Each strSupersederInfo In arrSupersedersList
			strSupersederID= Split(strSupersederInfo,"|")(0)
			tLog.log vbTab & " by " & strSupersederID & " - " & dictUpdateTitles(strSupersederID)
		Next
	Next
	tLog.log "-----------End Supersedence Info---------------"

End Sub 'PrintSupersedenceInfo

Function Is64 
	Dim objWMIService, colItems, objItem
	Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
	Set colItems = objWMIService.ExecQuery("Select SystemType from Win32_ComputerSystem")    
	For Each objItem In colItems
		If InStr(LCase(objItem.SystemType), "x64") > 0 Then
			Is64 = True
		Else
			Is64 = False
		End If
	Next
End Function


Private Function UnicodeToAscii(ByRef pStr)
	Dim x,conv,strOut
	For x = 1 To Len(pStr)
		conv = Mid(pstr,x,1)
		conv = Asc(conv)
		conv = Chr(conv)
		strOut = strOut & conv
	Next
	UnicodeToAscii = strOut
End Function 'UnicodeToAscii


Function CreateErrorCodesDict
	Dim dictErrorCodes : Set dictErrorCodes = CreateObject("Scripting.Dictionary")
	
	dictErrorCodes.Add "0xf0800","CBS_E_INTERNAL_ERROR"
	dictErrorCodes.Add "0xf0801","CBS_E_NOT_INITIALIZED"
	dictErrorCodes.Add "0xf0802","CBS_E_ALREADY_INITIALIZED"
	dictErrorCodes.Add "0xf0803","CBS_E_INVALID_PARAMETER"
	dictErrorCodes.Add "0xf0804","CBS_E_OPEN_FAILED"
	dictErrorCodes.Add "0xf0805","CBS_E_INVALID_PACKAGE"
	dictErrorCodes.Add "0xf0806","CBS_E_PENDING"
	dictErrorCodes.Add "0xf0807","CBS_E_NOT_INSTALLABLE"
	dictErrorCodes.Add "0xf0808","CBS_E_IMAGE_NOT_ACCESSIBLE"
	dictErrorCodes.Add "0xf0809","CBS_E_ARRAY_ELEMENT_MISSING"
	dictErrorCodes.Add "0xf080A","CBS_E_REESTABLISH_SESSION"
	dictErrorCodes.Add "0xf080B","CBS_E_PROPERTY_NOT_AVAILABLE"
	dictErrorCodes.Add "0xf080C","CBS_E_UNKNOWN_UPDATE"
	dictErrorCodes.Add "0xf080D","CBS_E_MANIFEST_INVALID_ITEM"
	dictErrorCodes.Add "0xf080E","CBS_E_MANIFEST_VALIDATION_DUPLICATE_ATTRIBUTES"
	dictErrorCodes.Add "0xf080F","CBS_E_MANIFEST_VALIDATION_DUPLICATE_ELEMENT"
	dictErrorCodes.Add "0xf0810","CBS_E_MANIFEST_VALIDATION_MISSING_REQUIRED_ATTRIBUTES"
	dictErrorCodes.Add "0xf0811","CBS_E_MANIFEST_VALIDATION_MISSING_REQUIRED_ELEMENTS"
	dictErrorCodes.Add "0xf0812","CBS_E_MANIFEST_VALIDATION_UPDATES_PARENT_MISSING"
	dictErrorCodes.Add "0xf0813","CBS_E_INVALID_INSTALL_STATE"
	dictErrorCodes.Add "0xf0814","CBS_E_INVALID_CONFIG_VALUE"
	dictErrorCodes.Add "0xf0815","CBS_E_INVALID_CARDINALITY"
	dictErrorCodes.Add "0xf0816","CBS_E_DPX_JOB_STATE_SAVED"
	dictErrorCodes.Add "0xf0817","CBS_E_PACKAGE_DELETED"
	dictErrorCodes.Add "0xf0818","CBS_E_IDENTITY_MISMATCH"
	dictErrorCodes.Add "0xf0819","CBS_E_DUPLICATE_UPDATENAME"
	dictErrorCodes.Add "0xf081A","CBS_E_INVALID_DRIVER_OPERATION_KEY"
	dictErrorCodes.Add "0xf081B","CBS_E_UNEXPECTED_PROCESSOR_ARCHITECTURE"
	dictErrorCodes.Add "0xf081C","CBS_E_EXCESSIVE_EVALUATION"
	dictErrorCodes.Add "0xf081D","CBS_E_CYCLE_EVALUATION"
	dictErrorCodes.Add "0xf081E","CBS_E_NOT_APPLICABLE"
	dictErrorCodes.Add "0xf081F","CBS_E_SOURCE_MISSING"
	dictErrorCodes.Add "0xf0820","CBS_E_CANCEL"
	dictErrorCodes.Add "0xf0821","CBS_E_ABORT"
	dictErrorCodes.Add "0xf0822","CBS_E_ILLEGAL_COMPONENT_UPDATE"
	dictErrorCodes.Add "0xf0823","CBS_E_NEW_SERVICING_STACK_REQUIRED"
	dictErrorCodes.Add "0xf0824","CBS_E_SOURCE_NOT_IN_LIST"
	dictErrorCodes.Add "0xf0825","CBS_E_CANNOT_UNINSTALL"
	dictErrorCodes.Add "0xf0826","CBS_E_PENDING_VICTIM"
	dictErrorCodes.Add "0xf0827","CBS_E_STACK_SHUTDOWN_REQUIRED"
	dictErrorCodes.Add "0xf0900","CBS_E_XML_PARSER_FAILURE"
	dictErrorCodes.Add "0xf0901","CBS_E_MANIFEST_VALIDATION_MULTIPLE_UPDATE_COMPONENT_ON_SAME_FAMILY_NOT_ALLOWED"
	dictErrorCodes.Add "0x240001","WU_S_SERVICE_STOP"
	dictErrorCodes.Add "0x00240002","WU_S_SELFUPDATE"
	dictErrorCodes.Add "0x00240003","WU_S_UPDATE_ERROR"
	dictErrorCodes.Add "0x00240004","WU_S_MARKED_FOR_DISCONNECT"
	dictErrorCodes.Add "0x00240005","WU_S_REBOOT_REQUIRED"
	dictErrorCodes.Add "0x00240006","WU_S_ALREADY_INSTALLED"
	dictErrorCodes.Add "0x00240007","WU_S_ALREADY_UNINSTALLED"
	dictErrorCodes.Add "0x00240008","WU_S_ALREADY_DOWNLOADED"
	dictErrorCodes.Add "0x80240001","WU_E_NO_SERVICE"
	dictErrorCodes.Add "0x80240002","WU_E_MAX_CAPACITY_REACHED"
	dictErrorCodes.Add "0x80240003","WU_E_UNKNOWN_ID"
	dictErrorCodes.Add "0x80240004","WU_E_NOT_INITIALIZED"
	dictErrorCodes.Add "0x80240005","WU_E_RANGEOVERLAP"
	dictErrorCodes.Add "0x80240006","WU_E_TOOMANYRANGES"
	dictErrorCodes.Add "0x80240007","WU_E_INVALIDINDEX"
	dictErrorCodes.Add "0x80240008","WU_E_ITEMNOTFOUND"
	dictErrorCodes.Add "0x80240009","WU_E_OPERATIONINPROGRESS"
	dictErrorCodes.Add "0x8024000A","WU_E_COULDNOTCANCEL"
	dictErrorCodes.Add "0x8024000B","WU_E_CALL_CANCELLED"
	dictErrorCodes.Add "0x8024000C","WU_E_NOOP"
	dictErrorCodes.Add "0x8024000D","WU_E_XML_MISSINGDATA"
	dictErrorCodes.Add "0x8024000E","WU_E_XML_INVALID"
	dictErrorCodes.Add "0x8024000F","WU_E_CYCLE_DETECTED"
	dictErrorCodes.Add "0x80240010","WU_E_TOO_DEEP_RELATION"
	dictErrorCodes.Add "0x80240011","WU_E_INVALID_RELATIONSHIP"
	dictErrorCodes.Add "0x80240012","WU_E_REG_VALUE_INVALID"
	dictErrorCodes.Add "0x80240013","WU_E_DUPLICATE_ITEM"
	dictErrorCodes.Add "0x80240016","WU_E_INSTALL_NOT_ALLOWED"
	dictErrorCodes.Add "0x80240017","WU_E_NOT_APPLICABLE"
	dictErrorCodes.Add "0x80240018","WU_E_NO_USERTOKEN"
	dictErrorCodes.Add "0x80240019","WU_E_EXCLUSIVE_INSTALL_CONFLICT"
	dictErrorCodes.Add "0x8024001A","WU_E_POLICY_NOT_SET"
	dictErrorCodes.Add "0x8024001B","WU_E_SELFUPDATE_IN_PROGRESS"
	dictErrorCodes.Add "0x8024001D","WU_E_INVALID_UPDATE"
	dictErrorCodes.Add "0x8024001E","WU_E_SERVICE_STOP"
	dictErrorCodes.Add "0x8024001F","WU_E_NO_CONNECTION"
	dictErrorCodes.Add "0x80240020","WU_E_NO_INTERACTIVE_USER"
	dictErrorCodes.Add "0x80240021","WU_E_TIME_OUT"
	dictErrorCodes.Add "0x80240022","WU_E_ALL_UPDATES_FAILED"
	dictErrorCodes.Add "0x80240023","WU_E_EULAS_DECLINED"
	dictErrorCodes.Add "0x80240024","WU_E_NO_UPDATE"
	dictErrorCodes.Add "0x80240025","WU_E_USER_ACCESS_DISABLED"
	dictErrorCodes.Add "0x80240026","WU_E_INVALID_UPDATE_TYPE"
	dictErrorCodes.Add "0x80240027","WU_E_URL_TOO_LONG"
	dictErrorCodes.Add "0x80240028","WU_E_UNINSTALL_NOT_ALLOWED"
	dictErrorCodes.Add "0x80240029","WU_E_INVALID_PRODUCT_LICENSE"
	dictErrorCodes.Add "0x8024002A","WU_E_MISSING_HANDLER"
	dictErrorCodes.Add "0x8024002B","WU_E_LEGACYSERVER"
	dictErrorCodes.Add "0x8024002C","WU_E_BIN_SOURCE_ABSENT"
	dictErrorCodes.Add "0x8024002D","WU_E_SOURCE_ABSENT"
	dictErrorCodes.Add "0x8024002E","WU_E_WU_DISABLED"
	dictErrorCodes.Add "0x8024002F","WU_E_CALL_CANCELLED_BY_POLICY"
	dictErrorCodes.Add "0x80240030","WU_E_INVALID_PROXY_SERVER"
	dictErrorCodes.Add "0x80240031","WU_E_INVALID_FILE"
	dictErrorCodes.Add "0x80240032","WU_E_INVALID_CRITERIA"
	dictErrorCodes.Add "0x80240033","WU_E_EULA_UNAVAILABLE"
	dictErrorCodes.Add "0x80240034","WU_E_DOWNLOAD_FAILED"
	dictErrorCodes.Add "0x80240035","WU_E_UPDATE_NOT_PROCESSED"
	dictErrorCodes.Add "0x80240036","WU_E_INVALID_OPERATION"
	dictErrorCodes.Add "0x80240037","WU_E_NOT_SUPPORTED"
	dictErrorCodes.Add "0x80240038","WU_E_WINHTTP_INVALID_FILE"
	dictErrorCodes.Add "0x80240039","WU_E_TOO_MANY_RESYNC"
	dictErrorCodes.Add "0x80240040","WU_E_NO_SERVER_CORE_SUPPORT"
	dictErrorCodes.Add "0x80240041","WU_E_SYSPREP_IN_PROGRESS"
	dictErrorCodes.Add "0x80240042","WU_E_UNKNOWN_SERVICE"
	dictErrorCodes.Add "0x80240FFF","WU_E_UNEXPECTED"
	dictErrorCodes.Add "0x80241001","WU_E_MSI_WRONG_VERSION"
	dictErrorCodes.Add "0x80241002","WU_E_MSI_NOT_CONFIGURED"
	dictErrorCodes.Add "0x80241003","WU_E_MSP_DISABLED"
	dictErrorCodes.Add "0x80241004","WU_E_MSI_WRONG_APP_CONTEXT"
	dictErrorCodes.Add "0x80241FFF","WU_E_MSP_UNEXPECTED"
	dictErrorCodes.Add "0x80244000","WU_E_PT_SOAPCLIENT_BASE"
	dictErrorCodes.Add "0x80244001","WU_E_PT_SOAPCLIENT_INITIALIZE"
	dictErrorCodes.Add "0x80244002","WU_E_PT_SOAPCLIENT_OUTOFMEMORY"
	dictErrorCodes.Add "0x80244003","WU_E_PT_SOAPCLIENT_GENERATE"
	dictErrorCodes.Add "0x80244004","WU_E_PT_SOAPCLIENT_CONNECT"
	dictErrorCodes.Add "0x80244005","WU_E_PT_SOAPCLIENT_SEND"
	dictErrorCodes.Add "0x80244006","WU_E_PT_SOAPCLIENT_SERVER"
	dictErrorCodes.Add "0x80244007","WU_E_PT_SOAPCLIENT_SOAPFAULT"
	dictErrorCodes.Add "0x80244008","WU_E_PT_SOAPCLIENT_PARSEFAULT"
	dictErrorCodes.Add "0x80244009","WU_E_PT_SOAPCLIENT_READ"
	dictErrorCodes.Add "0x8024400A","WU_E_PT_SOAPCLIENT_PARSE"
	dictErrorCodes.Add "0x8024400B","WU_E_PT_SOAP_VERSION"
	dictErrorCodes.Add "0x8024400C","WU_E_PT_SOAP_MUST_UNDERSTAND"
	dictErrorCodes.Add "0x8024400D","WU_E_PT_SOAP_CLIENT"
	dictErrorCodes.Add "0x8024400E","WU_E_PT_SOAP_SERVER"
	dictErrorCodes.Add "0x8024400F","WU_E_PT_WMI_ERROR"
	dictErrorCodes.Add "0x80244010","WU_E_PT_EXCEEDED_MAX_SERVER_TRIPS"
	dictErrorCodes.Add "0x80244011","WU_E_PT_SUS_SERVER_NOT_SET"
	dictErrorCodes.Add "0x80244012","WU_E_PT_DOUBLE_INITIALIZATION"
	dictErrorCodes.Add "0x80244013","WU_E_PT_INVALID_COMPUTER_NAME"
	dictErrorCodes.Add "0x80244014","WU_E_PT_INVALID_COMPUTER_LSID"
	dictErrorCodes.Add "0x80244015","WU_E_PT_REFRESH_CACHE_REQUIRED"
	dictErrorCodes.Add "0x80244016","WU_E_PT_HTTP_STATUS_BAD_REQUEST"
	dictErrorCodes.Add "0x80244017","WU_E_PT_HTTP_STATUS_DENIED"
	dictErrorCodes.Add "0x80244018","WU_E_PT_HTTP_STATUS_FORBIDDEN"
	dictErrorCodes.Add "0x80244019","WU_E_PT_HTTP_STATUS_NOT_FOUND"
	dictErrorCodes.Add "0x8024401A","WU_E_PT_HTTP_STATUS_BAD_METHOD"
	dictErrorCodes.Add "0x8024401B","WU_E_PT_HTTP_STATUS_PROXY_AUTH_REQ"
	dictErrorCodes.Add "0x8024401C","WU_E_PT_HTTP_STATUS_REQUEST_TIMEOUT"
	dictErrorCodes.Add "0x8024401D","WU_E_PT_HTTP_STATUS_CONFLICT"
	dictErrorCodes.Add "0x8024401E","WU_E_PT_HTTP_STATUS_GONE"
	dictErrorCodes.Add "0x8024401F","WU_E_PT_HTTP_STATUS_SERVER_ERROR"
	dictErrorCodes.Add "0x80244020","WU_E_PT_HTTP_STATUS_NOT_SUPPORTED"
	dictErrorCodes.Add "0x80244021","WU_E_PT_HTTP_STATUS_BAD_GATEWAY"
	dictErrorCodes.Add "0x80244022","WU_E_PT_HTTP_STATUS_SERVICE_UNAVAIL"
	dictErrorCodes.Add "0x80244023","WU_E_PT_HTTP_STATUS_GATEWAY_TIMEOUT"
	dictErrorCodes.Add "0x80244024","WU_E_PT_HTTP_STATUS_VERSION_NOT_SUP"
	dictErrorCodes.Add "0x80244025","WU_E_PT_FILE_LOCATIONS_CHANGED"
	dictErrorCodes.Add "0x80244026","WU_E_PT_REGISTRATION_NOT_SUPPORTED"
	dictErrorCodes.Add "0x80244027","WU_E_PT_NO_AUTH_PLUGINS_REQUESTED"
	dictErrorCodes.Add "0x80244028","WU_E_PT_NO_AUTH_COOKIES_CREATED"
	dictErrorCodes.Add "0x80244029","WU_E_PT_INVALID_CONFIG_PROP"
	dictErrorCodes.Add "0x8024402A","WU_E_PT_CONFIG_PROP_MISSING"
	dictErrorCodes.Add "0x8024402B","WU_E_PT_HTTP_STATUS_NOT_MAPPED"
	dictErrorCodes.Add "0x8024402C","WU_E_PT_WINHTTP_NAME_NOT_RESOLVED"
	dictErrorCodes.Add "0x8024502D","WU_E_PT_SAME_REDIR_ID"
	dictErrorCodes.Add "0x8024502E","WU_E_PT_NO_MANAGED_RECOVER"
	dictErrorCodes.Add "0x8024402F","WU_E_PT_ECP_SUCCEEDED_WITH_ERRORS"
	dictErrorCodes.Add "0x80244030","WU_E_PT_ECP_INIT_FAILED"
	dictErrorCodes.Add "0x80244031","WU_E_PT_ECP_INVALID_FILE_FORMAT"
	dictErrorCodes.Add "0x80244032","WU_E_PT_ECP_INVALID_METADATA"
	dictErrorCodes.Add "0x80244033","WU_E_PT_ECP_FAILURE_TO_EXTRACT_DIGEST"
	dictErrorCodes.Add "0x80244034","WU_E_PT_ECP_FAILURE_TO_DECOMPRESS_CAB_FILE"
	dictErrorCodes.Add "0x80244035","WU_E_PT_ECP_FILE_LOCATION_ERROR"
	dictErrorCodes.Add "0x80244FFF","WU_E_PT_UNEXPECTED"
	dictErrorCodes.Add "0x80245001","WU_E_REDIRECTOR_LOAD_XML"
	dictErrorCodes.Add "0x80245002","WU_E_REDIRECTOR_S_FALSE"
	dictErrorCodes.Add "0x80245003","WU_E_REDIRECTOR_ID_SMALLER"
	dictErrorCodes.Add "0x80245FFF","WU_E_REDIRECTOR_UNEXPECTED"
	dictErrorCodes.Add "0x8024C001","WU_E_DRV_PRUNED"
	dictErrorCodes.Add "0x8024C002","WU_E_DRV_NOPROP_OR_LEGACY"
	dictErrorCodes.Add "0x8024C003","WU_E_DRV_REG_MISMATCH"
	dictErrorCodes.Add "0x8024C004","WU_E_DRV_NO_METADATA"
	dictErrorCodes.Add "0x8024C005","WU_E_DRV_MISSING_ATTRIBUTE"
	dictErrorCodes.Add "0x8024C006","WU_E_DRV_SYNC_FAILED"
	dictErrorCodes.Add "0x8024C007","WU_E_DRV_NO_PRINTER_CONTENT"
	dictErrorCodes.Add "0x8024CFFF","WU_E_DRV_UNEXPECTED"
	dictErrorCodes.Add "0x80248000","WU_E_DS_SHUTDOWN"
	dictErrorCodes.Add "0x80248001","WU_E_DS_INUSE"
	dictErrorCodes.Add "0x80248002","WU_E_DS_INVALID"
	dictErrorCodes.Add "0x80248003","WU_E_DS_TABLEMISSING"
	dictErrorCodes.Add "0x80248004","WU_E_DS_TABLEINCORRECT"
	dictErrorCodes.Add "0x80248005","WU_E_DS_INVALIDTABLENAME"
	dictErrorCodes.Add "0x80248006","WU_E_DS_BADVERSION"
	dictErrorCodes.Add "0x80248007","WU_E_DS_NODATA"
	dictErrorCodes.Add "0x80248008","WU_E_DS_MISSINGDATA"
	dictErrorCodes.Add "0x80248009","WU_E_DS_MISSINGREF"
	dictErrorCodes.Add "0x8024800A","WU_E_DS_UNKNOWNHANDLER"
	dictErrorCodes.Add "0x8024800B","WU_E_DS_CANTDELETE"
	dictErrorCodes.Add "0x8024800C","WU_E_DS_LOCKTIMEOUTEXPIRED"
	dictErrorCodes.Add "0x8024800D","WU_E_DS_NOCATEGORIES"
	dictErrorCodes.Add "0x8024800E","WU_E_DS_ROWEXISTS"
	dictErrorCodes.Add "0x8024800F","WU_E_DS_STOREFILELOCKED"
	dictErrorCodes.Add "0x80248010","WU_E_DS_CANNOTREGISTER"
	dictErrorCodes.Add "0x80248011","WU_E_DS_UNABLETOSTART"
	dictErrorCodes.Add "0x80248013","WU_E_DS_DUPLICATEUPDATEID"
	dictErrorCodes.Add "0x80248014","WU_E_DS_UNKNOWNSERVICE"
	dictErrorCodes.Add "0x80248015","WU_E_DS_SERVICEEXPIRED"
	dictErrorCodes.Add "0x80248016","WU_E_DS_DECLINENOTALLOWED"
	dictErrorCodes.Add "0x80248017","WU_E_DS_TABLESESSIONMISMATCH"
	dictErrorCodes.Add "0x80248018","WU_E_DS_SESSIONLOCKMISMATCH"
	dictErrorCodes.Add "0x80248019","WU_E_DS_NEEDWINDOWSSERVICE"
	dictErrorCodes.Add "0x8024801A","WU_E_DS_INVALIDOPERATION"
	dictErrorCodes.Add "0x8024801B","WU_E_DS_SCHEMAMISMATCH"
	dictErrorCodes.Add "0x8024801C","WU_E_DS_RESETREQUIRED"
	dictErrorCodes.Add "0x8024801D","WU_E_DS_IMPERSONATED"
	dictErrorCodes.Add "0x80248FFF","WU_E_DS_UNEXPECTED"
	dictErrorCodes.Add "0x80249001","WU_E_INVENTORY_PARSEFAILED"
	dictErrorCodes.Add "0x80249002","WU_E_INVENTORY_GET_INVENTORY_TYPE_FAILED"
	dictErrorCodes.Add "0x80249003","WU_E_INVENTORY_RESULT_UPLOAD_FAILED"
	dictErrorCodes.Add "0x80249004","WU_E_INVENTORY_UNEXPECTED"
	dictErrorCodes.Add "0x80249005","WU_E_INVENTORY_WMI_ERROR"
	dictErrorCodes.Add "0x8024A000","WU_E_AU_NOSERVICE"
	dictErrorCodes.Add "0x8024A002","WU_E_AU_NONLEGACYSERVER"
	dictErrorCodes.Add "0x8024A003","WU_E_AU_LEGACYCLIENTDISABLED"
	dictErrorCodes.Add "0x8024A004","WU_E_AU_PAUSED"
	dictErrorCodes.Add "0x8024A005","WU_E_AU_NO_REGISTERED_SERVICE"
	dictErrorCodes.Add "0x8024AFFF","WU_E_AU_UNEXPECTED"
	dictErrorCodes.Add "0x80242000","WU_E_UH_REMOTEUNAVAILABLE"
	dictErrorCodes.Add "0x80242001","WU_E_UH_LOCALONLY"
	dictErrorCodes.Add "0x80242002","WU_E_UH_UNKNOWNHANDLER"
	dictErrorCodes.Add "0x80242003","WU_E_UH_REMOTEALREADYACTIVE"
	dictErrorCodes.Add "0x80242004","WU_E_UH_DOESNOTSUPPORTACTION"
	dictErrorCodes.Add "0x80242005","WU_E_UH_WRONGHANDLER"
	dictErrorCodes.Add "0x80242006","WU_E_UH_INVALIDMETADATA"
	dictErrorCodes.Add "0x80242007","WU_E_UH_INSTALLERHUNG"
	dictErrorCodes.Add "0x80242008","WU_E_UH_OPERATIONCANCELLED"
	dictErrorCodes.Add "0x80242009","WU_E_UH_BADHANDLERXML"
	dictErrorCodes.Add "0x8024200A","WU_E_UH_CANREQUIREINPUT"
	dictErrorCodes.Add "0x8024200B","WU_E_UH_INSTALLERFAILURE"
	dictErrorCodes.Add "0x8024200C","WU_E_UH_FALLBACKTOSELFCONTAINED"
	dictErrorCodes.Add "0x8024200D","WU_E_UH_NEEDANOTHERDOWNLOAD"
	dictErrorCodes.Add "0x8024200E","WU_E_UH_NOTIFYFAILURE"
	dictErrorCodes.Add "0x8024200F","WU_E_UH_INCONSISTENT_FILE_NAMES"
	dictErrorCodes.Add "0x80242010","WU_E_UH_FALLBACKERROR"
	dictErrorCodes.Add "0x80242011","WU_E_UH_TOOMANYDOWNLOADREQUESTS"
	dictErrorCodes.Add "0x80242012","WU_E_UH_UNEXPECTEDCBSRESPONSE"
	dictErrorCodes.Add "0x80242013","WU_E_UH_BADCBSPACKAGEID"
	dictErrorCodes.Add "0x80242014","WU_E_UH_POSTREBOOTSTILLPENDING"
	dictErrorCodes.Add "0x80242015","WU_E_UH_POSTREBOOTRESULTUNKNOWN"
	dictErrorCodes.Add "0x80242016","WU_E_UH_POSTREBOOTUNEXPECTEDSTATE"
	dictErrorCodes.Add "0x80242017","WU_E_UH_NEW_SERVICING_STACK_REQUIRED"
	dictErrorCodes.Add "0x80242FFF","WU_E_UH_UNEXPECTED"
	dictErrorCodes.Add "0x80246001","WU_E_DM_URLNOTAVAILABLE"
	dictErrorCodes.Add "0x80246002","WU_E_DM_INCORRECTFILEHASH"
	dictErrorCodes.Add "0x80246003","WU_E_DM_UNKNOWNALGORITHM"
	dictErrorCodes.Add "0x80246004","WU_E_DM_NEEDDOWNLOADREQUEST"
	dictErrorCodes.Add "0x80246005","WU_E_DM_NONETWORK"
	dictErrorCodes.Add "0x80246006","WU_E_DM_WRONGBITSVERSION"
	dictErrorCodes.Add "0x80246007","WU_E_DM_NOTDOWNLOADED"
	dictErrorCodes.Add "0x80246008","WU_E_DM_FAILTOCONNECTTOBITS"
	dictErrorCodes.Add "0x80246009","WU_E_DM_BITSTRANSFERERROR"
	dictErrorCodes.Add "0x8024600A","WU_E_DM_DOWNLOADLOCATIONCHANGED"
	dictErrorCodes.Add "0x8024600B","WU_E_DM_CONTENTCHANGED"
	dictErrorCodes.Add "0x80246FFF","WU_E_DM_UNEXPECTED"
	dictErrorCodes.Add "0x8024D001","WU_E_SETUP_INVALID_INFDATA"
	dictErrorCodes.Add "0x8024D002","WU_E_SETUP_INVALID_IDENTDATA"
	dictErrorCodes.Add "0x8024D003","WU_E_SETUP_ALREADY_INITIALIZED"
	dictErrorCodes.Add "0x8024D004","WU_E_SETUP_NOT_INITIALIZED"
	dictErrorCodes.Add "0x8024D005","WU_E_SETUP_SOURCE_VERSION_MISMATCH"
	dictErrorCodes.Add "0x8024D006","WU_E_SETUP_TARGET_VERSION_GREATER"
	dictErrorCodes.Add "0x8024D007","WU_E_SETUP_REGISTRATION_FAILED"
	dictErrorCodes.Add "0x8024D008","WU_E_SELFUPDATE_SKIP_ON_FAILURE"
	dictErrorCodes.Add "0x8024D009","WU_E_SETUP_SKIP_UPDATE"
	dictErrorCodes.Add "0x8024D00A","WU_E_SETUP_UNSUPPORTED_CONFIGURATION"
	dictErrorCodes.Add "0x8024D00B","WU_E_SETUP_BLOCKED_CONFIGURATION"
	dictErrorCodes.Add "0x8024D00C","WU_E_SETUP_REBOOT_TO_FIX"
	dictErrorCodes.Add "0x8024D00D","WU_E_SETUP_ALREADYRUNNING"
	dictErrorCodes.Add "0x8024D00E","WU_E_SETUP_REBOOTREQUIRED"
	dictErrorCodes.Add "0x8024D00F","WU_E_SETUP_HANDLER_EXEC_FAILURE"
	dictErrorCodes.Add "0x8024D010","WU_E_SETUP_INVALID_REGISTRY_DATA"
	dictErrorCodes.Add "0x8024D011","WU_E_SELFUPDATE_REQUIRED"
	dictErrorCodes.Add "0x8024D012","WU_E_SELFUPDATE_REQUIRED_ADMIN"
	dictErrorCodes.Add "0x8024D013","WU_E_SETUP_WRONG_SERVER_VERSION"
	dictErrorCodes.Add "0x8024DFFF","WU_E_SETUP_UNEXPECTED"
	dictErrorCodes.Add "0x8024E001","WU_E_EE_UNKNOWN_EXPRESSION"
	dictErrorCodes.Add "0x8024E002","WU_E_EE_INVALID_EXPRESSION"
	dictErrorCodes.Add "0x8024E003","WU_E_EE_MISSING_METADATA"
	dictErrorCodes.Add "0x8024E004","WU_E_EE_INVALID_VERSION"
	dictErrorCodes.Add "0x8024E005","WU_E_EE_NOT_INITIALIZED"
	dictErrorCodes.Add "0x8024E006","WU_E_EE_INVALID_ATTRIBUTEDATA"
	dictErrorCodes.Add "0x8024E007","WU_E_EE_CLUSTER_ERROR"
	dictErrorCodes.Add "0x8024EFFF","WU_E_EE_UNEXPECTED"
	dictErrorCodes.Add "0x80243001","WU_E_INSTALLATION_RESULTS_UNKNOWN_VERSION"
	dictErrorCodes.Add "0x80243002","WU_E_INSTALLATION_RESULTS_INVALID_DATA"
	dictErrorCodes.Add "0x80243003","WU_E_INSTALLATION_RESULTS_NOT_FOUND"
	dictErrorCodes.Add "0x80243004","WU_E_TRAYICON_FAILURE"
	dictErrorCodes.Add "0x80243FFD","WU_E_NON_UI_MODE"
	dictErrorCodes.Add "0x80243FFE","WU_E_WUCLTUI_UNSUPPORTED_VERSION"
	dictErrorCodes.Add "0x80243FFF","WU_E_AUCLIENT_UNEXPECTED"
	dictErrorCodes.Add "0x8024F001","WU_E_REPORTER_EVENTCACHECORRUPT"
	dictErrorCodes.Add "0x8024F002","WU_E_REPORTER_EVENTNAMESPACEPARSEFAILED"
	dictErrorCodes.Add "0x8024F003","WU_E_INVALID_EVENT"
	dictErrorCodes.Add "0x8024F004","WU_E_SERVER_BUSY"
	dictErrorCodes.Add "0x8024FFFF","WU_E_REPORTER_UNEXPECTED"
	dictErrorCodes.Add "0x80247001","WU_E_OL_INVALID_SCANFILE"
	dictErrorCodes.Add "0x80247002","WU_E_OL_NEWCLIENT_REQUIRED"
	dictErrorCodes.Add "0x80247FFF","WU_E_OL_UNEXPECTED"

	Set CreateErrorCodesDict = dictErrorCodes
End Function 'CreateErrorCodesDict

Sub RunPostAndQuit
' solves quit without post issue
	RunFilesInDir("post")
	WScript.Quit
End Sub 'RunPostAndQuit



Class MD5er
	' A simple and slow vbscript based MD5 hasher
	' Do not feed too much data in :)
	Private BITS_TO_A_BYTE
	Private BYTES_TO_A_WORD
	Private BITS_TO_A_WORD
	Private m_lOnBits(30)
	Private m_l2Power(30)
	
	
	Private Sub Class_Initialize
		BITS_TO_A_BYTE = 8 
		BYTES_TO_A_WORD = 4 
		BITS_TO_A_WORD = 32
		
		m_lOnBits(0) = CLng(1) 
		m_lOnBits(1) = CLng(3) 
		m_lOnBits(2) = CLng(7) 
		m_lOnBits(3) = CLng(15) 
		m_lOnBits(4) = CLng(31) 
		m_lOnBits(5) = CLng(63) 
		m_lOnBits(6) = CLng(127) 
		m_lOnBits(7) = CLng(255) 
		m_lOnBits(8) = CLng(511) 
		m_lOnBits(9) = CLng(1023) 
		m_lOnBits(10) = CLng(2047) 
		m_lOnBits(11) = CLng(4095) 
		m_lOnBits(12) = CLng(8191) 
		m_lOnBits(13) = CLng(16383) 
		m_lOnBits(14) = CLng(32767) 
		m_lOnBits(15) = CLng(65535) 
		m_lOnBits(16) = CLng(131071) 
		m_lOnBits(17) = CLng(262143) 
		m_lOnBits(18) = CLng(524287) 
		m_lOnBits(19) = CLng(1048575) 
		m_lOnBits(20) = CLng(2097151) 
		m_lOnBits(21) = CLng(4194303) 
		m_lOnBits(22) = CLng(8388607) 
		m_lOnBits(23) = CLng(16777215) 
		m_lOnBits(24) = CLng(33554431) 
		m_lOnBits(25) = CLng(67108863) 
		m_lOnBits(26) = CLng(134217727) 
		m_lOnBits(27) = CLng(268435455) 
		m_lOnBits(28) = CLng(536870911) 
		m_lOnBits(29) = CLng(1073741823) 
		m_lOnBits(30) = CLng(2147483647) 
		
		m_l2Power(0) = CLng(1) 
		m_l2Power(1) = CLng(2) 
		m_l2Power(2) = CLng(4) 
		m_l2Power(3) = CLng(8) 
		m_l2Power(4) = CLng(16) 
		m_l2Power(5) = CLng(32) 
		m_l2Power(6) = CLng(64) 
		m_l2Power(7) = CLng(128) 
		m_l2Power(8) = CLng(256) 
		m_l2Power(9) = CLng(512) 
		m_l2Power(10) = CLng(1024) 
		m_l2Power(11) = CLng(2048) 
		m_l2Power(12) = CLng(4096) 
		m_l2Power(13) = CLng(8192) 
		m_l2Power(14) = CLng(16384) 
		m_l2Power(15) = CLng(32768) 
		m_l2Power(16) = CLng(65536) 
		m_l2Power(17) = CLng(131072) 
		m_l2Power(18) = CLng(262144) 
		m_l2Power(19) = CLng(524288) 
		m_l2Power(20) = CLng(1048576) 
		m_l2Power(21) = CLng(2097152) 
		m_l2Power(22) = CLng(4194304) 
		m_l2Power(23) = CLng(8388608) 
		m_l2Power(24) = CLng(16777216) 
		m_l2Power(25) = CLng(33554432) 
		m_l2Power(26) = CLng(67108864) 
		m_l2Power(27) = CLng(134217728) 
		m_l2Power(28) = CLng(268435456) 
		m_l2Power(29) = CLng(536870912) 
		m_l2Power(30) = CLng(1073741824) 

	End Sub 'Class_Initialize
	
	
	Public Property Get GetMD5(str)
		GetMD5 = MD5(str)
	End Property 'GetMD5
	
	
	Private Function LShift(lValue, iShiftBits) 
	   If iShiftBits = 0 Then 
	      LShift = lValue 
	      Exit Function 
	   ElseIf iShiftBits = 31 Then 
	      If lValue And 1 Then 
	         LShift = &H80000000 
	      Else 
	         LShift = 0 
	      End If 
	
	      Exit Function 
	   ElseIf iShiftBits < 0 Or iShiftBits > 31 Then 
	      Err.Raise 6 
	   End If 
	
	   If (lValue And m_l2Power(31 - iShiftBits)) Then 
	      LShift = ((lValue And m_lOnBits(31 - (iShiftBits + 1))) * m_l2Power(iShiftBits)) Or &H80000000 
	   Else 
	      LShift = ((lValue And m_lOnBits(31 - iShiftBits)) * m_l2Power(iShiftBits)) 
	   End If 
	End Function 'LShift
	
	
	Private Function RShift(lValue, iShiftBits) 
	   If iShiftBits = 0 Then 
	      RShift = lValue 
	      Exit Function 
	   ElseIf iShiftBits = 31 Then 
	      If lValue And &H80000000 Then 
	         RShift = 1 
	      Else 
	         RShift = 0 
	      End If 
	      Exit Function 
	   ElseIf iShiftBits < 0 Or iShiftBits > 31 Then 
	      Err.Raise 6 
	   End If 
	   
	   RShift = (lValue And &H7FFFFFFE) \ m_l2Power(iShiftBits) 
	
	   If (lValue And &H80000000) Then 
	      RShift = (RShift Or (&H40000000 \ m_l2Power(iShiftBits - 1))) 
	   End If 
	End Function 'RShift
	
	
	Private Function RotateLeft(lValue, iShiftBits) 
	   RotateLeft = LShift(lValue, iShiftBits) Or RShift(lValue, (32 - iShiftBits)) 
	End Function 'RotateLeft
	
	
	Private Function AddUnsigned(lX, lY) 
	   Dim lX4 
	   Dim lY4 
	   Dim lX8 
	   Dim lY8 
	   Dim lResult 
	   
	   lX8 = lX And &H80000000 
	   lY8 = lY And &H80000000 
	   lX4 = lX And &H40000000 
	   lY4 = lY And &H40000000 
	   
	   lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF) 
	   
	   If lX4 And lY4 Then 
	      lResult = lResult Xor &H80000000 Xor lX8 Xor lY8 
	   ElseIf lX4 Or lY4 Then 
	      If lResult And &H40000000 Then 
	         lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8 
	      Else 
	         lResult = lResult Xor &H40000000 Xor lX8 Xor lY8 
	      End If 
	   Else 
	      lResult = lResult Xor lX8 Xor lY8 
	   End If 
	   
	   AddUnsigned = lResult 
	End Function 'AddUnsigned
	
	Private Function F(x, y, z) 
	   F = (x And y) Or ((Not x) And z) 
	End Function 'F
	
	Private Function G(x, y, z) 
	   G = (x And z) Or (y And (Not z)) 
	End Function 'G
	
	Private Function H(x, y, z) 
	   H = (x Xor y Xor z) 
	End Function 'H
	
	Private Function I(x, y, z) 
	   I = (y Xor (x Or (Not z))) 
	End Function 'I
	
	Private Sub FF(a, b, c, d, x, s, ac) 
	   a = AddUnsigned(a, AddUnsigned(AddUnsigned(F(b, c, d), x), ac)) 
	   a = RotateLeft(a, s) 
	   a = AddUnsigned(a, b) 
	End Sub 'FF
	
	Private Sub GG(a, b, c, d, x, s, ac) 
	   a = AddUnsigned(a, AddUnsigned(AddUnsigned(G(b, c, d), x), ac)) 
	   a = RotateLeft(a, s) 
	   a = AddUnsigned(a, b) 
	End Sub 'GG
	
	Private Sub HH(a, b, c, d, x, s, ac) 
	   a = AddUnsigned(a, AddUnsigned(AddUnsigned(H(b, c, d), x), ac)) 
	   a = RotateLeft(a, s) 
	   a = AddUnsigned(a, b) 
	End Sub 'HH
	
	Private Sub II(a, b, c, d, x, s, ac) 
	   a = AddUnsigned(a, AddUnsigned(AddUnsigned(I(b, c, d), x), ac)) 
	   a = RotateLeft(a, s) 
	   a = AddUnsigned(a, b) 
	End Sub 'II
	
	Private Function ConvertToWordArray(sMessage) 
	   Dim lMessageLength 
	   Dim lNumberOfWords 
	   Dim lWordArray() 
	   Dim lBytePosition 
	   Dim lByteCount 
	   Dim lWordCount 
	   
	   Const MODULUS_BITS = 512 
	   Const CONGRUENT_BITS = 448 
	   
	   lMessageLength = Len(sMessage) 
	   
	   lNumberOfWords = (((lMessageLength + ((MODULUS_BITS - CONGRUENT_BITS) \ BITS_TO_A_BYTE)) \ (MODULUS_BITS \ BITS_TO_A_BYTE)) + 1) * (MODULUS_BITS \ BITS_TO_A_WORD) 
	   ReDim lWordArray(lNumberOfWords - 1) 
	   
	   lBytePosition = 0 
	   lByteCount = 0 
	   Do Until lByteCount >= lMessageLength 
	      lWordCount = lByteCount \ BYTES_TO_A_WORD 
	      lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE 
	      lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(Asc(Mid(sMessage, lByteCount + 1, 1)), lBytePosition) 
	      lByteCount = lByteCount + 1 
	   Loop 
	
	   lWordCount = lByteCount \ BYTES_TO_A_WORD 
	   lBytePosition = (lByteCount Mod BYTES_TO_A_WORD) * BITS_TO_A_BYTE 
	
	   lWordArray(lWordCount) = lWordArray(lWordCount) Or LShift(&H80, lBytePosition) 
	
	   lWordArray(lNumberOfWords - 2) = LShift(lMessageLength, 3) 
	   lWordArray(lNumberOfWords - 1) = RShift(lMessageLength, 29) 
	   
	   ConvertToWordArray = lWordArray 
	End Function 'ConvertToWordArray
	
	Private Function WordToHex(lValue) 
	   Dim lByte 
	   Dim lCount 
	   
	   For lCount = 0 To 3 
	      lByte = RShift(lValue, lCount * BITS_TO_A_BYTE) And m_lOnBits(BITS_TO_A_BYTE - 1) 
	      WordToHex = WordToHex & Right("0" & Hex(lByte), 2) 
	   Next 
	End Function 'WordToHex
	
	Private Function MD5(sMessage) 
	   Dim x 
	   Dim k 
	   Dim AA 
	   Dim BB 
	   Dim CC 
	   Dim DD 
	   Dim a 
	   Dim b 
	   Dim c 
	   Dim d 
	   
	   Const S11 = 7 
	   Const S12 = 12 
	   Const S13 = 17 
	   Const S14 = 22 
	   Const S21 = 5 
	   Const S22 = 9 
	   Const S23 = 14 
	   Const S24 = 20 
	   Const S31 = 4 
	   Const S32 = 11 
	   Const S33 = 16 
	   Const S34 = 23 
	   Const S41 = 6 
	   Const S42 = 10 
	   Const S43 = 15 
	   Const S44 = 21 
	
	   x = ConvertToWordArray(sMessage) 
	   
	   a = &H67452301 
	   b = &HEFCDAB89 
	   c = &H98BADCFE 
	   d = &H10325476 
	
	   For k = 0 To UBound(x) Step 16 
	      AA = a 
	      BB = b 
	      CC = c 
	      DD = d 
	   
	      FF a, b, c, d, x(k + 0), S11, &HD76AA478 
	      FF d, a, b, c, x(k + 1), S12, &HE8C7B756 
	      FF c, d, a, b, x(k + 2), S13, &H242070DB 
	      FF b, c, d, a, x(k + 3), S14, &HC1BDCEEE 
	      FF a, b, c, d, x(k + 4), S11, &HF57C0FAF 
	      FF d, a, b, c, x(k + 5), S12, &H4787C62A 
	      FF c, d, a, b, x(k + 6), S13, &HA8304613 
	      FF b, c, d, a, x(k + 7), S14, &HFD469501 
	      FF a, b, c, d, x(k + 8), S11, &H698098D8 
	      FF d, a, b, c, x(k + 9), S12, &H8B44F7AF 
	      FF c, d, a, b, x(k + 10), S13, &HFFFF5BB1 
	      FF b, c, d, a, x(k + 11), S14, &H895CD7BE 
	      FF a, b, c, d, x(k + 12), S11, &H6B901122 
	      FF d, a, b, c, x(k + 13), S12, &HFD987193 
	      FF c, d, a, b, x(k + 14), S13, &HA679438E 
	      FF b, c, d, a, x(k + 15), S14, &H49B40821 
	   
	      GG a, b, c, d, x(k + 1), S21, &HF61E2562 
	      GG d, a, b, c, x(k + 6), S22, &HC040B340 
	      GG c, d, a, b, x(k + 11), S23, &H265E5A51 
	      GG b, c, d, a, x(k + 0), S24, &HE9B6C7AA 
	      GG a, b, c, d, x(k + 5), S21, &HD62F105D 
	      GG d, a, b, c, x(k + 10), S22, &H2441453 
	      GG c, d, a, b, x(k + 15), S23, &HD8A1E681 
	      GG b, c, d, a, x(k + 4), S24, &HE7D3FBC8 
	      GG a, b, c, d, x(k + 9), S21, &H21E1CDE6 
	      GG d, a, b, c, x(k + 14), S22, &HC33707D6 
	      GG c, d, a, b, x(k + 3), S23, &HF4D50D87 
	      GG b, c, d, a, x(k + 8), S24, &H455A14ED 
	      GG a, b, c, d, x(k + 13), S21, &HA9E3E905 
	      GG d, a, b, c, x(k + 2), S22, &HFCEFA3F8 
	      GG c, d, a, b, x(k + 7), S23, &H676F02D9 
	      GG b, c, d, a, x(k + 12), S24, &H8D2A4C8A 
	         
	      HH a, b, c, d, x(k + 5), S31, &HFFFA3942 
	      HH d, a, b, c, x(k + 8), S32, &H8771F681 
	      HH c, d, a, b, x(k + 11), S33, &H6D9D6122 
	      HH b, c, d, a, x(k + 14), S34, &HFDE5380C 
	      HH a, b, c, d, x(k + 1), S31, &HA4BEEA44 
	      HH d, a, b, c, x(k + 4), S32, &H4BDECFA9 
	      HH c, d, a, b, x(k + 7), S33, &HF6BB4B60 
	      HH b, c, d, a, x(k + 10), S34, &HBEBFBC70 
	      HH a, b, c, d, x(k + 13), S31, &H289B7EC6 
	      HH d, a, b, c, x(k + 0), S32, &HEAA127FA 
	      HH c, d, a, b, x(k + 3), S33, &HD4EF3085 
	      HH b, c, d, a, x(k + 6), S34, &H4881D05 
	      HH a, b, c, d, x(k + 9), S31, &HD9D4D039 
	      HH d, a, b, c, x(k + 12), S32, &HE6DB99E5 
	      HH c, d, a, b, x(k + 15), S33, &H1FA27CF8 
	      HH b, c, d, a, x(k + 2), S34, &HC4AC5665 
	   
	      II a, b, c, d, x(k + 0), S41, &HF4292244 
	      II d, a, b, c, x(k + 7), S42, &H432AFF97 
	      II c, d, a, b, x(k + 14), S43, &HAB9423A7 
	      II b, c, d, a, x(k + 5), S44, &HFC93A039 
	      II a, b, c, d, x(k + 12), S41, &H655B59C3 
	      II d, a, b, c, x(k + 3), S42, &H8F0CCC92 
	      II c, d, a, b, x(k + 10), S43, &HFFEFF47D 
	      II b, c, d, a, x(k + 1), S44, &H85845DD1 
	      II a, b, c, d, x(k + 8), S41, &H6FA87E4F 
	      II d, a, b, c, x(k + 15), S42, &HFE2CE6E0 
	      II c, d, a, b, x(k + 6), S43, &HA3014314 
	      II b, c, d, a, x(k + 13), S44, &H4E0811A1 
	      II a, b, c, d, x(k + 4), S41, &HF7537E82 
	      II d, a, b, c, x(k + 11), S42, &HBD3AF235 
	      II c, d, a, b, x(k + 2), S43, &H2AD7D2BB 
	      II b, c, d, a, x(k + 9), S44, &HEB86D391 
	   
	      a = AddUnsigned(a, AA) 
	      b = AddUnsigned(b, BB) 
	      c = AddUnsigned(c, CC) 
	      d = AddUnsigned(d, DD) 
	   Next 
	   
	   MD5 = LCase(WordToHex(a) & WordToHex(b) & WordToHex(c) & WordToHex(d)) 
	End Function 'MD5

End Class 'MD5er



' :::VBLib:TaniumNamedArg:Begin:::
Class TaniumNamedArg
	' Private m_dictTypes
	Private m_value
	Private m_name
	Private m_defaultValue
	Private m_exampleValue
	Private m_helpText
	Private m_CompanionArgName
	Private m_bIsOptional
	Private m_libVersion
	Private m_libName
	Private m_bErr
	Private m_errMessage
	Private m_translationFunctionRef
	Private m_validationFunctionRef
	Private IS_STRING
	Private IS_DOUBLE
	Private IS_INTEGER
	Private IS_YESNOTRUEFALSE
	Private m_arrTypeFlags
	Private m_bUnescape

	Private Sub Class_Initialize
		' No Constants inside a class
		IS_STRING = 0
		IS_DOUBLE = 1
		IS_INTEGER = 2
		IS_YESNOTRUEFALSE = 3
		m_libVersion = "6.5.314.4216"
		m_libName = "TaniumNamedArg"
		' Set m_dictTypes = CreateObject("Scripting.Dictionary")
		m_arrTypeFlags = Array()
		' Keep this set to whatever the highest type CONST value is
		ReDim m_arrTypeFlags(IS_YESNOTRUEFALSE)
		m_defaultValue = ""
		m_CompanionArgName = ""
		m_bIsOptional = True
		m_value = ""
		m_errMessage = ""
		m_helpText = "Descriptive Help Text Here"
		m_exampleValue = ""
		m_bUnescape = True
		m_bErr = False
		' Can supply any function to change the input however
		' it is needed, before it is placed into any argument
		' container
		Set m_translationFunctionRef = Nothing
		' Same for validation
		Set m_validationFunctionRef = Nothing
    End Sub

	Private Sub Class_Terminate
		' Set m_dictTypes = Nothing
	End Sub

    Public Property Get ErrorState
    	ErrorState = m_bErr
    End Property

    Public Property Get ErrorMessage
    	ErrorMessage = m_errMessage
    End Property

	Public Property Let TranslationFunctionReference(ByRef func)
		' Allow consumer to set a translation function as a property
		' Do this by doing
		' Set x = GetRef("MyFunctionName")
		' <thisobject>.TranslationFunctionReference = x
		If CheckVarType(func,vbObject) Then
			Set m_translationFunctionRef = func
		End If
		ErrorCheck
	End Property 'TranslationFunctionReference

	Public Property Let ValidationFunctionReference(ByRef func)
		' Allow consumer to set a validation function as a property
		' Do this by doing
		' Set x = GetRef("MyFunctionName")
		' <thisobject>.TranslationFunctionReference = x
		If CheckVarType(func,vbObject) Then
			Set m_validationFunctionRef = func
		End If
		ErrorCheck
	End Property 'ValidationFunctionReference

	Public Property Let ArgValue(value)
		If Not m_validationFunctionRef Is Nothing Then
			If Not m_validationFunctionRef(value) Then
				m_bErr = True
				m_errMessage = "Error: Using supplied Validation Function, argument was not valid"
			End If
		End If
		If Not m_translationFunctionRef Is Nothing Then
			value = m_translationFunctionRef(value)
		End If
		On Error Resume Next
		m_value = CStr(value)
		If Err.Number <> 0 Then
			m_bErr = True
			m_errMessage = "Error: Could not convert parameter value to string ("&Err.Description&")"
		End If
		On Error Goto 0

		SetArgValueByType(value)
		ErrorCheck
	End Property 'ArgValue

	Public Property Get ArgValue
		ArgValue = m_value
	End Property 'ArgValue

	Public Property Get ArgName
		ArgName = m_name
	End Property 'Name

	Public Property Get UnescapeFlag
		UnescapeFlag = m_bUnescape
	End Property 'UnescapeFlag

	Public Property Let UnescapeFlag(bUnescapeFlag)
		If VarType(bUnescapeFlag) = vbBoolean Then
			m_bUnescape = bUnescapeFlag
		Else
			m_bErr = True
			m_errMessage = "Error: The argument unescape flag must be a boolean value"
		End If
		ErrorCheck
	End Property 'UnescapeFlag

	Public Property Let ArgName(value)
		m_name = GetString(value)
	End Property 'DefaultValue

	Public Property Get HelpText
		HelpText = m_helpText
	End Property 'HelpText

	Public Property Let HelpText(value)
		m_helpText = value
	End Property 'HelpText

	Public Property Get ExampleValue
		ExampleValue = m_exampleValue
	End Property 'ExampleValue

	Public Property Let ExampleValue(value)
		m_exampleValue = value
	End Property 'ExampleValue

	Public Property Let DefaultValue(value)
		m_defaultValue = value
		If m_value = "" Then
			SetArgValueByType(value)
		End If
		ErrorCheck
	End Property 'DefaultValue

	Public Property Get DefaultValue
		DefaultValue = m_defaultValue
	End Property 'DefaultValue

	Public Property Let CompanionArgumentName(strOtherArgName)
		If CheckVarType(strOtherArgName,vbString) Then
			m_CompanionArgName = strOtherArgName
		End If
		ErrorCheck
	End Property 'CompanionArgumentName

	Public Property Get CompanionArgumentName
		CompanionArgumentName = m_CompanionArgName
	End Property 'CompanionArgumentName

	Public Property Let IsOptional(b)
		If CheckVarType(b,vbBoolean) Then
			m_bIsOptional = b
		End If
		ErrorCheck
	End Property 'IsOptional

	Public Property Get IsOptional
		IsOptional = m_bIsOptional
	End Property 'IsOptional

	' Set input type to string (default case)
	Public Property Let RequireDecimal(b)
		SetTypeArrayVal b,IS_DOUBLE
		ErrorCheck
	End Property 'RequireDecimal

	Public Property Get RequireDecimal
		RequireDecimal = m_arrTypeFlags(IS_DOUBLE)
	End Property 'RequireDecimal

	Public Property Let RequireString(b)
		SetTypeArrayVal b,IS_STRING
		ErrorCheck
	End Property 'RequireString

	Public Property Get RequireString
		RequireString = m_arrTypeFlags(IS_STRING)
	End Property 'RequireString

	Public Property Let RequireInteger(b)
		SetTypeArrayVal b,IS_INTEGER
		ErrorCheck
	End Property 'RequireInteger

	Public Property Get RequireInteger
		RequireInteger = m_arrTypeFlags(IS_INTEGER)
	End Property 'RequireInteger

	Public Property Let RequireYesNoTrueFalse(b)
		SetTypeArrayVal b,IS_YESNOTRUEFALSE
		ErrorCheck
	End Property 'RequireYesNoTrueFalse

	Public Property Get RequireYesNoTrueFalse
		RequireYesNoTrueFalse = m_arrTypeFlags(IS_YESNOTRUEFALSE)
	End Property 'RequireYesNoTrueFalse

	Private Function CheckVarType(var,typeNum)
		CheckVarType = True
		If VarType(var) <> typeNum Then
			 m_bErr = True
			 m_errMessage = "Error: Tried to set " & var & " to an invalid var type: " & typeNum
			 CheckVarType = False
		End If
	End Function 'CheckVarType

	Private Sub SetArgValueByType(value)
	' Looks at value to determine if it's an OK value
		Dim theSetFlag, i
		For i = 0 To UBound(m_arrTypeFlags)
			If m_arrTypeFlags(i) Then
				theSetFlag = i
			End If
		Next

		Select Case theSetFlag
			Case IS_STRING
				m_value = GetString(value)
			Case IS_DOUBLE
				m_value = GetDouble(value)
			Case IS_INTEGER
				m_value = GetInteger(value)
			Case IS_YESNOTRUEFALSE
				m_value = GetYesNoTrueFalse(value)
			Case Else
				m_bErr = True
				m_errMessage = "Error: Could not reliably determine flag type (please update library types): " & theSetFlag
		End Select
		ErrorCheck
	End Sub 'SetArgValueByType

	Private Function GetString(value)
		GetString = False
		On Error Resume Next
		value = CStr(value)
		If Err.Number <> 0 Then
			m_bErr = True
			m_errMessage = "Error: Could not convert value to string ("&Err.Description&")"
		Else
			GetString = value
		End If
		On Error Goto 0

	End Function 'GetString

	Private Function GetYesNoTrueFalse(value)
		GetYesNoTrueFalse = "" ' would be invalid as boolean
		On Error Resume Next
		value = CStr(value)
		If Err.Number <> 0 Then
			m_bErr = True
			m_errMessage = "Error: Could not convert value to string ("&Err.Description&")"
		End If
		On Error Goto 0
		value = LCase(value)

		Select Case value
			Case "yes"
				GetYesNoTrueFalse = True
			Case "true"
				GetYesNoTrueFalse = True
			Case "no"
				GetYesNoTrueFalse = False
			Case "false"
				GetYesNoTrueFalse = False
			Case Else
				m_bErr = True
				m_errMessage = "Error: Argument "&Chr(34)&m_name&Chr(34)&" requires Yes or No as input value, was given: " &value
		End Select
	End Function 'GetYesNoTrueFalse

	Private Function GetDouble(value)
		GetDouble = False
		If Not IsNumeric(value) Then
			m_bErr = True
			m_errMessage = "Error: argument "&m_name&" with value " & value & " is set to Decimal type but is not able to be converted to a number."
			Exit Function
		End If

		value = CDbl(value)
		If Err.Number <> 0 Then
			m_bErr = True
			m_errMessage = "Error: argument "&m_name&" with value " & value & " could not be converted to a Double, decimal value. ("&Err.Description&")"
		Else
			GetDouble = value
		End If
	End Function 'GetDouble

	Private Function GetInteger(value)
		' If value is an integer (or a string that can be an integer), store it
		' default case is to not accept value
		GetInteger = False
		' first character could be a dollar sign which is convertible
		' this is the case which occurs when a tanium command line has an invalid parameter spec
		Dim intDollar
		intDollar = InStr(value,"$")
		If intDollar > 0 And Len(value) > 1 Then ' only if more than one char
			value = Right(value,Len(value) - 1)
		End If
		If VarType(value) = vbString Then
			If Not IsNumeric(value) Then
				m_bErr = True
				m_ErrMessage = m_libName& " Error: " & value & " could not be converted to a number."
			End If
			Dim conv
			On Error Resume Next
			conv = CStr(CLng(value))
			If Err.Number <> 0 Then
				m_bErr = True
				m_ErrMessage = m_libName & " Error: " & value & " could not be converted to an integer. - max size is +/-2,147,483,647. ("&Err.Description&")"
			End If
			On Error Goto 0
			If conv = value Then
				GetInteger = CLng(value)
			End If
		ElseIf VarType(value) = vbLong Or VarType(value) = vbInteger Then
			GetInteger = CLng(value)
		Else
		 ' some non-string, non-numeric value
			m_bErr = True
			m_ErrMessage = m_libName & " Error: argument could not be converted to an integer, was type "&TypeName(value)
		End If
		ErrorCheck
	End Function 'GetInteger

	Private Sub SetTypeArrayVal(b,typeConst)
		If CheckVarType(b,vbBoolean) Then
			' clear all other types
			ClearTypesArray
			' Set this type
			m_arrTypeFlags(typeConst) = b
			' Ensure all others are false, with
			' potential default back to generic string
			CheckTypesArray
		End If
	End Sub 'SetTypeArrayVal

	Private Sub ClearTypesArray
		Dim i
		For i = 0 To UBound(m_arrTypeFlags)
			m_arrTypeFlags(i) = False
		Next
	End Sub 'ClearTypesArray

	Private Sub CheckTypesArray
		' If all values are false, default back to 'string'
		' There is no intention to support multiple types for single arg
		Dim i, bState
		bState = False
		For i = 0 To UBound(m_arrTypeFlags)
			bState = bState Or m_arrTypeFlags(i)
		Next

		If bState = False Then 'All are false, revert to string
			m_arrTypeFlags(IS_STRING) = True
		End If
	End Sub 'CheckTypesArray

	Public Sub ErrorClear
		m_bErr = False
		m_errMessage = ""
	End Sub

	Private Sub ErrorCheck
		' Call on all Lets
		If m_bErr Then
			Err.Raise vbObjectError + 1978, m_libName, m_errMessage
		End If
	End Sub 'ErrorCheck

End Class 'TaniumNamedArg
' :::VBLib:TaniumNamedArg:End:::

' :::VBLib:TaniumNamedArgsParser:Begin:::
Class TaniumNamedArgsParser
	Private m_args
	Private m_dict
	Private m_intMaxlines
	Private m_strWarning
	Private m_libVersion
	Private m_libName
	Private m_arrHelpArgs
	Private m_programDescription
	Private m_bErr
	Private m_errMessage

	Private Sub Class_Initialize
		Set m_args = WScript.Arguments.Named
		m_libVersion = "6.5.314.4216"
		m_libName = "TaniumNamedArgsParser"
		Set m_dict = CreateObject("Scripting.Dictionary")
		m_dict.CompareMode = vbTextCompare
		m_arrHelpArgs = Array("/?","/help","help","-h","--help")
    End Sub

    Public Property Get ErrorState
    	ErrorState = m_bErr
    End Property

    Public Property Get ErrorMessage
    	ErrorMessage = m_errMessage
    End Property

    Public Sub AddArg(arg)
    	' Arg is a TaniumNamedArg argument object
    	On Error Resume Next
    	' Very simple way to check if it's the right object type
    	Dim name, b
    	b = arg.RequireYesNoTrueFalse
    	If Err.Number <> 0 Then
    		' This is not a valid object
			m_bErr = True
			m_errMessage = "Error: Not operating on a TaniumNamedArg object, ("&Err.Description&")"
    	End If
    	' The behavior of named arguments in vbscript is to keep only the first argument's value
    	' if multiple same-named arguments are passed in
    	' so we will do the same, without an error message
    	If Not m_dict.Exists(arg.ArgName) Then
    		m_dict.Add arg.ArgName,arg
   		End If
    	ErrorCheck
    End Sub 'AddArg

    Public Property Get LibVersion
    	LibVersion = m_libVersion
    End Property

    Public Function GetArg(strName)
    	' Returns the Arg Object if it exists
    	Set GetArg = Nothing
    	If VarType(strName) = vbString Then
    		If m_dict.Exists(strName) Then
    			Set GetArg = m_dict.Item(strName)
    		End If
    	End If
    End Function 'GetArg

	Public Property Get ProgramDescription
		HelpText = m_programDescription
	End Property 'HelpText

	Public Property Let ProgramDescription(strDescription)
		m_programDescription = strDescription
	End Property 'HelpText

	Public Property Get WasPassedIn(strArgName)
		WasPassedIn = False
		If WScript.Arguments.Named.Exists(strArgname) Then
			WasPassedIn = True
		End If
	End Property 'WasPassedIn

	Public Property Get AddedArgsArray
		'Returns array of all added args
		'Can loop over and recognize by arg.ArgName
		'for direct access of a single arg, use GetArg method
		Dim strKey,j,size
		If m_dict.Count = 0 Then
			size = 0
		Else
			size = m_dict.Count - 1
		End If

		ReDim outArr(size) ' variable size
		j = 0
		For Each strKey In m_dict.Keys
			Set outArr(j) =  m_dict.Item(strKey)
			j = j + 1
		Next
		AddedArgsArray = outArr
	End Property 'AddedArgsArray

	Public Sub Parse
		Dim colNamedArgs
		Set colNamedArgs = WScript.Arguments.Named
		' Loop through all supplied named arguments
		Dim externalArg,internalArg,internalArgName

		If m_dict.Count = 0 Then
			m_bErr = True
			m_errMessage = "Error: NamedArgsParser had zero arguments added, nothing to parse!"
		Else
			' Even if no command line arguments are input, if any arguments have a default value, they are considered
			' arguments
			For Each internalArg In m_dict.Keys
				If m_dict.Item(internalArg).DefaultValue <> "" Then
					On Error Resume Next ' some argvalue sets will raise
					m_dict.Item(internalArg).ArgValue = m_dict.Item(internalArg).DefaultValue
					If Err.Number <> 0 Then
						m_bErr = True
						m_errMessage = "Error: Could not set argument "&m_dict.Item(internalArg).ArgName&" value to it's default value "& m_dict.Item(internalArg).DefaultValue&" - " &Err.Description
						Err.Clear
					End If
					On Error Goto 0
				End If
			Next
			' next overwrite any default values with arguments passed in
			For Each internalArg In m_dict.Keys
				If colNamedArgs.Exists(internalArg) Then
					On Error Resume Next ' some argvalue sets will raise
					Dim strProvidedArg
					If m_dict.Item(internalArg).UnescapeFlag Then
						strProvidedArg = Trim(Unescape(colNamedArgs.Item(internalArg)))
					Else
						strProvidedArg = Trim(colNamedArgs.Item(internalArg))
					End If
					m_dict.Item(internalArg).ArgValue = strProvidedArg
					If Err.Number <> 0 Then
						m_bErr = True
						m_errMessage = "Error: Argument "&m_dict.Item(internalArg).ArgName&" was set to invalid data, message was "&Err.Description
						Err.Clear
					End If
					On Error Goto 0
				End If
			Next

			' Go through Dictionary and sanity check
			' Check if required arguments are all there
			For Each internalArgName In m_dict.Keys
				Set internalArg = m_dict.Item(internalArgName)
				If internalArg.IsOptional = False And ( internalArg.ArgName = "" Or internalArg.ArgValue = "" ) Then
					m_bErr = True
					m_errMessage = "Error: Required argument "&internalArg.ArgName&" has no value"
				End If
			Next

			' Go through dictionary and check that any arguments required by others do exist
			Dim requiredArg
			For Each internalArgName In m_dict.Keys
				Set internalArg = m_dict.Item(internalArgName)
				If internalArg.CompanionArgumentName <> "" Then
					If Not m_dict.Exists(internalArg.CompanionArgumentName) Then
						' A required argument does not exist in the dictionary
						m_bErr = True
						m_errMessage = "Error: Argument " & internalArg.ArgName & " Requires " _
							& internalArg.CompanionArgumentName & ", which was not specified"
					Else
						If m_dict.Item(internalArg.CompanionArgumentName).ArgValue = "" Then
							m_bErr = True
							m_errMessage = "Error: Argument " & internalArg.ArgName & " Requires " _
								& internalArg.CompanionArgumentName & ", which has a null value"
						End If
					End If
				End If
			Next
		End If
		SanityCheck m_errMessage
		' Check for any help arguments
		If HasHelpArg Then
			PrintUsageAndQuit "Help invoked from command line"
		End If
	End Sub 'Parse

	Private Function HasHelpArg
		'Returns whether the command line arguments indicate help is needed
		HasHelpArg = False
		Dim args,arg,helpArg
		Set args = WScript.Arguments

		For Each arg In args
			arg = LCase(arg)
			For Each helpArg In m_arrHelpArgs
				If arg = helpArg Then
					HasHelpArg = True
				End If
			Next
		Next

	End Function 'HasHelpArg

	Public Sub PrintUsageAndQuit(strOptionalMessage)
	' Prints usage of script based on arguments, and quits
		If Not strOptionalMessage = "" Then
			WScript.Echo strOptionalMessage
		End If
		ArgEcho 0,m_programDescription
		ArgEcho 0,"Usage: "&WScript.ScriptName&" [arguments]"
		' Begin printing the dictionary in a clever way
		Dim argName,argObj,strRequiredBlock,strOptionalBlock
		Dim strArgData
		For Each argName In m_dict.Keys
			Set argObj = m_dict.Item(argName)
			strArgData = "/"&argObj.ArgName&":"&"<"&argObj.ExampleValue&">" _
				&" ("&argObj.HelpText&")"
			If argObj.CompanionArgumentName <> "" Then
				strArgData = strArgData&vbCrLf&AddTabs(2,"Requires argument /" _
					& argObj.CompanionArgumentName _
					& " to also be specified and set to a non-blank value")
			End If
			strArgData = strArgData & vbCrLf
			strArgData = AddTabs(1,strArgData)
			If argObj.IsOptional Then
				strOptionalBlock = strOptionalBlock _
					& strArgData
			Else
				strRequiredBlock = strRequiredBlock _
					& strArgData
			End If
		Next
		ArgEcho 1,"------REQUIRED------"
		ArgEcho 0,strRequiredBlock
		ArgEcho 1,"------OPTIONAL------"
		ArgEcho 0,strOptionalBlock
		WScript.Quit 1
	End Sub 'PrintUsageAndQuit

	Public Sub ErrorClear
		m_bErr = False
		m_errMessage = ""
	End Sub

	Private Sub ErrorCheck
		' Call on all Lets
		If m_bErr Then
			Err.Raise vbObjectError + 1978, m_libName, m_errMessage
		End If
	End Sub 'ErrorCheck

	Private Sub SanityCheck(strOptionalMessage)
		' This error check will not raise any errors
		' Instead, it will call PrintUsageAndQuit internally
		' if there is an error condition
		If m_bErr Then
			PrintUsageAndQuit "Argument Parse Error was: "&m_errMessage
		End If
	End Sub 'ErrorCheck

	Private Sub ArgEcho(intTabCount,str)
		WScript.Echo String(intTabCount,vbTab)&str
	End Sub 'ArgEcho

	Private Function AddTabs(intTabCount,str)
		AddTabs = String(intTabCount,vbTab)&str
	End Function 'ArgEcho

	Private Sub Class_Terminate
		Set m_args = Nothing
		Set m_dict = Nothing
	End Sub

End Class 'TaniumNamedArgsParser
' :::VBLib:TaniumNamedArgsParser:End:::


Function x64Fix
' This is a function which should be called before calling any vbscript run by 
' the Tanium client that needs 64-bit registry or filesystem access.
' It's for when we need to catch if a machine has 64-bit windows
' and is running in a 32-bit environment.
'  
' In this case, we will re-launch the sensor in 64-bit mode.
' If it's already in 64-bit mode on a 64-bit OS, it does nothing and the sensor 
' continues on
    
    Const WINDOWSDIR = 0
    Const HKLM = &h80000002
    
    Dim objShell: Set objShell = CreateObject("WScript.Shell")
    Dim objFSO: Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objSysEnv: Set objSysEnv = objShell.Environment("PROCESS")
    Dim objReg, objArgs, objExec
    Dim strOriginalArgs, strArg, strX64cscriptPath, strMkLink
    Dim strProgramFilesX86, strProgramFiles, strLaunchCommand
    Dim strKeyPath, strTaniumPath, strWinDir
    Dim b32BitInX64OS

    b32BitInX64OS = false

    ' we'll need these program files strings to check if we're in a 32-bit environment
    ' on a pre-vista 64-bit OS (if no sysnative alias functionality) later
    strProgramFiles = objSysEnv("ProgramFiles")
    strProgramFilesX86 = objSysEnv("ProgramFiles(x86)")
    ' tLog.log "Are the program files the same?: " & (LCase(strProgramFiles) = LCase(strProgramFilesX86))
    
    ' The windows directory is retrieved this way:
    strWinDir = objFso.GetSpecialFolder(WINDOWSDIR)
    'tLog.log "Windir: " & strWinDir
    
    ' Now we determine a cscript path for 64-bit windows that works every time
    ' The trick is that for x64 XP and 2003, there's no sysnative to use.
    ' The workaround is to do an NTFS junction point that points to the
    ' c:\Windows\System32 folder.  Then we call 64-bit cscript from there.
    ' However, there is a hotfix for 2003 x64 and XP x64 which will enable
    ' the sysnative functionality.  The customer must either have linkd.exe
    ' from the 2003 resource kit, or the hotfix installed.  Both are freely available.
    ' The hotfix URL is http://support.microsoft.com/kb/942589
    ' The URL For the resource kit is http://www.microsoft.com/download/en/details.aspx?id=17657
    ' linkd.exe is the only required tool and must be in the machine's global path.

    If objFSO.FileExists(strWinDir & "\sysnative\cscript.exe") Then
        strX64cscriptPath = strWinDir & "\sysnative\cscript.exe"
        ' tLog.log "Sysnative alias works, we're 32-bit mode on 64-bit vista+ or 2003/xp with hotfix"
        ' This is the easy case with sysnative
        b32BitInX64OS = True
    End If
    If Not b32BitInX64OS And objFSO.FolderExists(strWinDir & "\SysWow64") And (LCase(strProgramFiles) = LCase(strProgramFilesX86)) Then
        ' This is the more difficult case to execute.  We need to test if we're using
        ' 64-bit windows 2003 or XP but we're running in a 32-bit mode.
        ' Only then should we relaunch with the 64-bit cscript.
        
        ' If we don't accurately test 32-bit environment in 64-bit OS
        ' This code will call itself over and over forever.
        
        ' We will test for this case by checking whether %programfiles% is equal to
        ' %programfiles(x86)% - something that's only true in 64-bit windows while
        ' in a 32-bit environment
    
        ' tLog.log "We are in 32-bit mode on a 64-bit machine"
        ' linkd.exe (from 2003 resource kit) must be in the machine's path.
        
        strMkLink = "linkd " & Chr(34) & strWinDir & "\System64" & Chr(34) & " " & Chr(34) & strWinDir & "\System32" & Chr(34)
        strX64cscriptPath = strWinDir & "\System64\cscript.exe"
        ' tLog.log "Link Command is: " & strMkLink
        ' tLog.log "And the path to cscript is now: " & strX64cscriptPath
        On Error Resume Next ' the mklink command could fail if linkd is not in the path
        ' the safest place to put linkd.exe is in the resource kit directory
        ' reskit installer adds to path automatically
        ' or in c:\Windows if you want to distribute just that tool
        
        If Not objFSO.FileExists(strX64cscriptPath) Then
            ' tLog.log "Running mklink" 
            ' without the wait to completion, the next line fails.
            objShell.Run strMkLink, 0, true
        End If
        On Error GoTo 0 ' turn error handling off
        If Not objFSO.FileExists(strX64cscriptPath) Then
            ' if that cscript doesn't exist, the link creation didn't work
            ' and we must quit the function now to avoid a loop situation
            ' tLog.log "Cannot find " & strX64cscriptPath & " so we must exit this function and continue on"
            ' clean up
            Set objShell = Nothing
            Set objFSO = Nothing
            Set objSysEnv = Nothing
            Exit Function
        Else
            ' the junction worked, it's safe to relaunch            
            b32BitInX64OS = True
        End If
    End If
    If Not b32BitInX64OS Then
        ' clean up and leave function, we must already be in a 32-bit environment
        Set objShell = Nothing
        Set objFSO = Nothing
        Set objSysEnv = Nothing
        
        ' tLog.log "Cannot relaunch in 64-bit (perhaps already there)"
        ' important: If we're here because the client is broken, a sensor will
        ' run but potentially return incomplete or no values (old behavior)
        Exit Function
    End If
    
    ' So if we're here, we need to re-launch with 64-bit cscript.
    ' take the arguments to the sensor and re-pass them to itself in a 64-bit environment
    strOriginalArgs = ""
    Set objArgs = WScript.Arguments
    
    For Each strArg in objArgs
        strOriginalArgs = strOriginalArgs & " " & strArg
    Next
    ' after we're done, we have an unnecessary space in front of strOriginalArgs
    strOriginalArgs = LTrim(strOriginalArgs)
    
    ' If this is running as a sensor, we need to know the path of the tanium client
    strKeyPath = "Software\Tanium\Tanium Client"
    Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
    
    objReg.GetStringValue HKLM,strKeyPath,"Path", strTaniumPath

    ' tLog.log "StrOriginalArgs is:" & strOriginalArgs
    If objFSO.FileExists(Wscript.ScriptFullName) Then
        strLaunchCommand = Chr(34) & Wscript.ScriptFullName & Chr(34) & " " & strOriginalArgs
        ' tLog.log "Script full path is: " & WScript.ScriptFullName
    Else
        ' the sensor itself will not work with ScriptFullName so we do this
        strLaunchCommand = Chr(34) & strTaniumPath & "\VB\" & WScript.ScriptName & chr(34) & " " & strOriginalArgs
    End If
    ' tLog.log "launch command is: " & strLaunchCommand

    ' Note:  There is a timeout limit here of 1 hour, as extra protection for runaway processes
    Set objExec = objShell.Exec(strX64cscriptPath & " //T:3600 " & strLaunchCommand)
    
    ' skipping the two lines and space after that look like
    ' Microsoft (R) Windows Script Host Version
    ' Copyright (C) Microsoft Corporation
    '
    objExec.StdOut.SkipLine
    objExec.StdOut.SkipLine
    objExec.StdOut.SkipLine

    ' sensor output is all about stdout, so catch the stdout of the relaunched
    ' sensor
    WScript.Echo objExec.StdOut.ReadAll()
    
    ' critical - If we've relaunched, we must quit the script before anything else happens
    WScript.Quit
    ' Remember to call this function only at the very top
    
    ' Cleanup
    Set objReg = Nothing
    Set objArgs = Nothing
    Set objExec = Nothing
    Set objShell = Nothing
    Set objFSO = Nothing
    Set objSysEnv = Nothing
    Set objReg = Nothing
End Function 'x64Fix


Function UseScanSourceArgValidator(arg)
	UseScanSourceArgValidator = True
	
	If Not StringInCommaSeparatedList(arg, _
		"cab,systemdefault,internet,wsus,optimal" ) Then
		UseScanSourceArgValidator = False
	End If

End Function 'UseScanSourceArgValidator


Function StringInCommaSeparatedList(strToCheck,strCommaSeparated)
	StringInCommaSeparatedList = True
	
	Dim arrValidInput,strValid
	arrValidInput = Split(strCommaSeparated,",")
	strToCheck = LCase(strToCheck)
	Dim bInList
	bInList = False
	For Each strValid In arrValidInput
		If strToCheck = strValid Then
			bInList = True
		End If
	Next
	
	If Not bInList Then
		StringInCommaSeparatedList = False
	End If

End Function 'StringInCommaSeparatedList

Function NeverSupersedeListTranslator(strNeverSupersedeList)
	' Microsoft has a supersedence issue with certain updates
	' which block the application of windows 8.1 updates
	' Whatever is passed in for the never superseded list, add these
	' in as well

	Dim strNeverSupersedeDefaults
	strNeverSupersedeDefaults = "KB2919442,KB2969339"
	strNeverSupersedeList = LCase(strNeverSupersedeList)
	strNeverSupersedeDefaults = LCase(strNeverSupersedeDefaults)
	
	Dim arrList,strItem,strOut
	arrList = Split(strNeverSupersedeList,",")
	strOut = strNeverSupersedeDefaults
	For Each strItem In arrList
		' Some items are, by default, never superseded. Do not list them twice
		' if they are passed in
		If Not InStr(strNeverSupersedeDefaults,strItem) > 0 Then
			strOut = strOut&","&strItem
		End If
	Next

	NeverSupersedeListTranslator = strOut
End Function 'NeverSupersedeListTranslator

Sub ParseArgs(ByRef ArgsParser)

	' Pre- and Post- directory locations
	Dim objFSO
	Dim strPrePostPrefix,strFileDir,strExtension,strFolderName
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strFileDir = Replace(WScript.ScriptFullName,WScript.ScriptName,"")
	strExtension = objFSO.GetExtensionName(WScript.ScriptFullName)
	strFolderName = Replace(WScript.ScriptName,"."&strExtension,"")
	strPrePostPrefix = strFileDir&strFolderName
	
	ArgsParser.ProgramDescription = "Performs a Windows Patch install operation. Typically triggered as a " _ 
		& "Tanium Action. Will Log output to the ContentLogs subfolder of the Tools folder. Note: All " _
		& "command line arguments default to 'sticky' - they are stored in the registry and retrieved on " _
		& "subsequent calls to the script. Integrates with Tanium Maintenance Window content, and can run scripts in the " _
		& strPrePostPrefix&"\Pre"&Chr(34)&" and "&Chr(34)&strPrePostPrefix&"\Post"&Chr(34)&" directories before and after " _
		& "execution, respectively."

	Dim objWaitTimeArg
	Set objWaitTimeArg = New TaniumNamedArg
	objWaitTimeArg.RequireInteger = True
	objWaitTimeArg.ArgName = "RandomWaitTimeInSeconds"
	objWaitTimeArg.HelpText = "Waits up to X seconds before scanning"
	objWaitTimeArg.ExampleValue = "15"
	objWaitTimeArg.DefaultValue = 0
	objWaitTimeArg.IsOptional = True
	ArgsParser.AddArg objWaitTimeArg
	
	Dim objOnlineScanRandomWaitTimeArg
	Set objOnlineScanRandomWaitTimeArg = New TaniumNamedArg
	objOnlineScanRandomWaitTimeArg.RequireInteger = True
	objOnlineScanRandomWaitTimeArg.ArgName = "OnlineScanRandomWaitTimeInSeconds"
	objOnlineScanRandomWaitTimeArg.HelpText = "Waits up to X seconds before scanning if scan is online based"
	objOnlineScanRandomWaitTimeArg.ExampleValue = "300"
	objOnlineScanRandomWaitTimeArg.DefaultValue = 300
	objOnlineScanRandomWaitTimeArg.IsOptional = True	
	ArgsParser.AddArg objOnlineScanRandomWaitTimeArg
	
	Dim objSupersededUpdatesDaysOldThresholdArg
	Set objSupersededUpdatesDaysOldThresholdArg = New TaniumNamedArg
	objSupersededUpdatesDaysOldThresholdArg.RequireInteger = True
	objSupersededUpdatesDaysOldThresholdArg.ArgName = "ShowSupersededUpdatesPublishedDaysOld"
	objSupersededUpdatesDaysOldThresholdArg.HelpText = "If update is superseded in the last X days, show the update anyway"
	objSupersededUpdatesDaysOldThresholdArg.ExampleValue = "30"
	objSupersededUpdatesDaysOldThresholdArg.IsOptional = True
	ArgsParser.AddArg objSupersededUpdatesDaysOldThresholdArg
	
	Dim objClearHistoryFlagArg
	Set objClearHistoryFlagArg = New TaniumNamedArg
	objClearHistoryFlagArg.RequireYesNoTrueFalse = True
	objClearHistoryFlagArg.ArgName = "ClearHistoryOnBadLine"
	objClearHistoryFlagArg.HelpText = "If history file has a bad line, clear it"
	objClearHistoryFlagArg.ExampleValue = "Yes,No"
	objClearHistoryFlagArg.DefaultValue = "No"
	objClearHistoryFlagArg.IsOptional = True
	ArgsParser.AddArg objClearHistoryFlagArg
	
	Dim objPrintSupersedenceInfoArg
	Set objPrintSupersedenceInfoArg = New TaniumNamedArg
	objPrintSupersedenceInfoArg.RequireYesNoTrueFalse = True
	objPrintSupersedenceInfoArg.ArgName = "PrintSupersedenceInfo"
	objPrintSupersedenceInfoArg.HelpText = "Print the Supersedence Tree"
	objPrintSupersedenceInfoArg.ExampleValue = "Yes,No"
	objPrintSupersedenceInfoArg.DefaultValue = "No"
	objPrintSupersedenceInfoArg.IsOptional = True
	ArgsParser.AddArg objPrintSupersedenceInfoArg

	Dim objDoNotSaveOptionsArg
	Set objDoNotSaveOptionsArg = New TaniumNamedArg
	objDoNotSaveOptionsArg.RequireYesNoTrueFalse = True
	objDoNotSaveOptionsArg.ArgName = "DoNotSaveOptions"
	objDoNotSaveOptionsArg.HelpText = "Do not save the command line arguments to the registry for next use"
	objDoNotSaveOptionsArg.ExampleValue = "Yes,No"
	objDoNotSaveOptionsArg.DefaultValue = "No"
	objDoNotSaveOptionsArg.IsOptional = True
	ArgsParser.AddArg objDoNotSaveOptionsArg
	
	Dim objDisableMicrosoftUpdateArg
	Set objDisableMicrosoftUpdateArg = New TaniumNamedArg
	objDisableMicrosoftUpdateArg.RequireYesNoTrueFalse = True
	objDisableMicrosoftUpdateArg.ArgName = "DisableMicrosoftUpdate"
	objDisableMicrosoftUpdateArg.HelpText = "Disables Microsoft Update additional scan info for online (non-cab based) scans"
	objDisableMicrosoftUpdateArg.ExampleValue = "Yes,No"
	objDisableMicrosoftUpdateArg.DefaultValue = "No"
	objDisableMicrosoftUpdateArg.IsOptional = True
	ArgsParser.AddArg objDisableMicrosoftUpdateArg
	
	Dim UseScanSourceValidateRef,objUseScanSourceArg
	Set UseScanSourceValidateRef = GetRef("UseScanSourceArgValidator")
	Set objUseScanSourceArg = New TaniumNamedArg
	objUseScanSourceArg.ArgName = "UseScanSource"
	objUseScanSourceArg.HelpText = "Changes scan source from cab file to service based. Only cab will scan not using a service. Take care using any service based scans from endpoints, who may all scan a single service. systemdefault looks at how windows is configured via group policy. wsus uses any group policy configured wsus server. Internet uses windows update (without microsoft update). Optimal uses windows update with Microsoft update, with fallback to the .cab file."
	objUseScanSourceArg.ExampleValue = "cab,systemdefault,internet,wsus,optimal"
	objUseScanSourceArg.DefaultValue = "cab"
	objUseScanSourceArg.IsOptional = True
	objUseScanSourceArg.ValidationFunctionReference = UseScanSourceValidateRef
	ArgsParser.AddArg objUseScanSourceArg
	
	Dim objConsiderSupersededUpdatesArg
	Set objConsiderSupersededUpdatesArg = New TaniumNamedArg
	objConsiderSupersededUpdatesArg.RequireYesNoTrueFalse = True
	objConsiderSupersededUpdatesArg.ArgName = "ConsiderSupersededUpdates"
	objConsiderSupersededUpdatesArg.HelpText = "Control consideration of superseded updates. Default is True. Supersedence is handled carefully and automatically"
	objConsiderSupersededUpdatesArg.ExampleValue = "Yes,No"
	objConsiderSupersededUpdatesArg.DefaultValue = "yes"
	objConsiderSupersededUpdatesArg.IsOptional = True
	ArgsParser.AddArg objConsiderSupersededUpdatesArg
	
	Dim objNeverSupersedeListArg,NeverSupersedeListTranslatorRef
	Set NeverSupersedeListTranslatorRef = GetRef("NeverSupersedeListTranslator")
	Set objNeverSupersedeListArg = New TaniumNamedArg
	objNeverSupersedeListArg.ArgName = "NeverSupersedeList"
	objNeverSupersedeListArg.HelpText = "Always consider updates in this list to be non-superseded"
	objNeverSupersedeListArg.ExampleValue = "KB292222,KB930284"
	objNeverSupersedeListArg.DefaultValue = "KB2919442,KB2969339" ' Translate adds values to what is passed in
	objNeverSupersedeListArg.TranslationFunctionReference = NeverSupersedeListTranslatorRef
	objNeverSupersedeListArg.IsOptional = True
	ArgsParser.AddArg objNeverSupersedeListArg

	Dim objShowTimingsArg
	Set objShowTimingsArg = New TaniumNamedArg
	objShowTimingsArg.RequireYesNoTrueFalse = True
	objShowTimingsArg.ArgName = "ShowTimings"
	objShowTimingsArg.HelpText = "Display additional timing information in output"
	objShowTimingsArg.ExampleValue = "Yes,No"
	objShowTimingsArg.DefaultValue = "no"
	objShowTimingsArg.IsOptional = True
	ArgsParser.AddArg objShowTimingsArg
	
	ArgsParser.Parse
	' The arguments should be successfully parsed, and handling of the arguments
	' is performed elsewhere in the script
	If ArgsParser.ErrorState Then
		ArgsParser.PrintUsageAndQuit ""
	End If
End Sub 'ParseArgs

Sub MakeSticky(ByRef ArgsParser, ByRef tContentReg)
	' Makes arguments persist in the registry
	
	' Loop through each passed in argument and write string value
	Dim argsArr
	argsArr = ArgsParser.AddedArgsArray
	Dim objParsedArg,objCLIArg
	For Each objParsedArg In argsArr
		If WScript.Arguments.Named.Exists(objParsedArg.ArgName) Then
			tLog.Log "Making arg '" & objParsedArg.ArgName & "' with value '" _
				& objParsedArg.ArgValue & "' sticky by updating registry"
			tContentReg.ErrorClear
			tContentReg.RegValueType = "REG_SZ"
			tContentReg.ValueName = objParsedArg.ArgName
			tContentReg.Data = CStr(objParsedArg.ArgValue) ' all patch management values are string
			tContentReg.Write ' Will actually Raise an error
			On Error Resume Next
			If tContentReg.ErrorState Then
				tLog.Log "Could not write data for argument " & objParsedArg.ArgName _
					& " into registry"
				tContentReg.ErrorClear
			End If
			On Error Goto 0
		End If
	Next
End Sub 'MakeSticky

Sub LoadDefaultConfig(ByRef ArgsParser, ByRef dictPatchManagementConfig)
	' for each argument parsed, there is a default value
	' load these default values into the config dictionary
	' After this load, default values are stomped by the read of config from
	' the Registry
	Dim objArg
	For Each objArg In ArgsParser.AddedArgsArray
		If objArg.DefaultValue <> "" Then
			If Not dictPatchManagementConfig.Exists(objArg.ArgName) Then
				On Error Resume Next ' some argvalue sets will raise
				dictPatchManagementConfig.Add objArg.ArgName,objArg.ArgValue
				If Err.Number <> 0 Then
					tLog.Log = "Error: Could not set argument "&m_dict.Item(internalArg).ArgName&" value to it's default value "& m_dict.Item(internalArg).DefaultValue&" - " &Err.Description
					Err.Clear
				End If
				On Error Goto 0
			End If
		End If
	Next

End Sub 'LoadDefaultConfig

Function TryFromDict(ByRef dict,key,ByRef fallbackValue)
	' Pulls from a dictionary if possible, falls back to whatever
	' is specified.
		If dict.Exists(key) Then
			If IsObject(dict.Item(key)) Then
				Set TryFromDict = dict.Item(key)
			Else
				TryFromDict = dict.Item(key)
			End If
		Else
			If IsObject(fallbackValue) Then
				Set TryFromDict = fallbackValue
			Else
				TryFromDict = fallbackValue
			End If
		End If
End Function 'TryFromDict


Sub LoadParsedConfig(ByRef ArgsParser, ByRef dictPatchManagementConfig)
	' In case the arguments were not 'made sticky' by writing to registry, we must
	' read the parsed arguments into the config dictionary
	' This should happen after the load of default config
	Dim objArg
	For Each objArg In ArgsParser.AddedArgsArray
		If objArg.ArgValue <> "" Then
			If Not dictPatchManagementConfig.Exists(objArg.ArgName) Then
				On Error Resume Next ' some argvalue sets will raise
				dictPatchManagementConfig.Add objArg.ArgName,objArg.ArgValue
				If Err.Number <> 0 Then
					tLog.Log = "Error: Could not set argument "&m_dict.Item(internalArg).ArgName&" value to it's parsed value "& m_dict.Item(internalArg).ArgValue&" - " &Err.Description
					Err.Clear
				End If
				On Error Goto 0
			Else
				On Error Resume Next ' some argvalue sets will raise
				If Not objArg.ArgValue = objArg.DefaultValue Then
					dictRegConfig.Item(objArg.ArgName) = objArg.ArgValue
				End If
				If Err.Number <> 0 Then
					WScript.Echo = "Error: Could not set argument "&m_dict.Item(internalArg).ArgName&" value to it's parsed value "& m_dict.Item(internalArg).ArgValue&" - " &Err.Description
					Err.Clear
				End If
				On Error Goto 0
			End If
		End If
	Next
End Sub 'LoadParsedConfig


Sub LoadRegConfig(ByRef tContentReg, ByRef dictPatchManagementConfig)
	' Loop through the registry and create config dictionary
	' only String types are added to config dict
	' ultimately, config dict is Value Name, String Data
	tContentReg.ErrorClear
	Dim dictVals, strValName, dictRet
	Set dictVals = tContentReg.ValuesDict
	For Each strValName In dictVals
		If Not strValName = "" Then
			If dictVals.Item(strValName) = "REG_SZ" Then ' consider only these	
				tContentReg.ValueName = strValName
				tContentReg.RegValueType = "REG_SZ"
				If Not dictPatchManagementConfig.Exists(strValName) Then
					dictPatchManagementConfig.Add strValName,tContentReg.Read
				Else
					dictPatchManagementConfig.Item(strValName) = tContentReg.Read
				End If
			End If
		End If
	Next
End Sub 'LoadRegConfig


Sub EchoConfig(ByRef dictPConfig)
	Dim strKey
	tLog.Log "Patch Management Config (Registry and / or default, and parsed values)"
	For Each strKey In dictPConfig
		tLog.Log strKey &" = "& dictPConfig.Item(strKey)
	Next
End Sub 'EchoConfig


' :::VBLib:TaniumContentLog:Begin:::
Class TaniumContentLog
	Private m_strLogDirectory
	Private m_intMaxDaysToKeep
	Private m_intMaxLogsToKeep
	Private m_libVersion
	Private m_libName
	Private m_bErr
	Private m_errMessage
	Private m_strLogFileDir
	Private m_strLogFileName
	Private m_strLogFilePath
	Private m_objLogTextFile
	Private m_strRFC822Bias
	Private m_strLogSep
	Private m_strLogSepReplacementText
	Private m_objShell
	Private m_objFSO
	Private LOGFILEFORAPPENDING
	Private m_defaultLogFileDir
	Private m_defaultLogFileName

	Private Sub Class_Initialize
		m_libVersion = "6.5.314.4217"
		m_libName = "TaniumContentLog"
		Set m_objShell = CreateObject("WScript.Shell")
		Set m_objFSO = CreateObject("Scripting.FileSystemObject")
		LOGFILEFORAPPENDING = 8
		m_intMaxDaysToKeep = 180
		m_intMaxLogsToKeep = 5
		m_strLogSep = "|"
		m_strLogSepReplacementText = "<pipechar>"
		m_defaultLogFileDir = VBLibGetTaniumDir("Tools\Content Logs")
		m_defaultLogFileName = WScript.ScriptName&".log"
		LogRotateCheck m_defaultLogFileDir,m_defaultLogFileName
		SetupLogFileDirAndName m_defaultLogFileDir,m_defaultLogFileName
		GetRFC822Bias
    End Sub

	Private Sub Class_Terminate
		'Set m_objShell = Nothing
		'Set m_objFSO = Nothing
		On Error Resume Next
		'm_objLogTextFile.Close()
		On Error Goto 0
		'Set m_objLogTextFile = Nothing
	End Sub

	Public Property Let LogFieldSeparator(strSep)
		m_strLogSep = strSep
	End Property

	Public Property Let LogFieldSeparatorReplacementString(strSep)
		' this is the text that is inserted into the string being logged if the actual
		' separator character is found in the string. Defualt is <pipechar>
		m_strLogSepReplacementText = strSep
	End Property

	Public Property Let MaxDaysToKeep(intDays)
		' Ensure this is integer
		m_intMaxDaysToKeep = GetInteger(intDays)
		ErrorCheck
	End Property

	Public Property Let MaxLogFilesToKeep(intLogFiles)
		' Ensure this is integer
		m_intMaxLogsToKeep = GetInteger(intLogFiles)
		ErrorCheck
	End Property

    Public Property Get LibVersion
    	LibVersion = m_libVersion
    End Property

    Public Property Get ErrorState
    	ErrorState = m_bErr
    End Property

    Public Property Get ErrorMessage
    	ErrorMessage = m_errMessage
    End Property

	Private Sub ErrorCheck
		' Call on all Lets
		If m_bErr Then
			Err.Raise vbObjectError + 1978, m_libName, m_errMessage
		End If
	End Sub 'ErrorCheck

	Public Sub ErrorClear
		m_bErr = False
		m_errMessage = ""
	End Sub

    Public Property Get LogFileName
    	' There is no corresponding Let, this is read-only
    	LogFileName = m_strLogFileName
    End Property

    Public Property Get LogFileDir
    	' There is no corresponding Let, this is read-only
    	LogFileDir = m_strLogFileDir
    End Property
    
	Private Function UnicodeToAscii(ByRef pStr)
		Dim x,conv,strOut
		For x = 1 To Len(pStr)
			conv = Mid(pstr,x,1)
			conv = Asc(conv)
			conv = Chr(conv)
			strOut = strOut & conv
		Next
		UnicodeToAscii = strOut
	End Function 'UnicodeToAscii

	Public Sub Log(strText)
	' This function writes a timestamp and a string to a log file
	' whose object (objTextFile with FORAPPENDING on) is passed in
	' this way, the function writes an already-open file without
	' closing it over and over.
	' make sure to include all support functions and close the file
	' when done
	' and then call the logrotator function

		If Not VarType(strText) = vbString Then
			m_bErr = True
			m_errMessage = "Error: Cannot log, string to log is not a string"
			ErrorCheck
			Exit Sub
		End If
		' Temporarily not writing to unicode files, recognize a better solution
		strText = UnicodeToAscii(strText)
		WScript.Echo strText

		'log fields are separated by the | character
		'so strings passed in must have the pipe character replaced
		strText = Replace(strText,m_strLogSep,m_strLogSepReplacementText)
		On Error Resume Next
		m_objLogTextFile.WriteLine(vbTimeToRFC822(Now(),m_strRFC822Bias)&"|"&strText)
		If Err.Number <> 0 Then
			m_bErr = True
        	m_errMessage = "Content Log: Text was unable to be written: " & Err.Description & " - text variable type is: " & VarType(strText)
        End If
        On Error Goto 0
		ErrorCheck
	End Sub 'ContentLog


' Consider calling this function inside init?
' instead of doing a delete
' where to rotate?
	Public Default Function InitWithArgs(strLogFileDir,strLogFileName)
	' Deliberately make it very non-obvious how to change
	' the log path and directory. This should almost never
	' be changed from the defaults
	' to do this, use the following syntax:
	' Set x = New(TaniumContentLog)(<your_dir>,<your_filename>)
		LogRotateCheck VBLibGetTaniumDir(strLogFileDir),strLogFileName
		SetupLogFileDirAndName VBLibGetTaniumDir(strLogFileDir),strLogFileName
		' now delete the old log location
		Dim strPathToDelete
		strPathToDelete = m_objFSO.BuildPath(m_defaultLogFileDir,m_defaultLogFileName)
		On Error Resume Next
		m_objFSO.DeleteFile strPathToDelete, True
		If Err.Number <> 0 Then
			WScript.Echo "Warning: Overrode log path, but could not delete default log file location"
		End If
		On Error Goto 0
		Set InitWithArgs = Me
	End Function

	Private Sub SetupLogFileDirAndName(strDir,strFileName)
		' Tries to create the log file directory
		m_strLogFilePath = m_objFSO.BuildPath(strDir,strFileName)
		If Not m_objFSO.FolderExists(strDir) Then
			On Error Resume Next
			m_objFSO.CreateFolder strDir
			If Err.Number <> 0 Then
				m_bErr = True
				m_errMessage = m_libName& " Error: Could not create log file directory: " & strDir&", " & Err.Description
			End If
			On Error Goto 0
		End If

		On Error Resume Next
		Set m_objLogTextFile = m_objFSO.OpenTextFile(m_strLogFilePath,LOGFILEFORAPPENDING,True)
		If Err.Number <> 0 Then
			m_bErr = True
			m_errMessage = m_libName& " Error: Could not open or create log file " & m_strLogFilePath&", " & Err.Description
		End If
		On Error Goto 0
		ErrorCheck
	End Sub 'SetupLogFileDirAndName

	Private Function vbTimeToRFC822(myDate, offset)
	' must be set so that month is displayed with US/English abbreviations
	' as per the standard
		Dim intOldLocale
		intOldLocale = GetLocale()
		SetLocale 1033 'Require month prefixes to be us/english

		Dim myDay, myDays, myMonth, myYear
		Dim myHours, myMinutes, myMonths, mySeconds

		myDate = CDate(myDate)
		myDay = WeekdayName(Weekday(myDate),true)
		myDays = zeroPad(Day(myDate), 2)
		myMonth = MonthName(Month(myDate), true)
		myYear = Year(myDate)
		myHours = zeroPad(Hour(myDate), 2)
		myMinutes = zeroPad(Minute(myDate), 2)
		mySeconds = zeroPad(Second(myDate), 2)

		vbTimeToRFC822 = myDay&", "& _
		                              myDays&" "& _
		                              myMonth&" "& _
		                              myYear&" "& _
		                              myHours&":"& _
		                              myMinutes&":"& _
		                              mySeconds&" "& _
		                              offset
		SetLocale intOldLocale
	End Function 'vbTimeToRFC822

	Private Function VBLibGetTaniumDir(strSubDir)
	'GetTaniumDir with GeneratePath, works in x64 or x32
	'looks for a valid Path value
	'for use inside VBLib classes

		Dim keyNativePath, keyWoWPath, strPath

		keyNativePath = "HKLM\Software\Tanium\Tanium Client"
		keyWoWPath = "HKLM\Software\Wow6432Node\Tanium\Tanium Client"

	    ' first check the Software key (valid for 32-bit machines, or 64-bit machines in 32-bit mode)
	    On Error Resume Next
	    strPath = m_objShell.RegRead(keyNativePath&"\Path")
	    On Error Goto 0

	  	If strPath = "" Then
	  		' Could not find 32-bit mode path, checking Wow6432Node
	  		On Error Resume Next
	  		strPath = m_objShell.RegRead(keyWoWPath&"\Path")
	  		On Error Goto 0
	  	End If

	  	If Not strPath = "" Then
			If strSubDir <> "" Then
				strSubDir = "\" & strSubDir
			End If

			If m_objFSO.FolderExists(strPath) Then
				If Not m_objFSO.FolderExists(strPath & strSubDir) Then
					''Need to loop through strSubDir and create all sub directories
					GeneratePath strPath & strSubDir, m_objFSO
				End If
				VBLibGetTaniumDir = strPath & strSubDir & "\"
			Else
				' Specified Path doesn't exist on the filesystem
				m_errMessage = "Error: " & strPath & " does not exist on the filesystem"
				m_bErr = True
			End If
		Else
			m_errMessage = "Error: Cannot find Tanium Client path in Registry"
			m_bErr = False
		End If
	End Function 'VBLibGetTaniumDir

	Private Function GeneratePath(pFolderPath, fso)
		GeneratePath = False
		If Not fso.FolderExists(pFolderPath) Then
			If GeneratePath(fso.GetParentFolderName(pFolderPath), fso) Then
				GeneratePath = True
				Call fso.CreateFolder(pFolderPath)
			End If
		Else
			GeneratePath = True
		End If
	End Function 'GeneratePath

	Private Function GetInteger(value)
		' If value is an integer (or a string that can be an integer), store it
		' default case is to not accept value
		GetInteger = False
		' first character could be a dollar sign which is convertible
		' this is the case which occurs when a tanium command line has an invalid parameter spec
		Dim intDollar
		intDollar = InStr(value,"$")
		If intDollar > 0 And Len(value) > 1 Then ' only if more than one char
			value = Right(value,Len(value) - 1)
		End If
		If VarType(value) = vbString Then
			If Not IsNumeric(value) Then
				m_bErr = True
				m_ErrMessage = m_libName& " Error: " & value & " could not be converted to a number."
			End If
			Dim conv
			On Error Resume Next
			conv = CStr(CLng(value))
			If Err.Number <> 0 Then
				m_bErr = True
				m_ErrMessage = m_libName & " Error: " & value & " could not be converted to an integer. - max size is +/-2,147,483,647. ("&Err.Description&")"
			End If
			On Error Goto 0
			If conv = value Then
				GetInteger = CLng(value)
			End If
		ElseIf VarType(value) = vbLong Or VarType(value) = vbInteger Then
			GetInteger = CLng(value)
		Else
		 ' some non-string, non-numeric value
			m_bErr = True
			m_ErrMessage = m_libName & " Error: argument could not be converted to an integer, was type "&TypeName(value)
		End If
		ErrorCheck
	End Function 'GetInteger

	Private Sub LogRotateCheck(strLogFileDir,strLogFileName)
	' This function will rotate log files
	' the function takes days to keep and max number of files

	' Logs will rotate when the currently written log file
	' is max days old / intMaxFiles days old

	' example: max days is 180 and max files is 5
	' if the current log file is 36 days old
	' then rotate it.  Each log file contains 36 days of data

	' rotating it means renaming it to filename.log.0.log
	' where the digit between the dots is 0->maxFiles

		Dim strLogToRotateFilePath,objLogFile
		Dim dtmLogFileCreationDate,intLogFileDaysOld
		Dim strLogFileExtension,strCheckFilePath,i

		strLogToRotateFilePath = m_objFSO.BuildPath(strLogFileDir,strLogFileName)
		If Not m_objFSO.FileExists(strLogToRotateFilePath) Then
			Exit Sub
		End If
		Set objLogFile = m_objFSO.GetFile(strLogToRotateFilePath)

		dtmLogFileCreationDate = objLogFile.DateCreated
		intLogFileDaysOld = Round(Abs(DateDiff("s",Now(),dtmLogFileCreationDate)) / 86400,0)
		If Now() - dtmLogFileCreationDate > m_intMaxDaysToKeep / m_intMaxLogsToKeep Then 'rotate time
			WScript.Echo "Rotating Content Log File " & strLogToRotateFilePath
			strLogFileExtension = m_objFSO.GetExtensionName(strLogToRotateFilePath)

			' in case of file name collision, which shouldn't happen, we will append a date stamp
			If (m_objFSO.FileExists(strLogToRotateFilePath)) Then
				For i = m_intMaxLogsToKeep To 0 Step -1
					' rotated log file looks like m_strLogFilePath.0.<extension>
					strCheckFilePath = strLogToRotateFilePath&"."&m_intMaxLogsToKeep&"."&strLogFileExtension
					If m_objFSO.FileExists(strCheckFilePath) Then
						On Error Resume Next
						m_objFSO.DeleteFile strCheckFilePath,True ' force
						If Err.Number <> 0 Then
							WScript.Echo "Error: Could not delete " & strCheckFilePath
						End If
						On Error Goto 0
					Else ' start rotating
						strCheckFilePath = strLogToRotateFilePath&"."&i&"."&strLogFileExtension
						If m_objFSO.FileExists(strCheckFilePath) Then
							On Error Resume Next
							m_objFSO.DeleteFile	strLogToRotateFilePath&"."&i+1&"."&strLogFileExtension, True
							Err.Clear
							m_objFSO.MoveFile strCheckFilePath, strLogToRotateFilePath&"."&i+1&"."&strLogFileExtension
							' log.4.log now moves to log.5.log
							If Err.Number <> 0 Then
								WScript.Echo "Error: Could not move check file "&strCheckFilePath&", "&Err.Description
							End If
							On Error Goto 0
						End If
					End If
				Next
			' finally, we have a clear spot - there should be no log.1.log
			On Error Resume Next
			m_objFSO.DeleteFile strLogToRotateFilePath&".1."&strLogFileExtension, True
			Err.Clear
			m_objFSO.MoveFile strLogToRotateFilePath,strLogToRotateFilePath&".1."&strLogFileExtension
			If Err.Number <> 0 Then
				WScript.Echo "Error: Could not move log to rotate " & strLogToRotateFilePath&", " & Err.Description
			End If
			On Error Goto 0
			Else
				' Consider doing m_bErr and raising error, but this should not block
				WScript.Echo "Error: Log Rotator cannot find log file " & strLogToRotateFilePath
			End If
		End If
	End Sub 'LogRotator

	Private Sub GetRFC822Bias
	' This function returns a string which is a
	' timezone bias for RFC822 format
	' considers daylight savings
	' we choose 4 digits and a sign (+ or -)

		Dim objWMIService,colTimeZone,objTimeZone

		Dim intTZBiasInMinutes,intTZBiasMMHH,strSign,strReturnString

		Set objWMIService = GetObject("winmgmts:" _
		    & "{impersonationLevel=impersonate}!\\.\root\cimv2")
		Set colTimeZone = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")

		For Each objTimeZone in colTimeZone
		    intTZBiasInMinutes = objTimeZone.CurrentTimeZone
		Next

		' The offset is explicitly signed
		If intTZBiasInMinutes < 0 Then
			strSign = "-"
		Else
			strSign = "+"
		End If

		intTZBiasMMHH = Abs(intTZBiasInMinutes)
		intTZBiasMMHH = zeroPad(CStr(Int(CInt(intTZBiasMMHH)/60)),2) _
			&zeroPad(CStr(intTZBiasMMHH Mod 60),2)
		m_strRFC822Bias = strSign&intTZBiasMMHH

		'Cleanup
		Set colTimeZone = Nothing
		Set objWMIService = Nothing

	End Sub 'GetRFC822Bias

	Private Function zeroPad(m, t)
	   zeroPad = String(t-Len(m),"0")&m
	End Function 'zeroPad

End Class 'TaniumContentLog
' :::VBLib:TaniumContentLog:End:::

' :::VBLib:TaniumRandomSeed:Begin:::
Class TaniumRandomSeed

	Private m_bErr
	Private m_errMessage
	Private m_strFoundKey
	Private m_intComputerID
	Private m_RandomSeedVal
	Private m_libVersion
	Private m_objShell

	Private Sub Class_Initialize
		m_libVersion = "6.2.314.3262"
		m_strFoundKey = ""
		m_intComputerID = ""
		m_RandomSeedVal = ""
		Set m_objShell = CreateObject("WScript.Shell")
		m_errMessage = ""
		m_bErr = False
		FindClientKey
		GetComputerID
		GetRandomSeed
		TaniumRandomize
    End Sub
	
	Private Sub Class_Terminate
		Set m_objShell = Nothing
	End Sub
	    
    Public Property Get RandomSeedValue
    	RandomSeedValue = m_RandomSeedVal
    End Property
    
    Public Sub TaniumRandomize
    	If Not m_RandomSeedVal = "" Then
    		Randomize(m_RandomSeedVal)
    	Else
    		m_bErr = True
    		m_errMessage = "Error: Could not randomize with a blank Random Seed Value"
    	End If
    End Sub

    Public Property Get LibVersion
    	LibVersion = m_libVersion
    End Property

    Public Property Get ErrorState
    	ErrorState = m_bErr
    End Property      
    
    Public Property Get ErrorMessage
    	ErrorMessage = m_errMessage
    End Property
    
	Public Sub ErrorClear
		m_bErr = False
		m_errMessage = ""
	End Sub
	
	Private Sub FindClientKey
		Dim keyNativePath, keyWoWPath, strPath

		keyNativePath = "Software\Tanium\Tanium Client"
		keyWoWPath = "Software\Wow6432Node\Tanium\Tanium Client"

	    ' first check the Software key (valid for 32-bit machines, or 64-bit machines in 32-bit mode)
	    On Error Resume Next
	    strPath = m_objShell.RegRead("HKLM\"&keyNativePath&"\Path")
	    On Error Goto 0
		m_strFoundKey = "HKLM\"&keyNativePath
	 
	  	If strPath = "" Then
	  		' Could not find 32-bit mode path, checking Wow6432Node
	  		On Error Resume Next
	  		strPath = m_objShell.RegRead("HKLM\"&keyWoWPath&"\Path")
	  		On Error Goto 0
			m_strFoundKey = "HKLM\"&keyWoWPath
	  	End If
	End Sub 'FindClientKey
	
	Private Sub GetComputerID
		If Not m_strFoundKey = "" Then
			On Error Resume Next
			m_intComputerID = m_objShell.RegRead(m_strFoundKey&"\ComputerID")
			If Err.Number <> 0 Then
				m_bErr = True
				m_errMessage = "Error: Could not read ComputerID value"
			End If
			On Error Goto 0
			m_intComputerID = ReinterpretSignedAsUnsigned(m_intComputerID)
		Else
		    m_bErr = True
    		m_errMessage = "Error: Could not retrieve computer ID value, blank registry path"
    	End If
	End Sub
	
	Private Sub GetRandomSeed
		Dim timerNum
		timerNum = Timer()
		If m_intComputerID <> "" Then
			If timerNum < 1 Then
				m_RandomSeedVal = (m_intComputerID / Timer() * 10 )
			Else
				m_RandomSeedVal = m_intComputerID / Timer
			End If
		Else
		    m_bErr = True
    		m_errMessage = "Error: Could not calculate Tanium Random Seed, blank computer ID value"
    	End If	
	End Sub

	Private Function ReinterpretSignedAsUnsigned(ByVal x)
		  If x < 0 Then x = x + 2^32
		  ReinterpretSignedAsUnsigned = x
	End Function 'ReinterpretSignedAsUnsigned
	
End Class 'TaniumRandomSeed
' :::VBLib:TaniumRandomSeed:End:::


' :::VBLib:TaniumContentRegistry:Begin:::
Class TaniumContentRegistry
	Private m_strFoundKey
	Private m_objShell
	Private m_bErr
	Private m_errMessage
	Private m_val
	Private m_subKey
	Private m_type
	Private m_data
	Private m_libVersion
	Private m_libName	

	Private Sub Class_Initialize
		m_libVersion = "6.2.314.3262"
		m_libName = "TaniumContentRegistry"
		m_strFoundKey = ""
		m_subKey = "/"
		m_val = ""
		m_type = ""
		m_data = ""
		Set m_objShell = CreateObject("WScript.Shell")
		FindClientKey
		m_errMessage = ""
		m_bErr = False
    End Sub
	
	Private Sub Class_Terminate
		Set m_objShell = Nothing
	End Sub
    
    Public Property Get ErrorState
    	ErrorState = m_bErr
    End Property
	
	Public Sub ResetState
		Class_Initialize
	End Sub 'ResetState
	
    Public Property Get LibVersion
    	LibVersion = m_libVersion
    End Property
    
    Public Property Get ErrorMessage
    	ErrorMessage = m_errMessage
    End Property

    Public Property Let Data(valData)
		m_data = valData
    End Property
    
    Public Property Let ValueName(valName)
    	If StringCheck(valName) Then
			m_val = valName
		Else
			m_bErr = True
			m_errMessage = "Error: Invalid registry value name, was not a string"
			ErrorCheck
		End If
    End Property
    
	Public Property Let RegValueType(strType)
		Dim bOK
		bOK = False
		Select Case (strType)
			Case "REG_SZ"
				bOK = True
			Case "REG_DWORD"
				bOK = True
			Case "REG_QWORD"
				bOK = True
			Case "REG_BINARY"
				bOK = True
			Case "REG_MULTI_SZ"
				bOK = True
			Case "REG_EXPAND_SZ"
				bOK = True
			Case Else
				m_bErr = True
				m_errMessage = "Error: Invalid registry value data type ("&strType&")"
		End Select
		
		If bOK Then
			m_type = strType
		Else
			m_type = ""
			ErrorCheck
		End If

	End Property
	
	Public Property Get ClientRootKey
		If Not InStr(m_strFoundKey,"Tanium\Tanium Client") > 0 Then
			m_strFoundKey = "Unknown"
			m_bErr = True
			m_errMessage = "Error: Cannot find Tanium Client Registry Key"
			ClientRootKey = ""
		Else
			ClientRootKey = m_strFoundKey
		End If
	End Property
	
	Public Property Get ValuesDict
		' Returns a dictionary object, key is name, value is friendly type
		' Value | REG_SZ
		' note that Value may be a 'null' value, equal to "". This will trigger
		' an error in the tContentReg object when values are read or written to.
		' A Null value is not supported.
		
		Const HKEY_LOCAL_MACHINE = &H80000002
		Dim objReg,arrValueNames(),arrValueTypes()
		Dim arrFriendlyValueTypeNames,strKeyPath,intReturn
		arrFriendlyValueTypeNames = Array("","REG_SZ","REG_EXPAND_SZ","REG_BINARY", _
								"REG_DWORD","REG_MULTI_SZ")
		Set objReg = GetObject("winmgmts:\\.\root\default:StdRegProv")
		' Remove 'HLKM\'
		strKeyPath = Right(m_strFoundKey,Len(m_strFoundKey) - 5)&"\Content\"&m_subKey
		intReturn = objReg.EnumValues(HKEY_LOCAL_MACHINE, strKeyPath, arrValueNames, arrValueTypes)

		Dim i,name,valueType,dictVals
		Set dictVals = CreateObject("Scripting.Dictionary")
		If intReturn = 0 Then 
			For i = 0 To UBound(arrValueNames)
				name = arrValueNames(i)
				If Not dictVals.Exists(name) Then
					dictVals.Add name,arrFriendlyValueTypeNames(arrValueTypes(i))
				End If
			Next
		End If
		Set ValuesDict = dictVals
	End Property 'ValuesDict

	Public Property Get SubKeysArray
		Const HKEY_LOCAL_MACHINE = &H80000002
		Dim objReg,arrKeys,intReturn

		Set objReg = GetObject("winmgmts:\\.\root\default:StdRegProv")
		' Remove 'HLKM\'	
		strKeyPath = Right(m_strFoundKey,Len(m_strFoundKey) - 5)&"\Content\"&strSubKey
		intReturn = objReg.EnumKey(HKEY_LOCAL_MACHINE, strKeyPath, arrKeys)
		If intReturn = 0 Then
			SubKeysArray = arrKeys
		Else
			SubKeysArray = Array()
		End If
	End Property 'SubKeysArray
		
	Public Property Let ClientSubKey(strSubKey)
    	If StringCheck(strSubKey) Then
			m_subKey = EnsureSuffix(strSubKey,"\")
		Else
			m_bErr = True
			m_errMessage = "Error: Invalid registry subkey name, was not a string"
			ErrorCheck
		End If
	End Property
	
	Private Sub FindClientKey
		Dim keyNativePath, keyWoWPath, strPath, strDeleteTest

		keyNativePath = "Software\Tanium\Tanium Client"
		keyWoWPath = "Software\Wow6432Node\Tanium\Tanium Client"

	    ' first check the Software key (valid for 32-bit machines, or 64-bit machines in 32-bit mode)
	    On Error Resume Next
	    strPath = m_objShell.RegRead("HKLM\"&keyNativePath&"\Path")
	    On Error Goto 0
		m_strFoundKey = "HKLM\"&keyNativePath
	 
	  	If strPath = "" Then
	  		' Could not find 32-bit mode path, checking Wow6432Node
	  		On Error Resume Next
	  		strPath = m_objShell.RegRead("HKLM\"&keyWoWPath&"\Path")
	  		On Error Goto 0
			m_strFoundKey = "HKLM\"&keyWoWPath
	  	End If
	End Sub 'FindClientKey
	
	Private Sub CheckReady
		Dim arrReadyItems,item
		arrReadyItems = Array(m_strFoundKey,m_subKey,m_type,m_val)
		For Each item In arrReadyItems
			If item = "" Then
				m_bErr = True
				m_errMessage = "Error: Tried to commit but key, type, or value is not set. Default (blank) value names not supported."
			End If
		Next
	End Sub 'CheckReady
	
	Public Sub ErrorClear
		m_bErr = False
		m_errMessage = ""
	End Sub

	Private Sub ErrorCheck
		' Call on all Lets
		If m_bErr Then
			Err.Raise vbObjectError + 1978, m_libName, m_errMessage
		End If
	End Sub 'ErrorCheck
	
	Public Function Write
		CheckReady
		Dim res
		If m_data = "" Then
			m_bErr = True
			m_errMessage = "Error: Tried to commit but key, type, value, or data is not set"
		End If
		If Not m_bErr Then
			Dim errDesc
			If Not SubKeyExists Then CreateSubKey
			On Error Resume Next
			res = m_objShell.RegWrite(m_strFoundKey&"\Content\"&m_subKey&m_val,m_data,m_type)
			If Err.Number <> 0 Then
				errDesc = Err.Description
				On Error Goto 0
				m_bErr = True
				m_errMessage = "Error: Could not Write Data to "&m_strFoundKey&"\Content\"&m_subKey&EnsureSuffix(m_val, "\")&": "&errDesc
			End If
			On Error Goto 0
		End If
		Write = res
		ErrorCheck
	End Function

	Public Function DeleteVal
		Dim res
		CheckReady
		res = ""
		If Not m_bErr Then
			Dim errDesc, errNum
			On Error Resume Next
			res = m_objShell.RegDelete(m_strFoundKey&"\Content\"&m_subKey&m_val)
			If Err.Number <> 0 Then
				errDesc = Err.Description
				On Error Goto 0
				m_bErr = True
				m_errMessage = "Error: Could not Delete Value "&m_strFoundKey&"\Content\"&m_subKey&EnsureSuffix(m_val, "\")&": "&errDesc
			End If
		End If
		DeleteSubKey = res
		ErrorCheck
	End Function	
	
	Public Function SubKeyExists
		Dim num
		On Error Resume Next
		res = m_objShell.RegRead(m_strFoundKey&"\Content\"&m_subKey)
		num = Err.Number
		On Error Goto 0
		If num <> 0 Then
			SubKeyExists = False
		Else
			SubKeyExists = True
		End If
	End Function
	
	Public Function CreateSubKey
		Dim strKey,strCreateKey,res
		
		res = m_objShell.RegWrite(m_strFoundKey&"\Content\","")
		On Error Goto 0
		strCreateKey = m_strFoundKey&"\Content\"
		On Error Resume Next
		For Each strKey In Split(m_subKey,"\")
			If strKey <> "" Then
				strCreateKey = strCreateKey&EnsureSuffix(strKey,"\")
				res = m_objShell.RegWrite(strCreateKey,"")
				If Err.Number <> 0 Then
					errDesc = Err.Description			
					m_bErr = True
					m_errMessage = "Error: Registry Key Create Failure for "&strCreateKey&": "&errDesc
				End If
			End If
		Next
		On Error Goto 0
		CreateSubKey = m_bErr
		ErrorCheck
	End Function

	Public Function DeleteSubKey
		Dim strKey,strCreateKey,res,arr,i,j
		strCreateKey = m_strFoundKey&"\Content\"
		ReDim arr(UBound(Split(m_subKey,"\")))
		i = 0
		For Each strKey In Split(m_subKey,"\")
			If strKey <> "" Then
				strCreateKey = strCreateKey&EnsureSuffix(strKey,"\")
				arr(i) = strCreateKey
				i = i + 1
			End If
		Next
		On Error Resume Next
		For j = i To 0 Step -1
			If Trim(arr(j)) <> "" Then
				res = m_objShell.RegDelete(arr(j))
				If Err.Number <> 0 Then
					errDesc = Err.Description
					m_bErr = True
					m_errMessage = "Error: Registry Key Delete Failure for "&strCreateKey&": "&errDesc
				End If
			End If
		Next
		DeleteSubKey = m_bErr
		ErrorCheck
	End Function
	
	Public Function Read
		CheckReady
		Dim res,errDesc
		If Not m_bErr Then
			On Error Resume Next
			res = m_objShell.RegRead(m_strFoundKey&"\Content\"&m_subKey&m_val)
			If Err.Number <> 0 Then
				errDesc = Err.Description			
				m_bErr = True
				m_errMessage = "Error: Registry Read Failure for "&m_strFoundKey&"\Content\"&m_subKey&m_val&": "&errDesc
			End If
			On Error Goto 0
		End If
		Read = res ' no value will return ""
		ErrorCheck
	End Function

	Private Function StringCheck(inVar)
		If VarType(inVar) = vbString Then
			StringCheck = True
		Else
    		m_bErr = True
    		m_errMessage = "Error: Invalid input, must be a string"		
			StringCheck = False
		End If
	End Function
	
	Private Function EnsureSuffix(strIn,strSuffix)
		If Not Right(strIn,Len(strSuffix)) = strSuffix Then
			EnsureSuffix = strIn&strSuffix
		Else
			EnsureSuffix = strIn
		End If
	End Function 'EnsureSuffix
End Class 'TaniumContentRegistry
' :::VBLib:TaniumContentRegistry:End:::