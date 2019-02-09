'Tanium File Version:2.2.2.0011

Option Explicit

'use 64-bit WUA scanner on 64-bit OS
x64Fix

' allow override
RunOverride

' Global classes
Dim tLog
Set tLog = New TaniumContentLog
tLog.Log "----------------Beginning Patch Install----------------"

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

RandomSleep TryFromDict(dictPConfig,"RandomInstallWaitTimeInSeconds",240)

' Consider maintenance Windowing
' a window is only valid if it is 3 days long. Longer windows are considered errors
Dim intMaxWindowLengthInDays
intMaxWindowLengthInDays = 3 

QuitIfConfiguredToObeyMaintenanceWindows intMaxWindowLengthInDays

' check for / run 'pre' files
RunFilesInDir("pre")
' put files in a directory called install-patches\pre
' candidates would be popup notifiers, checks to continue

' Some Globals needed througout the script
Dim wuaService, wuaNeedsStop
Dim strPatchToolsDir,strToolsDir
Dim objFSO,intLocale,dtmNowUTC,strScanDir
Dim arrFilesInDirAtStart,l,strFileToDeletePath,strUseScanSource
Dim bPatchApprovalAware,bApprovedPatchStillRequired,strInstallingCurrentlyTextFilePath
Dim bNeedsRebootAwakenings,intFailureCount,bHasGlobalSuccess


'' - Advanced patch / patch approval related setup - ''
' look to see whether approved patches should be considered
' This results in modified reboot behavior
bPatchApprovalAware = IsPatchApprovalAware
' Examine state of queued reboot jobs and active lists
Dim dictDormantRebootLists : Set dictDormantRebootLists = CreateObject("Scripting.Dictionary")
' Pull all reboot state data into a dictionary of dictionaries
Dim dictSystemRebootValues : Set dictSystemRebootValues = CreateObject("Scripting.Dictionary")
GetSystemRebootValues(dictSystemRebootValues)
' Determine which patch lists have dormant reboot values
Set dictDormantRebootLists = GetActiveRebootForLists
' Determine which Approval Lists still contain required patches
bApprovedPatchStillRequired = False ' global set in function below
'' - end advanced patch / patch approval setup - ''

' Directory to store patches in can be overridden via command line. This can be relative path or full path.
Dim strPatchesDir
strPatchesDir = GetTaniumDir(PatchesDirTranslator(TryFromDict(dictPConfig,"PatchesDir","Tools\Patches")))

' will hold a list of paths to files in the patches directory
arrFilesInDirAtStart = GetFilesListInPatchesDirOrQuit(strPatchesDir)

' can set global locale code via content
' the LocaleID string value
intLocale = GetTaniumLocale()
SetLocale(intLocale) ' sets locale options (date, commas and decimals, etc ...)

' get current time in UTC
dtmNowUTC = DateAdd("n",-GetTZBias,Now())


' set up and delete (if exists) the installingcurrently file
Set objFSO = CreateObject("Scripting.FileSystemObject")

strScanDir = GetTaniumDir("Tools\Scans")
strInstallingCurrentlyTextFilePath = strScanDir & "installingcurrently.txt"

strPatchToolsDir = GetTaniumDir("Tools\Patch Tools")

' Find valid cab files to scan against
' if CustomCabSupport registry value is set
' also look for windows XP patch scan cab files
Dim dictCabs
' Check if custom cab scanning is on
Set dictCabs = GetAllCabs

' check and remember service status
' and put service status back the way it was after scan is done
CheckWindowsUpdate() 
bHasGlobalSuccess = False
Dim strCabPath
For Each strCabPath In dictCabs.Keys
	RunInstallForCab strCabPath,dictCabs.Item(strCabPath)
Next

If bHasGlobalSuccess Then
	tLog.log "Re-running patch scan to update patch data"
	AccessRunPatchScan
End If

' If Failure Count is at threshold value, or if there are no approved patches still required
' call reboot with end user tools to wake any dormant jobs related to approval lists
' which do not have any required patches in them to start the reboot process
Dim dictApprovalListsAndRequiredStatus : Set dictApprovalListsAndRequiredStatus = GetApprovalListsWithRequiredStatus
Dim strApprovalListForRebootGUID
If intFailureCount >= TryFromDict(dictPConfig,"MaxFailuresThreshold",5) Or Not bApprovedPatchStillRequired Then
	For Each strApprovalListForRebootGUID In dictApprovalListsAndRequiredStatus.Keys
		' activate only those dormant lists where no updates are required
		If dictApprovalListsAndRequiredStatus(strApprovalListForRebootGUID) = "NoRequiredPatches" Then
			If Not bIgnoreEUT Then
				tLog.log "Awakening any reboot jobs queued for Approval List: " & strApprovalListForRebootGUID
				AwakenEUTRebootJobs(strApprovalListForRebootGUID)
			End If
		End If
	Next
End If


If bApprovedPatchStillRequired Then
	tLog.log "Install and patch scan completed, " _
		& "but an approved patch is still required."
End If

tLog.log "Sleeping for four seconds"
WScript.Sleep(4000)
StopWindowsUpdate() 'if necessary

'Delete all files, regardless of success
For Each strFileToDeletePath In arrFilesInDirAtStart
	If objFSO.FileExists(strFileToDeletePath) Then
		objFSO.DeleteFile strFileToDeletePath, True
		tLog.log "Deleted " & strFileToDeletePath
	End If
Next


tLog.log "Done deleting any patch files that existed at start of patch install job"
RunFilesInDir("post")
' put files in a directory called install-patches\post
' candidates would be things like reboot if necessary, warn user of reboot

WScript.Quit

Function RunInstallForCab(strCabPath,strCabName)
	' Examine files in directory and decide if their filenames match up with
	' those listed in a cab file.
	' XP Custom Support cabs are deployed with a filename which is different than that
	' which is in the cab file.
	tLog.log "Scanning with definitions in " & strCabName
	' CHECK SCAN SOURCE (as specified in run-patch-scan.vbs)
	strUseScanSource = LCase(TryFromDict(dictPConfig,"UseScanSource","cab"))
	If Not strUseScanSource = "" Then
		tLog.log "Using scan source: " & strUseScanSource
	Else
		' default to using cab setting
		strUseScanSource = "cab"
		tLog.log "Using scan source: " & strUseScanSource & " (default)"
	End If
	
	
	'CHECK CAB FILE
	strToolsDir = GetTaniumDir("Tools")
	If Not objFSO.FileExists(strCabPath)  And strUseScanSource = "cab" Then
		tLog.log strCabPath & " not deployed and using cab as scan source, quitting"
		Exit Function
	End If

	Dim UpdateSession,UpdateServiceManager,UpdateService,UpdateSearcher
	Dim updatesToInstall,SearchResult,Updates,bFailureCountResetNeeded
	Dim bScanSourceOverrideToCab,bDisableMicrosoftUpdate
	
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
	
	'GET CLIENT UPDATE LIST
	Set UpdateSession = CreateObject("Microsoft.Update.Session")
	UpdateSession.UserLocale = intLocale ' Changes output language
	UpdateSession.WebProxy.AutoDetect = True ' Set proxy to use IE autodetect settings
	
	Set UpdateServiceManager = CreateObject("Microsoft.Update.ServiceManager")
	
	bDisableMicrosoftUpdate = TryFromDict(dictPConfig,"DisableMicrosoftUpdate",False)	
	If Not strUseScanSource = "cab" And Not bScanSourceOverrideToCab Then
		If Not bDisableMicrosoftUpdate Then
			' add microsoft update if we're using an online scan source
			On Error Resume Next
			Set UpdateService = UpdateServiceManager.AddService2("7971f918-a847-4430-9279-4a52d1efe18d",7,"")
			If Err.Number <> 0 Then
				tLog.log "Could not set update service to Microsoft Update, Error was " & Err.Number
				Err.Clear
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
			Exit Function
		Else ' cab is there
			Set UpdateService = UpdateServiceManager.AddScanPackageService("Offline Sync Service", strCabPath)
			If IsEmpty(UpdateService) Then
				tLog.log "Error creating Offline Sync Service object (Windows Update may be disabled or bad cab file), quitting"
				Exit Function
			End If
		End If
	End If
	
	UpdateSearcher.IncludePotentiallySupersededUpdates = True
	
	' Determine scan method for installs applicability
	If bScanSourceOverrideToCab Then
		UpdateSearcher.ServerSelection = 3
		UpdateSearcher.ServiceID = UpdateService.ServiceID	' only set when cab		
	Else
		Select Case strUseScanSource
			Case "systemdefault"
				UpdateSearcher.ServerSelection = 0
			Case "wsus"
				UpdateSearcher.ServerSelection = 1
			Case "internet"
				UpdateSearcher.ServerSelection = 2
			Case "optimal" ' this is online with Microsoft Update
				UpdateSearcher.ServerSelection = 3 ' now requires a serviceID
				UpdateSearcher.ServiceID = UpdateService.ServiceID
			Case "cab"
				UpdateSearcher.ServerSelection = 3
				UpdateSearcher.ServiceID = UpdateService.ServiceID	' only set when cab					
			Case Else ' unknown option, defaults to optimal
				tLog.log "Unknown Scan Source reference, choosing local cab file"
				UpdateSearcher.ServerSelection = 3
				UpdateSearcher.ServiceID = UpdateService.ServiceID
				strUseScanSource = "cab"
			End Select
			
		' sleep if scan source was a service based, non-cab source
		If Not strUseScanSource = "cab" Then
			tLog.log "Scan source is online / service-based, sleeping for a pre-determined amount of time"
			RandomSleep(TryFromDict(dictPConfig,"OnlineScanRandomWaitTimeInSeconds",0))
		End If
	End If
		
	' Set UpdateSearcher = UpdateSession.CreateUpdateSearcher()
	' Allow superseded installs no matter what - they were pushed here
	' The scan allows for turning this behavior off / modifying what is defined as superseded
	
	Set SearchResult = UpdateSearcher.Search("Type='Software' and IsHidden=1 or IsHidden=0")
	
	' start cab fallback
	If Err.Number <> 0 Then ' Try scanning with cab file as failover
		tLog.log "Cannot complete patch scan via scan source " & strUseScanSource & ", Scan Error: " & Err.Number
		tLog.log "Retrying with offline cab file as backup scan source"
		'Retry scan with offline cab file as last resort
		If Not objFSO.FileExists(strCabPath) Then
			tLog.log "Scan Error: Cannot locate offline cab file at "&strCabPath&" when scan source is set to cab"
			' todo - cleanup if exiting here
			Exit Function
		End If
		Set UpdateService = UpdateServiceManager.AddScanPackageService("Offline Sync Service", strCabPath)
		' Reset update objects
		'Set UpdateServiceManager = CreateObject("Microsoft.Update.ServiceManager") 'redundant
		Set UpdateSearcher = UpdateSession.CreateUpdateSearcher()
		' UpdateSearcher.ServiceID = UpdateService.ServiceID 'redundant
		UpdateSearcher.IncludePotentiallySupersededUpdates = bShowSupersededUpdates
		If IsEmpty(UpdateService) Then
			tLog.log "Error creating Offline Sync Service object (Windows Update may be disabled or bad cab file), quitting"
			Exit Function
		End If
		UpdateSearcher.ServerSelection = 3
		UpdateSearcher.ServiceID = UpdateService.ServiceID	' only set when cab	
		Err.Clear
		Set SearchResult = UpdateSearcher.Search("Type='Software' and IsHidden=1 or IsHidden=0")
		If Err.Number <> 0 Then
			tLog.log "Cannot complete patch scan with offline cab file as backup scan source, Scan Error: " & Err.Number
			Err.Clear
			Exit Function
		Else ' no error scanning with cab as backup
			tLog.log "Scan failed with " & strUseScanSource & " as scan source, but was successful with offline cab file as backup"
		End If
		Err.Clear
	End If
	''' end cab fallback
	
	Set Updates = SearchResult.Updates
	
	If searchResult.Updates.Count = 0 Then	
	    tLog.log "There are no applicable updates."
	    Exit Function
	End If
	
	Set updatesToInstall = CreateObject("Microsoft.Update.UpdateColl")
	
	' create text file to log updates installed by Tanium
	
	Dim strInstallResultsPath,strInstallResultsReadablePath
	Dim intInstallResultsFileMode
	Dim intDesiredInstallResultsColumnCount,bBadInstallResultsLines,bGoodInstallResultsLine
	Dim arrInstallResultsLine,strInstallResultsLine,dictInstallResults
	Dim objInstallResultsTextFile,installationResult,strSep
	Dim I,J,K,urls,update,bundledUpdate,cacheFiles,hasFoundFile
	Dim contents,installer,bHasLocalSuccess,bHasLocalFailure,strSeverity,strKBArticles,strKBArticle
	Dim objInstallingCurrentlyTextFile,strTrulyUniqueID,strDownloadSize
	Dim words,files,file,filenames
		
	strSep = "|"
	
	strInstallResultsPath = strScanDir & "\installedresults.txt"
	strInstallResultsReadablePath = strScanDir & "\installedresultsreadable.txt"
	
	
	If Not objFSO.FileExists(strInstallResultsPath) Then
		objFSO.CreateTextFile strInstallResultsPath,True
	End If
	
	Set objInstallResultsTextFile = objFSO.OpenTextFile(strInstallResultsPath,1,True)
	
	' Read existing InstallResults snapshot into a dictionary object
	Set dictInstallResults = CreateObject("Scripting.Dictionary")
	
	' the InstallResults file is read and appended to versus being overwritten each time.
	' If we change the format of the file, we are probably changing column count. 
	' If the line entry does not match the column count, do not append and instead
	' overwrite the file later.
	' only unique lines are read into the dictionary object, and only unique lines
	' are written back.
	
	intDesiredInstallResultsColumnCount = 5
	bBadInstallResultsLines = False	
	While objInstallResultsTextFile.AtEndOfStream = False
		bGoodInstallResultsLine = True
		strInstallResultsLine = objInstallResultsTextFile.ReadLine
		arrInstallResultsLine = Split(strInstallResultsLine,"|")
		If IsArray(arrInstallResultsLine) Then
			If UBound(arrInstallResultsLine) <> (intDesiredInstallResultsColumnCount-1) Then
				tLog.log "bad line detected:" & strInstallResultsLine
				bGoodInstallResultsLine = False
				bBadInstallResultsLines = True
			End If
			If (Not dictInstallResults.Exists(strInstallResultsLine)) And bGoodInstallResultsLine Then
				dictInstallResults.Add strInstallResultsLine,1
			End If
		End If
	Wend
	
	If TryFromDict(dictPConfig,"ClearInstallResultsOnBadLine",False) And bBadInstallResultsLines Then ' overwrite file
		intInstallResultsFileMode = 2 ' overwrite
		tLog.log "Will overwrite InstallResults file"		
	Else
		intInstallResultsFileMode = 8 ' append
	End If
	
	' Will need to re-open for either writing or appending now that it's in dictionary
	objInstallResultsTextFile.Close
	Set objInstallResultsTextFile = objFSO.OpenTextFile(strInstallResultsPath,intInstallResultsFileMode,True)
	
	'Array to hold pretty names of results
	Dim arrInstallationResultCodes,intResultCode
	arrInstallationResultCodes = Array( "Not Started", "In Progress", "Succeeded", _
		                        "Succeeded With Errors", "Failed", "Aborted" )
	
	' Create a file to hold which patches are currently being installed
	' for use with a companion sensor

	Set objInstallingCurrentlyTextFile = objFSO.OpenTextFile(strInstallingCurrentlyTextFilePath,2,True)

	Dim hasher
	Set hasher = New MD5er

	Dim dictUpdateSearchResultIDtoDetails
	Set dictUpdateSearchResultIDtoDetails = CreateObject("Scripting.Dictionary")
	
	Dim strFileNameFromCab,strFilePath
	For I = 0 to searchResult.Updates.Count-1
	    
	    Set update = searchResult.Updates.Item(I)
	    
	    urls = ""
		
		If Not(update.IsInstalled) Then tLog.log update.Title & " Is missing and could be installed"

		' SCUP updates are never bundled
		
    	For K = 0 To update.DownloadContents.Count-1
			Set cacheFiles = CreateObject("Microsoft.Update.StringColl")
			hasFoundFile = False	    	
    		Set contents = update.DownloadContents.Item(K)
    		words = Split(contents.DownloadUrl, "/")
    		If urls = "" Then
    			urls = contents.DownloadUrl
    			filenames = words(UBound(words))
			Else
	    		urls = urls & " " & contents.DownloadUrl
	    		filenames = filenames & "," & words(UBound(words))
			End If
			For Each strFilePath In arrFilesInDirAtStart
				Set file = objFSO.GetFile(strFilePath)
				strFileNameFromCab = Right(contents.DownloadUrl, Len(file.Name))
				If InStr(LCase(update.Title),") custom support") > 0 Then ' XP Custom support update
					' Modify the comparison so that the update's filename inside the cab
					' which looks something like windowsxp-kb2953522-x86-custom-enu_c6820e06d430fa90266689e71652327c057737ea.exe
					strFileNameFromCab = GetCleanXPFileFromURL(contents.DownloadUrl)
					' tLog.log "Custom XP Support File Found, rewriting file name to " & strFileNameFromCab
				End If
				If LCase(strFileNameFromCab) = LCase(file.Name) Then
					hasFoundFile = True
					tLog.log "Found file to update ("&update.Title&"): " & file.Name
					On Error Resume Next 'may fail if already exists
					cacheFiles.Add(strPatchesDir & file.Name)
					On Error Goto 0
					tLog.log "copied into cache: " & strPatchesDir & file.Name	

					If bundledUpdate.EulaAccepted = False Then
						bundledUpdate.AcceptEula
					End If
				End If
			Next
			
			If hasFoundFile Then
				' This may fail if already exists
				On Error Resume Next
				
				bundledUpdate.CopyToCache(cacheFiles)
				If update.EulaAccepted = False Then 
					update.AcceptEula
				End If
				On Error Goto 0
				' get information about update and write to the currentlyinstalling file
			
				' pull severity
			    strSeverity = ""
			    If (IsNull(update.MsrcSeverity) or update.MsrcSeverity = "")  then 
			    	strSeverity = "None"
			    Else
					strSeverity = update.MsrcSeverity
			    End If
			    
				' pull kb articles
				strKBArticles = " "
				For Each strKBArticle In update.KBArticleIDs
					strKBArticles = strKBArticles & "KB"&strKBArticle & " "
				Next
				strKBArticles = Trim(strKBArticles)
				If strKBArticles = "" Then strKBArticles = "None"
					
				' build an ID value which is more unique than the GUID
			    ' a GUID can be shared between two updates which are different binaries
			    ' example: silverlight update for end users and silverlight update
			    ' for developers share the same GUID value
			    strDownloadSize = GetPrettyFileSize(update.MaxDownloadSize)

				strTrulyUniqueID = ""
				If ( Not IsNull(update.Identity.UpdateID) ) Then
					strTrulyUniqueID = hasher.GetMD5(update.Identity.UpdateID&urls&filenames&strDownloadSize)
				Else
					strTrulyUniqueID = "None"
				End If

				' write to currently installing file
				tLog.log "Will attempt to install " _
					&update.Title&", severity " _ 
					&strSeverity&",KB Article "&strKBArticles&" - Unique ID: "&strTrulyUniqueID
				objInstallingCurrentlyTextFile.WriteLine(UnicodeToAscii(update.Title&strSep _ 
					&strSeverity&strSep&strKBArticles&strSep&strTrulyUniqueID))
				
				updatesToInstall.Add(update)
				'updatesToInstall.Add(bundledUpdate)
			End If
    	Next

	' The traditional method - Bundled updates
	    For J = 0 To update.BundledUpdates.Count-1
	    	Set bundledUpdate = update.BundledUpdates.Item(J)
	    	
			Set cacheFiles = CreateObject("Microsoft.Update.StringColl")
			hasFoundFile = False
	    	
	    	For K = 0 To bundledUpdate.DownloadContents.Count-1
	    		Set contents = bundledUpdate.DownloadContents.Item(K)
	    		words = Split(contents.DownloadUrl, "/")
	    		If urls = "" Then
	    			urls = contents.DownloadUrl
	    			filenames = words(UBound(words))
				Else
		    		urls = urls & " " & contents.DownloadUrl
		    		filenames = filenames & "," & words(UBound(words))
				End If
				For Each strFilePath In arrFilesInDirAtStart
					Set file = objFSO.GetFile(strFilePath)

					strFileNameFromCab = Right(contents.DownloadUrl, Len(file.Name))
					If InStr(LCase(update.Title),") custom support") > 0 Then ' XP Custom support update
						' Modify the comparison so that the update's filename inside the cab
						' which looks something like windowsxp-kb2953522-x86-custom-enu_c6820e06d430fa90266689e71652327c057737ea.exe
						strFileNameFromCab = GetCleanXPFileFromURL(contents.DownloadUrl)
						' tLog.log "Custom XP Support File Found, rewriting file name to " & strFileNameFromCab
					End If
					If LCase(strFileNameFromCab) = LCase(file.Name) Then
						hasFoundFile = True
						tLog.log "Found file to update ("&update.Title&"): " & file.Name
						On Error Resume Next 'may fail if already exists
						cacheFiles.Add(strPatchesDir & file.Name)
						On Error Goto 0
						tLog.log "copied into cache: " & strPatchesDir & file.Name	
	
						If bundledUpdate.EulaAccepted = False Then
							bundledUpdate.AcceptEula
						End If
					End If
				Next
				
				If hasFoundFile Then
					' This may fail if already exists
					On Error Resume Next
					bundledUpdate.CopyToCache(cacheFiles)
					If update.EulaAccepted = False Then 
						update.AcceptEula
					End If
					On Error Goto 0
					' get information about update and write to the currentlyinstalling file
				
					' pull severity
				    strSeverity = ""
				    If (IsNull(update.MsrcSeverity) or update.MsrcSeverity = "")  then 
				    	strSeverity = "None"
				    Else
						strSeverity = update.MsrcSeverity
				    End If
				    
					' pull kb articles
					strKBArticles = " "
					For Each strKBArticle In update.KBArticleIDs
						strKBArticles = strKBArticles & "KB"&strKBArticle & " "
					Next
					strKBArticles = Trim(strKBArticles)
					If strKBArticles = "" Then strKBArticles = "None"
						
					' build an ID value which is more unique than the GUID
				    ' a GUID can be shared between two updates which are different binaries
				    ' example: silverlight update for end users and silverlight update
				    ' for developers share the same GUID value
				    strDownloadSize = GetPrettyFileSize(update.MaxDownloadSize)

					strTrulyUniqueID = ""
					If ( Not IsNull(update.Identity.UpdateID) ) Then
						strTrulyUniqueID = hasher.GetMD5(update.Identity.UpdateID&urls&filenames&strDownloadSize)
					Else
						strTrulyUniqueID = "None"
					End If
					
					Dim strUpdateShortDetails
					strUpdateShortDetails = update.Identity.UpdateID & strSep & _
						   	update.Title & strSep & _
						   	dtmNowUTC & strSep & _
						   	strTrulyUniqueID
   					'Key off the update ID, and not the unique ID, which must be reconstructed
   					' There is certainly no chance that the same update ID would be installed at the same
   					' time on the same machine. Unique ID is a construct created for exceedingly rare conditions
					If Not dictUpdateSearchResultIDtoDetails.Exists(update.Identity.UpdateID) Then
						dictUpdateSearchResultIDtoDetails.Add update.Identity.UpdateID, strUpdateShortDetails
					End If
		
					' write to currently installing file
					tLog.log "Will attempt to install " _
						&update.Title&", severity " _ 
						&strSeverity&",KB Article "&strKBArticles&" - Unique ID: "&strTrulyUniqueID
					objInstallingCurrentlyTextFile.WriteLine(UnicodeToAscii(update.Title&strSep _ 
						&strSeverity&strSep&strKBArticles&strSep&strTrulyUniqueID))

					updatesToInstall.Add(update)
					'updatesToInstall.Add(bundledUpdate)
				End If
	    	Next
	    Next
	Next
	
	If updatesToInstall.Count = 0 Then
		tLog.log "No updates found"
		' close text files
		objInstallingCurrentlyTextFile.Close
		If objFSO.FileExists(strInstallingCurrentlyTextFilePath) Then
			' delete this file in the unlikely event that it exists
			objFSO.DeleteFile strInstallingCurrentlyTextFilePath,True
		End If
		Exit Function
	End If
	
	'go through all updates, hopefully through cached files!
	tLog.log "Installing updates..."
	' Log that it's Tanium (as seen in WindowsUpdate.log)
	UpdateSession.ClientApplicationID = "Tanium Patch Install " & strPatchToolsVersion
	Set installer = updateSession.CreateUpdateInstaller()
	installer.ClientApplicationID = "Tanium Patch Install " & strPatchToolsVersion
	
	installer.Updates = updatesToInstall
	If TryFromDict(dictPConfig,"RunInteractively",False) Then
		Set installationResult = installer.RunWizard("Tanium Patch Deployment")
	Else
		Set installationResult = installer.Install()
	End If
	
	'Output results of install
	intResultCode = installationResult.ResultCode
	tLog.log "Installation Result: " & _
	arrInstallationResultCodes(intResultCode)
	tLog.log "Reboot Required: " & _ 
	installationResult.RebootRequired & vbCRLF 
	tLog.log "Listing of updates installed " & _
	 "and individual installation results:" 
	' Get current failure count
	tContentReg.ValueName = "PatchInstallFailureCount"
	On Error Resume Next
	intFailureCount = tContentReg.Read
	If Err.Number <> 0 Or Not IsNumeric(intFailureCount) Then
		intFailureCount = "Not Set"
		Err.Clear
		tContentReg.ErrorClear
	Else
		intFailureCount = CInt(intFailureCount)
	End If

	On Error Goto 0
	
	bFailureCountResetNeeded = False
	' Ensure failure count is set / a number
	If IsNumeric(intFailureCount) Then
		If Not CStr(CLng(intFailureCount)) = intFailureCount Then 'if not integer
			bFailureCountResetNeeded = True
		End If
	Else ' was not numeric at all, possibly blank - reset to 0
		bFailureCountResetNeeded = True
	End If

	tLog.log "Previous Failure Count is " & intFailureCount
	If bFailureCountResetNeeded Then
		intFailureCount = 0
		tContentReg.ValueName = "PatchInstallFailureCount"
		tContentReg.Data = intFailureCount
		On Error Resume Next
		tContentReg.Write
		If Err.Number <> 0 Then
			tLog.Log "Error: Could not reset install failure count to registry, error was " & Err.Description
			Err.Clear
			tContentReg.ErrorClear			
		End If
		On Error Goto 0
		tLog.log "(Re)-Initializing failure count, writing 0 to registry"
	End If
	
	' loop through results and determine overall success or failure for run
	bHasLocalSuccess = False
	bHasLocalFailure = False
	For I = 0 to updatesToInstall.Count - 1
		intResultCode = installationResult.GetUpdateResult(i).ResultCode
		' rescan / consider success if 'in progress', 'succeeded', or
		' 'succeeded with errors'
		If intResultCode = 1 Or intResultCode = 2 Or intResultCode = 3 Then
			bHasLocalSuccess = True
			bHasGlobalSuccess = True ' will trigger a re-scan after all cab installs
		' note any failures if Succeeded with Errors or Failed
		ElseIf intResultCode = 3 Or intResultCode = 4 Then
			bHasLocalFailure = True
		End If
		
		tLog.log I + 1 & "> " & _
		updatesToInstall.Item(i).Title & _
		": " & arrInstallationResultCodes(installationResult.GetUpdateResult(i).ResultCode) & _
		", hresult code: " & installationResult.GetUpdateResult(i).HResult
		' write to log
		Dim arrShortDetails,strDetailsForOutput
		If dictUpdateSearchResultIDtoDetails.Exists(updatesToInstall.Item(i).Identity.UpdateID) Then
			strDetailsForOutput = dictUpdateSearchResultIDtoDetails.Item(updatesToInstall.Item(i).Identity.UpdateID)
		Else
			WScript.Echo "Could not retrieve details of installed update for output file"
		End If
		arrShortDetails = Split(strDetailsForOutput,strSep)

		If UBound(arrShortDetails) > 2 Then
			objInstallResultsTextFile.WriteLine(UnicodeToAscii(arrShortDetails(0)&strSep _
				&arrShortDetails(1)&strSep&arrShortDetails(2)&strSep _
				&arrInstallationResultCodes(installationResult.GetUpdateResult(i).ResultCode)&strSep _
				&arrShortDetails(3)))
		Else
			tLog.Log "Could not write result to install results file, details on the result were incompatible"
		End If
	Next
	' close InstalledResults file
	objInstallResultsTextFile.Close
	
	' remove currently installing file
	objInstallingCurrentlyTextFile.Close
	If objFSO.FileExists(strInstallingCurrentlyTextFilePath) Then
		' delete this file in the unlikely event that it exists
		objFSO.DeleteFile strInstallingCurrentlyTextFilePath,True
	End If
	
	' note failure
	If bHasLocalFailure Then
		If Not IsNumeric(intFailureCount) Then intFailureCount = 0
		intFailureCount = intFailureCount + 1 ' global value
		tLog.log "Failure detected, new failure count is: " & intFailureCount
		tContentReg.ValueName = "PatchInstallFailureCount"
		tContentReg.Data = intFailureCount
		On Error Resume Next
		tContentReg.Write
		If Err.Number <> 0 Then
			tLog.Log "Error: Could not record install failure count to registry, error was " & Err.Description
			Err.Clear
			tContentReg.ErrorClear			
		End If
	End If
	

	' If there was at least one successful install and no failures,
	' change failurecount value to 0
	If bHasLocalSuccess And Not bHasLocalFailure Then
		tLog.log "Successful install with no failures, resetting failure count to 0"
		RecordFailureCount 0
	End If

	Dim fileToCopy
	' make readable
	set fileToCopy = objFSO.GetFile(strInstallResultsPath)
	filetoCopy.Copy strInstallResultsReadablePath,True

End Function 'RunInstallForCab

Function AccessRunPatchScan
	Dim strVBS,fso,strCommand,objShell,objScriptExec,strResults
	
	strVbs = GetTaniumDir("Tools") & "run-patch-scan.vbs"
	
	Set fso = WScript.CreateObject("Scripting.Filesystemobject")
	
	If fso.FileExists (strVbs) Then
		
		strCommand = "cscript " & Chr(34) & strVbs & Chr(34)
		
		Set objShell = CreateObject("WScript.Shell")
		Set objScriptExec = objShell.Exec (strCommand)
		strResults = objScriptExec.StdOut.ReadAll
		
		''output results from running patch scan
		tLog.log strResults
	Else
		tLog.log strVbs & " not found"
	End If
	
End Function


Function StopWindowsUpdate()
	If wuaNeedsStop Then 
		tLog.log "Stopping Windows Update service"
		Dim oShell
		Set oShell = WScript.CreateObject ("WScript.Shell")
		oShell.run "net stop wuauserv /y"
		Set oShell = Nothing
	End If	
End Function

Function CheckWindowsUpdate()
	'Check to see if Windows Update Service needs to be enabled and/or stopped at end
	Dim objWMIService, colComputer, objComputer, strService, colServices, objService
	Dim strServiceStatus, strServiceMode
	
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
	End If
End Function


Function GetAllCabs
	' Returns a dictionary of patch scan cabinet files
	Dim dictCabs,objFSO,objCabFolder,objFile,strToolsDir
	Dim strExtraCabDir,strCabPath
	Set dictCabs = CreateObject("Scripting.Dictionary")
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strToolsDir = GetTaniumDir("Tools")
	strExtraCabDir = strToolsDir&"ExtraPatchCabs"
	If CustomCabSupportEnabled Then
		tLog.Log "Custom Patch Scan Cab Support is enabled"
		If objFSO.FolderExists(strExtraCabDir) Then
			tLog.log "Extra Cab Folder exists at "&Chr(34)&strExtraCabDir&Chr(34)&", looking for additional cab files"
			Set objCabFolder = objFSO.GetFolder(strExtraCabDir)
			For Each objFile In objCabFolder.Files
				If LCase(Right(objFile.Name,4)) = ".cab" Then
					If Not dictCabs.Exists(objFile.Path) Then
						dictCabs.Add objFile.Path,objFile.Name
						tLog.log "Found extra cab file " & objFile.Path
					End If
				End If
			Next
		Else
			tLog.log "Extra Cab Folder "&Chr(34)&strExtraCabDir&Chr(34)&" not found, looking for additional cab files"
		End If
	End If
	
	' Now add the default distributed wsusscn2.cab
	strCabPath = strToolsDir & "wsusscn2.cab"
	If objFSO.FileExists(strCabPath) And Not dictCabs.Exists(strCabPath) Then
		dictCabs.Add strCabPath,"wsusscn2.cab"
	End If
	Set GetAllCabs = dictCabs
End Function 'GetAllCabs

Function GetCleanXPFileFromURL(strURL)
	' a custom support URL may look like
	' http://download.windowsupdate.com/msdownload/update/csa/secu/2014/05/windowsxp-kb2953522-x86-custom-enu_c6820e06d430fa90266689e71652327c057737ea.exe
	' return only windowsxp-kb2953522-x86-custom-enu.exe
	
	Dim intLastDotPos,intLastUnderscorePos,strUselessData
	Dim intLastForwardSlashPos,strFileName
	
	intLastForwardSlashPos = InStrRev(strURL,"/")
	strFileName = Right(strURL,Len(strURL)-intLastForwardSlashPos)
	intLastDotPos = InStrRev(strFileName,".")
	intLastUnderscorePos = InStrRev(strFileName,"_")
	strUselessData = Mid(strFileName,intLastUnderscorePos,intLastDotPos-intLastUnderscorePos)
	strFileName = Replace(strFileName,strUselessData,"")

	GetCleanXPFileFromURL = strFileName
End Function 'GetCleanXPFileFromURL


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
End Function

Function EnsureSuffix(strIn,strSuffix)

	If Not Right(strIn,Len(strSuffix)) = strSuffix Then
		EnsureSuffix = strIn&strSuffix
	Else
		EnsureSuffix = strIn
	End If
	
End Function 'EnsureSuffix

Function RemoveSuffix(strIn,strSuffix)

	If Right(strIn,Len(strSuffix)) = strSuffix Then
		RemoveSuffix = Left(strIn,Len(strIn)-Len(strSuffix))
	Else
		RemoveSuffix = strIn
	End If
	
End Function 'RemoveSuffix

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

Sub RecordFailureCount(intFailureCount)
		tContentReg.ValueName = "PatchInstallFailureCount"
		tContentReg.Data = intFailureCount
		On Error Resume Next
		tContentReg.Write
		If Err.Number <> 0 Then
			tLog.Log "Error: Could not record install failure count ("&intFailureCount&") in registry, error was " & Err.Description
			Err.Clear
			tContentReg.ErrorClear
		End If
End Sub 'RecordFailureCount

Function GetFilesListInPatchesDirOrQuit(strPatchesDir)
	Dim arrFilesInDirAtStart
	arrFilesInDirAtStart = Array()

	tLog.Log "Examining "&Chr(34)&strPatchesDir&Chr(34)&" for patches to install. Any found files will be deleted after the installation is started and completed." 
	Dim objFSO,folder,files,file
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If Not objFSO.FolderExists(strPatchesDir) Then
		tLog.Log "Patches folder " & strPatchesDir & " does not exist, cannot determine files contained within at start, quitting."
		WScript.Quit
	End If
	Set folder = objFSO.GetFolder(strPatchesDir)
	Set files = folder.Files

	If files.Count = 0 Then
		tLog.log "No patches to install"
		If LCase(TryFromDict(dictPConfig,"PostOption","Always")) = "always" Then
			tLog.log "Running post actions"
			RunFilesInDir("post")
		End If
		WScript.Quit
	End If
		
	' If any files exist, appropriately size files array
	l = 0
	ReDim Preserve arrFilesInDirAtStart(files.Count - 1)
	' fill array - snapshot in time of files to be deleted after job
	tLog.Log "Patch Install starting for following files:"
	For Each file In files
		tLog.Log file.Path
		arrFilesInDirAtStart(l) = file.Path
		l = l + 1
	Next
	GetFilesListInPatchesDirOrQuit = arrFilesInDirAtStart
End Function 'GetFilesListInPatchesDirOrQuit


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

Function MaintenanceWindowEnabledAndValid
' This function echos back any bad values for maintenance window

	Const HKLM = &h80000002
	
	Dim strMaintenanceWindowRegKey, strComputer
	Dim strFrequency, strStartDay, strStartTime
	Dim strEndDay, strEndTime, strEnabled
	Dim objReg,bIsValidAndEnabled
	
	' Assume results are valid
	bIsValidAndEnabled = True
	
	' Set up access to registry via WMI
	strComputer = "."
	
	Set objReg = _
		GetObject("winmgmts:{impersonationLevel=impersonate}!\\" _ 
		& strComputer & "\root\default:StdRegProv")
	tLog.log "Maintenance Window Value Validation"
	strMaintenanceWindowRegKey = GetTaniumRegistryPath & "\MaintenanceWindow"
	If Not RegKeyExists(objReg, HKLM, strMaintenanceWindowRegKey) Then
		tLog.log "Maintenance Window is <Not Defined>"
		bIsValidAndEnabled = False
		' Cleanup
		Set objReg = Nothing		
		Exit Function
	End If
	
	objReg.GetStringValue HKLM, strMaintenanceWindowRegKey, "Enabled", strEnabled
	If LCase(strEnabled) = "no" Then
		' In this case, the window is disabled, so allow patch install to continue
		bIsValidAndEnabled = False
	Else
		If IsNull(strEnabled) Or strEnabled = "" Then
			tLog.log "Invalid 'Enabled' value: not set"
			bIsValidAndEnabled = False
		ElseIf Not ValidYesNo(strEnabled) Then
			tLog.log "Invalid 'Enabled' value: " & strEnabled
			bIsValidAndEnabled = False
		End If
	End If

	objReg.GetStringValue HKLM, strMaintenanceWindowRegKey, "Frequency", strFrequency
	If IsNull (strFrequency) Or strFrequency = "" Then 
		tLog.log "Invalid 'Frequency' value: not set"
		bIsValidAndEnabled = False
	ElseIf Not ValidFrequency(strFrequency) Then
		tLog.log "Invalid 'Frequency' value: " & strFrequency
		bIsValidAndEnabled = False
	End If
		
	objReg.GetStringValue HKLM, strMaintenanceWindowRegKey, "StartDay", strStartDay
	If IsNull (strStartDay) Or strStartDay = "" Then 
		tLog.log "Invalid 'StartDay' value: not set"
		bIsValidAndEnabled = False
	ElseIf Not ValidDay(strStartDay) Then
		tLog.log "Invalid 'StartDay' value: " & strStartDay
		bIsValidAndEnabled = False
	End If

	objReg.GetStringValue HKLM, strMaintenanceWindowRegKey, "StartTime", strStartTime
	If IsNull (strStartTime) Or strStartTime = "" Then 
		tLog.log "Invalid 'StartTime' value: not set"
		bIsValidAndEnabled = False
	ElseIf Not ValidTime(strStartTime) Then
		tLog.log "Invalid 'StartTime' value: " & strStartTime
		bIsValidAndEnabled = False
	End If
	
	objReg.GetStringValue HKLM, strMaintenanceWindowRegKey, "EndDay", strEndDay
	If IsNull (strEndDay) Or strEndDay = "" Then 
		tLog.log "Invalid 'EndDay' value: not set"
		bIsValidAndEnabled = False
	ElseIf Not ValidDay(strEndDay) Then
		tLog.log "Invalid 'EndDay' value: " & strEndDay
		bIsValidAndEnabled = False
	End If
	
	objReg.GetStringValue HKLM, strMaintenanceWindowRegKey, "EndTime", strEndTime
	If IsNull (strEndTime) Or strEndTime = "" Then 
		tLog.log "Invalid 'EndTime' value: not set"
		bIsValidAndEnabled = False
	ElseIf Not ValidTime(strEndTime) Then
		tLog.log "Invalid 'EndTime' value: " & strEndTime
		bIsValidAndEnabled = False
	End If
	
	MaintenanceWindowEnabledAndValid = bIsValidAndEnabled
	' Cleanup
	Set objReg = Nothing
End Function 'MaintenanceWindowEnabledAndValid

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


Function CustomCabSupportEnabled
	' Looks at registry to determine whether custom cab support is enabled
	' This is built for Windows XP patching post official support
	CustomCabSupportEnabled = False
	tContentReg.ValueName = "CustomCabSupport"
	On Error Resume Next
	strCustomCabSupportVal = LCase(tContentReg.Read)
	If LCase(strCustomCabSupport) = "true" Or LCase(strCustomCabSupport) = "yes" Then
		CustomCabSupportEnabled = True
	End If
	On Error Goto 0
End Function 'CustomCabSupportEnabled


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


Function ValidDay(strDay)
'' This function will check that a day passed in
'' is one of the allowed values
	
	Select Case LCase(strDay)
		Case "sun"
			ValidDay = True
		Case "mon"
			ValidDay = True
		Case "tue"
			ValidDay = True
		Case "wed"
			ValidDay = True
		Case "thu"
			ValidDay = True
		Case "fri"
			ValidDay = True
		Case "sat"
			ValidDay = True
		Case Else
			ValidDay = False
	End Select
End Function 'ValidDay

Function ValidYesNo(strYesNo)
'' This function will check that a value passed in
'' is yes or no as expected
	
	Select Case LCase(strYesNo)
		Case "yes"
			ValidYesNo = True
		Case "no"
			ValidYesNo = True
		Case Else
			ValidYesNo = False
	End Select
	
End Function 'ValidYesNo

Function ValidFrequency(strFrequency)
'' This function will check that a frequency value
'' passed in is what is expected
	
	Select Case LCase(strFrequency)
		Case "every week"
			ValidFrequency = True
		Case "first week"
			ValidFrequency = True
		Case "second week"
			ValidFrequency = True
		Case "third week"
			ValidFrequency = True
		Case "fourth week"
			ValidFrequency = True
		Case "every even"
			ValidFrequency = True
		Case "every odd"
			ValidFrequency = True						
		Case Else
			ValidFrequency = False
	End Select
	
End Function 'ValidFrequency

Function ValidTime(strTime)
'' This function ensures that the time passed in
'' follows a particular format

	Dim dateTest
	
	If Not InStr(strTime, ":") > 0 Then	
		ValidTime = False
		Err.Number = 0
		Exit Function
	End If
	
	On Error Resume Next ' this could error on bad time format

	dateTest = CDate(strTime)

	If Err.Number <> 0 Then
		ValidTime = False
		Err.Number = 0
		Exit Function
	End If
	On Error Goto 0
	
	ValidTime = True

End Function 'ValidTime

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

Function GetTZBias
' This functiong returns the number of minutes
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
	intCommandCount = CommandCount("cscript.exe","install-patches.vbs")
	
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
		tLog.log "Patch install not running, continuing"
	Else
		tLog.log "Patch install currently running, won't install concurrently - Quitting"
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


Sub QuitIfConfiguredToObeyMaintenanceWindows(intMaxWindowDays)
	If Not TryFromDict(dictPConfig,"IgnoreMaintenanceWindowing",False) Then
		' Set up Maintenance Window variables
		Dim dateNow
		dateNow = Now() ' capture this at start
		
		' max days a maintenance window can span
		' Will work with a value of 7
		intMaxWindowDays = 3
		' A 3 value means if it starts 2/1/2012 at midnight, the weekday abbrev
		' will not be greater than Sun for it to not report an error
		
		If MaintenanceWindowEnabledAndValid = True Then
			' If we are not in a maintenance window, we exit and do nothing
			If Not ( InMaintenanceWindow(dateNow, intMaxWindowDays, "") Or InMaintenanceWindow(dateNow, intMaxWindowDays, "Alt") ) Then
				tLog.log "Not in a Maintenance Window, Exiting"
				WScript.Quit
			Else
				tLog.log "We are in a maintenance window, installing patches"
			End If
		Else
			If TryFromDict(dictPConfig,"InstallWithoutMaintenanceWindowSet",False) = False Then
				tLog.log "Cannot find a valid maintenance window and it is required to be set"
				tLog.log "Quitting"
				WScript.Quit
			Else
				tLog.log "Maintenance Window invalid or disabled, installing patches"
			End If
		End If
	Else
		tLog.log "Maintenance Windowing is ignored / overridden"
	End If 'Ignore of maintenance windowing
End Sub 'QuitIfConfiguredToObeyMaintenanceWindows


Function InMaintenanceWindow(dateIn, intMaxWindowDays, strValPrefix)
' This function returns true or false indicating whether we're in
' a maintenance window
' dateIn should usually be Now(), not Date().  Hours are taken into account

	Const HKLM = &h80000002
	
	Dim strMaintenanceWindowRegKey, strComputer
	Dim strFrequency, strStartDay, strStartTime
	Dim intStartDay, intCurDay, intEndDay
	Dim strEndDay, strEndTime, strEnabled, strFrequencyWord
	Dim dateStartDate, dateStartDay, dateEndDate
	Dim strErrorMessage
	
	Dim objReg
	
	' Set up access to registry via WMI
	strComputer = "."
	
	Set objReg = _
		GetObject("winmgmts:{impersonationLevel=impersonate}!\\" _ 
		& strComputer & "\root\default:StdRegProv")

	strMaintenanceWindowRegKey = GetTaniumRegistryPath & "\MaintenanceWindow"
	If Not RegKeyExists(objReg, HKLM, strMaintenanceWindowRegKey) Then
		InMaintenanceWindow =  False
		tLog.log "Maintenance Window is <Not Defined>"
		' Cleanup
		Set objReg = Nothing		
		Exit Function
	End If
	
	objReg.GetStringValue HKLM, strMaintenanceWindowRegKey, "Enabled", strEnabled
	If IsNull(strEnabled) Or Not (LCase(strEnabled) = "no" Or LCase(strEnabled) = "yes") Then
		InMaintenanceWindow = False
		tLog.log "Error: Enabled value not set or malformed" & "  " & LCase(strEnabled)
		' Cleanup
		Set objReg = Nothing
		Exit Function
	End If
		
	If LCase(strEnabled) = "no" Then
		InMaintenanceWindow = False
		tLog.log strValPrefix&"Maintenance Window Disabled"
		' Cleanup
		Set objReg = Nothing
		Exit Function
	End If		
	
	' Assume the window is enabled, and let's pull values
	objReg.GetStringValue HKLM, strMaintenanceWindowRegKey, "Frequency", strFrequency
	If IsNull (strFrequency) Then 
		strFrequency = "Frequency not set"
		tLog.log "Error: " & strFrequency
		Exit Function
	End If
	objReg.GetStringValue HKLM, strMaintenanceWindowRegKey, strValPrefix&"StartDay", strStartDay
	If IsNull (strStartDay) Or DayAbbrevToInt(strStartDay) = False Then 
		strStartDay = strValPrefix&"StartDay not set or malformed"
		tLog.log "Error: " & strStartDay
		Exit Function
	End If	
	objReg.GetStringValue HKLM, strMaintenanceWindowRegKey, strValPrefix&"StartTime", strStartTime
	If IsNull (strStartTime) Or Not IsDate(strStartTime) Then 
		strStartTime = strValPrefix&"StartTime not set or malformed"
		tLog.log "Error: " & strStartTime
		Exit Function
	End If
	objReg.GetStringValue HKLM, strMaintenanceWindowRegKey, strValPrefix&"EndDay", strEndDay
	If IsNull (strEndDay) Or DayAbbrevToInt(strEndDay) = False Then 
		strEndDay = strValPrefix&"EndDay not set or malformed"
		tLog.log "Error: " & strEndDay	
		Exit Function
	End If
	objReg.GetStringValue HKLM, strMaintenanceWindowRegKey, strValPrefix&"EndTime", strEndTime
	If IsNull (strEndTime) Or Not IsDate (strEndTime) Then 
		strEndTime = strValPrefix&"EndTime not set or malformed"
		tLog.log "Error: " & strEndTime
		Exit Function
	End If
	
	
		' Some debug output
	'tLog.log "Based on the date " & FormatDateTime(dateIn,1) _
	'	 & " at " & FormatDateTime(dateIn,4) & " (military time)"
	'tLog.log "Calculating standard maintenance window starting " & _ 
	'	strStartDay & " at " & strStartTime & " and ending " & _ 
	'	strEndDay & " at " & strEndTime	& " with frequency " & _ 
	'	Chr(34) & strFrequency & Chr(34)
			
	
	' Start calculating start date.

	' StartDate is previous occurrence or next occurrence of a day in a week
	' (and that includes runtime day) plus start time.
	' EndDate uses StartDay as a base, and is the next occurrence
	' of a day (and that includes run time day) plus the end time added on.

	Dim bWantEvenWeek,bWantOddWeek
	bWantEvenWeek = False
	bWantOddWeek = False
	If LCase(strFrequency) = "every even" Then bWantEvenWeek = True
	If LCase(strFrequency) = "every odd" Then bWantOddWeek = True

	If bWantEvenWeek Or bWantOddWeek Then
		' Consider the Every Other case - essentially even and odd weeks of the month
		Dim intWeekNumOfYear,bIsEvenWeek,bIsOddWeek
		intWeekNumOfYear = DatePart("ww",dateIn)
		If intWeekNumOfYear Mod 2 = 0 Then
			bIsEvenWeek = True
		Else
			bIsOddWeek = True
		End If
	End If
	
	If LCase(strFrequency) = "every week" Or (bWantEvenWeek Or bWantOddWeek) Then	
	' The start day is either the previous Xday or the next.
		intCurDay = DatePart("w", dateIn) ' the current day of week in int form
		intStartDay = DayAbbrevToInt(strStartDay)
		intEndDay = DayAbbrevToInt(strEndDay)
		
		' If StartDay is behind us, find the previous weekday.
		' Else go forward.
		' save the day itself (minus time) so we can subtract
		' a week from that or not and re-use it later.
		If intStartDay < intCurDay Then
			dateStartDay = PreviousWeekDayToDate(dateIn, strStartDay)
			dateStartDate = PreviousWeekDayToDate(dateIn, strStartDay) + CDate(strStartTime)
		End If
		If intStartDay >= intCurDay Then
			dateStartDay = NextWeekDayToDate(dateIn, strStartDay)
			dateStartDate = NextWeekDayToDate(dateIn, strStartDay) + CDate(strStartTime)
		End If
		
		' test if we're spanning a week (friday to monday)
		' We are if if the startDay is numerically greater 
		' than the endDay (where Sunday is 1, Monday is 2, etc ...)
		If intStartDay > intEndDay Then
			'debug output
			' tLog.log "We're spanning a week because start " & strStartDay & _
			'	" is greater than end " & strEndDay
			' '
			' In this case, the maintenance window calculations fail
			' because when it hits sunday, it will move the startDay
			' to the current week.  So if it was Friday to Monday, when Sunday
			' comes along, it will be recalculated to be the next Friday
			' and the maintenance window would be incorrectly calculated.
			'
			' To fix, test if the current day is "between" the two days
			' As an example, Start = 7, a saturday.  End = 2, a Monday.
			' current day is Sunday = 1.  Only need to test if before
			' the end day.
			If intCurDay < intEndDay Then
				' And in this case, we just want the start date to go back a week.
				'debug output	
				'tLog.log "the current day is less than the end day"
				'tLog.log "The old start date is: " & dateStartDate
				dateStartDay = dateStartDay - 7
				dateStartDate = dateStartDate - 7
				'tLog.log "The new start date is: " & dateStartDate
				' The maintenance window can now span two different weeks.	
			End If
		End If
		
		' Calculate NextWeekDayToDate based on the previously calculated startday
		dateEndDate = NextWeekDayToDate(dateStartDay, strEndDay) + CDate(strEndTime)		
		' debug output
		'If dateStartDate < dateIn Then
		'	tLog.log "Window started " & dateStartDate & ", the last" _
		'		& " " & strStartDay & " before " & dateIn
		'Else		
		'	tLog.log "Window starts " & dateStartDate & ", the next" _
		'		& " " & strStartDay & " that will occur after " & dateIn
		'End If
		
		' Consider even / odd week
		If (bWantEvenWeek And bIsOddWeek) Or (bWantOddWeek And bIsEvenWeek) Then
			' add a week to the start day and end date, will happen next week
			dateStartDate = DateAdd("ww",1,dateStartDate)
			dateEndDate = DateAdd("ww",1,dateEndDate)
		End If
	Else
		' Calculate based on the first, second, third, fourth values		
		' strFrequency must be one of the following:
			' every week
			' first week
			' second week
			' third week
			' fourth week
			
		strFrequencyWord = Split(strFrequency, " ")(0)' this pulls the first word out
		If LCase(strFrequencyWord) = "first" Or LCase(strFrequencyWord) = "second" Or _ 
				LCase(strFrequencyWord) = "third" Or LCase(strFrequencyWord) = "fourth" Then
			dateStartDate = GetXthDayOfMonth(strFrequencyWord, strStartDay, dateIn) + CDate(strStartTime)
			dateEndDate = NextWeekDayToDate(GetXthDayOfMonth(strFrequencyWord, strStartDay, dateIn), strEndDay) + CDate(strEndTime)
			' debug output
			'tLog.log "Window starts " & dateStartDate & ", the " & strFrequencyWord _
			'	& " " & strStartDay & " in the month"
		Else
			strErrorMessage = "Error: The frequency value in the registry is " _ 
				& "corrupted: " & strFrequency
			tLog.log strErrorMessage ' this is not debug, leave on
			InMaintenanceWindow = False
			'Cleanup
			Set objReg = Nothing
			Exit Function
		End If ' End scrubbing frequency string		
	End If ' end calculating based on every or first, second, etc ...

	' Debug Output
	'tLog.log "Start Date is: " & dateStartDate & ", End Date is: " & dateEndDate
			
	' If the difference between the upcoming start date and the next end date 
	' is greater than MaxWindow, return an error string to the console 
	' indicating that
	If dateEndDate - dateStartDate > intMaxWindowDays Then
		strErrorMessage = "Error: The end date " & dateEndDate & _ 
			" is greater than the maximum of " & intMaxWindowDays & _ 
			" days away from the start date " & dateStartDate
		tLog.log strErrorMessage ' this is not debug, leave on
		InMaintenanceWindow = False
		'Cleanup
		Set objReg = Nothing
		Exit Function
	End If

	
	' Error checking if the start date is same as end date
	' or the end date comes before start date.
	If dateEndDate <= dateStartDate Then	
		strErrorMessage = "Error: The end date " & dateEndDate & " is less than " & _
			"or the same as the start date " & dateStartDate
		tLog.log strErrorMessage ' this is not debug, leave on
		InMaintenanceWindow = False
		'Cleanup
		Set objReg = Nothing
		Exit Function
	End If
	
	' Also check if the difference is greater than or equal to a week.	
	' This should never happen, it would indicate a bug.
	If (dateEndDate - dateStartDate) >= 7 Then
		strErrorMessage = "Error: The end date " & dateEndDate & " is greater than " & _
			"or equal to a week away from the start date " & dateStartDate
		tLog.log strErrorMessage ' this is not debug, leave on	
		InMaintenanceWindow = False
		'Cleanup
		Set objReg = Nothing
		Exit Function
	End If
	
	' debug output
	'tLog.log "We calculated based on " & dateIn & " and it's now " & Now()
	
	' Calculate if we're in a maintenance window
	If dateStartDate <= dateIn And dateEndDate >= dateIn Then
		InMaintenanceWindow = True
	Else
		InMaintenanceWindow = False
	End If
	
	' Cleanup
	Set objReg = Nothing
End Function 'InMaintenanceWindow

Function NextWeekDayToDate(dateIn, strThreeLtrWeekDay)
' This function returns the next occurance after an input date
' of the day of the week in 3 letter abbreviated string form
' For math purposes, Sunday is day 1, Monday is 2, etc ....
' This includes the current day if it matches

	Dim intWeekDay, intCurDay, intMagicNumber
	Dim dateStripped
	
	' dateIn does not need hours:minutes and this can cause bugs.
	dateStripped = StripTimeFromDate(dateIn)
	
	' Convert the passed in week day to an int value (sun is 1, etc ..)
	intWeekDay = DayAbbrevToInt(strThreeLtrWeekDay)
	intCurDay = DatePart("w", dateStripped) ' the current day of week in int form

	' if today is the weekday passed in, return that day.	
	If intWeekDay = intCurDay Then
		NextWeekDayToDate = dateStripped
	Else
		' go forward 7 days and subtract some days to get the next weekday.
				
		' This will return the next occurrence, but not including the current day
		' if it's a match (which was handled above)
		
		intMagicNumber = 7 - intWeekDay
		NextWeekDayToDate = dateStripped + 7 - (dateStripped + intMagicNumber) mod 7
	End If
End Function 'NextWeekDayToDate


Function PreviousWeekDayToDate(dateIn, strThreeLtrWeekDay)
' This function returns the previous occurance before an input date
' of the day of the week in 3 letter abbreviated string form
' For math purposes, Sunday is day 1, Monday is 2, etc ....
' This includes the current weekday if it matches
	
	Dim intWeekDay, intCurDay, intMagicNumber
	Dim dateStripped
	
	' dateIn does not need hours:minutes and this can cause bugs.
	dateStripped = StripTimeFromDate(dateIn)
		
	' Convert the passed in week day to an int value (sun is 1, etc ..)
	intWeekDay = DayAbbrevToInt(strThreeLtrWeekDay)
	intCurDay = DatePart("w", dateStripped) ' the current day of week in int form

	' This will return the previous ocurrance of the weekday, including today if
	' it matches.
	intMagicNumber = 7 - intWeekDay
	PreviousWeekDayToDate = dateStripped - (dateStripped + intMagicNumber) mod 7

	' These sets would return the previous dates without including the current
	' weekday as a match.
	' intMagicNumber = intWeekDay - 1
	' PreviousWeekDayToDate = dateStripped - 7 + (dateStripped+intMagicNumber) mod 7

End Function 'PreviousWeekDayToDate

Function DayAbbrevToInt(strDay)
' This function takes a day abbreviation
' returns the integer value of the weekday
' with Sunday as 1, etc ...

		Select Case LCase(strDay)
		Case "sun"
			DayAbbrevToInt = vbSunday	
		Case "mon"
			DayAbbrevToInt = vbMonday
		Case "tue"
			DayAbbrevToInt = vbTuesday
		Case "wed"
			DayAbbrevToInt = vbWednesday
		Case "thu"
			DayAbbrevToInt = vbThursday
		Case "fri"
			DayAbbrevToInt = vbFriday
		Case "sat"
			DayAbbrevToInt = vbSaturday
		Case Else
			DayAbbrevToInt = False
			' This shouldn't happen												
	End Select
End Function 'DayAbbrevToInt


Function GetXthDayOfMonth(strCleanFrequency, strDay, dateIn)
' This function will return the Xth Sunday, Monday, etc ... 
' of the month we're currently in as a date data type
' strFrequency must be:
	' first
	' second
	' third
	' fourth

' There can be a fifth day of the month, but if you schedule
' for the fifth day of a month, it won't happen every month

	Dim intWeekDay, dateStartOfCurMonth, dateXthDayOfMonth
	Dim intWeekNum
	Dim dateStripped
	
	' dateIn does not need hours:minutes and this can cause bugs.
	dateStripped = StripTimeFromDate(dateIn)
	
	' dateStarOfCurMonth is used for heavy duty date math
	dateStartOfCurMonth = DateSerial(Year(dateStripped), Month(dateStripped), 0)
	' and is actually the day before the first day of the current month
	
	' intWeekDay will hold the numerical value of the day passed in
	intWeekDay = DayAbbrevToInt(strDay)


	' Though this should be clean by the time it gets here
	Select Case LCase(strCleanFrequency)
		Case "first"
			intWeekNum = 1
		Case "second"
			intWeekNum = 2
		Case "third"
			intWeekNum = 3
		Case "fourth"
			intWeekNum = 4
		Case "else"
			' we can't interpret the frequency
			' so we never report that we're in the maintenance window
			GetXthDayOfMonth = False
			Exit Function
	End Select
	
	' calculate and return
	GetXthDayOfMonth = dateStartOfCurMonth + 7*intWeekNum - ((dateStartOfCurMonth-intWeekDay) Mod 7)
End Function 'GetXthDayOfMonth

Function StripTimeFromDate(dateWithHoursAndMinutes)
' This function removes the time from a date value with time in it
' by using datepart and reconstructing with dateserial
	StripTimeFromDate = DateSerial(DatePart("yyyy",dateWithHoursAndMinutes),DatePart("m",dateWithHoursAndMinutes),DatePart("d",dateWithHoursAndMinutes))
End Function 'StripTimeFromDate

Function DeleteInstallingCurrentlyFile
	' deletes the installingcurrently file if it exists

	Dim fso,strScanDir,strInstalilngCurrentlyTextFilePath,bDeleted
	 
	 bDeleted = False	
	
	' set up and delete (if exists) the installingcurrently file
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	strScanDir = GetTaniumDir("Tools\Scans")
	strInstallingCurrentlyTextFilePath = strScanDir & "\installingcurrently.txt"
	
	If fso.FileExists(strInstallingCurrentlyTextFilePath) Then
		' delete this file in the unlikely event that it exists
		fso.DeleteFile strInstallingCurrentlyTextFilePath,True
		bDeleted = True
	End If
	
	DeleteInstallingCurrentlyFile = bDeleted
	
End Function 'DeleteInstallingCurrentlyFile


Function ReinterpretSignedAsUnsigned(ByVal x)
	  If x < 0 Then x = x + 2^32
	  ReinterpretSignedAsUnsigned = x
End Function 'ReinterpretSignedAsUnsigned

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

Function IsPatchApprovalAware
' Returns true if a .dat file exists in the PatchApproval directory
' This is part of advanced patch / workbench
	Dim strWhitelistDir
	Dim bAware : bAware = False
	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objFolder, colFiles, approvalFile, strApprovalsDir
	strApprovalsDir = GetTaniumDir("Tools\PatchMgmt")
	strWhitelistDir = strApprovalsDir&"\Whitelist"
	bAware = False
	If objFSO.FolderExists(strWhitelistDir) Then	
		Set objFolder = objFSO.GetFolder(strWhitelistDir)
		Set colFiles = objFolder.Files
		For Each approvalFile In colFiles
			If Mid(approvalFile.name,len(approvalFile.name)-3,4) = ".dat" Then
				bAware = True
				Exit For
			End If
		Next				
	Else
		bAware = False
	End If
	IsPatchApprovalAware = bAware
End Function 'IsPatchApprovalAware

Function GetActiveRebootForLists
' Loop through end user reboot key and return
' dictionary of any 'dormant' reboot jobs
	Const DATE_UPDATED=0, REBOOT_STYLE=1, TEMPLATE=2, COUNTDOWN_TIME_IN_MIN=3, ADD_TIME_IN_MIN=4, _
	    ADD_TIMES_ALLOWED=5, REBOOT_MSG=6, ORIG_LAST_BOOT_TIME = 7, LAST_CHANCE_TIME = 8, ORIG_GUID = 9, DORMANT = 10
	
	Dim strRebootGUID,strRebootRegValue,dictAction,strItem,dictOut
	Set dictOut = CreateObject("Scripting.Dictionary")
	
	For Each strRebootGUID In dictSystemRebootValues.Keys ' A dictionary of dictionaries
		Set dictAction = dictSystemRebootValues.Item(strRebootGUID)
		For Each strRebootRegValue In dictAction.Keys
			' tLog.log strRebootRegValue&": "&dictAction.item(strRebootRegValue)
			If strRebootRegValue = DORMANT And LCase(dictAction.Item(strRebootRegValue)) = "true" Then
				' a dormant reboot list
				If Not dictOut.Exists(strRebootGUID) Then
					dictOut.Add strRebootGUID,1
				End If
			End If
		Next
	Next
		
	Set GetActiveRebootForLists = dictOut
	
End Function 'GetActiveRebootForLists


Function GetSystemRebootValues(table) 

	' Indexes for reboot information from registry
	Const DATE_UPDATED=0, REBOOT_STYLE=1, TEMPLATE=2, COUNTDOWN_TIME_IN_MIN=3, ADD_TIME_IN_MIN=4, _
	    ADD_TIMES_ALLOWED=5, REBOOT_MSG=6, ORIG_LAST_BOOT_TIME = 7, LAST_CHANCE_TIME = 8, ORIG_GUID = 9, DORMANT = 10
	' Types of reboot that can be directed
	Const IMMEDIATE_STYLE="immediate", COUNTDOWN_STYLE="countdown", SUGGEST_STYLE="suggest", ANNOY_STYLE="annoy"
	' Location of all of the reboot tools
	Const EUT_DIR="Tools\EUT"
	' Reg path for all of the reboot parameters
	Const REBOOT_MANAGEMENT_REG = "\RebootManagement"
	' Special GUID entry, for the current executing action
	Const CURRENT_ACTION_KEY = "current_action"

    Const HKLM = &h80000002
    
    Dim objReg,strRegPath,arrSubKeys,strKey
    Dim dictAction
    Dim strDateUpdated,strRebootStyle,strTemplate,strCountdownTimeInMin,strAddTimeInMin,strAddTimesAllowed, strRebootMessage, _
        strOrigLastBootTime, strLastChanceTime, strOrigGuid, strDormant

    Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")   
    strRegPath = GetTaniumRegistryPath & REBOOT_MANAGEMENT_REG

    If RegKeyExists(objReg, HKLM, strRegPath) Then
        objReg.EnumKey HKLM, strRegPath, arrSubKeys 

        If Not IsNull(arrSubKeys) Then
            For Each strKey in arrSubKeys
                Set dictAction = CreateObject("Scripting.Dictionary")
                strAddTimeInMin = ""
                objReg.getStringValue HKLM,strRegPath & "\" & strKey,"Dormant", strDormant
                objReg.GetStringValue HKLM,strRegPath & "\" & strKey,"DateUpdated", strDateUpdated 
                objReg.GetStringValue HKLM,strRegPath & "\" & strKey,"RebootStyle", strRebootStyle 
                objReg.GetStringValue HKLM,strRegPath & "\" & strKey,"Template", strTemplate
                objReg.GetStringValue HKLM,strRegPath & "\" & strKey,"CountdownTimeInMin", strCountdownTimeInMin 
                objReg.GetStringValue HKLM,strRegPath & "\" & strKey,"AddTimeInMin", strAddTimeInMin 
                objReg.GetStringValue HKLM,strRegPath & "\" & strKey,"AddTimesAllowed", strAddTimesAllowed              
                objReg.GetStringValue HKLM,strRegPath & "\" & strKey,"RebootMessage", strRebootMessage
                
                If strDormant = "true" Then 
                	dictAction.Add DORMANT, True
                Else 
                	dictAction.Add DORMANT, False
                End If 
                
                dictAction.Add DATE_UPDATED, strDateUpdated
                dictAction.Add REBOOT_STYLE, strRebootStyle
                ' set allowed empty values to empty string, if Null
                If IsNull(strTemplate) Then strTemplate = ""
                dictAction.Add TEMPLATE, strTemplate
                If IsNull(strCountdownTimeInMin) Then strCountdownTimeInMin = ""
                dictAction.Add COUNTDOWN_TIME_IN_MIN, strCountdownTimeInMin
                If IsNull(strAddTimeInMin) Then strAddTimeInMin = ""
                dictAction.Add ADD_TIME_IN_MIN, strAddTimeInMin
                If IsNull(strAddTimesAllowed) Then strAddTimesAllowed = ""
                dictAction.Add ADD_TIMES_ALLOWED, strAddTimesAllowed
                If IsNull(strRebootMessage) Then strRebootMessage = ""
                dictAction.Add REBOOT_MSG, strRebootMessage
                If (strKey = CURRENT_ACTION_KEY) then
                    objReg.GetStringValue HKLM,strRegPath & "\" & strKey,"OrigLastBootTime", strOrigLastBootTime 
                    objReg.GetStringValue HKLM,strRegPath & "\" & strKey,"LastChanceTime", strLastChanceTime
                    objReg.GetStringValue HKLM,strRegPath & "\" & strKey,"OrigGuid", strOrigGuid
                    
                    dictAction.Add ORIG_LAST_BOOT_TIME, strOrigLastBootTime
                    dictAction.Add LAST_CHANCE_TIME, strLastChanceTime
                    dictAction.Add ORIG_GUID, strOrigGuid
                End If              
                table.Add strKey, dictAction
            Next    
        End If
    End If  
End Function 'GetSystemRebootValues

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


Sub RunPostAndQuit
' solves quit without post issue
	RunFilesInDir("post")
	WScript.Quit
End Sub 'RunPostAndQuit

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

Function ApprovedPatchRequired 
' Returns true if any patch in any approval .dat file is required - if it's 'all
' the case where someone downloads for the specific list, look in the reboot
' key and get the active lists - approved patches required for all active
' build logic for active lists - use andrew's sensors - applicable patches for approved lists
' and do this for all as well so we do the same code in both places

	Dim bRequired : bRequired = False
	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objFolder, colFiles, approvalFile, objReadTextFile, strLine, strID
	Dim strToolsDir, strPatchResultsFile, strInstalledStatus, arrLine, strApprovalsDir
	Dim dictListsAndRequiredStatus
	Set dictListsAndRequiredStatus = CreateObject("Scripting.Dictionary")
	
	Dim dictApprovedPatchIDs : Set dictApprovedPatchIDs = CreateObject("Scripting.Dictionary")
	
	' read contents of all approval files into dictionary - ensuring unique values done once
	strToolsDir = GetTaniumDir("Tools")
	strPatchResultsFile = strToolsDir&"Scans\patchresultsreadable.txt"
	strApprovalsDir = strToolsDir&"PatchApproval"
	
	If Not objFSO.FileExists(strPatchResultsFile) Then
		ApprovedPatchRequired = False
		WScript.Echo "Warning: Cannot find patch results file to determine whether " _
			& " an approved patch is required"
		Exit Function
	End If
	
	bRequired = False
	If objFSO.FolderExists(strApprovalsDir) Then
		Set objFolder = objFSO.GetFolder(strApprovalsDir)
		Set colFiles = objFolder.Files
		For Each approvalFile In colFiles
			If Mid(approvalFile.name,len(approvalFile.name)-3,4) = ".dat" Then
				On Error Resume Next ' in case files are in flux during read in
				Set objReadTextFile = objFSO.OpenTextFile(approvalFile.Path,1)
				Do While Not objReadTextFile.AtEndOfStream
					strLine = Trim(objReadTextFile.ReadLine)
					strID = Split(strLine,"|")(0)
					' add IDs of appropriate length
					If Not dictApprovedPatchIDs.Exists(strID) And Len(strID) = 32 Then
						dictApprovedPatchIDs.add strID,1
					End If
				Loop
				objReadTextFile.Close
				On Error Goto 0
			End If
		Next

		' Read "Not Installed" patches to see if any are approved
		If objFSO.FileExists(strPatchResultsFile) Then
			Set objReadTextFile = objFSO.OpenTextFile(strPatchResultsFile,1)
			Do While Not objReadTextFile.AtEndOfStream
				strLine = LCase(Trim(objReadTextFile.ReadLine))
				arrLine = Split(strLine,"|")
				strID = ""
				strInstalledStatus = ""
				If UBound(arrLine) > 10 Then
					strInstalledStatus = arrLine(6)
					strID = arrLine(11)
				End If
				If strInstalledStatus = "not installed" And _
						dictApprovedPatchIDs.Exists(strID) Then
					bRequired = True
					Exit Do
				End If	
			Loop
		Else ' somehow the results file disappeared
			bRequired = False
		End If
	Else ' no Tools\PatchMgmt folder, patch approvals not in use on host
		bRequired = False
	End If
	
	ApprovedPatchRequired = bRequired
End Function 'ApprovedPatchRequired

Function GetApprovalListsWithRequiredStatus
' Returns a dictionary of names of approval lists and whether any patch is required
' or the list is fully installed.
' A dictionary object with a key value of List Name, data is either 
' "HasRequiredPatches" or "NoRequiredPatches"
' Also sets the global variable bApprovedPatchStillRequired

	Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim objFolder,colFiles,approvalFile,objReadTextFile,strLine,strID,strApprovalList
	Dim strToolsDir,strPatchResultsFile,strInstalledStatus,arrLine,strApprovalsDir
	Dim strPatchTitle,bRequiredInList
	Dim dictRequiredPatchIDs : Set dictRequiredPatchIDs = CreateObject("Scripting.Dictionary")
	Dim dictListsAndRequiredStatus
	Set dictListsAndRequiredStatus = CreateObject("Scripting.Dictionary")
	Dim objReg : Set objReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
	Const REBOOT_MANAGEMENT_REG = "\RebootManagement"
	Const HKLM = &h80000002
	
	Dim dictApprovedPatchIDs : Set dictApprovedPatchIDs = CreateObject("Scripting.Dictionary")
	
	strToolsDir = GetTaniumDir("Tools")
	strPatchResultsFile = strToolsDir&"Scans\patchresultsreadable.txt"
	strApprovalsDir = strToolsDir&"PatchApproval"
	
	If Not objFSO.FileExists(strPatchResultsFile) Then
		WScript.Echo "Warning: Cannot find patch results file to determine whether " _
			& " an approved patch is required"
		Set GetApprovalListsWithRequiredStatus = dictListsAndRequiredStatus	
		Exit Function
	End If

	' Read "Not Installed" patches to see if any are approved
	If objFSO.FileExists(strPatchResultsFile) Then
		Set objReadTextFile = objFSO.OpenTextFile(strPatchResultsFile,1)
		Do While Not objReadTextFile.AtEndOfStream
			strLine = LCase(Trim(objReadTextFile.ReadLine))
			arrLine = Split(strLine,"|")
			strID = ""
			strPatchTitle = ""
			strInstalledStatus = ""
			If UBound(arrLine) > 10 Then
				strInstalledStatus = arrLine(6)
				strID = arrLine(11)
				strPatchTitle = arrLine(0)
			End If
			If strInstalledStatus = "not installed" And _
					Not dictRequiredPatchIDs.Exists(strID) Then
				dictRequiredPatchIDs.Add strID,strPatchTitle
			End If
		Loop
	Else ' somehow the results file disappeared
		'
	End If
	
	' Loop through all approval .dat files and see if any IDs are in the
	' approved list. Note cases where the list has required patches or where the list
	' is fully installed

	Dim strRegPath,arrSubKeys,strKey,strDormant
    strRegPath = GetTaniumRegistryPath & REBOOT_MANAGEMENT_REG
    If RegKeyExists(objReg, HKLM, strRegPath) Then
        objReg.EnumKey HKLM, strRegPath, arrSubKeys 

	     If Not IsNull(arrSubKeys) Then
            For Each strKey in arrSubKeys
	            If InStr(strKey,"Patch_") Then
	            	 objReg.getStringValue HKLM,strRegPath & "\" & strKey,"Dormant", strDormant
	            	 If LCase(strDormant) = "true" Then
	            	 	strApprovalList = Mid(strKey,InStr(strKey,"_") + 1)
		            	Dim dWht,dBlk,whtLine,blkLine,arrBlk,bBlacklisted,i,patchCount
						Set dWht = GetDictionaryResults("Wht",strApprovalList,"Client")
						Set dBlk = GetDictionaryResults("Blk","AllLists","Client")
						arrBlk = dBlk.Keys()
						patchCount = 0
						For Each whtLine In dWht.Keys()
							bBlacklisted = False
							For i = 0 To UBound(arrBlk)
								If whtLine = arrBlk(i) Then
									bBlacklisted = True
								End If
							Next
							If Not bBlacklisted Then
								If Not InStr(whtLine,"Scan Error") And InStr(whtLine,"Not Installed") Then
									patchCount = patchCount + 1
								End If
							End If
						Next
					End If
					If patchCount > 0 Then
						' And only if the list is the active approval list we're updating against?
						dictListsAndRequiredStatus.Add strApprovalList,"HasRequiredPatches"
						bApprovedPatchStillRequired = True
					Else
						bApprovedPatchStillRequired = False
						dictListsAndRequiredStatus.Add strApprovalList,"NoRequiredPatches"
					End If
            	End If
            Next
         End If
	Else 
		' no Tools\PatchMgmt folder, patch approvals not in use on host
	End If
	
	Set GetApprovalListsWithRequiredStatus = dictListsAndRequiredStatus
	
End Function 'GetListsWithApprovedPatchesRequired

Function AwakenEUTRebootJobs(strGUID)
    Const EUT_DIR = "Tools\EUT", EUT_ACTIVATE_REBOOT="activate-dormant-reboots.vbs"

    Dim strActivateRebootScript, strActivateRebootCmd, objFSO
    strActivateRebootScript = GetTaniumDir(EUT_DIR) & EUT_ACTIVATE_REBOOT
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    If objFSO.FileExists(strActivateRebootScript) Then
        strActivateRebootCmd = "cscript " & Chr(34) & strActivateRebootScript & Chr(34) & " " &_
            Chr(34) & "/RebootGUID:" & strGUID & Chr(34) & " " &_
            Chr(34) & "/Prefix:Patch_" & Chr(34)
            
            LaunchProcess strActivateRebootCmd, True
    Else
        WScript.Echo "End User Tools not found.  Can not activate reboot."
    End If
    
End Function 'AwakenEUTRebootJobs

Function LaunchProcess(strCommand, bWait) 
    Dim objFSO, objShell, intResult

    Set objShell = WScript.CreateObject("WScript.Shell")
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    intResult = objShell.Run (strCommand,0,bWait)
    If intResult <> 0 Then
        WScript.Echo "Failed to run " & strCommand & " with error code " & intResult
    End If
End Function ' LaunchProcess

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
    ' WScript.Echo "Are the program files the same?: " & (LCase(strProgramFiles) = LCase(strProgramFilesX86))
    
    ' The windows directory is retrieved this way:
    strWinDir = objFso.GetSpecialFolder(WINDOWSDIR)
    'WScript.Echo "Windir: " & strWinDir
    
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
        ' WScript.Echo "Sysnative alias works, we're 32-bit mode on 64-bit vista+ or 2003/xp with hotfix"
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
    
        ' WScript.Echo "We are in 32-bit mode on a 64-bit machine"
        ' linkd.exe (from 2003 resource kit) must be in the machine's path.
        
        strMkLink = "linkd " & Chr(34) & strWinDir & "\System64" & Chr(34) & " " & Chr(34) & strWinDir & "\System32" & Chr(34)
        strX64cscriptPath = strWinDir & "\System64\cscript.exe"
        ' WScript.Echo "Link Command is: " & strMkLink
        ' WScript.Echo "And the path to cscript is now: " & strX64cscriptPath
        On Error Resume Next ' the mklink command could fail if linkd is not in the path
        ' the safest place to put linkd.exe is in the resource kit directory
        ' reskit installer adds to path automatically
        ' or in c:\Windows if you want to distribute just that tool
        
        If Not objFSO.FileExists(strX64cscriptPath) Then
            ' WScript.Echo "Running mklink" 
            ' without the wait to completion, the next line fails.
            objShell.Run strMkLink, 0, true
        End If
        On Error GoTo 0 ' turn error handling off
        If Not objFSO.FileExists(strX64cscriptPath) Then
            ' if that cscript doesn't exist, the link creation didn't work
            ' and we must quit the function now to avoid a loop situation
            ' WScript.Echo "Cannot find " & strX64cscriptPath & " so we must exit this function and continue on"
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
        
        ' WScript.Echo "Cannot relaunch in 64-bit (perhaps already there)"
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

    ' WScript.Echo "StrOriginalArgs is:" & strOriginalArgs
    If objFSO.FileExists(Wscript.ScriptFullName) Then
        strLaunchCommand = Chr(34) & Wscript.ScriptFullName & Chr(34) & " " & strOriginalArgs
        ' WScript.Echo "Script full path is: " & WScript.ScriptFullName
    Else
        ' the sensor itself will not work with ScriptFullName so we do this
        strLaunchCommand = Chr(34) & strTaniumPath & "\VB\" & WScript.ScriptName & chr(34) & " " & strOriginalArgs
    End If
    ' WScript.Echo "launch command is: " & strLaunchCommand

    ' Note:  There is typically a timeout of 1 hour, but this is eliminated here due to
    ' patch installs potentially taking a very long time
    Set objExec = objShell.Exec(strX64cscriptPath & " " & strLaunchCommand)
    
    ' skipping the two lines and space after that look like
    ' Microsoft (R) Windows Script Host Version
    ' Copyright (C) Microsoft Corporation
    '
    objExec.StdOut.SkipLine
    objExec.StdOut.SkipLine
    objExec.StdOut.SkipLine

    ' sensor output is all about stdout, so catch thae stdout of the relaunched
    ' sensor
    Wscript.Echo objExec.StdOut.ReadAll()
    
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
Function GetDictionaryResults(listType,listName,requestType)
 Dim dictApproved,patchMgmtDir,strApprovalFile,managedPatch,serverPatchMgmtDir,regPath,regValue
 If requestType = "Client" Then
	 If listType = "Wht" Then
	 	patchMgmtDir = GetTaniumDir("Tools") & "PatchMgmt\Whitelist"
	 End If
	 If listType = "Blk" Then
	 	patchMgmtDir = GetTaniumDir("Tools") & "PatchMgmt\Blacklist"
	 End If
 End If
 If requestType = "Server" Then
	Dim objReg : Set objReg = Getx64RegistryProvider()
	Const HKLM = &H80000002
	regPath = "Software\Wow6432Node\Tanium\PatchManagementController"
	objReg.GetStringValue HKLM,regPath,"InstallDir",regValue
 	serverPatchMgmtDir = regValue & "\"
 	If listType = "Wht" Then
 		patchMgmtDir = serverPatchMgmtDir & "Whitelist"
 	End If
 	If listType = "Blk" Then
 		patchMgmtDir = serverPatchMgmtDir & "Blacklist"
 	End If
 End If
 Dim objFS : Set objFS = CreateObject("Scripting.FileSystemObject")
 If Not objFS.FolderExists(patchMgmtDir) Then
 	objFS.CreateFolder(patchMgmtDir)
 End if 
 Dim dictOutput : Set dictOutput = CreateObject("Scripting.Dictionary")
 Dim approvalLists : Set approvalLists = objFS.GetFolder(patchMgmtDir)

If InStr(listName,"AllLists") > 0 Then
	For Each strApprovalFile In approvalLists.Files
		Set dictApproved = GetManagedPatches(strApprovalFile,"GUID",requestType)
		For Each managedPatch In dictApproved.Keys()
			If Not dictOutput.Exists(managedPatch) Then
				dictOutput.Add managedPatch, True
			End If
		Next
		Set dictApproved = GetManagedPatches(strApprovalFile,"Severity",requestType)
		For Each managedPatch In dictApproved.Keys()
			If Not dictOutput.Exists(managedPatch) Then
				dictOutput.Add managedPatch, True
			End If
		Next 
		Set dictApproved = GetManagedPatches(strApprovalFile,"RegEx",requestType)
		For Each managedPatch In dictApproved.Keys()
			If Not dictOutput.Exists(managedPatch) Then
				dictOutput.Add managedPatch, True
			End If
		Next   
	Next
Else
	strApprovalFile = patchMgmtDir & "\" & listName & ".dat"
	Set dictApproved = GetManagedPatches(strApprovalFile,"GUID",requestType)
	For Each managedPatch In dictApproved.Keys()
		If Not dictOutput.Exists(managedPatch) Then
			dictOutput.Add managedPatch, True
		End If
	Next
	Set dictApproved = GetManagedPatches(strApprovalFile,"Severity",requestType)
	For Each managedPatch In dictApproved.Keys()
		If Not dictOutput.Exists(managedPatch) Then
			dictOutput.Add managedPatch, True
		End If
	Next 
	Set dictApproved = GetManagedPatches(strApprovalFile,"RegEx",requestType)
	For Each managedPatch In dictApproved.Keys()
		If Not dictOutput.Exists(managedPatch) Then
			dictOutput.Add managedPatch, True
		End If
	Next  
End If

Set GetDictionaryResults = dictOutput
End Function 'GetDictionaryResults
Function GetManagedPatches(listName,entryType,requestType)
Dim whitelistDir,blacklistDir,readFile,readLine,dictReturn,dictTemp,taniumScans,allPatchList,objPatchList,strPatchTitle
Dim strPatchLine,strPatchValue,strStatus,strFileName,strPatchGuid,strURI,strSep,strGuid,strSeverity,strPatchSeverity,strRegex
Dim regPath,regValue,serverPatchMgmtDir,strPatchBulletin,strPatchDate,strPatchPackageSize,strPatchKB,strPatchCVE
Dim arrSplitRegex,strRegexColumn,strRegexValue
Dim fso : Set fso = WScript.CreateObject("Scripting.Filesystemobject")
Set dictReturn = CreateObject("Scripting.Dictionary")
Set dictTemp = CreateObject("Scripting.Dictionary")
If requestType = "Client" Then
	allPatchList = GetTaniumDir("Tools") & "Scans\patchresultsreadable.txt"
End If
If requestType = "Server" Then
	Dim objReg : Set objReg = Getx64RegistryProvider()
	Const HKLM = &H80000002
	regPath = "Software\Wow6432Node\Tanium\PatchManagementController"
	objReg.GetStringValue HKLM,regPath,"InstallDir",regValue
 	serverPatchMgmtDir = regValue & "\Data\"
 	allPatchList = serverPatchMgmtDir & "all-patches.dat"
End If
strSep = "|"

If fso.FileExists(listName) Then
	Set readFile = fso.OpenTextFile(listName,1)
	Do While Not InStr(readLine,entryType) > 0
		readLine = readFile.ReadLine
	Loop
	
	Do While Not readLine = "" And Not readFile.AtEndOfStream
		readLine = readFile.ReadLine
		If readLine = "" Then
			Exit Do
		End If
		If Not dictTemp.Exists(readLine) Then
			dictTemp.Add readLine, True
		End If	
	Loop

	Set objPatchList = fso.OpenTextFile (allPatchList,1)
	Do While Not objPatchList.AtEndOfStream
		strPatchLine = objPatchList.ReadLine
		If requestType = "Client" Then
			strPatchValue = Split (strPatchLine, strSep)
			strFileName = strPatchValue(5)
			strStatus = strPatchValue(6)
			strPatchGuid = strPatchValue(11)
			strURI = strPatchValue(4)
			strPatchSeverity = strPatchValue(1)
			strPatchTitle = strPatchValue(0)
			strPatchBulletin = strPatchValue(2)
			strPatchDate = strPatchValue(3)
			strPatchPackageSize = strPatchValue(8)
			strPatchKB = strPatchValue(9)
			strPatchCVE = strPatchValue(10)
		End If
		If requestType = "Server" Then
			strPatchValue = Split (strPatchLine, strSep)
			strPatchGuid = strPatchValue(8)
			strPatchSeverity = strPatchValue(1)
			strPatchTitle = strPatchValue(0)
			strPatchBulletin = strPatchValue(2)
			strPatchDate = strPatchValue(3)
			strPatchPackageSize = strPatchValue(5)
			strPatchKB = strPatchValue(6)
			strPatchCVE = strPatchValue(7)
		End If
		If entryType = "GUID" Then	
			For Each strGuid In dictTemp.Keys()
				If strGuid = strPatchGuid Then
					If Not dictReturn.Exists(strPatchLine) Then
						dictReturn.Add strPatchLine, True
					End If
				End If
			Next
		End If
		If entryType = "Severity" Then
			For Each strSeverity In dictTemp.Keys()
				If InStr(strSeverity,strPatchSeverity) > 0 Then
					If Not dictReturn.Exists(strPatchLine) Then
						dictReturn.Add strPatchLine, True
					End If
				End If
			Next
		End If
		If entryType = "RegEx" Then
			For Each strRegex In dictTemp.Keys()
				arrSplitRegex = Split(strRegex,"@")
				strRegexColumn = arrSplitRegex(0)
				strRegexValue = arrSplitRegex(1)
				If strRegexColumn = "Title" Then
					If RegExpMatch(strRegexValue,strPatchTitle,True,False) Then
						If Not dictReturn.Exists(strPatchLine) Then
							dictReturn.Add strPatchLine, True
						End If
					End If
				End If
				If strRegexColumn = "Bulletins" Then
					If RegExpMatch(strRegexValue,strPatchBulletin,True,False) Then
						If Not dictReturn.Exists(strPatchLine) Then
							dictReturn.Add strPatchLine, True
						End If
					End If
				End If
				If strRegexColumn = "Date" Then
					If RegExpMatch(strRegexValue,strPatchDate,True,False) Then
						If Not dictReturn.Exists(strPatchLine) Then
							dictReturn.Add strPatchLine, True
						End If
					End If
				End If
				If strRegexColumn = "Package Size" Then
					If RegExpMatch(strRegexValue,strPatchPackageSize,True,False) Then
						If Not dictReturn.Exists(strPatchLine) Then
							dictReturn.Add strPatchLine, True
						End If
					End If
				End If
				If strRegexColumn = "KB Article" Then
					If RegExpMatch(strRegexValue,strPatchKB,True,False) Then
						If Not dictReturn.Exists(strPatchLine) Then
							dictReturn.Add strPatchLine, True
						End If
					End If
				End If
				If strRegexColumn = "CVE ID" Then
					If RegExpMatch(strRegexValue,strPatchCVE,True,False) Then
						If Not dictReturn.Exists(strPatchLine) Then
							dictReturn.Add strPatchLine, True
						End If
					End If
				End If
				If strRegexColumn = "Unique ID" Then
					If RegExpMatch(strRegexValue,strPatchGUID,True,False) Then
						If Not dictReturn.Exists(strPatchLine) Then
							dictReturn.Add strPatchLine, True
						End If
					End If
				End If
			Next
		End If	
	Loop
	readFile.Close
	objPatchList.Close
End If
Set GetManagedPatches = dictReturn

End Function
Function RegExpMatch(strPattern,strToMatch,bGlobal,bIsCaseSensitive)

	Dim re
	Set re = New RegExp
	With re
	  .Pattern = strPattern
	  .Global = bGlobal
	  .IgnoreCase = Not bIsCaseSensitive
	End With
	
	RegExpMatch = re.Test(strToMatch)

End Function 'RegExpMatch
Function Getx64RegistryProvider
    ' Returns the best available registry provider:  32 bit on 32 bit systems, 64 bit on 64 bit systems
    Dim objWMIService, colItems, objItem, iArchType, objCtx, objLocator, objServices, objRegProv
    Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
    Set colItems = objWMIService.ExecQuery("Select SystemType from Win32_ComputerSystem")    
    For Each objItem In colItems
        If InStr(LCase(objItem.SystemType), "x64") > 0 Then
            iArchType = 64
        Else
            iArchType = 32
        End If
    Next
    
    Set objCtx = CreateObject("WbemScripting.SWbemNamedValueSet")
    objCtx.Add "__ProviderArchitecture", iArchType
    Set objLocator = CreateObject("Wbemscripting.SWbemLocator")
    Set objServices = objLocator.ConnectServer("","root\default","","",,,,objCtx)
    Set objRegProv = objServices.Get("StdRegProv")   
    
    Set Getx64RegistryProvider = objRegProv
End Function ' Getx64RegistryProvider


Function UseScanSourceArgValidator(arg)
	UseScanSourceArgValidator = True
	
	If Not StringInCommaSeparatedList(arg, _
		"cab,systemdefault,internet,wsus,optimal" ) Then
		UseScanSourceArgValidator = False
	End If

End Function 'UseScanSourceArgValidator

Function PostOptionArgValidator(arg)
	PostOptionArgValidator = True
	
	If Not StringInCommaSeparatedList(arg, _
		"always,patchonly" ) Then
		PostOptionArgValidator = False
	End If

End Function 'PostOptionArgValidator


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

Function PatchesDirValidator(arg)
	PatchesDirValidator = True
	Dim argTemp
	argTemp = arg 'GetTaniumDir will behave as pass by reference and modify the string

	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	On Error Resume Next
	If Not objFSO.FolderExists(GetTaniumDir(argTemp)) Then
	On Error Goto 0
		PatchesDirValidator = False
	Else
		tLog.Log "PatchesDir overridden, will look for patches in "&Chr(34)&GetTaniumDir(argTemp)&Chr(34)&", instead of the default directory " _
			&Chr(34)&GetTaniumDir("Tools\Patches")&Chr(34)
	End If

End Function 'PatchesDirValidator

Function PatchesDirTranslator(arg)

	Dim objFSO
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Dim strTaniumClientDir,intPos
	strTaniumClientDir = GetTaniumDir("")
	' Allow this to be fully specified, but internally use only the relative path
	intPos = InStr(1,arg,strTaniumClientDir,vbTextCompare)
	If intPos = 1 Then
		arg = Right(arg,Len(arg)-Len(strTaniumClientDir))
	End If
	PatchesDirTranslator = arg

End Function 'PatchesDirTranslator


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
	objWaitTimeArg.ArgName = "RandomInstallWaitTimeInSeconds"
	objWaitTimeArg.HelpText = "Waits up to X seconds before scanning"
	objWaitTimeArg.ExampleValue = "15"
	objWaitTimeArg.DefaultValue = 240
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
	
	Dim objPatchesDirArg,PatchesDirValidateRef,PatchesDirTranslateRef
	Set objPatchesDirArg = New TaniumNamedArg
	Set PatchesDirValidateRef = GetRef("PatchesDirValidator")
	Set PatchesDirTranslateRef = GetRef("PatchesDirTranslator")
	objPatchesDirArg.ArgName = "PatchesDir"
	objPatchesDirArg.HelpText = "Change the directory which hosts patch files to be installed. This must be a subdir of the Tanium Client directory."
	objPatchesDirArg.ExampleValue = "Tools\Patches"
	objPatchesDirArg.DefaultValue = "Tools\Patches"
	objPatchesDirArg.IsOptional = True
	objPatchesDirArg.ValidationFunctionReference = PatchesDirValidateRef	
	objPatchesDirArg.TranslationFunctionReference = PatchesDirTranslateRef
	ArgsParser.AddArg objPatchesDirArg

	Dim objIgnoreEUTFlagArg
	Set objIgnoreEUTFlagArg = New TaniumNamedArg
	objIgnoreEUTFlagArg.RequireYesNoTrueFalse = True
	objIgnoreEUTFlagArg.ArgName = "IgnoreEUT"
	objIgnoreEUTFlagArg.HelpText = "Ignore End User Tools reboot integration hooks"
	objIgnoreEUTFlagArg.ExampleValue = "Yes,No"
	objIgnoreEUTFlagArg.DefaultValue = "No"
	objIgnoreEUTFlagArg.IsOptional = True
	ArgsParser.AddArg objIgnoreEUTFlagArg

	Dim objInstallWithoutMaintenanceWindowSetFlagArg
	Set objInstallWithoutMaintenanceWindowSetFlagArg = New TaniumNamedArg
	objInstallWithoutMaintenanceWindowSetFlagArg.RequireYesNoTrueFalse = True
	objInstallWithoutMaintenanceWindowSetFlagArg.ArgName = "InstallWithoutMaintenanceWindowSet"
	objInstallWithoutMaintenanceWindowSetFlagArg.HelpText = "Do not require maintenance windows to be set. If set to Yes, if a window is set, observe it" _
		&" and abort patch install when outside window. If set to No, no patches are installed unless there is a set and valid window on the host." _
		&" This is a fail-closed approach to patch installations. Must have Maintenance Window content enabled."
	objInstallWithoutMaintenanceWindowSetFlagArg.ExampleValue = "Yes,No"
	objInstallWithoutMaintenanceWindowSetFlagArg.DefaultValue = "Yes"
	objInstallWithoutMaintenanceWindowSetFlagArg.IsOptional = True
	ArgsParser.AddArg objInstallWithoutMaintenanceWindowSetFlagArg

	Dim objIgnoreMaintenanceWindowingFlagArg
	Set objIgnoreMaintenanceWindowingFlagArg = New TaniumNamedArg
	objIgnoreMaintenanceWindowingFlagArg.RequireYesNoTrueFalse = True
	objIgnoreMaintenanceWindowingFlagArg.ArgName = "IgnoreMaintenanceWindowing"
	objIgnoreMaintenanceWindowingFlagArg.HelpText = "Do not consider maintenance windowing at all, and always install the patches."
	objIgnoreMaintenanceWindowingFlagArg.ExampleValue = "Yes,No"
	objIgnoreMaintenanceWindowingFlagArg.DefaultValue = "No"
	objIgnoreMaintenanceWindowingFlagArg.IsOptional = True
	ArgsParser.AddArg objIgnoreMaintenanceWindowingFlagArg


	Dim objRunInteractivelyFlagArg
	Set objRunInteractivelyFlagArg = New TaniumNamedArg
	objRunInteractivelyFlagArg.RequireYesNoTrueFalse = True
	objRunInteractivelyFlagArg.ArgName = "RunInteractively"
	objRunInteractivelyFlagArg.HelpText = "Allow patches to be installed via a wizard with an end user facing dialog. This will require user intervention to complete. The default WSUS / WUAPI install box will show. This is not typically done."
	objRunInteractivelyFlagArg.ExampleValue = "Yes,No"
	objRunInteractivelyFlagArg.DefaultValue = "No"
	objRunInteractivelyFlagArg.IsOptional = True
	ArgsParser.AddArg objRunInteractivelyFlagArg

	
	Dim objMaxFailuresThresholdArg
	Set objMaxFailuresThresholdArg = New TaniumNamedArg
	objMaxFailuresThresholdArg.RequireInteger = True
	objMaxFailuresThresholdArg.ArgName = "MaxFailuresThreshold"
	objMaxFailuresThresholdArg.HelpText = "Maximum number of failures tolerated before a reboot is allowed to be performed when performing an advacned patch related install"
	objMaxFailuresThresholdArg.ExampleValue = "5"
	objMaxFailuresThresholdArg.DefaultValue = 5
	objMaxFailuresThresholdArg.IsOptional = True
	ArgsParser.AddArg objMaxFailuresThresholdArg
	
	
	Dim objClearInstallResultsFlagArg
	Set objClearInstallResultsFlagArg = New TaniumNamedArg
	objClearInstallResultsFlagArg.RequireYesNoTrueFalse = True
	objClearInstallResultsFlagArg.ArgName = "ClearInstallResultsOnBadLine"
	objClearInstallResultsFlagArg.HelpText = "If the installresults file has a bad line, clear the entire file."
	objClearInstallResultsFlagArg.ExampleValue = "Yes,No"
	objClearInstallResultsFlagArg.DefaultValue = "No"
	objClearInstallResultsFlagArg.IsOptional = True
	ArgsParser.AddArg objClearInstallResultsFlagArg

	Dim objPostOptionArg,PostOptionValidateRef
	Set objPostOptionArg = New TaniumNamedArg
	Set PostOptionValidateRef = GetRef("PostOptionArgValidator")
	objPostOptionArg.ArgName = "PostOption"
	objPostOptionArg.HelpText = "Control the circumstances in which the scripts in the install-patches.vbs\Post directory - always, or only when patches are to be installed."
	objPostOptionArg.ExampleValue = "Always,PatchOnly"
	objPostOptionArg.DefaultValue = "PatchOnly"
	objPostOptionArg.IsOptional = True
	objPostOptionArg.ValidationFunctionReference = PostOptionValidateRef
	ArgsParser.AddArg objPostOptionArg

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
	objDoNotSaveOptionsArg.ExampleValue = "Yes"
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
