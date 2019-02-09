'Tanium File Version:2.2.2.0011

' uninstalls a patch by KB article
Option Explicit

x64Fix

' allow override
RunOverride

' Global classes
Dim tLog
Set tLog = New TaniumContentLog
tLog.Log "----------------Beginning Patch Uninstall----------------"

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

EnsureRunsOneCopy


'Argument handling
Dim ArgsParser
Set ArgsParser = New TaniumNamedArgsParser
ParseArgs ArgsParser


' Create a config - combination of default values and passed in arguments - for use in this script
Dim dictPConfig
Set dictPConfig = CreateObject("Scripting.Dictionary")
dictPConfig.CompareMode = vbTextCompare
' Load default values
LoadDefaultConfig ArgsParser,dictPConfig
' Read from Registry (parsed values are here now if it is 'sticky' - not an option in uninstall script)
LoadRegConfig tContentReg, dictPConfig
' Load parsed values - in case it not 'sticky'
LoadParsedConfig ArgsParser,dictPConfig

' check for / run 'pre' files
RunFilesInDir("pre")
' put files in a directory called uninstall-patch\pre
' candidates would be popup notifiers, checks to continue

' ----------------- arguments parsed -------------'
Dim dtmNow : dtmNow = Now()

'Globals needed to stop Windows Update service throughout script
Dim wuaService, wuaNeedsStop, wuaNeedsDisabled
' check state of update service
CheckWindowsUpdate()

Dim strUninstallResultsPath,strUninstallResultsReadablePath
Dim intUninstallResultsFileMode,strSep,strSepTwo,dictScanResultsByKB
Dim intDesiredUninstallResultsColumnCount,bBadUninstallResultsLines,bGoodUninstallResultsLine
Dim arrUninstallResultsLine,strUninstallResultsLine,dictUninstallResults
Dim objUninstallResultsTextFile,strResultMessage,strInfoFromScan,strOutLine
Dim strGUID,strTitle,arrResultLine,strScanDir,objFSO,bClearUninstallResultsOnBadLine
Dim bSuccess,dictScanResultsByColumnIndex,intColIndex
Dim dictRelevantPatchAndProductGUIDs,strTrulyUniqueID,strKB

Set objFSO = CreateObject("Scripting.FileSystemObject")
strSep = "|"
strSepTwo = "&&"
strScanDir = GetTaniumDir("Tools\Scans")

' create text file to log updates uninstalled by Tanium
' details about the kb being uninstalled are provided by
' the scanresults file
strUninstallResultsPath = strScanDir & "\uninstalledresults.txt"
strUninstallResultsReadablePath = strScanDir & "\uninstalledresultsreadable.txt"

If Not objFSO.FileExists(strUninstallResultsPath) Then
	objFSO.CreateTextFile strUninstallResultsPath,True
End If

Set objUninstallResultsTextFile = objFSO.OpenTextFile(strUninstallResultsPath,1,True)

' Read existing UninstallResults snapshot into a dictionary object
Set dictUninstallResults = CreateObject("Scripting.Dictionary")

' the UninstallResults file is read and appended to versus being overwritten each time.
' If we change the format of the file, we are probably changing column count. 
' If the line entry does not match the column count, do not append and instead
' overwrite the file later.
' only unique lines are read into the dictionary object, and only unique lines
' are written back.

intDesiredUninstallResultsColumnCount = 5
bBadUninstallResultsLines = False
While objUninstallResultsTextFile.AtEndOfStream = False
	bGoodUninstallResultsLine = True
	strUninstallResultsLine = objUninstallResultsTextFile.ReadLine
	arrUninstallResultsLine = Split(strUninstallResultsLine,"|")
	If IsArray(arrUninstallResultsLine) Then
		If UBound(arrUninstallResultsLine) <> (intDesiredUninstallResultsColumnCount-1) Then
			tLog.Log "bad line detected:" & strUninstallResultsLine
			bGoodUninstallResultsLine = False
			bBadUninstallResultsLines = True
		End If
		If (Not dictUninstallResults.Exists(strUninstallResultsLine)) And bGoodUninstallResultsLine Then
			dictUninstallResults.Add strUninstallResultsLine,1
		End If
	End If
Wend

If bClearUninstallResultsOnBadLine And bBadUninstallResultsLines Then ' overwrite file
	intUninstallResultsFileMode = 2 ' overwrite
	tLog.Log "will overwrite UninstallResults file"		
Else
	intUninstallResultsFileMode = 8 ' append
End If

' Will need to re-open for either writing or appending now that it's in dictionary
Set objUninstallResultsTextFile = objFSO.OpenTextFile(strUninstallResultsPath,intUninstallResultsFileMode,True)

Dim dictScanResultFiles,strResultFilePath
Set dictScanResultFiles = GetValidScanFiles
Set dictScanResultsByColumnIndex = CreateObject("Scripting.Dictionary")

' Anticipate any future support for non-KB number
Dim bByKb,strKBToUninstallNumbersOnly,strKBToUninstall
bByKb = False
strKBToUninstallNumbersOnly = TryFromDict(dictPConfig,"UninstallByKB","")
strKBToUninstall = "KB"&strKBToUninstallNumbersOnly
If strKBToUninstallNumbersOnly <> "" Then bByKb = True
If bByKB Then intColIndex = 9

For Each strResultFilePath In dictScanResultFiles.Keys
	' build dictScanResultsByKB for each scan results file
	BuildScanResultsDictByColumnIndex dictScanResultsByColumnIndex,strResultFilePath,intColIndex
Next

' Uninstall for Office Patches requires product GUID and Patch GUID
' Build up objects which will return the Product GUID and Patch GUIDs, comma separated
' for a particular kb article number (column index 9)
If bByKB Then
	Set dictRelevantPatchAndProductGUIDs = CreateObject("Scripting.Dictionary")
	BuildPatchGUIDToProductGUIDDict strKBToUninstall,intColIndex,dictScanResultsByColumnIndex, dictRelevantPatchAndProductGUIDs
Else
	tLog.Log "Unsupported Uninstall method"
	StopWindowsUpdate
	WScript.Quit
End If

Dim strPatchGUID,strProductGUID,strKey
If dictRelevantPatchAndProductGUIDs.Count > 0 Then
	bSuccess = True
	Dim bLocalSuccess
	' we must assume we will uninstall the patch using the product/patch GUID MSI method
	For Each strKey In dictRelevantPatchAndProductGUIDs.Keys ' a strSep separated patchguid|productguid
		' Require success for each update (AND together)
		strPatchGUID = Split(strKey,strSep)(0)
		strProductGUID = Split(strKey,strSep)(1)
		bLocalSuccess = UninstallMSIPatch(strProductGUID,strPatchGUID)
		bSuccess = bSuccess And bLocalSuccess
	Next
Else
	bSuccess = UninstallKB(strKBToUninstall,strKBToUninstallNumbersOnly) ' doesn't always return true if it was successful
End If

If bSuccess Then
	strResultMessage = "Success"
	tLog.Log "Uninstall appears to have been successful"
	tLog.Log "Re-running patch scan"
	AccessRunPatchScan ' necessary to stop any actions targeting that patch line
Else
	strResultMessage = "Failure"
	tLog.Log "Uninstall did not appear to have been successful"
End If

tLog.Log "Writing Uninstall Results File"
' Write uninstallresults log file
strInfoFromScan = dictScanResultsByColumnIndex.Item(strKBToUninstall)
' Possibly write multiple lines - each column ID is not guaranteed to have only one line
' in scan results. Split on second delimiter
For Each strOutLine In Split(strInfoFromScan,strSepTwo)
	' Log the GUID, title, Time, and result
	arrResultLine = Split(strOutLine,strSep)
	If UBound(arrResultLine) > 11 Then
		strTitle = arrResultLine(0)
		strGUID = arrResultLine(7)
		strTrulyUniqueID = arrResultLine(11)
		strKB = arrResultLine(9)
		tLog.Log strGUID&strSep&strTitle&strSep&CStr(dtmNow) _
			&strSep&strResultMessage&strSep&strTrulyUniqueID
		objUninstallResultsTextFile.WriteLine UnicodeToAscii(strGUID&strSep&strTitle&strSep&CStr(dtmNow) _
			&strSep&strResultMessage&strSep&strTrulyUniqueID)
	Else
		tLog.Log "Bad scan result line: " & strOutLine
	End If
Next

objUninstallResultsTextFile.Close
' Make readable file
objFSO.CopyFile strUninstallResultsPath,strUninstallResultsReadablePath
StopWindowsUpdate ' will stop only if necessary (was previously off)

RunFilesInDir("post")
' put files in a directory called uninstall-patch\post
' candidates would be things like reboot if necessary, warn user of reboot

' --- End Main Line ---- '
Function AccessRunPatchScan
	Dim strVBS,fso,strCommand,objShell,objScriptExec,strResults
	
	strVbs = GetTaniumDir("Tools") & "run-patch-scan.vbs"
	
	Set fso = WScript.CreateObject("Scripting.Filesystemobject")
	
	If fso.FileExists (strVbs) Then
		
		strCommand = "cscript //T:3600 " & Chr(34) & strVbs & Chr(34)
		
		Set objShell = CreateObject("WScript.Shell")
		Set objScriptExec = objShell.Exec (strCommand)
		strResults = objScriptExec.StdOut.ReadAll
		
		''output results from running patch scan
		tLog.Log strResults
	Else
		tLog.Log strVbs & " not found"
	End If
	
End Function

Sub BuildScanResultsDictByColumnIndex(ByRef dictScanResultsToBuild,strScanResultsReadablePath,intIndex)
' orders scan results
	Dim objFSO
	Dim strScanResultsLine,objScanResultsTextFile
	Dim arrScanResultsLine,strScanColumn

	' Builds a dictionary, keyed on the supplied column index, with the details line as the item
	' column index (such as KB number column) is not necessarily unique, so use a second delimiter 
	' and append other lines to the dictionary item
	
	Set objFSO = WScript.CreateObject("Scripting.FilesystemObject")
	
	If objFSO.FileExists(strScanResultsReadablePath) Then
		Set objScanResultsTextFile = objFSO.OpenTextFile(strScanResultsReadablePath,1)
		Do Until objScanResultsTextFile.AtEndOfStream = True
			strScanResultsLine = objScanResultsTextFile.ReadLine
			' To Key on KB, use 10th field
			arrScanResultsLine = Split(strScanResultsLine,strSep)
			If UBound(arrScanResultsLine) >= intIndex Then
				strScanColumn = arrScanResultsLine(intIndex)
			
				If Not dictScanResultsToBuild.Exists(strScanColumn) Then
					dictScanResultsToBuild.Add strScanColumn,strScanResultsLine
				Else
					strScanResultsLine = dictScanResultsToBuild.Item(strScanColumn)&strSepTwo&strScanResultsLine
					dictScanResultsToBuild.Item(strScanColumn) = strScanResultsLine
				End If
			Else
				tLog.Log "Index beyond available column boundary ("&intIndex&"), possibly Bad Scan Results Line: " _ 
					& strScanResultsLine
			End If
		Loop
	Else
		tLog.Log "No patch scan results found for "&strScanResultsReadablePath _
			&", may not be able to produce uninstall results file based on item passed in"
		Exit Sub
	End If
End Sub 'BuildScanResultsDictByColumnIndex


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


Function GetValidScanFiles
	' if Custom Cab Scan support is enabled
	' Returns a dictionary of valid scan results files
	' which are those that are based on cab files which exist
	Dim dictScanResultFiles,objFSO,objCabFolder,objFile,strToolsDir
	Dim strExtraCabDir,strResultsReadablePath,strScanDir,strCabPath
	Set dictScanResultFiles = CreateObject("Scripting.Dictionary")	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strToolsDir = GetTaniumDir("Tools")
	strScanDir = GetTaniumDir("Tools\Scans")
	If CustomCabSupportEnabled Then
		strExtraCabDir = strToolsDir&"ExtraPatchCabs"
		If objFSO.FolderExists(strExtraCabDir) Then
			Set objCabFolder = objFSO.GetFolder(strExtraCabDir)
			For Each objFile In objCabFolder.Files
				If LCase(Right(objFile.Name,4)) = ".cab" Then
					strResultsReadablePath = strScanDir&"patchresultsreadable-"&objFile.Name&".txt"
					If Not dictScanResultFiles.Exists(strResultsReadablePath) Then
						dictScanResultFiles.Add strResultsReadablePath,objFile.Name
					End If
				End If
			Next
		End If
	End If
	
	' always add the default distributed wsusscn2.cab
	strCabPath = strToolsDir & "wsusscn2.cab"
	If objFSO.FileExists(strCabPath) And Not dictScanResultFiles.Exists(strCabPath) Then
		dictScanResultFiles.Add strScanDir&"patchresultsreadable.txt","wsusscn2.cab"
	End If

	Set GetValidScanFiles = dictScanResultFiles
End Function 'GetValidScanFiles

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
		tLog.Log "Scan Error: Cannot find Windows Update (wuauserv)"
		WScript.Quit
	End If
	
	If strServiceStatus = "Stopped" Then
		tLog.Log "Windows Update is stopped, will stop after Patch Scan Complete"
		wuaNeedsStop = true
	End If
	
	If strServiceMode = "Disabled" Then
		tLog.Log "Attempting to change 'Windows Update' start mode to 'Manual'"
		tLog.Log "Return code: " & WuaService.ChangeStartMode("Manual")
		wuaNeedsDisabled = True
	End If

End Function

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
			tLog.Log "Error: " & strPath & " does not exist on the filesystem"
			GetTaniumDir = False
		End If
	Else
		tLog.Log "Error: Cannot find Tanium Client path in Registry"
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
		tLog.Log "Relaunching"
		strOriginalArgs = ""
		Set objArgs = WScript.Arguments
		
		For Each strArg in objArgs
		    strOriginalArgs = strOriginalArgs & " " & strArg
		Next
		' after we're done, we have an unnecessary space in front of strOriginalArgs
		strOriginalArgs = LTrim(strOriginalArgs)
	
		strLaunchCommand = Chr(34) & strFileDir&"override\"&strFileName & Chr(34) & " " & strOriginalArgs
		' tLog.Log "Script full path is: " & WScript.ScriptFullName
		
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
		tLog.Log objExec.StdOut.ReadAll()
		
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
		tLog.Log "Found subdirectory " & strSubDirArg
		Set objFolder = objFSO.GetFolder(strSubDir)
		Set objShell = CreateObject("WScript.Shell")
		For Each objFile In objFolder.Files
			strTargetExtension = Right(objFile.Name,3)
			If strTargetExtension = "vbs" Then
				tLog.Log "Running " & objFile.Path
				Set objExec = objShell.Exec(Chr(34)&WScript.FullName&Chr(34) & "//T:1800 " & Chr(34)&objFile.Path&Chr(34))
			
				' skipping the two lines and space after that look like
				' Microsoft (R) Windows Script Host Version
				' Copyright (C) Microsoft Corporation
				'
				objExec.StdOut.SkipLine
				objExec.StdOut.SkipLine
				objExec.StdOut.SkipLine
			
				' catch the stdout of the relaunched script
				tLog.Log objExec.StdOut.ReadAll()
			    Do While objExec.Status = 0
					WScript.Sleep 100
				Loop
				intResult = objExec.ExitCode
				If intResult <> 0 Then
					tLog.Log "Non-Zero exit code for file " & objFile.Path & ", Quitting"
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
Sub BuildPatchGUIDToProductGUIDDict(strKBsCommaSeparated,intColIndex,ByRef dictScanResultsByColumnIndex, ByRef dictGUIDs)

	' takes the kb articles, comma separated, the dictionary object of scan results from all multi-cab scan result files
	' and builds the dictGUIDs object, which is a list of all patch guids and their associated product GUIDs
	' that have any relation to the KB article passed in
	Dim dictKBsPassedIn
	Set dictKBsPassedIn = CreateObject("Scripting.Dictionary")
	dictKBsPassedIn.CompareMode = vbTextCompare
	Dim dictKBTitlesToArticleID: Set dictKBTitlesToArticleID = CreateObject("Scripting.Dictionary")
	dictKBTitlesToArticleID.CompareMode = vbTextCompare
	' first, map KB article to display name in the patch results files(s)
	Dim strKBToUninstall
	For Each strKBToUninstall In Split(strKBsCommaSeparated,",")
		If Not InStr(LCase(strKBToUninstall),"kb") = 1 Then
			strKBToUninstall = "KB"&strKBToUninstall ' KB prefix is required
		End If
		If Not dictKBsPassedIn.Exists(strKBToUninstall) Then
			dictKBsPassedIn.Add strKBToUninstall,""
		End If ' will make matching simple later in this function		
		If dictScanResultsByColumnIndex.Exists(strKBToUninstall) Then
			strTitle = Split(dictScanResultsByColumnIndex.Item(strKBToUninstall),strSep)(0)
			If Not dictKBTitlesToArticleID.Exists(strKBToUninstall) Then
				dictKBTitlesToArticleID.Add strTitle,strKBToUninstall
			End If
		Else
			tLog.Log "Could not find any record of kb number " & strKBToUninstall & " in the results file(s), cannot uninstall"
		End If
	Next
	
	' Next, loop through installer data to find patch and product information
	' If the 
	
	Const MSIINSTALLCONTEXT_ALL = 7 

	Dim oMSI,iContext,objProducts,objProd
	Dim dictProducts,dictPatches
	Dim allProducts,product,strProductCode,objPatches
	Dim objPatch,strUninstallName
	
	Set dictProducts = CreateObject("Scripting.Dictionary")
	Set dictPatches = CreateObject("Scripting.Dictionary")
	
	Set oMsi = CreateObject("WindowsInstaller.Installer")
	iContext = MSIINSTALLCONTEXT_ALL
	
	Set allProducts = oMSI.ProductsEx("","",4)
	
	For Each product In allProducts
		Set objPatches = oMSI.PatchesEx(product.ProductCode,"",4,1)
		For Each objPatch In objPatches
			For Each strUninstallName In dictKBTitlesToArticleID.Keys
				' If the title of the patch, in the results file, is the beginning of the Display name of the patch
				' and it's at least 3 characters long (failsafe), - OR - if the 
				' the kb article numbers are the display name of the msi patch
				If InStr(LCase(objPatch.PatchProperty("DisplayName")),LCase(strUninstallName)) = 1 And Len(objPatch.PatchProperty("DisplayName")) > 2 Or _
					dictKBsPassedIn.Exists(objPatch.PatchProperty("DisplayName")) Then
					tLog.Log "Found patch item " & strUninstallName & ", kb number " _
						&dictKBTitlesToArticleID.Item(strUninstallName)&", - matched installer patch name: " & objPatch.PatchProperty("DisplayName")
					tLog.Log "Patch Code: " & objPatch.PatchCode &",Product Code: "& product.ProductCode
					If Not dictGUIDs.Exists(objPatch.PatchCode&strSep&product.ProductCode) Then
						dictGUIDs.Add objPatch.PatchCode&strSep&product.ProductCode,""
					End If
				End If
			Next
		Next
	Next

End Sub 'BuildPatchGUIDToProductGUIDDictByColumnIndex

Function UninstallMSIPatch(strProductGUID,strPatchGUID)
	
	Dim objShell,objEnv,strWinDir,strCmd,intReturnCode
	
	Set objShell = WScript.CreateObject("WScript.Shell")
	Set objEnv = objShell.Environment("Process")
	tLog.Log "Performing uninstall of MSI based patch"
	objEnv("SEE_MASK_NOZONECHECKS") = 1
	strWinDir = UCase(objShell.ExpandEnvironmentStrings("%WINDIR%"))
	strCmd = strWinDir & "\system32\MsiExec.exe /package "&strProductGUID&" /uninstall " & strPatchGUID &" /qn /norestart REBOOT=ReallySuppress"
	tLog.Log "With Command " & strCmd
	intReturnCode = objShell.Run(strCmd, 0, TRUE)
	tLog.Log intReturnCode & " was return code of Office / MSI based patch uninstall job"
	If intReturnCode = 0 Or intReturnCode = 3010 Then ' 3010 is reboot required
		UninstallMSIPatch = True
	Else
		UninstallMSIPatch = False
	End If

End Function 'UninstallOfficePath

Function GetOSNameByVersion
' Returns the OS String by Version
' This is important for matching since the WMI class OS string
' is subject to Localization

	Dim objWMIService,colItems,objItem
	Dim strOS,strVersion

	Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
	Set colItems = GetObject("WinMgmts:root/cimv2").ExecQuery("select Version from win32_operatingsystem")    
	For Each objItem In colItems
		strVersion = objItem.Version ' like 6.2.9200
		strVersion = Left(strVersion,Len(strVersion) - 5)
	Next
	
	Select Case strVersion
		Case "6.2"
			strOS = "Windows 8 or Windows Server 2012"
		Case "6.1"
			strOS = "Windows 7 or Windows Server 2008 R2"
		Case "6.0"
			strOS = "Windows Vista or Windows Server 2008"
		Case "5.2"
			strOS = "Windows Server 2003 or XP64"
		Case "5.1"
			strOS = "Windows XP"
		Case "5.0"
			strOS = "Windows 2000"
		Case Else
			strOS = "Unknown OS"
	End Select
	
	GetOSNameByVersion = strOS
	
End Function 'GetOSNameByVersion

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


Function UninstallKB(strKB,strKBNumbersOnly)
' Chooses correct uninstall command for OS and bitness, returns
' run result
	
	Const WINDOWSDIR = 0
	Const HKLM = &h80000002	
	
	Dim intResult,objShell,objFSO,objReg,strWinDir,strKeyPath,strUninstallStringFromReg
	Dim strUninstallKeyPath,strUninstallCommand,strOSMajorVersion,sngOSMajorVersion,strRegKeyPath
	Dim arrRegKeys,strRegKey, bUninstallWasOK
	
	strOSMajorVersion = GetOSMajorVersion

	strUninstallStringFromReg = ""
	strUninstallKeyPath = "Software\Microsoft\Windows\CurrentVersion\Uninstall\"
	strRegKeyPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\Packages"
	Set objReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strWinDir = objFSO.GetSpecialFolder(WINDOWSDIR)

	If Not IsNumeric(strOSMajorVersion) Then
		tLog.Log "Uninstall function cannot determine OS version"
		UninstallKB = False
		Exit Function
	Else
		sngOSMajorVersion = CSng(strOSMajorVersion)
	End If
	
	Set objShell = CreateObject("WScript.Shell")
	
	If sngOSMajorVersion < 5.0 And sngOSMajorVersion >= 4.0 Then
		tLog.Log "Windows NT 4 uninstall not supported" ' WSH won't execute this script!
		UninstallKB = False
		Exit Function
	End If
	
	bUninstallWasOK = False ' assume failure of uninstall
	If sngOSMajorVersion >= 5 And sngOSMajorVersion < 6.0 Then
		tLog.Log "Uninstalling for Pre-Vista OS"
		' Pull uninstall key for pre-vista OS's
		' We are either looking for the key exactly as ...Uninstall/KB12456, or we will have to search
		' for ...Uninstall/KB12345-???, where ??? could be -IE8, or -v4, etc.
		If RegKeyExists(objReg,HKLM,strUninstallKeyPath & strKB) Then
			On Error Resume Next
			objReg.GetStringValue HKLM,strUninstallKeyPath & strKB,"UninstallString",strUninstallStringFromReg
			If Err.Number <> 0 Or strUninstallStringFromReg = "" Or IsNull(strUninstallStringFromReg) Then
				tLog.Log "Unexpected Error retrieving uninstall registry key"
				On Error Goto 0
				UninstallKB = False
				Exit Function
			End If
			On Error Goto 0
		Else 
			' We didn't find the key directly.  We'll have to go searching for the kind with the "-???" appended
			' Note, we didn't just search directly, because startsWith() could still point to the wrong KB, 
			' as the number digits change.
			objReg.EnumKey HKLM,strUninstallKeyPath,arrRegKeys
			If IsArray(arrRegKeys) Then
				For Each strRegKey In arrRegKeys
					If InStr(strRegKey, strKB&"-") = 1 Then 'starts with
						tLog.Log "Found KB at " & strRegKey
						On Error Resume Next
						objReg.GetStringValue HKLM,strUninstallKeyPath&"\"&strRegKey,"UninstallString",strUninstallStringFromReg
						If Err.Number <> 0 Or strUninstallStringFromReg = "" Or IsNull(strUninstallStringFromReg) Then
							tLog.Log "Unexpected Error retrieving uninstall registry key"
							On Error Goto 0
							UninstallKB = False
							Exit Function
						End If
						On Error Goto 0
					End if
				Next
			End If 						
		End If
	
		If strUninstallStringFromReg = "" Or IsNull(strUninstallStringFromReg) Then
			tLog.Log "Cannot obtain uninstall string from registry"
			UninstallKB = False
			Exit Function
		End If
		strUninstallCommand = strUninstallStringFromReg&" -u -q -z"
		tLog.Log "Trying Uninstall Command: " & strUninstallCommand
		intResult = objShell.Run(strUninstallCommand,0,True)
		If intResult = 0 Or intResult = 3010 Then bUninstallWasOK = True
		If intResult = 3010 Then tLog.Log "A reboot is required"
	End If ' Pre-Vista
	
	If sngOSMajorVersion = 6.0 Then
		tLog.Log "Uninstalling for Windows Vista or Server 2008"
		If RegKeyExists(objReg,HKLM,strRegKeyPath) Then
		' Looking for a key that looks like
		' Package_for_KB960803~31bf3856ad364e35~amd64~~6.0.1.0
			objReg.EnumKey HKLM,strRegKeyPath,arrRegKeys
			If IsArray(arrRegKeys) Then
				For Each strRegKey In arrRegKeys
					If InStr(strRegKey,"Package_for_"&strKB) = 1 Then 'starts with
						' we may try multiple uninstall commands.  Track if any were successful.
						tLog.Log "Found packages key in registry: "&strRegKey
						strUninstallCommand = strWinDir&"\system32\pkgmgr.exe /quiet /norestart /up:"&strRegKey
						tLog.Log "Trying Uninstall Command: " & strUninstallCommand
						intResult = objShell.Run(strUninstallCommand,0,True)
						If intResult = 0 Or intResult = 3010 Then bUninstallWasOK = True
						If intResult = 3010 Then tLog.Log "A reboot is required"						
					End If
				Next
			Else
				tLog.Log "Unexpected Error enumerating Packages keys, cannot continue"
				UninstallKB = False
				Exit Function
			End If
		Else
			tLog.Log "Unexpected Error locating Packages key, cannot continue"
		End If
	End If ' Vista / 2008
				
	If sngOSMajorVersion > 6.0 Then
		tLog.Log "Uninstalling for Windows 7 and above"
		strUninstallCommand = strWinDir&"\system32\wusa.exe /uninstall /kb:"&strKBNumbersOnly&" /quiet /norestart"
		tLog.Log "Trying Uninstall Command: " & strUninstallCommand
		intResult = objShell.Run(strUninstallCommand,0,True)
		If intResult = 0 Or intResult = 3010 Then bUninstallWasOK = True
		If intResult = 3010 Then tLog.Log "A reboot is required"	
	End If

	UninstallKB = bUninstallWasOK
End Function 'UninstallKB

Function StopWindowsUpdate()
	Dim oShell,objWMIService,colServices,objService,WuaService
	Dim strService,strServiceStatus,strServiceMode
	
	strService = "wuauserv"
	If wuaNeedsStop Or wuaNeedsDisabled Then 
		tLog.Log "Stopping Windows Update service"

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
			tLog.Log "Scan Error: Cannot find Windows Update (wuauserv)"
			WScript.Quit
		End If
		
		tLog.Log "Return code: " & WuaService.ChangeStartMode("Disabled")
	End If
End Function


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


Sub ReadRegConfig(ByRef tContentReg, ByRef dictPatchManagementConfig)
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
End Sub 'ReadRegConfig


Sub EchoConfig(ByRef dictPConfig)
	Dim strKey
	tLog.Log "Patch Management Config (Registry and / or default values)"
	For Each strKey In dictPConfig
		tLog.Log strKey &" = "& dictPConfig.Item(strKey)
	Next
End Sub 'EchoConfig


Function KBArticleTranslator(strUninstallByArg)
	' Must strip KB prefix off and accept numbers only
	strUninstallByArg = UCase(strUninstallByArg)
	If InStr(strUninstallByArg,"KB") = 1 Then
		KBArticleTranslator = Right(strUninstallByArg,Len(strUninstallByArg) - 2)
	Else
		KBArticleTranslator = strUninstallByArg
	End If
End Function 'KBArticleTranslator


Sub ParseArgs(ByRef ArgsParser)

	' Pre- and Post- directory locations
	Dim objFSO
	Dim strPrePostPrefix,strFileDir,strExtension,strFolderName
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strFileDir = Replace(WScript.ScriptFullName,WScript.ScriptName,"")
	strExtension = objFSO.GetExtensionName(WScript.ScriptFullName)
	strFolderName = Replace(WScript.ScriptName,"."&strExtension,"")
	strPrePostPrefix = strFileDir&strFolderName
	
	ArgsParser.ProgramDescription = "Performs a Windows Patch uninstall operation. Typically triggered as a " _ 
		& "Tanium Action. Will Log output to the ContentLogs subfolder of the Tools folder. Note: NO " _
		& "command line arguments are 'sticky' - unlike the scan and install scripts, they are NOT stored in the registry " _
		& "and retrieved on subsequent calls to the script. Integrates with Tanium Maintenance Window content, " _
		& "and can run scripts in the "&strPrePostPrefix&"\Pre"&Chr(34)&" and "&Chr(34)&strPrePostPrefix&"\Post"&Chr(34) _ 
		&" directories before and after execution, respectively."
	
	Dim objUninstallByKBArg,UninstallByKBRef
	Set objUninstallByKBArg = New TaniumNamedArg
	Set UninstallByKBRef = GetRef("KBArticleTranslator")
	objUninstallByKBArg.ArgName = "UninstallByKB"
	objUninstallByKBArg.HelpText = "Uninstalls a single update by KB article number."
	objUninstallByKBArg.ExampleValue = "KB239084"
	objUninstallByKBArg.IsOptional = False
	objUninstallByKBArg.TranslationFunctionReference = UninstallByKBRef
	ArgsParser.AddArg objUninstallByKBArg

	Dim objClearInstallResultsFlagArg
	Set objClearInstallResultsFlagArg = New TaniumNamedArg
	objClearInstallResultsFlagArg.RequireYesNoTrueFalse = True
	objClearInstallResultsFlagArg.ArgName = "ClearInstallResultsOnBadLine"
	objClearInstallResultsFlagArg.HelpText = "If the UninstallResults file has a bad line, clear the entire file."
	objClearInstallResultsFlagArg.ExampleValue = "Yes,No"
	objClearInstallResultsFlagArg.DefaultValue = "No"
	objClearInstallResultsFlagArg.IsOptional = True
	ArgsParser.AddArg objClearInstallResultsFlagArg
	
	ArgsParser.Parse
	' The arguments should be successfully parsed, and handling of the arguments
	' is performed elsewhere in the script
	If ArgsParser.ErrorState Then
		ArgsParser.PrintUsageAndQuit ""
	End If
End Sub 'ParseArgs

''' ---- Fix Function definition ---- '''
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