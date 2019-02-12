'========================================
' Dump Memory
'========================================

'@INCLUDE=utils/os/GetOSMajorVersion.vbs
'@INCLUDE=utils/reg/Getx64RegistryProvider.vbs

Option Explicit

Dim colNamedArguments, strOutputDir, strCurrentDir, objFSO, strMemoryChoiceFdPro, strMemoryChoiceWinPMem

Set colNamedArguments = WScript.Arguments.Named

If Not colNamedArguments.Exists("Dir") Then
    WScript.Echo "Quiting, must have output dir location"
    WScript.Quit 1
Else
    strOutputDir = Trim(unescape(colNamedArguments.Item("Dir")))
End If

If Not IsWinpmemSafe() Then
	WScript.Echo "Cannot safely collect memory on this build of Windows 10 with winpmem."
	WScript.Quit 1
End If

' verify that there is enough room
If Not VerifyDiskSpace2xMemorySize() Then
	WScript.Echo "Not enough disk space:  2X Memory Size of Free Space Required for Memory Dump and Compression"
	WScript.Quit 1
End If

strCurrentDir = Replace(WScript.ScriptFullName, WScript.ScriptName, "")

Set objFSO = CreateObject("Scripting.FileSystemObject")


strMemoryChoiceFdPro = "fdpro.exe"
strMemoryChoiceWinPMem = "winpmem.gb414603.exe"

If objFSO.FileExists(strCurrentDir & strMemoryChoiceFdPro) Then
	WScript.Echo "Found FDPro, running FDPro memory dump"
	RunFdPro strOutputDir, strCurrentDir, strMemoryChoiceFdPro
ElseIf objFSO.FileExists(strCurrentDir & strMemoryChoiceWinPMem) Then
	WScript.Echo "Found winpmem, running winpmem memory dump"
	RunWinPMem strOutputDir, strCurrentDir, strMemoryChoiceWinPMem
Else
	WScript.Echo "Could not find memory dump tools, exiting"
End If

Function RunFdPro(strDir, strCurrentDir, strFdPro)
	Dim objShell, objFSO, strOutputFile, strOutputCmd, strCommand
	Dim objShellExec, intTotal
	Set objShell = CreateObject("WScript.Shell")
	Set objFSO = CreateObject("Scripting.FileSystemObject")

	GeneratePath strDir, objFSO

	strOutputFile = strDir

	strCommand = Chr(34) & strCurrentDir & strFdPro & Chr(34) & " " & Chr(34) & strOutputFile &"\memorydump.hpak" & Chr(34) & " -compress"

	WScript.Echo "fdpro memory dump:"
	WScript.Echo "   command: " & strCommand
	WScript.Echo "   outfile: " & strOutputFile & "\memorydump.hpak"

	objShell.CurrentDirectory = strOutputFile

	If objFSO.FileExists(strCurrentDir & strFdPro) Then
		objShell.Run strCommand, 0, True
	Else
		WScript.Echo "Memory Dump Tool " & strFdPro & " does not exist, can not run"
	End If

End Function ' RunFdPro

Function RunWinPMem(strDir, strCurrentDir, strWinPMem)
	Dim objShell, objFSO, strOutputFile, strOutputCmd, strCommand
	Dim objShellExec, strComputer, strComputerName, colPageFiles, objPageFile, PageFile, objWMIService, DateTime
	strComputer = "."

	Set objShell = CreateObject("WScript.Shell")
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colPageFiles = objWMIService.ExecQuery("Select * from Win32_PageFileUsage")
	strComputerName = objShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )

	For Each objPageFile in colPageFiles
	    PageFile = objPageFile.Name
	    Exit For
	Next

	GeneratePath strDir, objFSO

	strOutputFile = strDir
	DateTime = Year(Now) & Right(0 & (Month(Now)),2) & Right(0 & (Day(Now)),2) & Hour(Now) & Minute(Now)
	strCommand = Chr(34) & strCurrentDir & strWinPMem & Chr(34) & " --format raw --volume_format raw --output " & Chr(34) & strOutputFile & "\" & strComputerName & "_" & DateTime & ".raw" & Chr(34)

	WScript.Echo "winpmem memory dump:"
	WScript.Echo "   command: " & strCommand
	WScript.Echo "   PageFile: " & PageFile
	WScript.Echo "   outfile: " & strOutputFile

	Set objShellExec = objShell.Exec(strCommand)
	WScript.Echo "   command output: " & objShellExec.StdOut.ReadAll

End Function ' Function

Function VerifyDiskSpace2xMemorySize()
	Dim objWMIService, colComputer, objComputer, colDisks, objDisk, bResult
	Dim dblMemSize, dblFreeSpace
	bResult = False

	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Set colComputer = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
	Set colDisks = objWMIService.ExecQuery("Select * from Win32_LogicalDisk where Caption = 'c:'")

	' get the size of the computer memory
	For Each objComputer in colComputer
   		dblMemSize = objComputer.TotalPhysicalMemory
   	Next

	For Each objDisk in colDisks
	  	If Not IsNull(objDisk.FreeSpace) Then
	    	dblFreeSpace = objDisk.FreeSpace
	  	End If
	Next

	WScript.Echo "Verifying Required Space:  Memory Size=" & dblMemSize & "  Disk Free Space=" & dblFreeSpace

	' if free space is greater than twice the memory size, then its OK to run the memory dump
	If dblFreeSpace > (dblMemSize * 2) Then bResult = True

   	VerifyDiskSpace2xMemorySize = bResult
End Function ' VerifyDiskSpace2xMemorySize

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



Function IsAtLeastWin10()
    Dim arrMajorVersion, intVersion, bResult
    
    bResult = False
    arrMajorVersion = Split(GetOSMajorVersion(), ".")
    
    If IsNumeric(arrMajorVersion(0)) Then 
        intVersion = Int(arrMajorVersion(0))
        
        If intVersion >= 10 Then
            bResult = True
        End If
    Else
        WScript.Echo "Error: Can not determine OS Version, cannot continue."
        WScript.Quit 1
    End If

    IsAtLeastWin10 = bResult
End Function ' IsAtLeastWin10

Function IsWinpmemSafe()
	If Not IsAtLeastWin10() Then
		IsWinpmemSafe = True
		Exit Function
	End If

	If Not IsDeviceGuardHVCIRunning() Then
		IsWinpmemSafe = True
		WScript.Echo "DeviceGuard HVCI is not running."
		Exit Function
	End If
	WScript.Echo "DeviceGuard HVCI is running."

	Const HKEY_LOCAL_MACHINE = &H80000002
	Const keyCurrentVersion = "Software\Microsoft\Windows NT\CurrentVersion"

	Set objReg = Getx64RegistryProvider()

	Dim nReleaseId, nCurrentBuild, nUBR, objReg
	If RegKeyExists(objReg, HKEY_LOCAL_MACHINE, keyCurrentVersion) Then
		Dim strReleaseId, strCurrentBuild
		objReg.GetStringValue HKEY_LOCAL_MACHINE, keyCurrentVersion, "ReleaseId", strReleaseId
		objReg.GetStringValue HKEY_LOCAL_MACHINE, keyCurrentVersion, "CurrentBuild", strCurrentBuild
		objReg.GetDWORDValue HKEY_LOCAL_MACHINE, keyCurrentVersion, "UBR", nUBR

		If IsNull(strReleaseId) Then
			WScript.Echo "Error: Failed to retrieve 'ReleaseId' from registry. Assuming unsafe for Winpmem."
			IsWinpmemSafe = False
			Exit Function
		End If

		If IsNull(strCurrentBuild) Then
			WScript.Echo "Error: Failed to retrieve 'CurrentBuild' from registry. Assuming unsafe for Winpmem."
			IsWinpmemSafe = False
			Exit Function
		End If

		If IsNull(nUBR) Then
			WScript.Echo "Error: Failed to retrieve 'UBR' from registry. Assuming unsafe for Winpmem."
			IsWinpmemSafe = False
			Exit Function
		End If

		nReleaseId = CInt(strReleaseId)
		nCurrentBuild = CInt(strCurrentBuild)
	Else
		WScript.Echo "Error: registry key '" & keyCurrentVersion & "' does not exist. Assuming unsafe for Winpmem."
		IsWinpmemSafe = False
		Exit Function
	End If

	WScript.Echo "ReleaseId: " & CStr(nReleaseId)
	WScript.Echo "CurrentBuild: " & CStr(nCurrentBuild)
	WScript.Echo "UBR: " & CStr(nUBR)

	If nReleaseId >= 1809 Then
		' Assume newer versions are unaffected by BSOD fixed by Sep 2018 rollups.
		IsWinpmemSafe = True
		Exit Function
	End If

	Select Case nReleaseId
		Case 1803
			If nUBR >= 320 Then
				IsWinpmemSafe = True
				Exit Function
			End If
		Case 1709
			If (nCurrentBuild = 16299) And (nUBR >= 699) Then
				IsWinpmemSafe = True
				Exit Function
			End If
		Case 1703
			If nUBR >= 1356 Then
				IsWinpmemSafe = True
				Exit Function
			End If
		Case 1607
			If nUBR >= 2515 Then
				IsWinpmemSafe = True
				Exit Function
			End If
	End Select

	' Default to false
	IsWinpmemSafe = False
End Function ' IsWinpmemSafe

Const DG_SecurityServicesRunning_None = 0
Const DG_SecurityServicesRunning_CredentialGuard = 1
Const DG_SecurityServicesRunning_HVCI = 2

Function IsDeviceGuardHVCIRunning()
	Dim objWMIService, colDeviceGuard, objDeviceGuard

	If Not IsAtLeastWin10() Then
		' DeviceGuard + HVCI requires at least Windows 10
		IsDeviceGuardHVCIRunning = False
		Exit Function
	End If

	On Error Resume Next
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\Microsoft\Windows\DeviceGuard")
	If Not Err.Number = 0 Then
		' Error encountered, so default to HVCI being enabled.
		WScript.Echo "Error while trying to open root\Microsoft\Windows\DeviceGuard WMI. Assuming that HVCI is enabled. " & CStr(Err.Number) & " - " & Err.Description
		IsDeviceGuardHVCIRunning = True
		Exit Function
	End If

	Set colDeviceGuard = objWMIService.ExecQuery("Select * from Win32_DeviceGuard")
	If Not Err.Number = 0 Then
		' Error encountered, so default to HVCI being enabled.
		WScript.Echo "Error while trying to query Win32_DeviceGuard. Assuming that HVCI is enabled. " & CStr(Err.Number) & " - " & Err.Description
		IsDeviceGuardHVCIRunning = True
		Exit Function
	End If
	On Error Goto 0

	' Enumerate all running services.
	For Each objDeviceGuard in colDeviceGuard
		Dim RunningService
		For Each RunningService in objDeviceGuard.SecurityServicesRunning
			If RunningService = DG_SecurityServicesRunning_HVCI Then
				IsDeviceGuardHVCIRunning = True
				Exit Function
			End If
		Next
	Next

	IsDeviceGuardHVCIRunning = False
End Function ' IsDeviceGuardHVCIRunning

Function RegKeyExists(objRegistry, sHive, sRegKey)
	Dim aValueNames, aValueTypes
	If objRegistry.EnumValues(sHive, sRegKey, aValueNames, aValueTypes) = 0 Then
		RegKeyExists = True
	Else
		RegKeyExists = False
	End If
End Function
'------------ INCLUDES after this line. Do not edit past this point -----
'- Begin file: utils/os/GetOSMajorVersion.vbs
' Used to return just the first 2 digits of the Windows version
' be aware that it is returned as string with "X.Y", not a number

' To include this file, copy/paste: INCLUDE=utils/os/GetOSMajorVersion.vbs


Function GetOSMajorVersion
	Dim strVersion,arrVersion
	
	strVersion = GetOSVersion()

	arrVersion = Split(strVersion,".")
	If UBound(arrVersion) >= 1 Then
		strVersion = arrVersion(0)&"."&arrVersion(1)
	End If

	GetOSMajorVersion = strVersion
End Function 'GetOSMajorVersion
'- End file: utils/os/GetOSMajorVersion.vbs
'- Begin file: utils/os/GetOSVersion.vbs
' Used to return the full version number on Windows

' To include this file, copy/paste: INCLUDE=utils/os/GetOSVersion.vbs


Function GetOSVersion
' Returns the OS Version
	Dim objWMIService,colItems,objItem
	Dim strVersion
	
	strVersion = Null

	Set objWMIService = GetObject("winmgmts:\\.\root\CIMV2")
	Set colItems = GetObject("WinMgmts:root/cimv2").ExecQuery("select Version from win32_operatingsystem")
	For Each objItem In colItems
		strVersion = objItem.Version ' like 6.2.9200
	Next
	
	If IsNull(strVersion) Then
		fRaiseError 5, "GetOSVersion", "Error:  Can not determine OS Version", False
	End If
	
	GetOSVersion = strVersion
End Function ' GetOSMajor
'- End file: utils/os/GetOSVersion.vbs
'- Begin file: utils/RaiseError.vbs
' To include this file, copy/paste: INCLUDE=utils/RaiseError.vbs

Function fRaiseError(errCode, errSource, errorMsg, RaiseError)
    If RaiseError Then
      On Error Resume Next
      Call Err.Raise(errCode, errSource, errorMsg)
      Exit Function
    Else
      WScript.Echo errorMsg
      Wscript.Quit
    End If
End Function
'- End file: utils/RaiseError.vbs
'- Begin file: utils/reg/Getx64RegistryProvider.vbs
' To include this file, copy/paste: INCLUDE=utils/reg/Getx64RegistryProvider.vbs

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
'- End file: utils/reg/Getx64RegistryProvider.vbs
