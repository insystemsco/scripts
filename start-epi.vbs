Option Explicit

'@INCLUDE=utils/settings/GetClientDir.vbs
'@INCLUDE=index/DeletePIDIfExists.vbs

Dim strTaniumEpiDir,fso, strDestPath,strExeName,strDestFilePath,objShell,execed

strTaniumEpiDir = GetClientDir() & "Tools\EPI\"
strExeName = "TaniumEndpointIndex.exe"

Set fso = WScript.CreateObject("Scripting.Filesystemobject")
Set objShell = CreateObject("WScript.Shell")
strDestFilePath = strTaniumEpiDir & strExeName
If (DeletePIDIfExists(fso,strTaniumEpiDir)) Then
	execed=ExeEPI(fso,objShell)
Else
	WScript.Echo "Tanium Endpoint Index already running"
End If

Function ExeEPI(ByRef fso, ByRef objShell)
	Dim intRes
	ExeEPI=False
	If (fso.FileExists(strDestFilePath)) Then
		Dim strCmd,configFilePath

		'set current directory to Index home directory
		objShell.CurrentDirectory = strTaniumEpiDir
		'now start it
		configFilePath=strTaniumEpiDir & "config.ini"
		strCmd = Chr(34)&strDestFilePath&Chr(34)&" -i -c " & Chr(34) & configFilePath & Chr(34)

		intRes = objShell.Run(strCmd,0,False)
		If intRes <> 0 Then
			WScript.Echo "Error: Could not start Tanium Endpoint Index"
		Else
			ExeEPI=True
			WScript.Echo "Tanium Endpoint Index Started"
		End If
	Else
		WScript.Echo "Error: Could not find Tanium Endpoint Index"
	End If
End Function
'------------ INCLUDES after this line. Do not edit past this point -----
'- Begin file: utils/settings/GetClientDir.vbs
' Returns the directory of the client
' Note:  GetClientDir always returns ending with a \
' To include this file, copy/paste: INCLUDE=utils/settings/GetClientDir.vbs


Function GetClientDir()
	Dim strResult
	
	strResult = GetEnvironmentValue("TANIUM_CLIENT_ROOT")
	
	If IsNull(strResult) Then 
		Dim objSh
		strResult = ""
		Set objSh = CreateObject("WScript.Shell")
	
		On Error Resume Next
		If strResult="" Then strResult=Eval("objSh.RegRead(""HKLM\Software\Tanium\Tanium Client\Path"")") : Err.Clear
		If strResult="" Then strResult=Eval("objSh.RegRead(""HKLM\Software\Wow6432Node\Tanium\Tanium Client\Path"")") : Err.Clear
		If strResult="" Then Err.Clear
		On Error Goto 0
	End If
	
	If strResult="" Then Call fRaiseError(5, "GetClientDir", _
		"TSE-Error:Can not locate client directory", False)
		
	If Right(strResult, 1) <> "\" Then strResult = strResult & "\"

	GetClientDir = strResult
End Function
'- End file: utils/settings/GetClientDir.vbs
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
'- Begin file: utils/os/GetEnvironmentValue.vbs
' Returns the passed in environment variable value.
' Returns either the value as set in the environment, or Null if not set

' To include this file, copy/paste: INCLUDE=utils/os/GetEnvironmentValue.vbs


Function GetEnvironmentValue(strName) 
	Dim objShell, strSub, strResult
	Set objShell = CreateObject("WScript.Shell")
	strSub = "%" & strName & "%"
	strResult = objShell.ExpandEnvironmentStrings(strSub)
	
	If strResult = strSub Then 
		GetEnvironmentValue = Null
	Else 
		GetEnvironmentValue = strResult
	End If
End Function ' GetEnvironmentValue
'- End file: utils/os/GetEnvironmentValue.vbs
'- Begin file: index/DeletePIDIfExists.vbs
Function DeletePIDIfExists(ByRef fso, ByVal strTaniumEpiDir)
	Dim strPidFile
	DeletePIDIfExists=False
	strPidFile = strTaniumEpiDir & ".pid"
	If (fso.FileExists(strPidFile)) Then
		Err.Clear
		On Error Resume Next
		fso.DeleteFile strPidFile
		On Error Goto 0
		If (fso.FileExists(strPidFile)) Then
			DeletePIDIfExists=False
		Else
			DeletePIDIfExists=True
		End If
	Else
		DeletePIDIfExists=True
	End If
End Function
'- End file: index/DeletePIDIfExists.vbs
