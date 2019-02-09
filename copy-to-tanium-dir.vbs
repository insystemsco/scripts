''# copy-to-tanium-dir.vbs takes in an argument that represents a Tanium folder
''# (i.e., a folder in the Tanium Client directory) that all files in the 
''# Package definition are copied to. 
''# 
''# If no argument is passed in, it defaults to the Tanium Client root directory
''# 
''# Sample:  copy-to-tanium-dir.vbs "Tools/Scan" 
''# Outcome: Files in the same directory as copy-to-tanium-dir.vbs are copied
''#          to the <Tanium Client>/Tools/Scan directory

Dim strTaniumDir

If Wscript.Arguments.Count > 0 Then
	'Get first argument, ignore all others
	Dim strArg
	strArg = WScript.Arguments(0)
	WScript.Echo "Using first argument to get a Tanium folder: " & strArg
	strTaniumDir = GetTaniumDir(strArg)
Else
	WScript.Echo "No argument supplied, using root directory"
	strTaniumDir = GetTaniumDir("")
End If

Set objShell = CreateObject("WScript.shell")
strCurrentDir = objShell.CurrentDirectory

Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder(strCurrentDir)
Set files = folder.Files

For Each file In files
	If file.Name <> WScript.ScriptName Then
		fso.CopyFile file.Path, strTaniumDir, True
	End If
Next

Wscript.Sleep(3000)

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
			WScript.Echo "Error: " & strPath & " does not exist on the filesystem"
			GetTaniumDir = False
		End If
	Else
		WScript.Echo "Error: Cannot find Tanium Client path in Registry"
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