'Tanium File Version:2.2.2.0011

' Deleted a previously queued patch from the /Patches directory

Option Explicit

Dim strPatchFiles, strPatchFile, strPatchDir, strFilePath, objFSO

If WScript.Arguments.Count <> 1 Then
	WScript.Echo "Usage:  delete-queued-patch.vbs {patch files to delete}"
	WScript.Quit
End if

strPatchFiles = Split(unescape(WScript.Arguments(0)), ",")
strPatchDir = GetTaniumDir("Tools\Patches")

Set objFSO = CreateObject("Scripting.FileSystemObject")

For Each strPatchFile In strPatchFiles
	strFilePath = strPatchDir & strPatchFile
	If objFSO.FileExists(strFilePath) Then
		objFSO.DeleteFile strFilePath, True
		WScript.Echo "Deleted: " & strFilePath
	Else 
		WScript.Echo "File not found: " & strFilePath
	End If
Next	


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
