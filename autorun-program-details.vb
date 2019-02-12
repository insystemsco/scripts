Option Explicit
'@INCLUDE=LogUtil.vbs
'@INCLUDE=ManAppsScanner.vbs
'@INCLUDE=CompareVersions.vbs
'@INCLUDE=GetTaniumComputerId.vbs
'@INCLUDE=GetTaniumDir.vbs
'@INCLUDE=IsRegExMatch.vbs
'@INCLUDE=LoadXmlDoc.vbs
'@INCLUDE=AppendToFile.vbs
'@INCLUDE=WinRegConstants.vbs
'@INCLUDE=GeneratePath.vbs
'@INCLUDE=GetFileVersion.vbs
'@INCLUDE=GetFileArchitecture.vbs
'@INCLUDE=GetWinOsName.vbs
'@INCLUDE=GetWinOsBits.vbs
'@INCLUDE=GetWinOsVersion.vbs

Dim LogLevel : LogLevel = 4

Dim sh  : Set sh  = CreateObject("WScript.Shell")
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")

Dim ScannerOptions : Set ScannerOptions = CreateObject("Scripting.Dictionary")
With ScannerOptions
    .Add "sh",  sh
    .Add "fso", fso
End With

' Setup User Options
Dim UserOptionList : Set UserOptionList = WScript.Arguments.Named
Dim UserOption
For Each UserOption In UserOptionList
    Select Case LCase(UserOption)
    Case "randomwaittimeinseconds"
        Dim WaitTime : WaitTime = UserOptionList.Item(UserOption)
        If CStr(CLng(WaitTime)) = WaitTime Then
            ScannerOptions.Add "MaxDelayInSeconds", WaitTime
        End If
    Case "loglevel"
        Dim NewLogLevel : NewLogLevel = UserOptionList.Item(UserOption)
        If CStr(CLng(NewLogLevel)) = NewLogLevel Then
            LogLevel = NewLogLevel
        End If
    End Select
Next

' Logging Functions for backwards compatability with new logging utility
Dim LogOptions : Set LogOptions = CreateObject("Scripting.Dictionary")
With LogOptions
    .Add "LOG_LEVEL", LogLevel
    .Add "LOG_FILENAME", "ManAppsScan.log"
End With
Dim LOGGER : Set LOGGER = (New LogUtil)(LogOptions)
Sub ERROR( Msg ) LOGGER.ERROR Msg : End Sub
Sub WARN( Msg )  LOGGER.WARN Msg  : End Sub
Sub INFO( Msg )  LOGGER.INFO Msg  : End Sub
Sub DEBUG( Msg ) LOGGER.DEBUG Msg : End Sub
Sub TRACE( Msg ) LOGGER.TRACE Msg : End Sub


Dim Scanner : Set Scanner = (New ManAppsScanner)(ScannerOptions)

' Dim RunDelay : RunDelay = Scanner.GetScannerOption("RunDelay")
' INFO "Waiting " & RunDelay & "ms before initiating scan."
' WScript.Sleep RunDelay

' Load Dat file information
Scanner.LoadDatFileInfo
' Filter out excluded applications
Scanner.FilterOutExcludedApplications
' Get appllication version and architecture info from local system
Scanner.GetInstalledAppInfo
' Filter out applications based on dat file restrictions
Scanner.FilterOutRestrictedApplications
' Write scan results to file
Scanner.WriteScanResultsToFile
' Create scan results "readable" file
Scanner.CreateReadableFile
'------------ INCLUDES after this line. Do not edit past this point -----
'- Begin file: LogUtil.vbs
'@SAFELIBINCLUDE
Class LogUtil
' Summary:  Class for running managing log messages and files.
' Requires: GeneratePath.vbs
'---------'---------'---------'---------'---------'---------'---------'---------
    Private sh
    Private fso
    Private LogUtilOptions

    Public Default Function Init(InitArgs)
        Set sh  = CreateObject("WScript.Shell")
        Set fso = CreateObject("Scripting.FileSystemObject")

        Set LogUtilOptions = CreateObject("Scripting.Dictionary")
        LogUtilOptions.CompareMode = VBTextCompare

        Dim DefaultOptions : Set DefaultOptions = CreateObject("Scripting.Dictionary")
        With DefaultOptions
            .Add "LOG_LEVEL", 3
            .Add "LOG_DIR", sh.CurrentDirectory
            .Add "LOG_FILENAME", "log_" & Right(Year(Date), 2) & Right("0" & Month(Date),2) & Right("0" & Day(Date),2) & ".log"
            .Add "VERBOSE", True
            .Add "APPEND", True
            .Add "MAX_LOG_SIZE", 1000000
            .Add "MAX_LOG_RETENTION", 5
        End With

        Dim CurrOpt
        For Each CurrOpt In DefaultOptions
            If InitArgs.Exists(CurrOpt) Then
                LogUtilOptions.Add CurrOpt, InitArgs.Item(CurrOpt)
            Else
                LogUtilOptions.Add CurrOpt, DefaultOptions.Item(CurrOpt)
            End If
        Next

        LogUtilOptions.Add "LOG_SESSION_ID", GenerateLogSessionID()

        If Not GeneratePath( fso, fso.GetAbsolutePathName(LogUtilOptions.Item("LOG_DIR")) ) Then
            ERROR "Logging to file disabled. Could not create directory: " & fso.GetAbsolutePathName(LogUtilOptions.Item("LOG_DIR"))
        Else
            LogUtilOptions.Add "LOG_FILE_PATH", LogUtilOptions.Item("LOG_DIR") & "\" &  LogUtilOptions.Item("LOG_FILENAME")
        End If

        INFO "==>> New Log Session Initiated With ID " & LogUtilOptions.Item("LOG_SESSION_ID")
        TRACE "Loggin session initialized with following options:"
        Dim OptName
        For Each OptName In LogUtilOptions
            If VarType(LogUtilOptions.Item(OptName)) >= 8192 Then
                TRACE "  " & OptName & " (Array) => " & Join(LogUtilOptions.Item(OptName), ", ")
            Else
                TRACE "  " & OptName & " => " & LogUtilOptions.Item(OptName)
            End If
        Next

        Set Init = Me
    End Function ' Init

    Private function GenerateLogSessionID()
        Dim TypeLib : Set TypeLib = CreateObject("Scriptlet.TypeLib")
        GenerateLogSessionID = Right(Mid(TypeLib.Guid, 2, 36), 5)
        Set TypeLib = Nothing
    End Function ' GenerateLogSessionID

    Public Function GetOption(OptName)
        GetOption = "ERROR: Invalid Option"
        If (LogUtilOptions.Exists(OptName)) Then
            GetOption = LogUtilOptions.Item(OptName)
        End If
    End Function ' GetOption

    Public Function SetOption(OptName, OptValue)
        SetOption = "ERROR: Invalid Option"

        If (LogUtilOptions.Exists(OptName)) Then
            TRACE "Setting Logging Option: " & OptName & " = " & OptValue
            LogUtilOptions.Item(OptName) = OptValue
            SetOption = OptValue
        End If
    End Function ' SetOption

    Private Function GetLogFileSize(LogFilePath)
        GetLogFileSize = 0
        If fso.FileExists(LogFilePath) Then
            Dim LogFile : Set LogFile = fso.GetFile(LogFilePath)
            GetLogFileSize = LogFile.Size
        End If
    End Function ' GetLogFileSize

    Private Sub RotateLogFiles()
        Dim count : count = 1
        Dim NewFileName, OrigFileName
        Do While (fso.FileExists(LogUtilOptions.Item("LOG_FILE_PATH") & "." & count) )
            count = count + 1
        Loop
        Dim i
        For i = count To 0 Step -1
            NewFileName = LogUtilOptions.Item("LOG_FILE_PATH") & "." & count
            If ( count-1 = 0 ) Then
                OrigFileName = LogUtilOptions.Item("LOG_FILE_PATH")
            Else
                OrigFileName = LogUtilOptions.Item("LOG_FILE_PATH") & "." & count-1
            End If

            If fso.FileExists( OrigFileName ) Then
                If count > LogUtilOptions.Item("MAX_LOG_RETENTION") Then
                    fso.DeleteFile OrigFileName
                Else
                    fso.MoveFile OrigFileName, NewFileName
                End If
            End If
        Next
    End Sub 'Rotate Log Files

    Private Sub WriteMsgToLog(EventTimeStamp, EventType, EventMsg)
        If ( LogUtilOptions.Exists("LOG_FILE_PATH") ) Then
            Dim LogFileSize : LogFileSize = GetLogFileSize(LogUtilOptions.Item("LOG_FILE_PATH"))
            If LogFileSize > LogUtilOptions.Item("MAX_LOG_SIZE") Then
                RotateLogFiles
            End If

            ' Log file is Unicode16 encoded
            Dim LogFile : Set LogFile = fso.OpenTextFile(LogUtilOptions.Item("LOG_FILE_PATH"), 8, True, -2)
            LogFile.Write FormatMsg(EventTimeStamp, EventType, EventMsg) & vbCRLF
            LogFile.Close
        End IF
    End Sub ' WriteMsgToLog

    Private Sub EchoMsgToSTDOUT(EventTimeStamp, EventType, EventMsg)
        WScript.Echo FormatMsg(EventTimeStamp, EventType, EventMsg)
    End Sub ' EchoMsgToSTDOUT

    Private Function FormatMsg(EventTimeStamp, EventType, EventMsg)
        FormatMsg = "["&EventTimeStamp&"] " _
                  & "("&LogUtilOptions.Item("LOG_SESSION_ID")&") " _
                  & "["&LEFT(EventType&Space(5), 5)&"] " _
                  & EventMsg
    End Function ' FormatMsg

    Public Sub ERROR(Msg)
        If LogUtilOptions.Item("LOG_LEVEL") >= 1 Then
            DIM EventTimeStamp : EventTimeStamp = Now()
            If LogUtilOptions.Exists("LOG_FILE_PATH") Then
                WriteMsgToLog EventTimeStamp, "ERROR", Msg
            End If
            If LogUtilOptions.Item("VERBOSE") Then
                EchoMsgToSTDOUT EventTimeStamp, "ERROR", Msg
            End If
        End If
    End Sub 'ERROR

    Public Sub WARN(Msg)
        If LogUtilOptions.Item("LOG_LEVEL") >= 2 Then
            DIM EventTimeStamp : EventTimeStamp = Now()
            If LogUtilOptions.Exists("LOG_FILE_PATH") Then
                WriteMsgToLog EventTimeStamp, "WARN", Msg
            End If
            If LogUtilOptions.Item("VERBOSE") Then
                EchoMsgToSTDOUT EventTimeStamp, "WARN", Msg
            End If
        End If
    End Sub 'WARN

    Public Sub INFO(Msg)
        If LogUtilOptions.Item("LOG_LEVEL") >= 3 Then
            DIM EventTimeStamp : EventTimeStamp = Now()
            If LogUtilOptions.Exists("LOG_FILE_PATH") Then
                WriteMsgToLog EventTimeStamp, "INFO", Msg
            End If
            If LogUtilOptions.Item("VERBOSE") Then
                EchoMsgToSTDOUT EventTimeStamp, "INFO", Msg
            End If
        End If
    End Sub 'INFO

    Public Sub DEBUG(Msg)
        If LogUtilOptions.Item("LOG_LEVEL") >= 4 Then
            DIM EventTimeStamp : EventTimeStamp = Now()
            If LogUtilOptions.Exists("LOG_FILE_PATH") Then
                WriteMsgToLog EventTimeStamp, "DEBUG", Msg
            End If
            If LogUtilOptions.Item("VERBOSE") Then
                EchoMsgToSTDOUT EventTimeStamp, "DEBUG", Msg
            End If
        End If
    End Sub 'DEBUG

    Public Sub TRACE(Msg)
        If LogUtilOptions.Item("LOG_LEVEL") >= 5 Then
            DIM EventTimeStamp : EventTimeStamp = Now()
            If LogUtilOptions.Exists("LOG_FILE_PATH") Then
                WriteMsgToLog EventTimeStamp, "TRACE", Msg
            End If
            If LogUtilOptions.Item("VERBOSE") Then
                EchoMsgToSTDOUT EventTimeStamp, "TRACE", Msg
            End If
        End If
    End Sub 'TRACE
End Class ' LogUtil
'- End file: LogUtil.vbs
'- Begin file: ManAppsScanner.vbs
'@SAFELIBINCLUDE
Class ManAppsScanner
' Summary:  Class for running Managed Applications scan and writes results to
'           scan results file.
' Requires: LogUtil.vbs
'           GeneratePath.vbs
'           GetTaniumDir.vbs
'           LoadXmlDoc.vbs
'           WinRegConstants.vbs
'           IsRegExMatch.vbs
'           CompareVersions.vbs
'           GetFileVersion.vbs
'           GetFileArchitecture.vbs
'           GetWinOsBits.vbs
'           GetWinOsVersion.vbs
'           GetWinOsName.vbs
'---------'---------'---------'---------'---------'---------'---------'---------
    Private sh
    Private fso
    Private ScannerOptions
    Private RestrictPrefix

    Public  ScanResults

    Public Default Function Init(InitArgs)
        Set sh  = InitArgs.Item("sh")
        Set fso = InitArgs.Item("fso")

        RestrictPrefix = "restrict-by-"

        Dim TaniumClientToolsDir : TaniumClientToolsDir = GetTaniumDir("Tools", sh)

        Set ScannerOptions = CreateObject("Scripting.Dictionary")
        ScannerOptions.CompareMode = VBTextCompare

        Dim DefaultOptions : Set DefaultOptions = CreateObject("Scripting.Dictionary")
        With DefaultOptions
            .Add "MaxDelayInSeconds",   0
            .Add "RunDelay",            0
            .Add "InputFieldSeperator", "|"
            .Add "DatFolderList",       Array( TaniumClientToolsDir & "Managed Applications\TaniumDatFiles" _
                                             , TaniumClientToolsDir & "Managed Applications\CustomDatFiles" )
            .Add "ScanResultsDir",      TaniumClientToolsDir & "Scans"
            .Add "ScanResultsFileName",     "maresults.txt"
            .Add "ResultsReadableFileName", "maresultsreadable.txt"
        End With

        Dim CurrOpt
        For Each CurrOpt In DefaultOptions
            If InitArgs.Exists(CurrOpt) Then
                ScannerOptions.Add CurrOpt, InitArgs.Item(CurrOpt)
            Else
                ScannerOptions.Add CurrOpt, DefaultOptions.Item(CurrOpt)
            End If
        Next

        Set ScanResults = CreateObject("Scripting.Dictionary")
        ScanResults.CompareMode = VBTextCompare
        With ScanResults
            .Add "ExclusionList",   CreateObject("Scripting.Dictionary")
            .Add "DatFileInfo",     CreateObject("Scripting.Dictionary")
            .Add "RegistryIndex",   CreateObject("Scripting.Dictionary")
            .Add "InstalledInfo",   CreateObject("Scripting.Dictionary")

            .Item("ExclusionList").CompareMode  = VBTextCompare
            .Item("DatFileInfo").CompareMode    = VBTextCompare
            .Item("InstalledInfo").CompareMode  = VBTextCompare
        End With

        ' Disabling Random Wait functionality as it seems to have outlived its usefulness.
        ' SetRunDelay Int(ScannerOptions.Item("MaxDelayInSeconds"))
        ' WScript.Sleep ScannerOptions.Item("RunDelay")

        ScannerOptions.Add "ScanResultFullPath" _
                         , fso.GetAbsolutePathName( ScannerOptions.Item("ScanResultsDir") & "\" & ScannerOptions.Item("ScanResultsFileName") )

        ScannerOptions.Add "ScanResultReadableFullPath" _
                         , fso.GetAbsolutePathName( ScannerOptions.Item("ScanResultsDir") & "\" & ScannerOptions.Item("ResultsReadableFileName") )

        ScannerOptions.Add "AuditOnlyString" _
                         , "audit_only"

        If Not GeneratePath( fso, fso.GetAbsolutePathName(ScannerOptions.Item("ScanResultsDir")) ) Then
            ERROR "Could not create directory: " & fso.GetAbsolutePathName(ScannerOptions.Item("ScanResultsDir"))
        End If
        Dim DatFolder
        For Each DatFolder In ScannerOptions.Item("DatFolderList")
            If Not GeneratePath( fso, fso.GetAbsolutePathName(DatFolder) ) Then
                ERROR "Could not create directory: " & fso.GetAbsolutePathName(ScannerOptions.Item("ScanResultsDir"))
            End If
        Next

        DEBUG "Scanner initialized with following options: "
        Dim OptName
        For Each OptName In ScannerOptions
            If VarType(ScannerOptions.Item(OptName)) >= 8192 Then
                DEBUG "  " & OptName & " (Array) => " & Join(ScannerOptions.Item(OptName), ", ")
            Else
                DEBUG "  " & OptName & " => " & ScannerOptions.Item(OptName)
            End If
        Next

        Set Init = Me
    End Function

    Public Sub SetRunDelay( waitTime )
        If waitTime = null Then
            Exit Sub
        Else
            Randomize(GenRndSeed())
            waitTime = waitTime * 1000 ' convert to milliseconds
            ScannerOptions.Item("RunDelay") = Int( ( waitTime + 1 ) * Rnd )
        End If
    End Sub ' SetRunDelay

    Public Function GenRndSeed()
        Dim strComputer : strComputer = "."
        Dim objWMIService : Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
        Dim colItems : Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfFormattedData_PerfOS_Memory",,48)
        Dim objItem
        For Each objItem in colItems
            GenRndSeed = objItem.AvailableBytes
            Exit For
        Next
    End Function ' GenRndSeed

    Public Sub SetDatFieldSeperator( newDatFieldDeperator )
        ScannerOptions.Item("DatFieldSeperator") = newDatFieldDeperator
    End Sub  ' SetDatFieldSeperator

    Public Function GetScannerOption( OptionName )
        GetScannerOption = Null
        If ScannerOptions.Exists(OptionName) Then
            GetScannerOption = ScannerOptions.Item(OptionName)
        End If
    End Function ' GetScannerOption

    Public Sub LoadDatFileInfo()
        TRACE "Sub Called: LoadDatFileData"

        INFO "Parsing dat files"
        Dim DatFolder
        For Each DatFolder In ScannerOptions.Item("DatFolderList")
            INFO "Searching for .dat files in " & Chr(34)&DatFolder&Chr(34)
            ParseDatFilesInDir DatFolder
        Next
    End Sub ' LoadDatFileInfo

    Public Sub FilterOutExcludedApplications( )
        Dim ExclusionList   : Set ExclusionList = ScanResults.Item("ExclusionList")
        Dim DatFileInfo     : Set DatFileInfo   = ScanResults.Item("DatFileInfo")
        Dim EntryId, ManagedAppName

        For Each EntryId In DatFileInfo
            ManagedAppName = DatFileInfo.Item(EntryId).Item("name")
            If ExclusionList.Exists(ManagedAppName) Then
                DEBUG "Removing excluded app from managed applicaiton list: " & EntryId
                DatFileInfo.Remove(EntryId)
            End If
        Next
    End Sub ' FilterOutExcludedApplications

    Public Sub FilterOutRestrictedApplications( )
        Dim ListId
        For Each ListId In (ScanResults.Item("DatFileInfo"))
            Dim DatFileInfo : Set DatFileInfo = ScanResults.Item("DatFileInfo")

            Dim AppName, OsVer, OsName, OsArch, InstVer
            AppName = DatFileInfo.Item(ListId).Item("name")
            OsName  = GetWinOsName()
            OsVer   = GetWinOsVersion()
            OsArch  = GetWinOsBits()
            InstVer = ScanResults.Item("InstalledInfo").Item(ListId).Item("Installed-Version")

            Dim resOsName, resOsVer, resOsArch, rxOsArch, resMinVer, resMaxVer
            resOsName = DatFileInfo.Item(ListId).Item(RestrictPrefix & "os-name")
            resOsVer  = DatFileInfo.Item(ListId).Item(RestrictPrefix & "os-version")
            resOsArch = DatFileInfo.Item(ListId).Item(RestrictPrefix & "os-arch")
            rxOsArch  = "x?"&OsArch&"\-?(bit)?"
            resMinVer = DatFileInfo.Item(ListId).Item(RestrictPrefix & "min-version")
            resMaxVer = DatFileInfo.Item(ListId).Item(RestrictPrefix & "max-version")

            DEBUG "Checking for Restrictions: " & ListId

            Dim IsRestricted : IsRestricted = False
            If Len(resOsName) > 0 _
            AND Not IsRegExMatch( OsName, resOsName ) Then
                INFO ListId & " is restricted by OS Name"
                DEBUG "  " & OsName & " did not match " & resOsName
                IsRestricted = True
            ElseIf Len(resOsArch) > 0 _
            AND Not IsRegExMatch( resOsArch, rxOsArch ) Then
                INFO ListId & " is restricted by OS Architecture"
                DEBUG "  " & OsArch & " did not match " & resOsArch
                IsRestricted = True
            ElseIf Len(resOsVer) > 0 _
            AND Not IsRegExMatch( OsVer, "^"&resOsVer&"\.?" ) Then
                INFO ListId & " is restricted by OS Version"
                DEBUG "  " & OsVer & " did not match " & resOsVer
                IsRestricted = True
            ElseIf Len(resMinVer) > 0 _
            AND Not InstVer = "0.0.0.0" _
            AND Not CompareVersions(resMinVer, InstVer ) = 2 Then
                INFO ListId & " is restricted by Minimum Version"
                DEBUG "  " & InstVer & " is less than " & resMinVer
                IsRestricted = True
            ElseIf Len(resMaxVer) > 0 _
            AND Not InstVer = "0.0.0.0" _
            AND Not CompareVersions(resMaxVer, InstVer) = 1 Then
                INFO ListId & " is restricted by Maximum Version"
                DEBUG "  " & InstVer & " is greater than " & resMaxVer
                IsRestricted = True
            Else
                DEBUG "   No applicable restrictions present"
            End If

            If IsRestricted Then
                INFO "Removing Restricted Entry From Scan Results: " & AppName
                DatFileInfo.Remove(ListId)
            End If
        Next
    End Sub ' FilterOutRestrictedApplications

    Public Sub GetInstalledAppInfo( )
        Dim DatFileInfo : Set DatFileInfo   = ScanResults.Item("DatFileInfo")
        Dim InstalledInfo : Set InstalledInfo = ScanResults.Item("InstalledInfo")

        IndexAllUsedRegistryKeys

        Dim ListId
        For Each ListId In DatFileInfo
            Dim DatFileEntry : Set DatFileEntry = DatFileInfo.Item(ListId)

            If InstalledInfo.Exists(ListId) Then
                WARN "Duplicate entry found when retrieving installed application info: " & ListId
                Exit Sub
            Else
                InstalledInfo.Add ListId, CreateObject("Scripting.Dictionary")
                InstalledInfo.Item(ListId).CompareMode = VBTextCompare
            End If
            With InstalledInfo.Item(ListId)
                .Add("Installed-Arch"),     "n/a"
                .Add("Installed-Version"),  "0.0.0.0"
                .Add("Installed-Status"),   "Not Installed"
            End With

            Select Case DatFileEntry.Item("match-type")
            Case "0"
                SetInstalledVerAndArchFromRegistryIndex ListId

                If InstalledInfo.Item(ListId).Item("Installed-Version") <> "0.0.0.0" Then
                    SetInstalledAppStatus ListId
                End If
            Case "1"
                Dim ExpandedPath : ExpandedPath = ExpandPath(DatFileEntry.Item("file-path"))
                If fso.FileExists( ExpandedPath ) Then
                    DEBUG "File Found: " & ExpandedPath
                    SetInstalledVerAndArchFromFileHeaders ListId
                    SetInstalledAppStatus ListId
                Else
                    DEBUG "File Not Found: " & ExpandedPath
                End If
            Case Else
                WARN "Unable to search for installed application information. Unrecognized match type."
            End Select
        Next
    End Sub ' GetInstalledAppInfo

    Public Sub WriteScanResultsToFile( )
        TRACE "Sub Called: WriteScanResultsToFile"
        Dim IFS : IFS = ScannerOptions.Item("InputFieldSeperator")
        Dim FilePath : FilePath = ScannerOptions.Item("ScanResultFullPath")

        INFO "Writing Scan Results File: " & FilePath
        If fso.FileExists(FilePath) Then
            DEBUG "Deleting old scan result file: " & FilePath
            fso.DeleteFile FilePath, True
        End If

        Dim ResultFileEntry, InstalledInfoEntry, DatFileInfoEntry, ListId
        For Each ListId In (ScanResults.Item("DatFileInfo"))
            Set InstalledInfoEntry = ScanResults.Item("InstalledInfo").Item(ListId)
            Set DatFileInfoEntry   = ScanResults.Item("DatFileInfo").Item(ListId)

            If InstalledInfoEntry.Item("Installed-Status") = "Not Installed" Then
                ResultFileEntry = DatFileInfoEntry.Item("Name") & IFS _
                                & DatFileInfoEntry.Item("Publisher") & IFS _
                                & InstalledInfoEntry.Item("Installed-Arch") & IFS _
                                & InstalledInfoEntry.Item("Installed-Version") & IFS _
                                & DatFileInfoEntry.Item("Install-Latest-Version") & IFS _
                                & InstalledInfoEntry.Item("Installed-Status") & IFS _
                                & DatFileInfoEntry.Item("Install-File-Name") & IFS _
                                & DatFileInfoEntry.Item("Install-Download") & IFS _
                                & DatFileInfoEntry.Item("Install-Command") & IFS _
                                & DatFileInfoEntry.Item("Install-Hash")
            Else
                ResultFileEntry = DatFileInfoEntry.Item("Name") & IFS _
                                & DatFileInfoEntry.Item("Publisher") & IFS _
                                & InstalledInfoEntry.Item("Installed-Arch") & IFS _
                                & InstalledInfoEntry.Item("Installed-Version") & IFS _
                                & DatFileInfoEntry.Item("Update-Latest-Version") & IFS _
                                & InstalledInfoEntry.Item("Installed-Status") & IFS _
                                & DatFileInfoEntry.Item("Update-File-Name") & IFS _
                                & DatFileInfoEntry.Item("Update-Download") & IFS _
                                & DatFileInfoEntry.Item("Update-Command") & IFS _
                                & DatFileInfoEntry.Item("Update-Hash")
            End If

            Dim ResultType
            If InStr(1, ResultFileEntry, IFS&"audit_only"&IFS, 1) Then
                ResultType = "Package_Required"
            Else
                ResultType = "Actionable"
            End If
            ResultFileEntry = ResultFileEntry & IFS & ResultType

            DEBUG "  Scan Result Entry: " & ResultFileEntry
            AppendToFile fso _
                       , ResultFileEntry & vbCRLF _
                       , FilePath
        Next
    End Sub ' WriteScanResultsToFile

    Public Sub CreateReadableFile( )
        TRACE "Sub Called: CreateReadableFile"
        Dim ScanResultFilePath  : ScanResultFilePath = ScannerOptions.Item("ScanResultFullPath")
        Dim ReadableFilePath    : ReadableFilePath   = ScannerOptions.Item("ScanResultReadableFullPath")

        INFO "Copying file from " & Chr(34)&ScanResultFilePath&Chr(34) & " to " & Chr(34)&ReadableFilePath&Chr(34)
        If fso.CopyFile(ScanResultFilePath, ReadableFilePath, True) Then
            DEBUG "Copy Successful"
        End If
    End Sub ' CreateReadableFile


    ' == Private Subs and Functions
    '---------'---------'---------'---------'---------'---------'---------'-----
    Private Sub ParseDatFilesInDir( DirPath )
        TRACE "Sub Called: ParseDatFilesInDir"
        If Not fso.FolderExists(DirPath) Then
            WARN "Cannot find firectory: " & DirPath
            Exit Sub
        End If

        Dim DirObj  : Set DirObj  = fso.GetFolder(DirPath)

        Dim objCurrFile
        For Each objCurrFile In DirObj.Files
            If LCase(fso.GetExtensionName(objCurrFile.Name)) = "dat" Then
                ParseDatFile objCurrFile
            End If
        Next

        INFO "Exclusion Count: " & ScanResults.Item("ExclusionList").Count
        INFO "Managed App Count: " & ScanResults.Item("DatFileInfo").Count
    End Sub ' ParseDatFilesInDir

    Private Sub ParseDatFile( ByRef objDatFile )
        TRACE "Sub Called: ParseDatFile"
        INFO "Reading Dat File: " & objDatFile.Name
        Dim xmlDoc : Set xmlDoc = LoadXmlDoc( objDatFile.Path )
        If TypeName(xmlDoc) = "DOMDocument" Then
            Dim ExcludedAppNode
            For Each ExcludedAppNode In xmlDoc.SelectNodes("//excluded-application")
                AddAppToExclusionList ExcludedAppNode
            Next
            For Each ExcludedAppNode In xmlDoc.SelectNodes("//*/excluded-application")
                AddAppToExclusionList ExcludedAppNode
            Next
            Dim ManagedAppNode
            For Each ManagedAppNode In xmlDoc.SelectNodes("//managed-application")
                DEBUG "Managed Application: " & ManagedAppNode.SelectSingleNode("./name").Text
                AddAppToDatFileInfo ManagedAppNode
            Next
            For Each ManagedAppNode In xmlDoc.SelectNodes("//*/managed-application")
                DEBUG "Managed Application: " & ManagedAppNode.SelectSingleNode("./name").Text
                AddAppToDatFileInfo ManagedAppNode
            Next
        End If
    End Sub ' ParseDatFile

    Private Sub AddAppToExclusionList( ByRef objAppNode )
        TRACE "Sub Called: AddAppToExclusionList"
        If TypeName(objAppNode) <> "IXMLDOMElement" Then
            DEBUG "AddAppToExclusionList expected first parameter to be "&Chr(34)&"IXMLDOMElement"&Chr(34)&" type but was passed variable type " &Chr(34)& TypeName(objAppNode) &Chr(34)& " instead."
            Exit Sub
        End If

        Dim ExclusionList   : Set ExclusionList = ScanResults.Item("ExclusionList")

        Dim NameOfApp : NameOfApp = objAppNode.GetAttribute("name")
        If Not ExclusionList.Exists(NameOfApp) Then
            ExclusionList.Add NameOfApp, 0
        Else
            WARN NameOfApp & " previously added to exclusion list. Ignoring duplicate entry."
        End If
    End Sub ' AddAppToExclusionList

    Private Sub AddAppToDatFileInfo( objAppNode )
        TRACE "Sub Called: AddAppToDatFileInfo"
        If TypeName(objAppNode) <> "IXMLDOMElement" Then
            DEBUG "AddAppToDatFileInfo expected first parameter to be "&Chr(34)&"IXMLDOMElement"&Chr(34)&" type but was passed variable type " &Chr(34)& TypeName(objAppNode) &Chr(34)& " instead."
            Exit Sub
        End If

        Dim DatFileInfo    : Set DatFileInfo = ScanResults.Item("DatFileInfo")

        Dim ManagedAppInfo : Set ManagedAppInfo = ExtractManAppsDataFromXmlNode( objAppNode )
        If HasMinRequiredManAppFields( ManagedAppInfo ) Then
            WARN "Missing one or more required fields. Skipping application."
            Exit Sub
        End If

        Dim ListId : ListId = MakeDatFileInfoId( ManagedAppInfo )
        If DatFileInfo.Exists(ListId) Then
            WARN "Ignoring duplicate application: " & ListId
        Else
            DatFileInfo.Add ListId, ManagedAppInfo
        End If
    End Sub ' AddAppToDatFileInfo

    Private Function ExtractManAppsDataFromXmlNode( objXmlNode )
        TRACE "Function Called: ExtractManAppsDataFromXmlNode"
        Set ExtractManAppsDataFromXmlNode = CreateObject("Scripting.Dictionary")
        ExtractManAppsDataFromXmlNode.CompareMode = VBTextCompare

        If TypeName(objXmlNode) <> "IXMLDOMElement" Then
            DEBUG "ExtractManAppsDataFromXmlNode expected first parameter to be "&Chr(34)&"IXMLDOMElement"&Chr(34)&" type but was passed variable type " &Chr(34)& TypeName(XmlNode) &Chr(34)& " instead."
            Exit Function
        End If

        With ExtractManAppsDataFromXmlNode
            .Add "name",        GetTextFromXmlNode( objXmlNode.SelectSingleNode("./name") )
            .Add "publisher",   GetTextFromXmlNode( objXmlNode.SelectSingleNode("./publisher") )
            .Add "match-type",  GetTextFromXmlNode( objXmlNode.SelectSingleNode("./match/type") )
        End With

        Dim NodeTest, NodeRoot
        Set NodeTest = objXmlNode.SelectSingleNode("./update")
        If NodeTest Is Nothing Then
            NodeRoot = "./latest-version/"
        Else
            NodeRoot = "./update/latest-version/"
        End If
        With ExtractManAppsDataFromXmlNode
            .Add "update-latest-version",   GetTextFromXmlNode( objXmlNode.SelectSingleNode(NodeRoot & "version") )
            .Add "update-file-name",        GetTextFromXmlNode( objXmlNode.SelectSingleNode(NodeRoot & "file-name") )
            .Add "update-hash",             GetTextFromXmlNode( objXmlNode.SelectSingleNode(NodeRoot & "hash") )
            .Add "update-download",         GetTextFromXmlNode( objXmlNode.SelectSingleNode(NodeRoot & "download") )
            .Add "update-command",          GetTextFromXmlNode( objXmlNode.SelectSingleNode(NodeRoot & "command") )
        End With

        Set NodeTest = objXmlNode.SelectSingleNode("./install")
        If NodeTest Is Nothing Then
            NodeRoot = "./latest-version/"
        Else
            NodeRoot = "./install/latest-version/"
        End If
        With ExtractManAppsDataFromXmlNode
            .Add "install-latest-version",   GetTextFromXmlNode( objXmlNode.SelectSingleNode(NodeRoot & "version") )
            .Add "install-file-name",        GetTextFromXmlNode( objXmlNode.SelectSingleNode(NodeRoot & "file-name") )
            .Add "install-hash",             GetTextFromXmlNode( objXmlNode.SelectSingleNode(NodeRoot & "hash") )
            .Add "install-download",         GetTextFromXmlNode( objXmlNode.SelectSingleNode(NodeRoot & "download") )
            .Add "install-command",          GetTextFromXmlNode( objXmlNode.SelectSingleNode(NodeRoot & "command") )
        End With

        Select Case ExtractManAppsDataFromXmlNode.Item("match-type")
        Case "0"
            With ExtractManAppsDataFromXmlNode
                .Add "reg-key",               GetTextFromXmlNode( objXmlNode.SelectSingleNode("./match/reg-key") )
                .Add "enum-keys",             GetTextFromXmlNode( objXmlNode.SelectSingleNode("./match/enum-keys") )
                .Add "value",                 GetTextFromXmlNode( objXmlNode.SelectSingleNode("./match/value") )
                .Add "data-match-regex",      GetTextFromXmlNode( objXmlNode.SelectSingleNode("./match/data-match-regex") )
                .Add "current-version-value", GetTextFromXmlNode( objXmlNode.SelectSingleNode("./match/current-version-value") )
            End With
        Case "1"
            With ExtractManAppsDataFromXmlNode
                .Add "version-type", GetTextFromXmlNode( objXmlNode.SelectSingleNode("./match/version-type") )
                .Add "file-path",    GetTextFromXmlNode( objXmlNode.SelectSingleNode("./match/file") )
            End With
        End Select

        ' Restrictions
        Dim RestrictionNode
        For Each RestrictionNode In objXmlNode.SelectNodes("./restriction")
            Dim Attribute, AttribName, AttribValue
            For Each Attribute In RestrictionNode.Attributes
                AttribName = Attribute.Name
                AttribValue = Attribute.Value
                Exit For
            Next
            If ExtractManAppsDataFromXmlNode.Exists(RestrictPrefix&AttribName) Then
                ExtractManAppsDataFromXmlNode.Item(RestrictPrefix&AttribName) = AttribValue
            Else
                ExtractManAppsDataFromXmlNode.Add(RestrictPrefix&AttribName), AttribValue
            End If
        Next

        ' Make sure that audit only commands match the expected case
        If UCase( ExtractManAppsDataFromXmlNode.Item("install-command") ) _
         = UCase( ScannerOptions.Item("AuditOnlyString") ) _
        Then
            ExtractManAppsDataFromXmlNode.Item("install-command") = ScannerOptions.Item("AuditOnlyString")
        End If
        If UCase( ExtractManAppsDataFromXmlNode.Item("update-command") ) _
         = UCase( ScannerOptions.Item("AuditOnlyString") ) _
        Then
            ExtractManAppsDataFromXmlNode.Item("update-command") = ScannerOptions.Item("AuditOnlyString")
        End If

        Dim Key
        TRACE "Extracted the following Data: "
        For Each Key in ExtractManAppsDataFromXmlNode
            TRACE "  " & Key & " => " & ExtractManAppsDataFromXmlNode.Item(Key)
        Next
    End Function ' ExtractManAppsDataFromXmlNode

    Private Function GetTextFromXmlNode( XmlNode )
        TRACE "Function Called: GetTextFromXmlNode"
        GetTextFromXmlNode = ""

        If TypeName(XmlNode) <> "IXMLDOMElement" Then
            DEBUG "GetTextFromXmlNode expected first parameter to be "&Chr(34)&"IXMLDOMElement"&Chr(34)&" type but was passed variable type " &Chr(34)& TypeName(XmlNode) &Chr(34)& " instead."
            Exit Function
        End If

        If Not XmlNode Is Nothing Then
            GetTextFromXmlNode = XmlNode.Text
        End If
    End Function ' GetTextFromXmlNode

    Private Function MakeDatFileInfoId( ManAppInfo )
        TRACE "Function Called: MakeDatFileInfoId"

        If ManAppInfo.Item("install-latest-version") = ManAppInfo.Item("update-latest-version") Then
            MakeDatFileInfoId = ManAppInfo.Item("name") & " version " & ManAppInfo.Item("install-latest-version")
        Else
            MakeDatFileInfoId = ManAppInfo.Item("name") & " install version " & ManAppInfo.Item("install-latest-version") & ", update version " & ManAppInfo.Item("update-latest-version")
        End If
    End Function ' MakeDatFileInfoId

    Private Function HasMinRequiredManAppFields( ManAppInfoDict )
        TRACE "Function Called: HasMinRequiredManAppFields"
        HasMinRequiredManAppFields = False
        If TypeName(ManAppInfoDict) <> "Dictionary" Then
            DEBUG "HasMinRequiredManAppFields expected first parameter to be "&Chr(34)&"Dictionary"&Chr(34)&" type but was passed variable type " &Chr(34)& TypeName(ManAppInfoDict) &Chr(34)& " instead."
            Exit Function
        End If

        Dim RequiredFields : Set RequiredFields = CreateObject("Scripting.Dictionary")
        With RequiredFields
            .Add "name", 0
            .Add "publisher", 0
            .Add "update-latest-version", 0
            .Add "update-file-name", 0
            .Add "update-download", 0
            .Add "update-command", 0
            .Add "install-latest-version", 0
            .Add "install-file-name", 0
            .Add "install-download", 0
            .Add "install-command", 0
        End With

        Dim MissingCount : missingCount = 0
        Dim Field
        For Each Field In RequiredFields
            If ManAppInfoDict.Item(Field) = "" Then
                WARN "Missing required field: " & Field
                MissingCount = MissingCount + 1
            End IF
        Next
        If MissingCount > 0 Then
            HasMinRequiredManAppFields = True
        End If
    End Function ' HasMinRequiredManAppFields

    Private Sub IndexAllUsedRegistryKeys( )
        TRACE "Sub Called: IndexAllUsedRegistryKeys"

        Dim UniqueRegKeyList : Set UniqueRegKeyList = CreateObject("Scripting.Dictionary")

        Dim ListId, RegHive, RegKeyPath, Arch
        For Each ListId In ScanResults.Item("DatFileInfo")
            Dim CurrDatInfo : Set CurrDatInfo = ScanResults.Item("DatFileInfo").Item(ListId)
            If CurrDatInfo.Item("match-type") = 0 Then
                    Dim FullValuePath
                If CurrDatInfo.Item("enum-keys") = 1 Then
                    FullValuePath = CurrDatInfo.Item("reg-key") & "\**\" & CurrDatInfo.Item("value")

                    If Not UniqueRegKeyList.Exists(FullValuePath)  Then
                        DEBUG "Adding key " & Chr(34)&FullValuePath&Chr(34) & " to recursive registry index schedule."
                        UniqueRegKeyList.Add FullValuePath, CreateObject("Scripting.Dictionary")
                        With UniqueRegKeyList.Item(FullValuePath)
                            .Add "FullRegKey",  CurrDatInfo.Item("reg-key")
                            .Add "ValueName",   CurrDatInfo.Item("value")
                        End With
                    Else
                        DEBUG "Registry key for " & ListId & " already scheduled to be indexed."
                    End If
                ElseIf CurrDatInfo.Item("enum-keys") = 0 Then
                    DEBUG "Indexing key " & Chr(34)&CurrDatInfo.Item("reg-key")&Chr(34)
                    ParseRegistryPath CurrDatInfo.Item("reg-key")  _
                                    , RegHive _
                                    , RegKeyPath

                    For Each Arch in Array(64, 32)
                        IndexRegKey RegHive _
                                  , RegKeyPath _
                                  , CurrDatInfo.Item("value") _
                                  , Arch
                    Next
                End If
            Else
                DEBUG "No registry search to be performed on " & ListId & " with match type " & CurrDatInfo.Item("match-type")
            End If
        Next

        For Each FullValuePath In UniqueRegKeyList
            ParseRegistryPath UniqueRegKeyList.Item(FullValuePath).Item("FullRegKey") _
                            , RegHive _
                            , RegKeyPath

            For Each Arch in Array(64, 32)
                IndexRegKeyRecursive RegHive _
                                   , RegKeyPath _
                                   , UniqueRegKeyList.Item(FullValuePath).Item("ValueName") _
                                   , Arch
            Next
        Next
    End Sub ' IndexAllUsedRegistryKeys

    Private Sub IndexRegKey( RegHive, RegKeyPath, ValueName, RegArch )
        TRACE "Sub Called: IndexRegKey"

        Dim objCtx, objReg, SubKey, SubKeyList
        Set objCtx = CreateContextObject( RegArch )
        Set objReg = CreateRegistryObject( objCtx )

        Dim ValueData : ValueData = GetRegistryStringValue( objCtx _
                                                          , objReg _
                                                          , RegHive _
                                                          , RegKeyPath _
                                                          , ValueName )

        If Not IsNull(ValueData) Then
            Dim IndexId : IndexId = "["&RegArch&"-bit] " & RegKeyPath&"\"&SubKey&"\"&ValueName
            If Not ScanResults.Item("RegistryIndex").Exists(IndexId) Then
                ScanResults.Item("RegistryIndex").Add IndexId, CreateObject("Scripting.Dictionary")
                With ScanResults.Item("RegistryIndex").Item(IndexId)
                    .Add "Value-Data", ValueData
                    .Add "Hive",       RegHive
                    .Add "Key-Path",   RegKeyPath&"\"&SubKey
                    .Add "Arch",       RegArch
                End With
            ElseIf ScanResults.Item("RegistryIndex").Item(IndexId).Item("Value-Data") = ValueData Then
                WARN "Duplicate Index Detected with Identical Value"
            Else
                WARN "Duplicate Index Detected: " & IndexId
                WARN "  Existing Value:    " & ScanResults.Item("RegistryIndex").Item(IndexId).Item("Value-Data")
                WARN "  Conflicting Value: " & ValueData
            End If
        End If
    End Sub ' IndexRegKey

    Private Sub IndexRegKeyRecursive( RegHive, RegKeyPath, ValueName, RegArch )
        TRACE "Sub Called: IndexRegKeyRecursive"

        Dim objCtx, objReg, SubKey, SubKeyList
        Set objCtx = CreateContextObject( RegArch )
        Set objReg = CreateRegistryObject( objCtx )
        SubKeyList = GetWinRegSubKeys( objCtx _
                                     , objReg _
                                     , RegHive _
                                     , RegKeyPath )
        If Not IsNull(SubKeyList) Then
            For Each SubKey In SubKeyList
                IndexRegKey RegHive _
                          , RegKeyPath&"\"&SubKey _
                          , ValueName _
                          , RegArch
            Next
        End If
    End Sub ' IndexRegKeyRecursive

    Private Sub SetInstalledVerAndArchFromFileHeaders( EntryId )
        TRACE "Sub Called: SetInstalledVerAndArchFromFileHeaders"
        If Not ScanResults.Item("InstalledInfo").Exists(EntryId) Then
            WARN "Invalid Entry Id: " & EntryId
            Exit Sub
        End If
        Dim InstalledInfoEntry : Set InstalledInfoEntry = ScanResults.Item("InstalledInfo").Item(EntryId)
        Dim DatFileInfoEntry   : Set DatFileInfoEntry   = ScanResults.Item("DatFileInfo").Item(EntryId)

        Dim ExpandedPath, FileArch, FileVer
        ExpandedPath = ExpandPath( DatFileInfoEntry.Item("file-path") )
        Select Case GetFileArchitecture(ExpandedPath)
            Case "x86"
                FileArch = "32-bit"
            Case "x64"
                FileArch = "64-bit"
            Case Else
                FileArch = "n/a"
        End Select

        FileVer = GetFileVersion(fso, ExpandedPath, DatFileInfoEntry.Item("version-type"))
        If Len(FileVer) < 1 Then
            FileVer = "0.0.0.0"
        End If

        DEBUG "Version and Architecture gathered from file headers: " & ExpandedPath
        DEBUG "  Version: " & FileVer
        DEBUG "  Architecture: " & FileArch

        InstalledInfoEntry.Item("Installed-Arch")    = FileArch
        InstalledInfoEntry.Item("Installed-Version") = FileVer
    End Sub ' SetInstalledVerAndArchFromFileHeaders

    Private Sub SetInstalledVerAndArchFromRegistryIndex( EntryId )
        TRACE "Sub Called: SetInstalledVerAndArchFromRegistryIndex"
        If Not ScanResults.Item("InstalledInfo").Exists(EntryId) Then
            WARN "Invalid Entry Id: " & EntryId
            Exit Sub
        End If
        Dim InstalledInfoEntry : Set InstalledInfoEntry = ScanResults.Item("InstalledInfo").Item(EntryId)
        Dim DatFileInfoEntry   : Set DatFileInfoEntry   = ScanResults.Item("DatFileInfo").Item(EntryId)

        Dim RegIndexId
        For Each RegIndexId In ScanResults.Item("RegistryIndex")
            Dim RegIndexEntry : Set RegIndexEntry = ScanResults.Item("RegistryIndex").Item(RegIndexId)
            Dim IsValueMatch : IsValueMatch = IsRegExMatch( RegIndexEntry.Item("Value-Data") _
                                                          , DatFileInfoEntry.Item("data-match-regex") )

            Dim MatchHive, MatchKeyPath
            ParseRegistryPath DatFileInfoEntry.Item("reg-key") _
                            , MatchHive _
                            , MatchKeyPath

            Dim IsKeyMatch
            If DatFileInfoEntry.Item("enum-keys") = 1 Then
                IsKeyMatch = StrComp( fso.GetParentFolderName(RegIndexEntry.Item("Key-Path")) _
                                    , MatchKeyPath _
                                    , vbTextCompare)
            Else
                If Not Right(RegIndexEntry.Item("Key-Path"), 1) = "\" Then
                    RegIndexEntry.Item("Key-Path") = RegIndexEntry.Item("Key-Path")&"\"
                End If
                If Not Right(MatchKeyPath, 1) = "\" Then
                    MatchKeyPath = MatchKeyPath&"\"
                End If
                IsKeyMatch = StrComp( RegIndexEntry.Item("Key-Path") _
                                    , MatchKeyPath _
                                    , vbTextCompare)
            End If

            If IsValueMatch And IsKeyMatch = 0 Then
                DEBUG "Found Registry Match for " & EntryId
                DEBUG "  Registry Entry: " & RegIndexId
                Dim RegArch, RegHive, RegKeyPath, ValueName, objCtx, objReg
                RegArch    = RegIndexEntry.Item("Arch")
                RegHive    = RegIndexEntry.Item("Hive")
                RegKeyPath = RegIndexEntry.Item("Key-Path")
                ValueName  = DatFileInfoEntry.Item("current-version-value")
                Set objCtx = CreateContextObject( RegArch )
                Set objReg = CreateRegistryObject( objCtx )
                Dim thisVer : thisVer = GetRegistryStringValue( objCtx _
                                                              , objReg _
                                                              , RegHive _
                                                              , RegKeyPath _
                                                              , ValueName )

                If CompareVersions(InstalledInfoEntry.Item("Installed-Version"), thisVer) = 2 Then
                    InstalledInfoEntry.Item("Installed-Arch")    = RegArch&"-bit"
                    InstalledInfoEntry.Item("Installed-Version") = thisVer
                End If
            End If
        Next
    End Sub ' SetInstalledVerAndArchFromRegistryIndex

    Private Sub SetInstalledAppStatus( EntryId )
        TRACE "Sub Called: SetInstalledAppStatus"
        If Not ScanResults.Item("InstalledInfo").Exists(EntryId) Then
            WARN "Invalid Entry Id: " & EntryId
            Exit Sub
        End If
        If Not ScanResults.Item("DatFileInfo").Exists(EntryId) Then
            WARN "Invalid Entry Id: " & EntryId
            Exit Sub
        End If

        Dim InstalledInfoEntry : Set InstalledInfoEntry = ScanResults.Item("InstalledInfo").Item(EntryId)
        Dim DatFileInfoEntry   : Set DatFileInfoEntry = ScanResults.Item("DatFileInfo").Item(EntryId)

        Dim InstalledVersion, LatestVersion
        InstalledVersion    = InstalledInfoEntry.Item("Installed-Version")
        LatestVersion       = DatFileInfoEntry.Item("update-latest-version")

        Dim CompResult : CompResult = CompareVersions( InstalledVersion, LatestVersion )
        Select Case CompResult
        Case 0  ' Versions are equal
            InstalledInfoEntry.Item("Installed-Status") = "Up to Date"
        Case 1  ' Installed version is greater
            InstalledInfoEntry.Item("Installed-Status") = "Up to Date"
        Case 2  ' Latest version is greater
            InstalledInfoEntry.Item("Installed-Status") = "Out of Date"
        End Select
    End Sub ' SetInstalledAppStatus

    Private Function ExpandPath( FilePath )
        ' When cmd is launched from a 32-bit application (like the Tanium Client)
        ' some enviroment variables, such as %PROGRAMFILES% do not expand to the
        ' values expected.
        Dim NewPath :
        NewPath = sh.ExpandEnvironmentStrings(FilePath)
        If  IsRegExMatch(FilePath, "%PROGRAMFILES%") _
        And IsRegExMatch(NewPath, "\\Program Files \(x86\)\\") _
        Then
            NewPath = Replace( NewPath, "Program Files (x86)", "Program Files" )
        End If
        ExpandPath = NewPath
    End Function ' ExpandPath

    ' == Registry Subs and Functions
    '---------'---------'---------'---------'---------'---------'---------'-----
    Private Function FindRegKeyWithValue( RegHive, RegKeyPath, MatchValueName, MatchRegExQuery, RegArch )
        TRACE "Function Called: FindRegKeyWithValue"
        FindRegKeyWithValue = Null
        Dim ConextObject, RegistryObject, SubKey, SubKeyList

        Set ConextObject    = CreateContextObject( RegArch )
        Set RegistryObject  = CreateRegistryObject( ConextObject )
        SubKeyList = GetWinRegSubKeys( ConextObject _
                                     , RegistryObject _
                                     , RegHive _
                                     , RegKeyPath )
        If IsNull(SubKeyList) Then
            Exit Function
        End If
        Dim ValueData
        For Each SubKey In SubKeyList
             ValueData = GetRegistryStringValue( ConextObject _
                                               , RegistryObject _
                                               , RegHive _
                                               , RegKeyPath&"\"&SubKey _
                                               , MatchValueName )
            If Not IsNull(ValueData) Then
                If IsRegExMatch( ValueData, MatchRegExQuery ) Then
                    FindRegKeyWithValue = SubKey
                End If
            End If
        Next
    End Function ' FindRegKeyWithValue

    Private Sub ParseRegistryPath( FullRegPath, ByRef RegHive, ByRef RegKeyPath )
        TRACE "Sub Called: ParseRegistryPath"
        Dim RegKeyParts, PathLen
        RegKeyParts = Split(FullRegPath, "\")
        PathLen     = (Len(FullRegPath) - Len(RegKeyParts(0)) -1)

        RegHive     = GetHiveConst( RegKeyParts(0) )
        RegKeyPath  = Right( FullRegPath, PathLen )
    End Sub ' ParseRegistryPath

    Private Function CreateContextObject( Arch )
        TRACE "Function Called: CreateContextObject"
        Dim objCtx : Set objCtx = CreateObject("WbemScripting.SWbemNamedValueSet")
        With objCtx
            .Add "__ProviderArchitecture", CInt(Arch)
            .Add "__RequiredArchitecture", TRUE
        End With
        Set CreateContextObject = objCtx
    End Function ' CreateContextObject

    Private Function CreateRegistryObject(objCtx)
        TRACE "Function Called: CreateRegistryObject"
        Dim objLocator  : Set objLocator    = CreateObject("Wbemscripting.SWbemLocator")
        Dim objServices : Set objServices   = objLocator.ConnectServer("","root\default","","",,,,objCtx)

        Set CreateRegistryObject = objServices.Get("StdRegProv")
    End Function ' CreateRegistryObject

    Private Function GetRegistryStringValue( objCtx, objReg, RegHive, RegKeyPath, RegValueName )
        TRACE "Function Called: GetRegistryStringValue"
        Dim InParams, OutParams

        Set InParams = objReg.Methods_("GetStringValue").InParameters
        With InParams
            .hDefKey        = RegHive
            .sSubKeyName    = RegKeyPath
            .sValueName     = RegValueName
        End With

        Set OutParams = objReg.ExecMethod_( "GetStringValue" _
                                          , InParams _
                                          , _
                                          , objCtx )

        GetRegistryStringValue = OutParams.sValue
    End Function ' GetRegistryStringValue

    Private Function GetWinRegSubKeys(objCtx, objReg, RegHive, RegKeyPath)
        TRACE "Function Called: GetWinRegSubKeys"
        Dim InParams, OutParams

        Set InParams  = objReg.Methods_("EnumKey").Inparameters
        With InParams
            .Hdefkey = RegHive
            .sSubkeyname = RegKeyPath
        End With

        Set OutParams = objReg.ExecMethod_( "EnumKey" _
                                           , InParams _
                                           , _
                                           , objCtx )
        If Err.Number <> 0 Then
            ERROR "Failed to enumerate registry keys of key " & Chr(34)&RegKeyPath&Chr(34)
            ERROR Err.Description & "(" & Err.Number & ")"
        End If
        GetWinRegSubKeys = OutParams.sNames
    End Function ' GetWinRegSubKeys
End Class
'- End file: ManAppsScanner.vbs
'- Begin file: CompareVersions.vbs
'@SAFELIBINCLUDE
Function CompareVersions( ByVal verNum1, ByVal verNum2 )
' Sumamry:  Perform a case insensitive comparison of two versions
'           (well, really, any two strings). Letters are treated asgreater than
'           numbers, so a23 > 123
' Returns:  1 if first version is greater
'           2 if second version is greater
'           0 if two vrsions are equal
'---------'---------'---------'---------'---------'---------'---------'---------
    CompareVersions = 0
    Dim ver1 : ver1 = Trim(verNum1)
    Dim ver2 : ver2 = Trim(verNum2)
    
    Dim arrVerParts1 : arrVerParts1 = Split( verNum1, "." )
    Dim arrVerParts2 : arrVerParts2 = Split( verNum2, "." )

    Dim posCount
    If UBound( arrVerParts1 ) < UBound( arrVerParts2 ) Then
        posCount = UBound( arrVerParts2 )
    Else
        posCount = UBound( arrVerParts1 )
    End If
 
    Dim i, verPart1, verPart2
    For i = 0 To posCount
        If i <= UBound(arrVerParts1) Then
            verPart1 = arrVerParts1(i)
        Else
            verPart1 = 0
        End If
        If i <= UBound(arrVerParts2) Then
            verPart2 = arrVerParts2(i)
        Else
            verPart2 = 0
        End If
        
        If IsNumeric(verPart1) And IsNumeric(verPart2) Then
            If ( CInt(verPart1) > CInt(verPart2) ) Then
                CompareVersions = 1
                Exit Function
            ElseIf ( CInt(verPart1) < CInt(verPart2) ) Then
                CompareVersions = 2
                Exit Function
            End If
        Else
            Select Case StrComp(verPart1, verPart2, 1)
                Case 1
                    CompareVersions = 1
                    Exit Function
                Case -1
                    CompareVersions = 2
                    Exit Function
            End Select
        End If
    Next
End Function ' CompareVersions
'- End file: CompareVersions.vbs
'- Begin file: GetTaniumComputerId.vbs
'@SAFELIBINCLUDE
Function GetTaniumComputerId(sh)
    On Error Resume Next
    GetTaniumComputerId=Eval("sh.Environment(""Process"")(""TANIUM_CLIENT_ROOT"")") : Err.Clear
    If GetTaniumComputerId="" Then GetTaniumComputerId=Eval("sh.RegRead(""HKLM\Software\Tanium\Tanium Client\ComputerID"")") : Err.Clear
    If GetTaniumComputerId="" Then GetTaniumComputerId=Eval("sh.RegRead(""HKLM\Software\Wow6432Node\Tanium\Tanium Client\ComputerID"")") : Err.Clear
    If GetTaniumComputerId="" Then GetTaniumComputerId=Eval("sh.RegRead(""HKLM\Software\McAfee\Real Time\ComputerID"")") : Err.Clear
    If GetTaniumComputerId="" Then GetTaniumComputerId=Eval("sh.RegRead(""HKLM\Software\Wow6432Node\McAfee\Real Time\ComputerID"")") : Err.Clear
    If GetTaniumComputerId="" Then Err.Clear
    On Error Goto 0
End Function ' GetTaniumComputerId
'- End file: GetTaniumComputerId.vbs
'- Begin file: GetTaniumDir.vbs
'@SAFELIBINCLUDE
Function GetTaniumDir(strSubDir, sh)
    On Error Resume Next
    GetTaniumDir=Eval("sh.Environment(""Process"")(""TANIUM_CLIENT_ROOT"")") : Err.Clear
    If GetTaniumDir="" Then GetTaniumDir=Eval("sh.RegRead(""HKLM\Software\Tanium\Tanium Client\Path"")") : Err.Clear
    If GetTaniumDir="" Then GetTaniumDir=Eval("sh.RegRead(""HKLM\Software\Wow6432Node\Tanium\Tanium Client\Path"")") : Err.Clear
    If GetTaniumDir="" Then GetTaniumDir=Eval("sh.RegRead(""HKLM\Software\McAfee\Real Time"")") : Err.Clear
    If GetTaniumDir="" Then GetTaniumDir=Eval("sh.RegRead(""HKLM\Software\Wow6432Node\McAfee\Real Time\Path"")") : Err.Clear
    If GetTaniumDir="" Then Err.Clear: Exit Function
    On Error Goto 0
    GetTaniumDir=GetTaniumDir & "\" & strSubDir & "\"
End Function ' GetTaniumDir
'- End file: GetTaniumDir.vbs
'- Begin file: IsRegExMatch.vbs
'@SAFELIBINCLUDE
Function IsRegExMatch( StringToTest, RegExQueryString )
    Dim RegX : Set RegX	=   New RegExp
    With RegX
        .Pattern        =   RegExQueryString
        .Global         =   True
        .IgnoreCase     =   True
    End With

	IsRegExMatch = RegX.Test( StringToTest )
End Function ' IsRegExMatch
'- End file: IsRegExMatch.vbs
'- Begin file: LoadXmlDoc.vbs
'@SAFELIBINCLUDE
Function LoadXmlDoc( XmlDocPath )
' Summary:  Loads an XML document and returns the XMLDoc object
' Requires: Logging.vbs
'---------'---------'---------'---------'---------'---------'---------'---------
    Set LoadXmlDoc = CreateObject("Microsoft.XMLDOM")
    With LoadXmlDoc
        .Async = "False"
        .setProperty "SelectionLanguage", "XPath"
    End With

    On Error Resume Next
    LoadXmlDoc.load( XmlDocPath )
    If (LoadXmlDoc.parseError.errorCode <> 0) Then
        Dim xmlLoadErr : Set xmlLoadErr = LoadXmlDoc.parseError
        ERROR "Failed to load XML file " &Chr(34)&XmlDocPath&Chr(34) & " with reason: " & xmlLoadErr.reason
        
        Set LoadXmlDoc = Nothing
    End If
    Err.Clear
    On Error GoTO 0
End Function ' LoadXmlDoc
'- End file: LoadXmlDoc.vbs
'- Begin file: AppendToFile.vbs
'@SAFELIBINCLUDE
Sub AppendToFile( fso_, FileEntry, FilePath )
' Summary:  Appends text to a given file
' Requires: Logging.vbs
'---------'---------'---------'---------'---------'---------'---------'---------
    TRACE "Sub Called: AppendToFile"
    Dim fso, objFile
    set fso = fso_

    On Error Resume Next
    If Not fso.FileExists( FilePath ) Then
        TRACE "Creating File: " & FilePath
        Set objFile = fso.CreateTextFile( FilePath, True )

        If Err.Number <> 0 Then
            ERROR "Failed to create new file: " & FilePath
            ERROR Err.Description & " ( erro number " & Err.Number & ")"
        End If
    Else
        TRACE "Openning File: " & FilePath
        Set objFile = fso.OpenTextFile( FilePath, 8 )

        If Err.Number <> 0 Then
            ERROR "Failed to open existing file: " & FilePath
            ERROR Err.Description & " ( erro number " & Err.Number & ")"
        End If
    End If

    objFile.Write FileEntry
    objFile.Close
    
    Err.Clear
    On Error GoTo 0
End Sub ' AppendToFile
'- End file: AppendToFile.vbs
'- Begin file: WinRegConstants.vbs
'@SAFELIBINCLUDE
Const REG_HIVE_CLASSES_ROOT     = &H80000000
Const REG_HIVE_CURRENT_USER     = &H80000001
Const REG_HIVE_LOCAL_MACHINE    = &H80000002
Const REG_HIVE_USERS            = &H80000003

Const REG_VALUE_STRING          = 1
Const REG_VALUE_EXPANDED_STRING = 2
Const REG_VALUE_BINARY          = 3
Const REG_VALUE_DWORD           = 4
Const REG_VALUE_MULTI_STRING    = 7

Function GetHiveConst( HiveName )
    Select Case HiveName
    Case "HKEY_CLASSES_ROOT"
        GetHiveConst = REG_HIVE_CLASSES_ROOT
    Case "HKCR"
        GetHiveConst = REG_HIVE_CLASSES_ROOT
    Case "HKEY_CURRENT_USER"
        GetHiveConst = REG_HIVE_CURRENT_USER
    Case "HKCU"
        GetHiveConst = REG_HIVE_CURRENT_USER
    Case "HKEY_LOCAL_MACHINE"
        GetHiveConst = REG_HIVE_LOCAL_MACHINE
    Case "HKLM"                 
        GetHiveConst = REG_HIVE_LOCAL_MACHINE
    Case "HKEY_USERS"           
        GetHiveConst = REG_HIVE_USERS
    Case "HKU"                  
        GetHiveConst = REG_HIVE_USERS
    End Select
End Function ' GetHiveConst
'- End file: WinRegConstants.vbs
'- Begin file: GeneratePath.vbs
'@SAFELIBINCLUDE
Function GeneratePath(fso, DirPath)
' Summary:  Recursively generates the the given directory path.
' Returns:  True if the directory already exists or was successfully Created
'           False if full directory path does not exist when function completes
'---------'---------'---------'---------'---------'---------'---------'---------
    GeneratePath = False

    If Not fso.FolderExists(DirPath) Then
        If GeneratePath( fso, fso.GetParentFolderName(DirPath) ) Then
            GeneratePath = True
            fso.CreateFolder(DirPath)
        End If
    Else
        GeneratePath = True
    End If
End Function 'GeneratePath
'- End file: GeneratePath.vbs
'- Begin file: GetFileVersion.vbs
'@SAFELIBINCLUDE
Function GetFileVersion(fso, FilePath, VersionType)
    GetFileVersion = ""

    Dim appsh : Set appsh = CreateObject("Shell.Application")
    Dim FileName : FileName = fso.GetFileName(FilePath)
    Dim FolderPath : FolderPath = Left(FilePath, Len(FilePath)-Len(FileName))

    Dim HeaderName
    Select Case LCase(VersionType)
    Case "file version"
        HeaderName = "file version"
    Case "file"
        HeaderName = "file version"
    Case "product version"
        HeaderName = "product version"
    Case "product"
        HeaderName = "product version"
    Case Else
        HeaderName = "file version"
    End Select

    If fso.FileExists(FilePath) Then
        GetFileVersion = "No version information available"
        Dim objFolder : Set objFolder = appsh.Namespace(FolderPath)
        Dim objFolderItem : Set objFolderItem = objFolder.ParseName(FileName)
        Dim i
        For i = 0 To 300
            Dim Header : Header = objFolder.GetDetailsOf(objFolder.Items, i)
            If LCase(Header) = HeaderName Then
                GetFileVersion= objFolder.GetDetailsOf(objFolderItem, i)
                Exit Function
            End If
        Next
    End If
End Function
'- End file: GetFileVersion.vbs
'- Begin file: GetFileArchitecture.vbs
'@SAFELIBINCLUDE
Function GetFileArchitecture(FilePath)
    ' The Optional PE Header, which immediately follows the PE HEader, begins with a 2-byte magic code
    ' that representing the architecture (0x010B for PE32, 0x020B for PE64, 0x0107 ROM)
    ' By locating and reading this value we can determin the architecture of the file
    On Error Resume Next
    Dim oStream : Set oStream = createobject("ADODB.Stream")
    oStream.type = 1
    oStream.open()
    oStream.loadfromfile(FilePath)
    Dim oBin : set oBin = createobject("Microsoft.XMLDOM").CreateElement("binObj")
    oBin.datatype = "bin.hex"

    ' First, we get the 4 byte address that is the pointer to the PE Header
    oStream.position = 60
    oBin.nodetypedvalue = oStream.read(4)

    ' Next we set the new position in the stream at Optional PE header and get the magic code value
    oStream.position = 1 * ("&H" & mid(oBin.text, 1, 2)) _
                     + 256 * ("&H" & mid(oBin.text, 3, 2)) _
                     + 65536 * ("&H" & mid(oBin.text, 5, 2)) _
                     + 16777216 * ("&H" & mid(oBin.text, 7, 2)) _
                     + 25
    oBin.nodetypedvalue = oStream.read(1)

    ' Finally, convert value to friendly string
    If oBin.text = "02" Then
        GetFileArchitecture = "x64"
    ElseIf oBin.text = "01" Then
       GetFileArchitecture = "x86"
    Else
        GetFileArchitecture = "n/a"
    End If
    On Error GoTo 0
End Function
'- End file: GetFileArchitecture.vbs
'- Begin file: GetWinOsName.vbs
'@SAFELIBINCLUDE
Function GetWinOsName( )
	Dim WMIService : Set WMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Dim OsSet : Set OsSet = WMIService.ExecQuery ("Select * from Win32_OperatingSystem",,48)

	Dim SetItem
	For Each SetItem In OsSet
		Dim Property
		GetWinOsName = SetItem.Caption
	Next
End Function ' GetWinOsName
'- End file: GetWinOsName.vbs
'- Begin file: GetWinOsBits.vbs
'@SAFELIBINCLUDE
Function GetWinOsBits()
	GetWinOsBits = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2:Win32_Processor='cpu0'").AddressWidth
End Function ' GetWinOsBits
'- End file: GetWinOsBits.vbs
'- Begin file: GetWinOsVersion.vbs
'@SAFELIBINCLUDE
Function GetWinOsVersion( )
	Dim WMIService : Set WMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
	Dim OsSet : Set OsSet = WMIService.ExecQuery ("Select * from Win32_OperatingSystem",,48)

	Dim SetItem
	For Each SetItem In OsSet
		Dim Property
		GetWinOsVersion = SetItem.Version
	Next
End Function ' GetWinOsVersion
'- End file: GetWinOsVersion.vbs
