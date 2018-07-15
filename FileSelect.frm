VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.ocx"
Begin VB.Form FileSelect 
   Caption         =   "File Select"
   ClientHeight    =   4230
   ClientLeft      =   1620
   ClientTop       =   1830
   ClientWidth     =   5565
   Icon            =   "FileSelect.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   5565
   Begin VB.TextBox txtFile 
      Height          =   1815
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   1680
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "FileSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'cdlOFNOverwritePrompt  'prompt Overwite if exists
'cdlOFNExplorer use Explorer like dialog box
'cdlOFNLongNames USE long filenames
'cdlOFNHideReadOnly hide read only check box
'cdlOFNCreatePrompt prompt to create if doesnt exist
Private Sub Form_Load()
'    dlgFile.FileName = "default.ini"
    dlgFile.InitDir = App.Path
'Decription 1|Filter1|Description 2|Filter2
    dlgFile.Filter = "All Files (*.*)|*.*"
    dlgFile.FilterIndex = 1 'use filter1 as the default
    dlgFile.DialogTitle = "Default Settings"
        dlgFile.Flags = _
        cdlOFNOverwritePrompt + _
        cdlOFNLongNames + _
        cdlOFNHideReadOnly + _
        cdlOFNExplorer
    dlgFile.CancelError = True
End Sub

'Ask user which filename to use
Public Function AskFileName(FileReq As String, Optional FileMustExist As Boolean, Optional ByRef Cancel As Boolean) As String
Dim FileFound As Boolean
Dim Count As Long
Dim errmsg As String

'IF the specified cmd file not found, ask for a file
Do
'Set Flags if not already set, may be from previous pass
    If FileMustExist And ((dlgFile.Flags And cdlOFNFileMustExist) <> cdlOFNFileMustExist) Then
        dlgFile.Flags = dlgFile.Flags + cdlOFNFileMustExist
    End If
'Unset flags if file does not need to exists (SaveAs)
    If FileMustExist = False And ((dlgFile.Flags And cdlOFNFileMustExist) = cdlOFNFileMustExist) Then
        dlgFile.Flags = dlgFile.Flags - cdlOFNFileMustExist
    End If
    Select Case FileReq
'cmd file is only here for testing, program does not use this call
    Case Is = "IniFileName", "cmdIniFileName"
        Call SetFileName(FileReq)
        dlgFile.Filter = "Initialisation Files (*.ini)|*.ini" _
        & "|All Files (*.*)|*.*"
        dlgFile.DialogTitle = "Initialisation File"
'do not require to remove overwrite propmt (only actioned with SaveAs)
        On Error Resume Next
        dlgFile.ShowOpen
    Case Is = "NmeaLogFile"
        Call SetFileName(FileReq)
        dlgFile.Filter = "Log Files (*." & _
        ExtFromFullName(dlgFile.FileName) & ")|*." & _
        ExtFromFullName(dlgFile.FileName) & "|All Files (*.*)|*.*"
        dlgFile.DialogTitle = "NMEA Input Data Log File"
'do not require to remove overwrite propmt (only actioned with SaveAs)
        On Error Resume Next
        dlgFile.ShowOpen
    Case Is = "OutputFileName"
        Call SetFileName(FileReq)
        dlgFile.Filter = _
         "NMEA (*.nmea)|*.nmea" _
        & "|CSV (*.csv)|*.csv" _
        & "|Google Earth (*.kml;*.kmz)|*.kml;*.kmz" _
        & "|Google Maps (*.kml)|*.kml" _
        & "|HTML (*.html,*.htm)|*.html;*.htm" _
        & "|ASP (*.aspx)|*.aspx" _
        & "|All Files (*.*)|*.*"
'set default filter
        Select Case LCase$(ExtFromFullName(dlgFile.FileName))
        Case Is = "nmea"
            dlgFile.FilterIndex = 1
        Case Is = "csv"
            dlgFile.FilterIndex = 2
        Case Is = "kml"
            dlgFile.FilterIndex = 3
        Case Is = "html", "htm"
            dlgFile.FilterIndex = 5
        Case Is = "aspx"
            dlgFile.FilterIndex = 6
        Case Else
            dlgFile.FilterIndex = 7
        End Select
        dlgFile.DialogTitle = "Output File Name"
'do not require to remove overwrite propmt (only actioned with SaveAs)
        On Error Resume Next
        dlgFile.ShowSave
    Case Is = "NmeaReadFile"
        Call SetFileName(FileReq)
        dlgFile.Filter = "Log Files (*." & _
        ExtFromFullName(dlgFile.FileName) & ")|*." & _
        ExtFromFullName(dlgFile.FileName) & "|All Files (*.*)|*.*"
        dlgFile.DialogTitle = "NMEA Data File"
'do not require to remove overwrite propmt (only actioned with SaveAs)
        On Error Resume Next
        dlgFile.ShowOpen
    Case Is = "TagTemplateReadFile"
        Call SetFileName(FileReq)
        dlgFile.Filter = _
         "Google Earth Template (*.kml)|*.kml" _
        & "|Google Maps Template (*.kml)|*.kml" _
        & "|HTML Template (*.html,*.htm)|*.html;*.htm " _
        & "|ASP Template (*.aspx)|*.aspx " _
        & "|All Files (*.*)|*.*"
'set default filter
        Select Case LCase$(ExtFromFullName(dlgFile.FileName))
        Case Is = "kml"
            dlgFile.FilterIndex = 1
        Case Is = "html", "htm"
            dlgFile.FilterIndex = 3
        Case Is = "aspx"
            dlgFile.FilterIndex = 4
        Case Else
            dlgFile.FilterIndex = 5
        End Select
        dlgFile.DialogTitle = "Output Template File"
'we do not need to remove overwrite propmt (only actioned with SaveAs)
        On Error Resume Next
        dlgFile.ShowOpen
'if user cancels without selecting a file, file will still be not set
'if directory is valid
    Case Else   'just used for testing
        Call SetFileName(FileReq)
        dlgFile.DialogTitle = "Not called by AskFileName (Testing only)"
        On Error Resume Next
        dlgFile.ShowOpen

    End Select
        
    If err.Number = 0 Then
'dlgFile.FileName contains dev+Path+Name selected
'Unchanged if not selected (if set on load, it will not be null)
        AskFileName = dlgFile.FileName
     Else
        If err.Number = cdlCancel Then
' The user canceled the replace
'return the default for this file
            AskFileName = dlgFile.InitDir & dlgFile.FileName
            MsgBox "Current " & dlgFile.DialogTitle & " is" & vbCrLf _
            & AskFileName, vbInformation, "File Select"
            Cancel = True   'return to caller
        Else
        ' Unknown error.
            MsgBox "Error " & Format$(err.Number) & _
            " selecting file." & vbCrLf & _
            err.Description
        End If
        On Error GoTo 0
    End If

'If (dlgFile.Flags And cdlOFNFileMustExist) = cdlOFNFileMustExist Then
'MsgBox "File must exist"
'End If

'MsgBox AskFileName
    If cmdJna = True Then
        TestFileSelect.Show
        TestFileSelect.Text1.Text = AskFileName
        TestFileSelect.Text2.Text = FileReq
        TestFileSelect.Command1(0).SetFocus
    End If
'check if must exists anyway
    If FileMustExist = False Then Exit Do
        
    FileFound = FileExists(AskFileName)
    If FileFound = False Then
        errmsg = "Not Found"
    Else
        If FileReq = "IniFileName" Then
            If IsIniFile(AskFileName) = False Then
                errmsg = "Has incorrect format"
                FileFound = False
            End If
        End If
    End If
    
    If FileFound = True Then
        Exit Do
    Else
        Count = Count + 1
        MsgBox AskFileName & vbCrLf & errmsg & " (" & Count & ")" & vbCrLf, _
            vbOKOnly, "File Not Found Error"
        If Count >= 4 Then
            MsgBox "Exiting AisDecoder after 4 retries" & vbCrLf & "You must select an existing file for successful initialisation", _
            vbOKOnly + vbCritical, "File Select Error"
            If StartupLogFileCh <> 0 Then Call WriteStartUpLog("Does not exist, Terminating")
            Call Terminate
            Exit Function
        End If
    End If
Loop Until FileFound = True
    
    Unload Me
End Function

'returns the current file name that is in use
'Called by SetFileName, in turn called by AskFileNAme
Private Function GetFileName(FileReq As String) As String
    Select Case FileReq
    Case Is = "StartupLogFile"
        GetFileName = StartupLogFile
    Case Is = "cmdIniFileName"
        GetFileName = cmdIniFileName
    Case Is = "IniFileName"
        GetFileName = IniFileName
    Case Is = "ErrorLogFile"
        GetFileName = ErrorLogFile
    Case Is = "NmeaLogFile"
        GetFileName = NmeaLogFile
    Case Is = "OutputFileName"
        GetFileName = OutputFileName
    Case Is = "OverlayOutputFileName"
        GetFileName = OverlayOutputFileName
    Case Is = "OverlayTemplateReadFile"
        GetFileName = OutputFileName
    Case Is = "TagTemplateReadFile"
        GetFileName = TagTemplateReadFile
    Case Is = "VesselsFileName"
        GetFileName = VesselsFileName
    Case Is = "TrappedMsgsFileName"
        GetFileName = TrappedMsgsFileName
    Case Is = "NmeaReadFile"
        GetFileName = NmeaReadFile
    Case Is = "FtpLocalFileName"
        GetFileName = FtpLocalFileName
    Case Is = "FtpRemoteFileName"
        GetFileName = FtpRemoteFileName
    Case Is = "TcpLoginFileName"
        GetFileName = TcpLoginFileName
    End Select
End Function

'Tne Calling function must set  NmeaLogFile = SetFileName(filereqd)
'Checks the Current Path exists, if not replace with default
'Checks the Current FileName set, if not replace with default
'on exit the full file name should be valid
'Also called by AskFileName to display the current file name
Public Function SetFileName(FileReq As String) As String
Dim FileName As String
Dim FileDate As String
    
    Select Case FileReq
'could be a rollover, checks if rollover on this file
    Case Is = "NmeaLogFile"
        FileName = GetFileName(FileReq)
        If FolderExists(PathFromFullName(FileName)) Then
            dlgFile.InitDir = PathFromFullName(FileName) _
            & "\"   'Important
        Else
            dlgFile.InitDir = DefaultPath(FileReq)
        End If
'True removes any RollOver Date from the filename
        dlgFile.FileName = NameFromFullPath(FileName, , True)
        If dlgFile.FileName = "" Then
            dlgFile.FileName = DefaultName(FileReq)
'default name may include rollover date
            dlgFile.FileName = NameFromFullPath(dlgFile.FileName, , True)
        End If
'insert new rollover date (if required)
        If TreeFilter.Check1(3).Value <> 0 Then  'rollover
            FileDate = Format$(Now(), "yyyy-mm-dd")
            dlgFile.FileName = _
            ExtendFullName(dlgFile.FileName, "_" _
            & Format$(FileDate, "yyyymmdd"))
        Else
            FileDate = ""
        End If
        NmeaLogFileDate = FileDate
        If IsFileInUse(dlgFile.InitDir & dlgFile.FileName) Then
           FileName = UniqueFileName(dlgFile.InitDir & dlgFile.FileName)
'MsgBox "Set File Name (Rollover added)" & vbCrLf & FileName
'False leaves any RollOver Date from the filename
            dlgFile.FileName = NameFromFullPath(FileName, , False)
        End If
        SetFileName = dlgFile.InitDir & dlgFile.FileName
'could be a rollover, checks if rollover on this file
    Case Is = "OutputFileName", "OverlayOutputFileName"
        FileName = GetFileName(FileReq)
        If FolderExists(PathFromFullName(FileName)) Then
            dlgFile.InitDir = PathFromFullName(FileName) _
            & "\"   'Important
        Else
            dlgFile.InitDir = DefaultPath(FileReq)
        End If
'True removes any RollOver Date from the filename
'MsgBox FileName
        dlgFile.FileName = NameFromFullPath(FileName, , True)
        If dlgFile.FileName = "" Then dlgFile.FileName = DefaultName(FileReq)
'insert new rollover date (if required)
        If OutputFileRollover(1) = True Then 'rollover
            FileDate = Format$(Now(), "yyyy-mm-dd")
            dlgFile.FileName = _
            ExtendFullName(dlgFile.FileName, "_" _
            & Format$(FileDate, "yyyymmdd"))
        Else
            FileDate = ""
        End If
        OutputFileDate = FileDate
        If IsFileInUse(dlgFile.InitDir & dlgFile.FileName) Then
           FileName = UniqueFileName(dlgFile.InitDir & dlgFile.FileName)
'MsgBox "Set File Name (Rollover added)" & vbCrLf & FileName
'False leaves any RollOver Date from the filename
            dlgFile.FileName = NameFromFullPath(FileName, , False)
        End If
        SetFileName = dlgFile.InitDir & dlgFile.FileName
'        If FileReq = "OverlayOutputFileName" Then
'MsgBox SetFileName & ":" & Len(SetFileName)
'            SetFileName = ExtendFullName(SetFileName, "_" & CStr(GetCurrentProcessId))
'MsgBox SetFileName & ":" & Len(SetFileName)
'        End If
        
'no rollover Full File name
    Case Is = "NmeaReadFile", "VesselsFileName", _
    "TrappedMsgsFileName", "ErrorLogFile" _
    , "TagTemplateReadFile", "IniFileName", "ExitLogFile" _
    , "StartupLogFile", "cmdIniFileName", "OverlayTemplateReadFile", "TcpLoginFileName"
        FileName = GetFileName(FileReq)
        If FolderExists(PathFromFullName(FileName)) Then
            dlgFile.InitDir = PathFromFullName(FileName) _
            & "\"   'Important
        Else
            dlgFile.InitDir = DefaultPath(FileReq)
        End If
'True removes any RollOver Date from the filename
        dlgFile.FileName = NameFromFullPath(FileName, , True)
        If dlgFile.FileName = "" Then dlgFile.FileName = DefaultName(FileReq)
        SetFileName = dlgFile.InitDir & dlgFile.FileName
    
'FtpLocalFileName is always same as OutputFileName
    Case Is = "FtpLocalFileName"
        FileName = GetFileName(FileReq)
        If FolderExists(PathFromFullName(FileName)) Then
            dlgFile.InitDir = PathFromFullName(FileName) _
            & "\"   'Important
        Else
            dlgFile.InitDir = DefaultPath(FileReq)
        End If
'FtpLocalFileName is always same as OutputFileName
        dlgFile.FileName = DefaultName(FileReq)
        SetFileName = dlgFile.InitDir & dlgFile.FileName
    
'The filename can include the directory with / as delimiter
    Case Is = "FtpRemoteFileName"
        FileName = GetFileName(FileReq)
        If FolderExists(PathFromFullName(FileName, "/")) Then
            dlgFile.InitDir = PathFromFullName(FileName, "/") _
            & "/"   'Important
        Else
            dlgFile.InitDir = DefaultPath(FileReq)
        End If
        dlgFile.FileName = NameFromFullPath(FileName, "/")
        If dlgFile.FileName = "" Then dlgFile.FileName = DefaultName(FileReq)
        SetFileName = dlgFile.InitDir & dlgFile.FileName
    End Select
    
'Carry out any actions needed as a result of the file
'name changes
    Select Case FileReq
'reset the appropriate nmea,csv or tag file names
'must be done at the end otherwise it is set to the overlayoutputfile
    Case Is = "OutputFileName"
            Call SetOutputFileNames(SetFileName)
'We must check if the Template file needs parsing again
'Now checked every Start in StartInputOutput as it may have been changed
'    Case Is = "TagTemplateReadFile"
'        Call NmeaRcv.ReadTagTemplate(SetFileName)
    End Select


'refresh the files list
If Files.Visible Then Call Files.RefreshData

'Used to debug FileSelect
'If cmdJna = True Then
'    TestFileSelect.Show
'    TestFileSelect.Text1.Text = SetFileName
'    TestFileSelect.Text2.Text = FileReq
'    TestFileSelect.Command1(1).SetFocus
'End If

End Function

'returns Default FileName
Private Function DefaultName(FileReq As String) As String
    Select Case FileReq
    Case Is = "cmdIniFileName"
        DefaultName = ""
    Case Is = "IniFileName"
        DefaultName = "default.ini"
    Case Is = "ErrorLogFile"
        DefaultName = "error.log"
    Case Is = "NmeaLogFile"
        DefaultName = "nmea.log"
    Case Is = "OutputFileName"
'only output to file on channel(1)
'try using last name first
        Select Case ChannelFormat(1)
        Case Is = "nmea"
            DefaultName = "output.nmea"
        Case Is = "csv"
            DefaultName = "output.csv"
        Case Is = "tag"
            DefaultName = NameFromFullPath(TagTemplateReadFile)
        End Select
    Case Is = "OverlayOutputFileName"
'call in Name then path with a "." delimeter extracts only the name
'without both the path and the .ext
        DefaultName = PathFromFullName(NameFromFullPath(TagTemplateReadFile), ".") & "_link.kml"
    Case Is = "TagTemplateReadFile"
        DefaultName = "data.xml"
    Case Is = "VesselsFileName"
        DefaultName = "Vessels.dat"
    Case Is = "TrappedMsgsFileName"
        DefaultName = "TrappedMsgs.dat"
    Case Is = "NmeaReadFile"
        DefaultName = "nmea.log"
    Case Is = "StartupLogFile"
        DefaultName = "AisDecoderStartup.log"
    Case Is = "ExitLogFile"
        DefaultName = "AisDecoderExit.log"
    Case Is = "FtpLocalFileName"
        DefaultName = NameFromFullPath(OutputFileName)
    Case Is = "FtpRemoteFileName"
        DefaultName = NameFromFullPath(FtpLocalFileName)
    Case Is = "TcpLoginFileName"
        DefaultName = "TcpLogin.cmd"
    End Select
End Function

'returns the default path
Public Function DefaultPath(FileReq As String) As String
    Select Case FileReq
'/Settings
    Case Is = "IniFileName", "cmdIniFileName"
        DefaultPath = Environ("APPDATA") & "\Arundale\Ais Decoder\Settings\"
'/Logs
    Case Is = "ErrorLogFile", "NmeaLogFile", "NmeaReadFile"
        DefaultPath = Environ("APPDATA") & "\Arundale\Ais Decoder\Logs\"
'/Output
Case Is = "OutputFileName", "FtpLocalFileName", "OverlayOutputFileName"
        DefaultPath = Environ("APPDATA") & "\Arundale\Ais Decoder\Output\"
'/Templates
    Case Is = "TagTemplateReadFile", "OverlayTemplateReadFile"
        DefaultPath = Environ("APPDATA") & "\Arundale\Ais Decoder\Templates\"
'/Files
    Case Is = "VesselsFileName", "TrappedMsgsFileName", "TcpLoginFileName"
        DefaultPath = Environ("APPDATA") & "\Arundale\Ais Decoder\Files\"
'/Temp
        Case Is = "StartupLogFile"
        DefaultPath = LongFileName(Environ("TEMP") & "\")
    Case Is = "ExitLogFile"
        DefaultPath = LongFileName(Environ("TEMP") & "\")
'FTP
    Case Is = "FtpRemoteFileName"
        DefaultPath = TreeFilter.Text1(7)
        If Right$(TreeFilter.Text1(7), 1) <> "/" _
        And Left$(TreeFilter.Text1(10), 1) <> "/" Then _
        DefaultPath = DefaultPath & "/"
        DefaultPath = DefaultPath & TreeFilter.Text1(10) & "/"
    End Select
End Function

Function SetOutputFileNames(FileName As String)
'Set appropriate individual output file names to same
'name as output file, must be done when output file name has been changed
    Select Case ChannelFormat(1)
    Case Is = "nmea"
        OutputFileNameNmea = OutputFileName
    Case Is = "csv"
        OutputFileNameCsv = OutputFileName
    Case Is = "tag"
        OutputFileNameTagged = OutputFileName
'Output file can be .kml or .kmz
        OverlayOutputFileName = Replace(OutputFileName, ".kmz", ".kml")
        OverlayOutputFileName = Replace(OverlayOutputFileName, ".kml", "_link.kml")
'MsgBox OverlayOutputFileName
    End Select
    TreeFilter.lblOutputFile = NameFromFullPath(FileName)
End Function

