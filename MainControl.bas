Attribute VB_Name = "MainControl"
Option Explicit
'Copyright@ 2009-17 Neal Arundale
Public debugCount As Long
Public DebugTreeFilter As Boolean

Private Const PI = 3.14159265358979 'used for calculating range
'My Declarations followMaxNmeaDecodeListCount
Public Const NMEABUF_MAX = 10000 'NmeaBuf - all sentences here first
Public Const MAXRCVLISTCOUNT = 2000 'Nmea Rcv Display
Public Const MINNMEAOUTBUFSIZE = 100 'NmeaOutBuf - all sentences not yet processed by scheduler
Public Const MAXNMEADECODELISTCOUNT_Rcv = 3000  'If not displaing scheduled output
Public Const MAXNMEADECODELISTCOUNT_Sched = 30000  'If displaying schedule output
Public MaxNmeaDecodeListCount As Long   'Filtered Summary Display
Public MaxOutputDisplayCount As Long    'const causes compile error !
Public Const MYFMT = "@@@@@@@@@@@@@@@@@@@@@@@@"
Public Const MILLISECS_PER_MIN = 60000
Public Const SPEED_INTERVAL = 6000     'millisecs

'Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
'see http://support.microsoft.com/kb/224816
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function ShellExecute _
                            Lib "SHELL32.DLL" _
                            Alias "ShellExecuteA" ( _
                            ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) _
                            As Long

'http://www.jrsoftware.org/ishelp/index.php?topic=setup_appmutex
Private Declare Function CreateMutex Lib "kernel32" _
        Alias "CreateMutexA" _
       (ByVal lpMutexAttributes As Long, _
        ByVal bInitialOwner As Long, _
        ByVal lpName As String) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Dim hMutex As Long

Public DownloadURL As String
Public MyNewVersion As New clsNewVersion

'SetTreeViewBackColor - Change the background color of a TreeView control
' see http://www.devx.com/vb2themax/Tip/19099
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As _
    Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd _
    As Long, ByVal nIndex As Long) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd _
    As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function IsNTAdmin Lib "advpack.dll" (ByVal dwReserved As Long, ByRef lpdwReserved As Long) As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long

Public Const GWL_STYLE = -16&
Public Const TVM_SETBKCOLOR = 4381&
Public Const TVS_HASLINES = 2&
'-----------------------------------------------
'http://visualbasic.about.com/b/2005/10/11/a-globalizing-trick-for-vb-6.htm
'see http://support.microsoft.com/?kbid=221435 for list
Private Declare Function GetUserDefaultLCID% Lib "kernel32" ()
'-----------------------------------------------
'-----------------------------------------------
'http://vbnet.mvps.org/index.html?code/locale/gettimezonebias.htm
Private Const TIME_ZONE_ID_UNKNOWN As Long = 0
Private Const TIME_ZONE_ID_STANDARD As Long = 1
Private Const TIME_ZONE_ID_DAYLIGHT As Long = 2
Private Const TIME_ZONE_ID_INVALID As Long = &HFFFFFFFF

Public Type SYSTEMTIME
   wYear         As Integer
   wMonth        As Integer
   wDayOfWeek    As Integer
   wDay          As Integer
   wHour         As Integer
   wMinute       As Integer
   wSecond       As Integer
   wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
   Bias As Long
   StandardName(0 To 63) As Byte  'unicode (0-based)
   StandardDate As SYSTEMTIME
   StandardBias As Long
   DaylightName(0 To 63) As Byte  'unicode (0-based)
   DaylightDate As SYSTEMTIME
   DaylightBias As Long
End Type

Private Declare Function GetTimeZoneInformation Lib "kernel32" _
    (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long

'--------------------------
'For NowUtc()
Public Declare Sub GetSystemTime Lib "kernel32.dll" (lpSystemTime As SYSTEMTIME)

'Used to see if a ini file exists in ExpEnvStr
Declare Function ExpandEnvironmentStrings _
   Lib "kernel32" Alias "ExpandEnvironmentStringsA" _
   (ByVal lpSrc As String, ByVal lpDst As String, _
   ByVal nSize As Long) As Long

Public AisMsgs As New Collection
Public Vessels As New Collection
'Public MmsiTags As New Collection  'Not used (added to vessels)
Public TrappedMsgs As New Collection

Public AllFields() As String    'all fields if AllCsv to be output
Public AllFieldsNo As Long     'number of fields in this message

Public Type FieldArrayDef
    MsgKey As String    '1
    Source As String    '2
    Member As String    '3
    From As Long        '4
    reqbits As Long     '5
    Arg As String       '6
    Arg1 As String      '7
    Column As Long      '8
'    MsgType As String
'    Dac As String
'    Fi As String
'    FiId As String
    Tag As String       '9
    Valdes As String    '10
'    Value As String
End Type

Public Type ShipDef
    Mmsi As String
    Name As String
    RcvTime As String
    Lat As Single  'Set if sentence is !AIVDO
    Lon As Single
End Type
    
Private Type MungeCurr
    Value As Currency
End Type

Private Type Munge2Long
    LowVal As Long
    HighVal As Long
End Type

Private Type defProcessing
    Suspended As Boolean
    NmeaBuf As Boolean
    Scheduler As Boolean
    InputOptions As Boolean
    Paused As Boolean
End Type

Private Type defGis
    PositionOK As Boolean    'For this vessel we have a valid position
    Heading As String   'heading if n/a then COG
    Course As String   'last COG
    IScale As String     '1+-10% for each meter +-100
    Color As String
    Name As String
End Type

Private Type defTrack
    TCPCallCountExit As Long
    TCPArrival As Long
    TCPDataRcvLFCount As Long
    TCPDataRcvLoops As Long
End Type
Public Track As defTrack
'This array is zero based and maps directly to treefilter.msflexgrid
'which is the ais fields. The first element represents the
'first fixed row (header). The second element will be blank if there
'are no fields in msflexgrid
Public FieldArray() As FieldArrayDef
Public FieldArrayFromTo(0 To 63, 0 To 1) 'aismsgtype 0 is other, from/to)

Public LicenseStatus As String
Public LicenseLevel As String

' no longer used Public WebCurrentInstallFileName As String       'Installation executable to download
' ditto Public TempInstallFile As String       'Installation exe on PC

 'Debug Variables
Public HexDumpFileName As String    'Debug
Public HexDumpFileCh As Long
Public HexDump As Boolean   'In registry \Settings\HexDump = "True"

Public IniFileName As String        'path & filename
Public IniFileCh As Long

Public ErrorLogFile As String
Public ErrorLogFileCh As Long       'nmea log

Public CacheLogFile As String
Public CacheLogFileCh As Long       'Cache Log
Public CacheRecord As Long

Public NmeaLogFile As String        'Used to Set InputLogFile.Name
Public NmeaLogFileDate As String    'date was opened, blank unless rolling over

Public OutputFileName As String
Public ZipOutputFile As Boolean     'When created the Output filename will have .unzipped appended
'Public OutputFileCh As Long
Public OutputFileDate As String     'date was opened, blank unless rolling over
Public OutputDataFile As HugeBinaryFile
Public OutputFileErr As Boolean     'True if we cant open it

Public OverlayOutputFileName As String  'full path
Public OverlayOutputFileCh As Long

Public VesselsFileName As String
Public VesselsFileSyncTime As SYSTEMTIME 'Checked by TimeZoneTimer every sec, updated hourly
Public VesselTagsHeader As String   'TagArray(0,0),TagArray(n,0) eg "receivedtime_0_2,vesselname_0_2"

Public TrappedMsgsFileName As String
Public TrappedMsgsFileSaveTime As Single

Public TagTemplateReadFile As String
Public TagTemplateReadFileCh As Long
Public OverlayTemplateReadFile As String    'Currently hard coded

Public ShellFileName As String
Public ShellFileCh As Long

Public NmeaReadFile As String    'Nmea file were reading InputFile.Name

'Public StartupLogFile As String    'Moved to modGeneral
'Public StartupLogFileCh As Long    'Moved to modGeneral
Public ExitLogFile As String
Public ExitLogFileCh As Long

Public FtpUserName As String
Public FtpIniPasswordEncrypted As String     'Encrypted written back on save unless changed
Public FtpIniPasswordDecrypted As String
Public FtpPasswordEncrypted As String        'Current Encrypted Password
Public FtpPassword As String                'Current decrypted password
Public FtpRemoteFileName As String  'includes domain etc
Public FtpLocalFileName As String   'excludes path

Public UserLicenceFileName As String 'The encrypted file name (used in show files)

Public TcpLoginFileName As String   'TCP command file (TcpLogin.cmd)

Public TimeLocal As String
Public TimeUTC As String
'Public TimeOutput As String 'Local or GMT

Public NmeaBufXoff As Boolean   'True = dont send any more data - don't call InputBuffersToNmeaBuf
Public NmeaBuf() As String
Public NmeaBufNxtIn As Long
Public NmeaBufNxtOut As Long
Public NmeaBufUsed As Long

Public cBytesRx As Currency
Public Received As Double
Public LastReceived As Double   'used to calculate speed
Public LastSpeed As Double
Public ForecastSpeed As Double
Public Rejected As Double
'Public Waiting As Double   'Now NmeaBufUsed
Public Processed As Double
Public Filtered As Double
Public Outputted As Double
Public NamedVessels As Double
Public SchedOut As Double

Public TotalRcv As Double
Public TotalRejected As Double
Public TotalRcvTime As Double

'This alters the message count step to debug overflows at high receive rates
'Normally this will be set to 1
Public Const MESSAGE_COUNT_STEP = 1

Public LastTimeProcessed As Single 'last time this function was called (secs past midnight)
Public DecoderStartTime As String   'Now() locale

Public ListFilters(0 To 4) As New ListFilter
Public Outputs(1 To 2) As New Output    'Channels, 1=File,2=UDP

Public cmdStart As Boolean      'start passed on command line
Public cmdNoWindow As Boolean
Private cmdMinimize As Boolean
Public cmdIniFileName As String
Public cmdJna As Boolean
'V142 Public QueryQuit As Boolean   'True = query progam exit
Public EditOptions As Boolean   'True if user allowed to change options
Public TestDac As Boolean       'True if mapping
Private arry() As String        'To create Long array
Public DacMap(0 To 1) As Long         '0=From,1=to
Public DisableNmeaFillBitsError As Boolean  'True = don't check if false (assume = true)
Public LastProgramUpdateCheckTime As SYSTEMTIME

'these two are created for the current sentence at the same time
Public clsSentence As New clsInputSentence   'create
Public clsCb As New clsInputCB   'create
Public GroupSentence As New clsGroupSentence    'applies to all sentences in this group

Public clsField As New clsOutputField
Public NmeaWords() As String    'cant keep dynamic array in class module
'Public NmeaCrc As String    'crc is removed from nmeaword containing crc
Public CbWords() As String    'cant keep dynamic array in class module
Public GroupWords() As String    'cant keep dynamic array in clsGroupSentence

Public TagArray() As String  '(Row,Col) Cols 0=Tag,1=Value,2=RcvTime,3=Min,4=Max,5-Name(for csvhead)
                                'The Tag is is the Unique FieldTag (not the user's tag name)
Public LastSentence(1 To 9) As String 'last multi-part ID nos
'only completed by TagsFromFields so we can output previous parts
'*Public NmeaArray(1 To 100) As String    'BufferedNmea Output from Scheduler for each mmsi
'Public NmeaArrayNo As Long
Public NmeaOutBuf() As String
Public NmeaOutBufCount As Long

Public Gis As defGis

'Public IconHeading As String   'heading if n/a then COG
'Public IconCourse As String   'last COG
'Public IconScale As String     '1+-10% for each meter +-100
'Public IconColor As String
'Public IconName As String
'Public LatOk As Boolean     'true if we have a lat for GIS output
'Public LonOK As Boolean     'true if we have a lon

Public DateTimeOutputFormat As String    'dd/mm/yyyy hh:nn:ss
Public DecimalSeparator As String   'Used for lat/lon range check

'separately defined to speed output & set up if treefilter
'0 True if Either channel(1) or channel(2) is true
'option1 or check1 clicked
Public NmeaOutput(0 To 2) As Boolean        'one of these must be true
Public CsvOutput(0 To 2) As Boolean         'true if tagged csv
Public CsvAll(0 To 2) As Boolean            'output all fields as csv
Public TaggedOutput(0 To 2) As Boolean
Public ShellOn(0 To 2) As Boolean           'Sill shell on output close
Public TaggedOutputOn(1 To 2) As Boolean    'true when first output
Public ChannelMethod(1 To 2) As String      'file,udp
Public ChannelFormat(1 To 2) As String      'nmea,csv,tag
Public ChannelEncoding(1 To 2) As String    'none,xml,kml
Public OutputFileRollover(1 To 2) As Boolean    'true if rolling over
Public NoDataOutput(1 To 2) As Boolean      'Since Output File Opened
Public CsvDelim(0 To 2) As String           '0 set if 1 & 2 the same
Public ChannelOutput(0 To 2) As Boolean 'if method output
'Public ChannelOutputOk(0 To 2) As Boolean   'this group is ok to output
                                            '0 =ok on at least one channel
Public DisplayOutput(1 To 2) As Boolean 'if display output on channel
                        'only displays if outputting to file or udp
Public RangeReq(0 To 2) As Boolean
Public ScheduledReq(0 To 2) As Boolean
Public FenReq(0 To 2) As Boolean
Public GisReq(0 To 2) As Boolean
Public TimeStampReq(0 To 2) As Boolean
Public TagsReq(0 To 2) As Boolean 'csv or tagged output
Public MethodOutput(0 To 2) As Boolean  'If udp or file output
Public FtpOutput(0 To 2) As Boolean 'True/false
'dont think LatOK/LonOK are used
Public VesselLatMax As Single   'Will be rejected by ChkSentenceFilter if
Public VesselLatMin As Single   'clssentence.VesselLat/Lon outside range
Public VesselLonMax As Single   'Note applies to Static vessel messages as well
Public VesselLonMin As Single
Public ChkVesselLatMax As Boolean   'Because we may want to filter Lat or Lon=0 or even
Public ChkVesselLatMin As Boolean   'Lat =91, Lon=181 (Not available) we need to
Public ChkVesselLonMax As Boolean   'set a flag if the Lat or lon requires checking
Public ChkVesselLonMin As Boolean   'and they must be checked separately for max & min
Public OutputFileNameNmea As String
Public OutputFileNameCsv As String
Public OutputFileNameTagged As String
Public SpawnGisOk As Boolean    'kml file is being created
Public TagTemplateHead(1 To 2) As String
Public TagTemplateContent(1 To 2) As String
Public TagTemplateTail(1 To 2) As String
Public OverlayReq(1 To 2) As Boolean 'set when TagTemplateContent is parsed
Public CsvHead(1 To 2) As String
Public CsvHeadOn(1 To 2) As Boolean
Public TagChr(0 To 1) As String     'tag open & close characters <> or []
Public CacheVessels As Boolean      'If True or Changed to True, update Vessels.dat
Public CacheTrappedMsgs As Boolean
Public SchedKeyLen As Long

Public MyShip As ShipDef    'This is the current VDO ship

Public SerialPortCount As Long  'No of serial ports on PC

'Public DecodingState As Long    '9=Loading
                                '0=Available
                                '1=ProcessingNmeaBuf
                                '2=Processing AisMsgs (Scheduler)
                                '3=DecodeList (Summary)
                                '4=RcvList (NmeaInput)
                                '5=TrappedMessages
                                '6=CreatingTag
                                '7=DecodeSentence
Public Processing As defProcessing
'Public ProcessSuspended As Boolean
'Public SuspendProcessScheduler As Boolean
'Public SuspendProcessList As Boolean
'Public SuspendProcessOptionsInput As Boolean
'Public SuspendProcessOutput As Boolean  'fix expression too complex
'v131 Public ProcessingSentence As Boolean   'In module ProcessSentence
Public FtpUploadExecuting As Boolean
Public LastTimeOutput As String 'Last time scheduler run

'seconds user has requested between last and next output
Public UserScheduledSecs As Long 'Next is the calculated freq after adjusting for output time
'Adjustment required to slow down scheduled frequency if
'user request is too fast for output to complete
'Public UserFieldTagName As String   'Input from frmFieldInput

Public ScheduledFreqAdj As Long 'NOT in use
Public strDebug As String

'This ONLY set by clicking the command buttons - Used to control program flow
'so it is faster to be referenced by - "If program state = ? then"
'rather than calling a subroutine to work out the state from the buttons
Public InputState As Long     '0=Stopped,1=Running(NotPaused),2=Paused

Dim addComma As Boolean 'test v129 compatibility
                                                                
Sub Main()
Dim InstalledBuildNo As Long    'this is the app that is currently running
Dim LastUpdatedBuildNo As Long  'last build no updated for this user
Dim oFileSys As New FileSystemObject
Dim oFolder As Folder
Dim FallBackFile As String
Dim NewUserFile As String
Dim HKCUFile As String
Dim kb As String
Dim StartUpCommand As String
Dim cmdOptions() As String
Dim i As Long
Dim j As Long
Dim ch As Long  'Temp channel for Encrypt
Dim ReloadVersion As String 'true to test NewVersion download and stop current version

'ReloadVersion = "3.3.143" 'Used to test NewVersion

    On Error GoTo Main_Error
Debug.Print "==== Load ===="
#If jnasetup = True Then
    cmdJna = True
#End If
    StartupLogFile = FileSelect.SetFileName("StartupLogFile") 'Loads FileSelect then Files
    Unload Files        'V142 Do not leave loaded - we must have NmeaRcv as first form in the forms
    Unload FileSelect   'V142 Collection so QueryUnload event is triggered in NmeaRcv form
                        'V142 Otherwise you get the Program is not responding message
    
    StartupLogFileCh = FreeFile
    LogFileCh = -1  'If LogFile is used create a new file
'create a new startup log file
    Open StartupLogFile For Output As #StartupLogFileCh
    Call WriteStartUpLog("Log Opened at " & Now())
    Call WriteStartUpLog("This version is " & App.EXEName & "(" & App.Major & "." & App.Minor & "." & App.Revision & ")")
        
    Set MyNewVersion = New clsNewVersion
    DownloadURL = "arundale.com/docs/ais/"
    Call WriteStartUpLog("Checking " & DownloadURL & " for later version")
    
    Call MyNewVersion.CheckNewVersion(DownloadURL, ReloadVersion) 'True to reload
    
    If MyNewVersion.Downloaded = True Then
'V3.4.143 Must exit calling program otherwise install will fail because this exe is referenced
        Call WriteStartUpLog("Downloaded " & MyNewVersion.DownloadedURL)
'V3.4.143        Call Terminate
    End If
'V3.4.143 from here
'Must exit calling program otherwise install will fail because this exe is referenced
    If MyNewVersion.NewVersion <> MyNewVersion.ThisVersion Or ReloadVersion <> "" Then
        If MyNewVersion.NewVersion <> MyNewVersion.ThisVersion Then
            Call WriteStartUpLog("New version changed - terminating AisDecoder " & MyNewVersion.ThisVersion)
        End If
        If ReloadVersion <> "" Then
            Call WriteStartUpLog("Reloading latest version - terminating AisDecoder " & ReloadVersion)
        End If
        Call Terminate
        Exit Sub
    Else
         Call WriteStartUpLog("New version not changed " & MyNewVersion.NewVersion)
   End If
'V3.4.143 to here
    Set MyNewVersion = Nothing
    
DebugTreeFilter = False     'enables a complete display of the treefilter
    
    Load frmSplash
    frmSplash.Show
    frmSplash.Refresh
    ReDim NmeaOutBuf(MINNMEAOUTBUFSIZE)
    

    If IsDate(Now()) = True Then
        Call WriteStartUpLog("Date is valid")
    Else
        Call WriteStartUpLog("Date is invalid")
    End If
#If False Then
    On Error Resume Next
        Call WriteStartUpLog("Date=" & Date)
        Call WriteStartUpLog("DateValue=" & DateValue(Date))
        Call WriteStartUpLog("Time=" & Time)
        Call WriteStartUpLog("TimeValue=" & TimeValue(Time))
        Call WriteStartUpLog("Date=" & Date)
        Call WriteStartUpLog("DateValue=" & DateValue(Date))
        Call WriteStartUpLog("Time=" & Time)
        Call WriteStartUpLog("TimeValue=" & TimeValue(Time))
    On Error GoTo 0
#End If
    Call WriteStartUpLog("StartupLogFile=" & StartupLogFile)
    Call WriteStartUpLog("")
    Call WriteStartUpLog("Environment Settings")
    i = 1
    Do
        Call WriteStartUpLog(Environ(i))
        i = i + 1
    Loop Until Environ(i) = ""
    
'There is no envirinment variable to get the ALL USERS\Application Data location
'internationally Application data is re-named si the english must NOT be used

    Call WriteStartUpLog("GetSpecialFolder CSIDL_COMMON_APPDATA = " & GetSpecialFolderA(CSIDL_COMMON_APPDATA))
    Call WriteStartUpLog("Windows Version = " & GetVersion1())
    Call WriteStartUpLog("User Default LocaleID = " & GetUserDefaultLCID())
    DecimalSeparator = GetDecimalSeparator
    Call WriteStartUpLog("Decimal Separator = " & DecimalSeparator)

    EditOptions = CBool(IsNTAdmin(ByVal 0&, ByVal 0&))
    Call WriteStartUpLog("User has Administrator Rights = " & EditOptions)
    EditOptions = True  'from v3.3.138
    Call WriteStartUpLog("User allowed to edit options = " & EditOptions)

    Call WriteStartUpLog("***************")
    Call WriteStartUpLog("Initial Registry Settings")
    Call PrintRegistry

'to allow setup to detect running program
'to release terminate must be called
    hMutex = CreateMutex(0&, 0&, "AisDecoder")

'See is we have any command line arguments
    Call WriteStartUpLog("Command$=" & Command$)
    StartUpCommand = LCase$(Command$)
'Stop   'to test startup command
'start /nowindow /ini=C:\AMS_Decoder.ini
'    StartUpCommand = "start /nowindow /ini=""C:\Documents and Settings\jna\Application Data\Arundale\yuyama\AMS_Decoder.ini"""
    If StartUpCommand <> "" Then
        cmdOptions = Split(StartUpCommand, "/")
        For i = 0 To UBound(cmdOptions)
            j = InStr(1, cmdOptions(i), "=")
            If j = 0 Then j = Len(cmdOptions(i))
            Select Case Trim$(Left$(cmdOptions(i), j))
            Case Is = ""
            Case Is = "start"
                cmdStart = True
            Case Is = "nowindow", "nowindows"
                cmdNoWindow = True
            Case Is = "minimise", "minimize"
                cmdMinimize = True
            Case Is = "jna"
                cmdJna = True
            Case Is = "ini="
                cmdIniFileName = Mid$(cmdOptions(i), j + 1)
'remove " from file name (if any)
                cmdIniFileName = Replace(cmdIniFileName, """", "")
'We cannot check if file exists yet because if a new user
'the file may not yet have been placed in the current users directory
            End Select
        Next i
'        If InStr(1, Startupcommand, "start", vbTextCompare) Then cmdStart = True
'        If InStr(1, Startupcommand, "nowindow", vbTextCompare) Then cmdNoWindow = True
    End If

'cmdIniFileName = "GoogleEarth.ini" 'testing only
'If cmdNoWindow = False Then Files.Show

'cmdIniFileName = "test.ini"
'cmdStart = True    'test while using vbe
'cmdNoWindow = True
    Call WriteStartUpLog("InstalledVersion=" & App.Major & "." & App.Minor & "." & App.Revision)
    InstalledBuildNo = App.Major * 2 ^ 8 + App.Minor * 2 ^ 4 + App.Revision
    Call WriteStartUpLog("InstalledBuildNo=" & InstalledBuildNo)
'First ensure weve a default.ini file available for TreeFilter
'The FallBack file is the only one guaranteed to be downloaded with
'any new version of the file
'see if weve an .ini file for this machine (FallBack file)
    FallBackFile = QueryValue(HKEY_LOCAL_MACHINE, "Software\Arundale\" & App.EXEName & "\Settings", "FallBackInitialisationFile")
    Call WriteStartUpLog("FallBackFile=" & FallBackFile)
    If FileExists(FallBackFile) = False Then
        MsgBox FallBackFile & " does not exist" & vbCrLf & "Exiting AisDecoder", _
        vbOKOnly + vbCritical, "Initialisation Error"
        Call WriteStartUpLog("Does not exist, Terminating")
        Unload frmSplash        'V142
        Call Terminate
        Exit Sub
    End If

'see if weve an .ini file defined for Current user
     If IsIniFile(cmdIniFileName) Then
        HKCUFile = cmdIniFileName
    Else
        HKCUFile = QueryValue(HKEY_CURRENT_USER, "Software\Arundale\" & App.EXEName & "\Settings", "InitialisationFile")
    End If
'MsgBox HKCUFile
    Call WriteStartUpLog("HKCU Initialisation File=" & HKCUFile)
'if no register key, it must be the first time for this user
    If IsIniFile(HKCUFile) = False Then
        Call WriteStartUpLog("Valid Current User Initialisation File Not Found")
        NewUserFile = QueryValue(HKEY_LOCAL_MACHINE, "Software\Arundale\" & App.EXEName & "\Settings", "NewUserInitialisationFile")
        HKCUFile = Environ("APPDATA") & "\Arundale\Ais Decoder\Settings\" & NameFromFullPath(NewUserFile)
'v106        HKCUFile = "%APPDATA%\Arundale\Ais Decoder\Settings\" & NameFromFullPath(NewUserFile)
        Call WriteStartUpLog("Current User File reset to New User File =" & HKCUFile)
    
    Else
'Stop
    End If
    If IsIniFile(HKCUFile) = True Then
        Call WriteStartUpLog(HKCUFile & " exists in Current User")

'yes
'MsgBox "got " & HKCUFile & " in Current User"
    Else
        Call WriteStartUpLog(HKCUFile & " not found or invalid")
        Call WriteStartUpLog("No Initialsation File defined for Current User")
'this must be the first time this user has used the program
'see if we have an "All Users initialisation file defined
'if we havn't this will be picked up when we try and load it in treefilter
'which will use the FallBack file to start the program
        NewUserFile = QueryValue(HKEY_LOCAL_MACHINE, "Software\Arundale\" & App.EXEName & "\Settings", "NewUserInitialisationFile")
        Call WriteStartUpLog("HKLM NewUserInitialisationFile Found=" & NewUserFile)
'set the Current User initialisation file to same as All Users
'It may have been changed from default.ini
'%APPDATA%\Arundale\Ais Decoder\Settings\default.ini
        HKCUFile = Environ("APPDATA") & "\Arundale\Ais Decoder\Settings\" & NameFromFullPath(NewUserFile)
        Call WriteStartUpLog("HKCUFile set to New User =" & HKCUFile)
'MsgBox HKCUFile
'ensure key exists
        CreateNewKey HKEY_CURRENT_USER, "Software\Arundale\" & App.EXEName & "\Settings"
        Call WriteStartUpLog("HKCU\Settings checked/created")
        SetKeyValue HKEY_CURRENT_USER, "Software\Arundale\" & App.EXEName & "\Settings", "InitialisationFile", HKCUFile, REG_EXPAND_SZ
        Call WriteStartUpLog("HKCU\Settings\InitialisationFile checked/created =" & HKCUFile)
        SetKeyValue HKEY_CURRENT_USER, "Software\Arundale\" & App.EXEName & "\Settings", "CurrentUserPath", "%APPDATA%\Arundale\Ais Decoder", REG_EXPAND_SZ
        Call WriteStartUpLog("HKCU\Settings\CurrentUserPath checked/created =" & "%APPDATA%\Arundale\Ais Decoder")
        SetKeyValue HKEY_CURRENT_USER, "Software\Arundale\" & App.EXEName & "\Settings", "InstalledIniFileDateTime", Date & " " & Time(), REG_SZ
        Call WriteStartUpLog("HKCU\Settings\InstalledIniFileDateTime checked/created =" & Date$ & " " & Time$())
        SetKeyValue HKEY_CURRENT_USER, "Software\Arundale\" & App.EXEName & "\Settings", "LastUpdatedBuildNo", InstalledBuildNo, REG_DWORD
        Call WriteStartUpLog("HKCU\Settings\LastUpdatedBuildNo set to Installed (exe) build no =" & InstalledBuildNo)
'copy all files in All Users to Current User (This is the first time for this user)
'MsgBox "Copying all files" & vbCrLf & _
"from " & Environ("ALLUSERSPROFILE") & "\Application Data\Arundale\Ais Decoder" & vbCrLf & _
"to   " & Environ("APPDATA") & "\Arundale\Ais Decoder" & vbCrLf
        On Error Resume Next
        Call WriteStartUpLog("set folder " & Environ("APPDATA") & "\Arundale")
        Set oFolder = oFileSys.CreateFolder(Environ("APPDATA") & "\Arundale")
        Call WriteStartUpLog("set folder " & Environ("APPDATA") & "\Arundale\Ais Decoder")
        Set oFolder = oFileSys.CreateFolder(Environ("APPDATA") & "\Arundale\Ais Decoder")
'        Set oFolder = oFileSys.CreateFolder(Environ("USERPROFILE") & "\My Documents\Ais Decoder")
'MsgBox oFolder.Attributes
        On Error GoTo Main_Error
'MsgBox "copy folder"
'The Special Folder has a trailing \
    Call WriteStartUpLog("Copy folder " & GetSpecialFolderA(CSIDL_COMMON_APPDATA) & "Arundale\Ais Decoder\*" _
    & " to " & Environ("APPDATA") & "\Arundale\Ais Decoder\")
        oFileSys.CopyFolder GetSpecialFolderA(CSIDL_COMMON_APPDATA) & "Arundale\Ais Decoder\*", _
            Environ("APPDATA") & "\Arundale\Ais Decoder\", _
            True
        Set oFileSys = Nothing
        Set oFolder = Nothing
'weve still not got an ini file (not in AllUsersProfile)
'force to FallBack
        If IsIniFile(HKCUFile) = False Then
            SetKeyValue HKEY_CURRENT_USER, "Software\Arundale\" & App.EXEName & "\Settings", "InitialisationFile", HKCUFile, REG_EXPAND_SZ
'MsgBox "Copying from " & FallBackFile & vbCrLf & "to " & HKCUFile & vbCrLf
            On Error Resume Next
            Set oFolder = oFileSys.CreateFolder(PathFromFullName(HKCUFile))
            On Error GoTo Main_Error
            oFileSys.CopyFile FallBackFile, HKCUFile, True
            Set oFileSys = Nothing
        End If
    End If
'At this point we should have a valid default.ini file in %APPDATA%
    IniFileName = HKCUFile

'We now check if any reserved files have changed if there has been a new
'version installed, if so we copy them to the current user
'MsgBox "Checking Last Update"
    LastUpdatedBuildNo = QueryValue(HKEY_CURRENT_USER, "Software\Arundale\" & App.EXEName & "\Settings", "LastUpdatedBuildNo")
    Call WriteStartUpLog("LastUpdatedBuildNo=" & LastUpdatedBuildNo)

    If InstalledBuildNo <= 902 Then
'V118 has big initialisation file changes - force overwriting default files
        InstalledBuildNo = 0
        Call WriteStartUpLog("Installed build = v118 - Forcing User File Update")
    End If
    
    If LastUpdatedBuildNo <> InstalledBuildNo Then
        Call WriteStartUpLog("LastUpdatedBuildNo (" & LastUpdatedBuildNo & ") <> InstalledBuildNo (" & InstalledBuildNo & ") - Updating User Files ")
'Check if AllUsers .ini files same as CurrentUser.ini files
        Call frmFiles.UpdateUserFiles
        If frmFiles.Cancel = False Then
            SetKeyValue HKEY_CURRENT_USER, "Software\Arundale\" & App.EXEName & "\Settings", "LastUpdatedBuildNo", InstalledBuildNo, REG_DWORD
            Call WriteStartUpLog("HKCU\Settings\LastUpdatBuildNo=" & InstalledBuildNo)
        End If
    Else
        Call WriteStartUpLog("LastUpdatedBuildNo (" & LastUpdatedBuildNo & ") = InstalledBuildNo (" & InstalledBuildNo & ") - Not Updating User Files ")
    End If

'If an initialisation file has been specified on the command line
'cmdIniFileName will contain the passed parameter (without any "")
    Call WriteStartUpLog("Checking CommandIniFileName " & cmdIniFileName)
    If cmdIniFileName <> "" Then
'set up default path (if not passed)
        kb = FileSelect.SetFileName("cmdIniFileName")
        If FileExists(kb) = True Then
            IniFileName = kb
            Call WriteStartUpLog("Using cmdIniFileName - " & kb)
        End If
    End If
    
'Check if Initialisation file exists
    Do Until FileExists(IniFileName) = True
        MsgBox "Initialisation File does not exist" _
        & vbCrLf & IniFileName & vbCrLf, vbOKOnly, "File not found"
        IniFileName = FileSelect.AskFileName("IniFileName", True)
    Loop
    
    Call WriteStartUpLog("Using IniFileName " & IniFileName)

'reset HKCU
    SetKeyValue HKEY_CURRENT_USER, "Software\Arundale\" & App.EXEName & "\Settings", "InitialisationFile", IniFileName, REG_EXPAND_SZ

'To start using the licence code the installation exe needs to set the Licence
'registry key to anything other than ""
    kb = QueryValue(HKEY_CURRENT_USER, "Software\Arundale\" & App.EXEName & "\Settings", "Licence")
    If kb <> "" Then
        Call LoadLicence
    Else
        UserLicence.sMaxInputFileSize = "0"    '"0"=Unlimited (ditto)
        UserLicence.sMaxOutputFileSize = "0"   '"0"=Unlimited (ditto)
        UserLicence.ExpiryDate = CDate(#2/28/2015#)
        UserLicence.ExpiryDate = "00:00:00"
        UserLicence.ComputerName = Environ$("ComputerName")
        UserLicence.UserName = Environ$("UserName")
    End If

#If False Then
'#mm/dd/yy# ensures date is in Code Format(Always US if using VB)   p793 Programming guide
'CDate(#mm/dd/yy#) returns date in Locale format
'Date returns date in Locale format
    If Date > CDate(#2/28/2015#) Then     'always US Format
        Call LoadLicence
    Else
        UserLicence.sMaxInputFileSize = "0"    '"0"=Unlimited (ditto)
        UserLicence.sMaxOutputFileSize = "0"   '"0"=Unlimited (ditto)
        UserLicence.ExpiryDate = CDate(#2/28/2015#)
        UserLicence.ComputerName = Environ$("ComputerName")
        UserLicence.UserName = Environ$("UserName")
    End If
#End If

'x64 problem after here
    
    Call WriteStartUpLog("Calling nowutc()")
    DecoderStartTime = NowUtc()    'used to determine if vessl name is cached
    Call WriteStartUpLog("DecoderStartTime=" & DecoderStartTime)

    'Call DisplayForms("Before Loading NmeaRcv")
    'Dont think we need to load NmeaRcv as it will be loaded when TreeFilter is loaded
    'Load NmeaRcv        'do first as TreeFilter will set up .ini variables
    'Call DisplayForms("After Loading NmeaRcv")
    'Why do we need to load ReceivedData ?
    'Load ReceivedData
    'I dont think we need to load Detail as it will be loaded when required ?
    'Load Detail
    'Call DisplayForms("After Loading Detail")

'must be done before TreeFilter loads TemplateFile if testing
'LicenseStatus = "Registered"

'first time IniFileName is set up
    IniFileName = QueryValue(HKEY_CURRENT_USER, "Software\Arundale\" & App.EXEName & "\Settings", "InitialisationFile")
'Copy initialisation file to startup log (so we can find where it's crashed
'if it crashes when loading treefilter
'final check to ensure it exists
    IniFileName = FileSelect.SetFileName("IniFileName")
    Unload Files        'V142 Program not responding issue
    Unload FileSelect   'V142 Program not responding issue
    Call FileToLog(IniFileName)

'Check if Registry allows unpriv user to edit
    If EditOptions = False Then
        If QueryValue(HKEY_CURRENT_USER, "Software\Arundale\" & App.EXEName & "\Settings", "EditOptions") = "Allowed" Then
        Call WriteStartUpLog("EditOptions Allowed (Overridden by Registry)")
            EditOptions = True
        End If
    End If

'Check if Registry requires Dac=0 processed as TestDac
    If TestDac = False Then
        If QueryValue(HKEY_CURRENT_USER, "Software\Arundale\" & App.EXEName & "\Settings", "TestDac") = "True" Then
            Call WriteStartUpLog("TestDac (Overridden by Registry)")
            kb = QueryValue(HKEY_CURRENT_USER, "Software\Arundale\" & App.EXEName & "\Settings", "DacMap")
            If kb <> "" Then    'Not nul
                arry = Split(kb, "-")
                If UBound(arry) = 1 Then
                    If IsNumeric(arry(0)) Then DacMap(0) = arry(0)
                    If IsNumeric(arry(1)) Then DacMap(1) = arry(1)
'No point in mapping
                    If DacMap(0) <> DacMap(1) Then
                        TestDac = True
                    End If
                End If
            End If
        End If
    End If

'Check if Registry requires Hex Dump
    If HexDump = False Then
        If QueryValue(HKEY_CURRENT_USER, "Software\Arundale\" & App.EXEName & "\Settings", "HexDump") = "True" Then
        Call WriteStartUpLog("HexDump (Overridden by Registry)")
            HexDump = True
        End If
    End If

'Check if TCP Login File exists
    TcpLoginFileName = FileSelect.SetFileName("TcpLoginFileName")
    If FileExists(TcpLoginFileName) Then
        Call WriteStartUpLog("TCP Login Command File is " & TcpLoginFileName)
        ch = FreeFile
        Open TcpLoginFileName For Input As #ch
        Do Until EOF(ch)
            Line Input #ch, kb
            Call WriteStartUpLog(vbTab & kb)
        Loop
        Close #ch
    Else
        Call WriteStartUpLog("No TCP Login Command File")
    End If

'TreeFilter must be loaded and .ini file read to see if
'we need to check for updates
Call WriteStartUpLog("Before Loading TreeFilter in .main")

'    If ExitSub Then Exit Sub
    Load NmeaRcv     'V142 Must be loaded before TreeFilter (Program is not responding problem)
    Load TreeFilter     'then get .ini file and set up variables Loads NmeaRcv then Files,FileSelect et al
    Call WriteStartUpLog("After Loading TreeFilter in .main")

'x64 error before here
'Check if weve a new version available on the web
    If TreeFilter.Check1(2).Value <> 0 Then Call CheckUpdates
'Start Timer after CheckUpdates otherwise the Timer will force a check of updates
    NmeaRcv.TimeZoneTimer.Enabled = True

'V142     QueryQuit = True    'if we get here we need to check if we actually want to quit
                    'if we try and unload NmeaRcv form
ReDim Multipart(9, 9)   'Created dynamic so we can clear it

    Call DecodeDefs.Initialise  'Load variables
'IconScale = "1"
'IconColor = "ffffffff"
    Gis.IScale = "1"
    Gis.Color = "ffffffff"

#If jnasetup Then
'    CacheTrappedMsgs = True
#End If
    If CacheTrappedMsgs Then Call ReadTrappedMsgs

'Call DisplayForms("Before Hide in .main")

'Ensure input/output display forms are hidden before start
    ReceivedData.Hide
    Output.Hide
    List.Hide
    Detail.Hide
'Stop
'cmdNoWindow = True 'debug Command line args    'debug
'cmdStart = True    'debug
    If cmdNoWindow = False Then
        NmeaRcv.Show
        If cmdMinimize Then
            NmeaRcv.WindowState = vbMinimized
        End If
    End If
If cmdJna Then
'    frmFTP.Show            'Debug FTP
'    TestFileSelect.Show    'Debug FileSelect
End If
    Call WriteStartUpLog("Final Registry Settings")
    Call PrintRegistry
    If cmdNoWindow = True Then
        Call FormsLog("StartUp", "Force nowindow at End of StartUp", False)
    Else
        Call FormsLog("StartUp", "End of StartUp")
    End If
    Call WriteStartUpLog("Initialisation finished")
    Call CloseStartupLogFile

'This MUST be done at the end because if it is a file input, any code after here will
'only be actioned after the file input has terminated
    If cmdStart = True Then
'If cmdStart = True then don't prompt for FileName if FileInput
        NmeaRcv.cbStart.Value = True 'trigger click event
    End If

    Unload frmSplash
    Call CheckExpiry
'MsgBox IniFileName & vbCrLf & TreeFilter.Text1(9).Text, , TreeFilter.Caption
Main_Exit:
'Call DisplayForms
Call IsAnyFileOpen(True)         'V142
'Call LogDebugCls        'V143  DO NOT HERE because there will be a class for each vessel in vessels
                                'and it will take for ever to log these
Call LogForms("Open forms at end of MainControl")           'V142 14/11/16
Call CloseStartupLogFile    'V142 14/11/16

Exit Sub

Main_Error:
    Select Case err.Number
    Case Else
        MsgBox "Error " & err.Number & " - " & err.Description, vbCritical, "Main_Error"
    End Select
    Resume Next 'Only in sub Main (otherwise application does not terminate)
'    Resume Main_Exit
    
End Sub

Public Sub CheckExpiry()
Dim ExpiryDays As Long
Dim kb As String

        If UserLicence.ExpiryDate <> "00:00:00" Then
            ExpiryDays = DateDiff("d", Date, UserLicence.ExpiryDate)
            If ExpiryDays <= 10 Then
                kb = "Your enhanced version of AisDecoder "
                Select Case ExpiryDays
                Case Is > 1
                    kb = kb & "expires in " & ExpiryDays & " days" & vbCrLf
                Case Is = 1
                    kb = kb & "expires tomorrow" & vbCrLf
                Case Is = 0
                    kb = kb & "expires today" & vbCrLf
                Case Is = -1
                    kb = kb & "expired yesterday" & vbCrLf
                Case Is < -1
                    kb = kb & "expired " & -ExpiryDays & " days ago" & vbCrLf
                End Select
            kb = kb & "For more details contact myself - neal@arundale.com"
            MsgBox kb, vbInformation, "Enhanced Version Expiry"
            End If
        End If

End Sub
'only called by ProcessNmeaBuf
'Separate calls for each part of multi-part sentences
Public Sub ProcessSentence(InputSentence As String)

'walk the input tree calling modules as required
Dim Level As Long   'must be here to retain value in recursive call
'level is not actually required but helps debugging recursion
Dim Pass As Boolean
Dim ScheduledDue As Boolean
Dim TagsFound As Boolean    'we have tags that require output
Dim Channel As Long
Dim InError As Boolean

'creates clsSentence, PayloadBytes() and Ship(names)

'Must have an error trap here, otherwise an unexpected error (in ProcessSentence) will go back to the
'DataArrival error trap. This causes ProcessSuspended to never get released
'All sentences then get held in the Waiting buffer
    On Error GoTo ProcessSentence_err
    Call DecodeSentence(InputSentence)
    If TreeFilter.Option1(2).Value <> 0 Then Call NmeaRcv.WriteInputLogFile(InputSentence)

'unfiltered list and detail
    If NmeaRcv.Option1(1).Value = True Then
'when displaying as received
        MaxNmeaDecodeListCount = MAXNMEADECODELISTCOUNT_Rcv
        Call List.AddToNmeaSummaryList  'This will be the new list
    End If
'output unfiltered detail
    If NmeaRcv.Option1(5).Value = True Then Call Detail.SentenceAndPayloadDetail(0)       'clsSentence, PayloadBytes required

'check if this sentence is required
'If we have part1 we can reject, if not we must accept
'the sentence here, and check the decode later when
'we have all parts (if required)
    Pass = ProcessSentenceFilter
    If Pass = True Then
        Pass = IsFilterOK(NmeaRcv.TreeView1.Nodes("InputFilter"))
    End If
    
    If Pass = True Then       'true if passed Partfilter
        Filtered = Filtered + MESSAGE_COUNT_STEP
'use to debug overflow        incr Filtered
#If jnasetup = True Then

'If clsSentence.AisMsgType = "1" Then    'And Len(clsSentence.AisPayload) * 6 <> 168 Then
'    Stop
'        If NmeaWords(1) <> NmeaWords(2) Then
'            If NmeaWords(6) <> "0" Then     'Single part only
'                MsgBox "NMEA fill bits should be 0"
'                InError = True
'            End If
'        Else
'            If NullToZero(NmeaWords(6)) <> ChrnoToFillBits(Len(clsSentence.AisPayload)) Then
'                InError = True
'            End If
'        End If
    
'        If clsSentence.AisMsgPartsComplete = True Then
'            If clsSentence.AisPayloadBits <> 424 Then
'                InError = True
'                clsSentence.PayloadReassemblerComments = clsSentence.AisPayloadBits
'            End If
'        End If
'    If Pass = False Then
'        NmeaRcv.Option1(2).Value = True
'        Call List.AddToNmeaSummaryList  'This will be the new list
'        NmeaRcv.Option1(2).Value = False
'        NmeaRcv.Option1(6).Value = True
'        Detail.SentenceAndPayloadDetail (0)
'        NmeaRcv.Option1(6).Value = False
'    End If
'End If
    
'We are only passing the sentences in error
'        Pass = InError
'        If InError = True Then
'            Outputted = Outputted + 1
'        End If
#End If
    Else    'Sentence failed filter
'test        Call WriteErrorLog(InputSentence)
    End If
'Generate error here & find out why precessing does not re-start
'Debug.Print 1 / 0
'AisSentence can be true abd if part2 only MsgType ="" (if sentence not complete)
If clsSentence.IsAisSentence = True And clsSentence.AisMsgPartsComplete = True Then
    If Pass = False And CLng(clsSentence.AisMsgType) <= 3 And clsSentence.AisPayloadFillBits > 0 Then
Debug.Print clsSentence.AisMsgType & ", fill=" & clsSentence.AisPayloadFillBits
'Stop
    Pass = True
    End If
End If
    
    If Pass = True Then       'true if first part

'output filtered to log, if required
        If TreeFilter.Option1(3).Value = True Then Call NmeaRcv.WriteInputLogFile(InputSentence)

'Output each message summary (List) if required
'if list required when scheduled buffer is output dont do here as well
        If NmeaRcv.Option1(2).Value = True Then
'when displaying as received
            MaxNmeaDecodeListCount = MAXNMEADECODELISTCOUNT_Rcv
            Call List.AddToNmeaSummaryList  'This will be the new list
        End If

'Display Detail Filtered if required (6) - (7) cannot also be true
        If NmeaRcv.Option1(6).Value = True Then Detail.SentenceAndPayloadDetail (0)      'clsSentence, PayloadBytes required
'if scheduled put all parts of message into scheduled buffer
'return if scheduled due to be output
'the scheduler will output the current message because it outputs
'based on the time the message is received
'note tags will not be applied if not scheduled and not ranged
'this could be changed here by also calling TagsInput
'if CsvOutput(0) or taggedoutput as well as RangeReq(0)
If cmdJna Then
    Call Testing.Variables
End If
'SchedulerInput puts sentence into Schedule Buffer (AisMsgs)
'If AllPartsComplete are true (on this message)
'returns True if Output is required.
        If ScheduledReq(0) Then
            ScheduledDue = SchedulerInput
        End If
'If SchedulerOutput is not called at the end of this routine
'(because the FTPUpload from the previous call has not completed)
'then next time will have been incremented by the schedule period
'in the Scheduler Input routine. This means the Schedulet does not keep
'cycling playing catchup

        If CacheTrappedMsgs Then Call TrappedMsgInput

'tags may be required even if not scheduled eg csv
'must be complete (even if last) and not $PGHP message
        If clsSentence.AisMsgPartsComplete _
        Or clsSentence.AisMsgType = "" Then
'Decode all fields if required (complete senetences only)
'        If CsvAll(0) = True Then Detail.SentenceAndPayloadDetail (4)
'Call Testing.Variables
'Testing.Visible = True
            For Channel = 1 To 2
'if scheduled we dont need the tags until scheduled is output
'This is where the decoded fields that are required are extracted
'from the complete sentence
                If TagsReq(Channel) = True And ScheduledReq(Channel) = False Then
                    TagsFound = TagsFromFields   'found a tag
                    Exit For    'only do once as output channels does all channels
                End If
            Next Channel
'Debug.Print InputSentence
'Debug.Print "Process " & Filtered & " " & Outputted
'output now for any channel not scheduled unless it is an encapulating sentence
            If clsSentence.EncapsulatingNmeaSentence = False Then
                Call OutputAll(False)   'False = on any non-scheduled channel
            End If

'Set up the new encapsulating sentence as a Group sentence
            If clsSentence.EncapsulatingNmeaSentence = True Then
                GroupSentence.NmeaSentence = clsSentence.NmeaSentence
            Else
'Clear the Group sentence when the first non encapsulating sentence is received
'This needs changing when the encapsulating sentence encompasses more than one
'following Ais sentence AND in
'V 144 Jan 2017 unless $AITAG(Jason time stamp)has been found make CSV output compatible with V129
                If Left$(GroupSentence.NmeaSentence, 6) <> "$AITAG" Then
                    Set GroupSentence = Nothing
                End If
            End If
        End If
'Debug.Print "OutputAll " & Filtered & " " & Outputted

'output tagged
'    If TagsFound = True Then Call OutputTagsIfRangeOk 'OutputCsv and Tagged (all channels)
'not scheduled output with Range filtering, output on mmsi change
'End If  'Filtered OK
    
'Outputing from the Scheduled buffer must be after all processing of the current message
'because current message details (clsSentence, PayloadBytes etc) will be overwritten
'MUST have completed previous upload
        If ScheduledDue = True And FtpUploadExecuting = False Then
 'debug    If AisMsgs.Count = 1 Then Stop
 '           If DecodingState And 1 Then Call SchedulerOutput("ProcessSentence") 'call AddToNmeaSummaryList, then SentenceAndPayloadDetail (if reqd)
        End If
    Else    'Does not pass input filter
'If called by ScheduledOutput all sentences will already have passed the input filter
'If not scheduled NmeaOutBuf will only contain the current sentence
        Call NmeaOutBufClear
'If block of sentences with same MMSI are not output (by OutputAll) the Tag values are
'not cleared if the sentences has been range filtered out
        Call ClearTagValues
    End If  'Filtered OK

    Processed = Processed + MESSAGE_COUNT_STEP
'use to debug overflow    incr Processed
'v131    ProcessingSentence = False     'Allows UDP input to continue
 'Debug.Print ProcessingSentence
Exit Sub
ProcessSentence_err:
    Select Case err.Number
    Case Is = 457
'Jason error 457 (31/8/2017) Key is already associted with an element
'Trap inserted V147 8/9/17
        Exit Sub    'dont process this sentence
    Case Else
'in error
        NmeaRcv.StatusBar.Panels(1).Text = "ProcessSentence, Error " & err.Number & " - " & clsSentence.NmeaSentence ' err.Description
Debug.Print err.Description
        NmeaRcv.ClearStatusBarTimer.Enabled = True
'        Call WriteErrorLog(StatusBar.Panels(1).Text & vbCrLf & clsSentence.NmeaSentence)
    End Select
    err.Clear
    Resume Next    'Quit routine
End Sub

'Checks clsSentence for shore,latlon, or gisok
'requires clssentence to be set for filternames
'returns false if test not passed
Function ChkSentenceFilter(FilterName As String, FilterKey As String) As Boolean
Dim Range As Single

    With clsSentence
        Select Case FilterName
    
        Case Is = "shore"   'Only accept of MMSI not shore station
            ChkSentenceFilter = False     'assume failure
            On Error GoTo Fail    'may be nul or string (fails)
            If CLng(clsSentence.AisMsgFromMmsi) >= 100000000 Then
                ChkSentenceFilter = True  'not a shore station
            End If

        Case Is = "latlonrange"  'Reject only position messages out of lat lon range
'lat/lon is always dot, force lat/lon & range to locale
'because Csng is locale aware
            Select Case .AisMsgType
            Case Is = 1, 2, 3, 4, 9, 11, 17, 18, 19, 21, 27
                ChkSentenceFilter = True  'assume passed
                If ChkVesselLatMin = True And .VesselLat < VesselLatMin Then ChkSentenceFilter = False
                If ChkVesselLonMin = True And .VesselLon < VesselLonMin Then ChkSentenceFilter = False
                If ChkVesselLatMax = True And .VesselLat > VesselLatMax Then ChkSentenceFilter = False
                If ChkVesselLonMax = True And .VesselLon > VesselLonMax Then ChkSentenceFilter = False
            Case Else
                ChkSentenceFilter = True  'passed
            End Select
#If False Then
'If filtering by Lat.Lon still let the vessel name msgs through
            If .AisMsgType = 5 Then
                ChkSentenceFilter = True
            ElseIf .AisMsgType = 24 And .AisMsgFiId = "A" Then
                ChkSentenceFilter = True
            Else
'Assume the range check passes, if any lat/lon requires checking and it does not pass then
'the sentence fails
                ChkSentenceFilter = True
                If ChkVesselLatMin = True And .VesselLat < VesselLatMin Then ChkSentenceFilter = False
                If ChkVesselLonMin = True And .VesselLon < VesselLonMin Then ChkSentenceFilter = False
                If ChkVesselLatMax = True And .VesselLat > VesselLatMax Then ChkSentenceFilter = False
                If ChkVesselLonMax = True And .VesselLon > VesselLonMax Then ChkSentenceFilter = False
            End If
#End If

'If .AisMsgType = 5 Or .AisMsgType = 24 And .AisMsgFiId = "A" Then
'    ChkSentenceFilter = True
'    If ChkVesselLatMin = True And CachedVessel.Lat < VesselLatMin Then ChkSentenceFilter = False
'    If ChkVesselLonMin = True And .VesselLon < VesselLonMin Then ChkSentenceFilter = False
'    If ChkVesselLatMax = True And .VesselLat > VesselLatMax Then ChkSentenceFilter = False
'    If ChkVesselLonMax = True And .VesselLon > VesselLonMax Then ChkSentenceFilter = False
'End If
        
        Case Is = "gis" 'accept position messages in range + static data messages
                    'reject everything else
            Select Case .AisMsgType
            Case Is = 1, 2, 3, 4, 9, 11, 17, 18, 19, 21, 27
                If clsSentence.AisPositionOK = True Then    'false if latlon 91 or 181
                    ChkSentenceFilter = True  'assume passed
                    If ChkVesselLatMin = True And .VesselLat < VesselLatMin Then ChkSentenceFilter = False
                    If ChkVesselLonMin = True And .VesselLon < VesselLonMin Then ChkSentenceFilter = False
                    If ChkVesselLatMax = True And .VesselLat > VesselLatMax Then ChkSentenceFilter = False
                    If ChkVesselLonMax = True And .VesselLon > VesselLonMax Then ChkSentenceFilter = False
                End If
            Case Is = 5, 21, 24
                ChkSentenceFilter = True 'passed
            End Select
        End Select
    End With
Exit Function

Fail:
    err.Clear
    On Error GoTo 0
End Function

Function LLtoSng(LatLon As String) As Single
    If LatLon <> "" Then
    LLtoSng = LatLon
    End If
End Function

'FilterName=tNode.Tag, FilterKey=tNode.Key
Function ChkInputFilter(FilterName As String, FilterKey As String) As Boolean
'requires clssentence to be set for filternames
'returns false if test not passed
'returns true on first and subsequent parts of multipart messages
'retain data between calls so we dont keep calling same routines
Dim FilterKeyArg As String
Dim arry() As String

'FilterKey is field 3, FilterName is Field 5
'The FilterKeyArgument is extracted from the FilterKey (which is unique)
'There are 2 formats
'1  AisMsg6Dac-1,4,AisMsg6Dac1Fi-0,Function 0,AisMsgDacFi,0,
'1  AIS,4,AisMsg1,AIS Message 1,AisMsg,0,   (filterkey=AisMsg1,FilterName=AisMsg,filterkeyarg=1, msg 1)
'2   Argument for AisMsgNo is added to the key delimited with a -
'   AisMsg24,4,AisMsg24Id-A,Type A,AisMsgFiID,0,    (arg=A)
'or NMEA$,4,NMEA$-$GPZDA,$GPZDA Date & Time,NMEA$,0, (arg=$GPZDA)
'Stop
    arry = Split(FilterKey, "-")
    If UBound(arry) = 0 Then    'Format 1
        FilterKeyArg = Replace(FilterKey, FilterName, "")   '=1
    Else
        FilterKeyArg = arry(1)  'Format 2
    End If

If clsSentence.IecFormat = "VDO" Then
'    Stop
End If

ChkInputFilter = False     'assume failure
'Decide which filter to use
Select Case FilterName  'this is node.tag (5th Field on the .ini file
Case Is = "InputFilter"     'root node must exist
    If clsSentence.NmeaSentence <> "" Then      'reject blank sentence is dumb
        ChkInputFilter = True
    End If
    If clsCb.Block <> "" Then   'Accept if only comment block
        ChkInputFilter = True
    End If
Case Is = "CRCerror"
    If clsSentence.CRCerrmsg <> "" Then ChkInputFilter = True
Case Is = "NMEA"                'Legacy before V131 24Nov14
    If clsSentence.CRCerrmsg = "" Then
        ChkInputFilter = True
    End If
Case Is = "AIS"
'allows CRC errors through
    If clsSentence.IsAisSentence = True Then
        If FilterKeyArg = Left$(clsSentence.NmeaSentenceType, Len(FilterKeyArg)) Then
            ChkInputFilter = True
        End If
    End If
Case Is = "NMEA!"  'Any NMEA sentence starting with a ! (V131 - 24Nov14 on)
'NMEA!-!=Pass all !xxxxx
'NMEA!-!AI=Pass all !AIxxx
'NMEA!-!AIVDO=Pass !AIVDO
    If FilterKeyArg = "" Then FilterKeyArg = "!"    'NMEA!=Pass all !xxxxx  'v149
    If FilterKeyArg = Left$(clsSentence.NmeaSentenceType, Len(FilterKeyArg)) Then
        ChkInputFilter = True
    End If
Case Is = "NMEA$"  'Any NMEA sentence starting with a $ (V131 - 24Nov14 on)
'NMEA$-$=Pass all $xxxxx
'NMEA$-$GP=Pass all $GPxxx
'NMEA$-$GPGSV=Pass $GPGSV
    If FilterKeyArg = "" Then FilterKeyArg = "$"    'NMEA$=Pass all $xxxxx  'v149
    If FilterKeyArg = Left$(clsSentence.NmeaSentenceType, Len(FilterKeyArg)) Then
        ChkInputFilter = True
    End If
Case Is = "IecTalker"
    If FilterKeyArg = clsSentence.IecTalkerID Then
        ChkInputFilter = True
    End If
Case Is = "IEC" 'IecFormat
'Stop
    If FilterKeyArg = "" Then ChkInputFilter = True
    If FilterKeyArg = clsSentence.IecFormat Then
        ChkInputFilter = True
    End If
Case Is = "$"  'Any NMEA sentence starting with a $ - legacy
    If Left$(FilterKey, Len(FilterKey)) _
    = Left$(clsSentence.NmeaSentenceType, Len(FilterKey)) Then
        ChkInputFilter = True
    End If
Case Is = "AisMsg", "VdmMsg""VdoMsg", "TharMsg" 'Filter Name (.tag on the node)
'This filters by ais msgtype
'Below shows how to set-up .ini file, if __VDM requires filtering
'IEC,4,VDM,__VDM AIS VHF data-link message,IEC,0,
'VDM,4,VdmMsg1,AIS Message 1,VdmMsg,0,
'repeat using all sub tree of AIS in .ini file
    If clsSentence.AisMsgType = FilterKeyArg Then ChkInputFilter = True
Case Is = "AisMsgAll"
    ChkInputFilter = True
Case Is = "AisMsg>27"
    On Error Resume Next    'may not be decimal
    If CInt(clsSentence.AisMsgType) > 27 _
    Then ChkInputFilter = True
    On Error GoTo 0
Case Is = "AisMsgFromMmsi"
'Remove FilterName from the Key - With a list the key is set as the filter name + MMSI
    If Replace(FilterKey, FilterName, "") = clsSentence.AisMsgFromMmsi Then
        ChkInputFilter = True
    End If
'    If Mid$(FilterKey, 15, Len(FilterKey) - 14) _
'    = clsSentence.AisMsgFromMmsi _
'    Then ChkInputFilter = True
Case Is = "AisMsgToMmsi"
    If Mid$(FilterKey, 13, Len(FilterKey) - 12) _
    = clsSentence.AisMsgToMmsi _
    Then ChkInputFilter = True
Case Is = "DacList"
    If Replace(FilterKey, FilterName, "") = clsSentence.AisMsgDac Then
        ChkInputFilter = True
    End If
Case Is = "AisMsgDac"
'If FilterName = "DacList" Then Stop
    FilterKeyArg = Mid$(FilterKey, InStrRev(FilterKey, "-") + 1, Len(FilterKey))
'If FilterKeyArg = "All" Then Stop
    If FilterKeyArg = "All" _
    Or FilterKeyArg = clsSentence.AisMsgDac _
    Then ChkInputFilter = True
Case Is = "FiList"
    If Replace(FilterKey, FilterName, "") = clsSentence.AisMsgFi Then
        ChkInputFilter = True
    End If
Case Is = "AisMsgDacFi"
    If Mid$(FilterKey, InStrRev(FilterKey, "-") + 1, Len(FilterKey) - 3) _
    = clsSentence.AisMsgFi _
    Then ChkInputFilter = True
Case Is = "AisMsgDacFiID", "ListFiId"   ''only called if FiId are created using list
    If Mid$(FilterKey, InStrRev(FilterKey, "-") + 1, Len(FilterKey) - 3) _
    = clsSentence.AisMsgFiId _
    Then ChkInputFilter = True
Case Is = "AisMsgFiID"      'message ID (type 24 A,B)
    If Mid$(FilterKey, InStrRev(FilterKey, "-") + 1, Len(FilterKey) - 3) _
    = clsSentence.AisMsgFiId _
    Then ChkInputFilter = True
Case Is = "DetailOut"
'this is where individual fields may be processed
'    MsgBox FilterKey
'if enumInputTree is called in ProcessSentence routine
Case Else
    MsgBox "Filter " & FilterName & " not found", vbCritical, "ChkInputFilter"
End Select
End Function

'ONLY called by DecodeSentence if comment block exists
'uses clsCb.block which will include
'\ at start and closing \ (if it exists)
'if ! or $ of nothing is used to close Comment block there will be
'no closing separator
'From clsCB.Block
'Creates CbWords(), clsCB.Crc and clsCB.errmsg
Public Function DecodeCommentBlock()  'Input is clscb.Block
Dim CbCs As Long   '=1 if Closing Separator
Dim CbWordNo As Long
Dim i As Long
Dim param As String
Dim ParameterCode As String
Dim Line As String  'Group line no
Dim Lines As String 'No of lines in group
Dim arry() As String

    If clsCb.Block = "" Then   'should not happen
        Exit Function
    End If
        
    If Right$(Mid$(clsCb.Block, 2), 1) = "\" Then
        CbCs = 1
    End If
'Exclude Closing separator from CRC check
    clsCb.errmsg = _
    NmeaCrcChk(Left$(clsCb.Block, Len(clsCb.Block) - CbCs))
    If CbCs = 0 Then
        If clsCb.errmsg <> "" Then
                clsCb.errmsg = clsCb.errmsg & ", "
        End If
        clsCb.errmsg = clsCb.errmsg & "No Closing separator"
    End If
    If Len(clsCb.Block) = CbCs + 1 Then
        If clsCb.errmsg <> "" Then
                clsCb.errmsg = clsCb.errmsg & ", "
        End If
        clsCb.errmsg = clsCb.errmsg & "Null length Comment"
    End If
    
'Dont decode comment block if bad CRC on comment block
    If clsCb.errmsg <> "" Then Exit Function
    CbWords = Split(clsCb.Block, ",")
'must have some content
    CbWords(0) = Mid$(CbWords(0), 2)    'remove starting \
'remove CRC and \ from last word
    clsCb.CbCrc = SplitCrc(CbWords(UBound(CbWords)))

    For CbWordNo = 0 To UBound(CbWords)
        i = InStr(1, CbWords(CbWordNo), ":")
        If i > 1 Then ParameterCode = Left$(CbWords(CbWordNo), i - 1)
        param = Right$(CbWords(CbWordNo), Len(CbWords(CbWordNo)) - i)
        Select Case ParameterCode
        Case Is = "c"
            clsCb.Time = param
        Case Is = "d"
            clsCb.Destination = param
        Case Is = "i"
            clsCb.Text = param
        Case Is = "s"
            clsCb.Source = param
        Case Is = "x"
            clsCb.Counter = param
        Case Is = "g"   'ExactEarth
'or g:g:1-2-0061
            arry = Split(param, "-")
            If UBound(arry) <= 2 Then
                clsCb.GroupLine = arry(0)
                clsCb.GroupLines = arry(1)
                clsCb.GroupId = arry(2)
            Else
'Invalid ParameterCode
                Call UnknownCb(CbWords(CbWordNo))
            End If
        Case Is = "G"
'Break up the group
            If IsGroup(ParameterCode, Line, Lines) Then
                clsCb.GroupId = param
                clsCb.GroupLine = Line
                clsCb.GroupLines = Lines
            Else
                Call UnknownCb(CbWords(CbWordNo))
            End If
        Case Else
'Invalid ParameterCode
            Call UnknownCb(CbWords(CbWordNo))
        End Select
    Next CbWordNo

End Function

'ParameterCode = xGy where x & y are integers
Public Function IsGroup(ByVal ParameterCode As String, Line As String, Lines As String) As Boolean
Dim i As Long
    If Left$(ParameterCode, 1) = "g" Then
'Try for ExactEarth standard
    
    Else
'Try for IEC standard
        i = InStr(1, ParameterCode, "G")
'Check at least 1 char before and after G
        If i >= 2 And Len(ParameterCode) >= i + 1 Then
'Check characters before and after G are numeric
            If IsNumeric(Left$(ParameterCode, i - 1)) = True _
            And IsNumeric(Mid$(ParameterCode, i + 1)) = True Then
                Line = Left$(ParameterCode, i - 1)
                Lines = Mid$(ParameterCode, i + 1)
                IsGroup = True
            End If
        End If
    End If
        

End Function

Private Function UnknownCb(Word As String)
'Invalid ParameterCode
    If clsCb.Unknown <> "" Then
        clsCb.Unknown = clsCb.Unknown & "|"
    End If
    clsCb.Unknown = clsCb.Unknown & Word
End Function
'Creates clsSentence
'Splits FullSentence into CommentBlock and NmeaSentence
'Calls DecodeCommentBlock if it exists
'creates clsSentence,NmeaWords, PayloadBytes() and looksup Vessels.dat
'calls PayloadReassembler to set PayloadBytes if Ais message
'NmeaOut should only return values for the NMEA part of the detail display
'Should be in .main and called by decode sentence
'Called by ProcessSentence,SchedulerOutput and Click on NmeaDecodedList or NmeaRcvList
'these decode the details of the current sentence using SentenceAndPayloadDetail.
'PayloadReassembler is called and loads all bits into the PayloadBytes array
'Details held are required for both the input filter and
'the list view and are held in clsInputSentence
''Static NmeaWords() As String
'PayloadReassembler will not write into class module variable
'Creates clsCb.block and .NmeaSentece
'The 2 parts are then processed by DecodeCommentBlock and SentenceAndPayloadDetail
'If only the List (& no other output) is required, Sentence and payload detail is
'never called
Public Function DecodeSentence(InputSentence As String)
Dim DacFrom As Long    'to start position of bits
'Dim Payload8Bits As Long
Dim CBFrom As Long  'Start of comment block
Dim CBTo As Long   'end of Comment block = 0 (excl delimeter) if not found
Dim NmeaFrom As Long   'Start of NMEA (incl $ or !)part of sentence = 0 if not found
Dim i As Long
Dim kb As String
Dim ThisShip As ShipDef

'Dim ChkDate As Date

'    If DecodingState <> 0 Then
'        Exit Function
'    End If
    If InputSentence = "" Then
        Exit Function
    End If
'Same as last sentence decoded
    If clsSentence.FullSentence = InputSentence Then Exit Function
'    DecodingState = DecodingState Or 64
    Call ClearInputSentence

With clsSentence
    .FullSentence = InputSentence
'Split the FullSentence into Comments and NMEA
'find start of nmea (IEC spec - ! or $ is always start of NMEA
    NmeaFrom = InStr(1, .FullSentence, "!")
    If NmeaFrom = 0 Then
        NmeaFrom = InStr(1, .FullSentence, "$")
    End If
    If NmeaFrom <> 0 Then   'start of NMEA found (! or $ always terminates CB)
'Extract NMEA sentence
        .NmeaSentence = Mid$(.FullSentence, NmeaFrom)
    End If
    
'Comments - sentence must start with \
    CBFrom = InStr(1, .FullSentence, "\")
    If CBFrom > 0 Then
'find closing \
        CBTo = InStr(CBFrom + 1, .FullSentence, "\")
        If CBTo > CBFrom Then   'no closing \ separator
'CBTo must exist so minimum must be at least 1
            clsCb.Block = Mid$(.FullSentence, CBFrom, CBTo - CBFrom + 1)
            clsCb.Block = ConvEscChrs(clsCb.Block)
            Call DecodeCommentBlock
'clsCB now contains the details of the comment block
        Else
'No comment block
            CBFrom = 0
            CBTo = 0
        End If
    End If
'Have we a NmeaPrefix (not a comment with only one \ separator)
'results in a prefix
    If NmeaFrom > CBTo + 1 Then
        .NmeaPrefix = Mid$(.FullSentence, CBTo + 1, NmeaFrom - 1)
    End If

'check we don't just have a comment block or no NMEA ! or $ introducer
    If .NmeaSentence = "" Then
        Exit Function
    End If
    
    If Left$(.NmeaSentence, 6) = "$AITAG" Then
        .NmeaSentence = Replace(.NmeaSentence, " ", ",")
    End If
'Note NmeaCRC check excludes first character ! or $
    .CRCerrmsg = NmeaCrcChk(.NmeaSentence)
    
'created into Public Variable becuase you cant have a dynamic array in a class module
    NmeaWords = Split(ConvEscChrs(.NmeaSentence), ",")  '350k/min
'Left$ to remove xxxxx from $GPTAG xxxxx
    .NmeaSentenceType = NmeaWords(0)
    If Mid$(.NmeaSentenceType, 2, 1) = "P" Then 'Proprietary
        .IecTalkerID = (Mid$(.NmeaSentenceType, 2, 1))
        .IecFormat = (Mid$(.NmeaSentenceType, 3))
    Else
        .IecTalkerID = (Mid$(.NmeaSentenceType, 2, 2))
        .IecFormat = (Mid$(.NmeaSentenceType, 4))
    End If

'By processing THAR as a CRC error you need to tick CRC errors on the input filter to
'output !PTHAR sentences. Because they are flagged IsAisSentence=true the sentence details
'are decoded and output
    If .IecFormat = "THAR" Then
        .CRCerrmsg = "Data Link HDLC CRC Error"
    End If

'Set if this NMEA sentence encapsulates a following AIS sentence
'IE Contains a time stamp for the following sentence
    Select Case .NmeaSentenceType
    Case Is = "$PGHP", "$AITAG"
        .EncapsulatingNmeaSentence = True
    End Select

'Get the last word of the NmeaSentence - before any timestamps etc
'This makes it easier to split the "proper" data from the added comments etc
    For i = 0 To UBound(NmeaWords)
        If InStr(1, NmeaWords(i), "*") > 0 Then
            .NmeaCrcWord = i
        End If
    Next i
'if CRC error all details are suspect (but display first word)
    .NmeaCrc = SplitCrc(NmeaWords(.NmeaCrcWord)) 'remove *crc
'    clsSentence.NmeaSentenceType = NmeaWords(0)
'Try and sort out the received time
'Assume comment bloc is correct first
    NmeaOutBufCount = NmeaOutBufCount + 1
'v128test MsgBox UBound(NmeaOutBuf) 'v128
    If NmeaOutBufCount - 1 > UBound(NmeaOutBuf) Then ReDim Preserve NmeaOutBuf(NmeaOutBufCount + MINNMEAOUTBUFSIZE)
    NmeaOutBuf(NmeaOutBufCount - 1) = .FullSentence
'Debug.Print "Decode(" & NmeaOutBufCount & ")*" & .NmeaCrc   '"NmeaOutBufClear"
    
    
    If clsCb.Time <> "" Then
        .NmeaRcvTime = UnixTimeToDate(clsCb.Time)
    End If
    
'Try for a prefix date windows or unix format
    If .NmeaRcvTime = "" Then
        If .NmeaPrefix <> "" Then
            If IsDate(.NmeaPrefix) Then
                .NmeaRcvTime = .NmeaPrefix
'V142            ElseIf IsNumeric(.NmeaPrefix) Then
            ElseIf IsNumeric(.NmeaPrefix) And IsLong(.NmeaPrefix) Then  'V142
                    .NmeaRcvTime = UnixTimeToDate(.NmeaPrefix)
            Else
'2014-01-23T12:00:00Z;  Format from Jeffrey van Gils
                kb = Replace(.NmeaPrefix, "T", " ")
                kb = Replace(kb, "Z;", "")
                If IsDate(kb) Then
                    .NmeaRcvTime = kb
                End If
            End If
        End If
    End If

'Try for a date added to the sentence, windows or unix format
'If CRC does not exist don't assume last word can be time stamp
'as sentence may be just passed through eg $AITAG
    If .NmeaRcvTime = "" And .NmeaCrcWord > 0 Then
        If UBound(NmeaWords) > .NmeaCrcWord Then
'try for a windows date first word after CRC (will be here if reading time stamped file
            If IsDate(NmeaWords(.NmeaCrcWord + 1)) Then 'v149
                .NmeaRcvTime = NmeaWords(.NmeaCrcWord + 1)  'v149
'else Try for a Windows date as last word
            ElseIf IsDate(NmeaWords(UBound(NmeaWords))) Then
                .NmeaRcvTime = NmeaWords(UBound(NmeaWords))
            Else
                If IsNumeric(NmeaWords(UBound(NmeaWords))) Then
                    .NmeaRcvTime = UnixTimeToDate(NmeaWords(UBound(NmeaWords)))
                End If
            End If
        End If
    End If
    
'If not valid force NowUtc()    'PositionTime : "1007/2014 06:42:23"
    If IsDate(.NmeaRcvTime) = False Then
        .NmeaRcvTime = NowUtc()  'locale format
    End If

'Make received time the current time
    If .NmeaRcvTime = "" Then
        .NmeaRcvTime = NowUtc()  'locale format
    End If

'IEC Encapsulated Data
    If Left$(.NmeaSentenceType, 1) = "!" Then
        .IsIecEncapsulated = IecEncapsulationCheck(.NmeaSentence)
    End If
    
'Must pass encapsulation test to be a valid ais sentence
    If .IsIecEncapsulated Then
'Must have correct no of words before crc for AIS (1 more than encapsulation)
        If .NmeaCrcWord = 6 Then
            Select Case .IecFormat
            Case Is = "VDM", "VDO", "THAR"
'Process "THAR" some fields WILL be incorrect
'IsAisSentence must be true to view detailed fields
                .IsAisSentence = True
            End Select
        End If
    End If
    
    If .IsAisSentence = True Then
'extract information available if incomplete
        .SentencePart = NmeaWords(2)
'PayloadReassembler must only be called once for each sentence, because
'It will clear the MultiPart buffer when all parts are complete
        .PayloadReassemblerErr = PayloadReassembler(.NmeaSentence, .PayloadReassemblerComments) 'all bits
'Returns error if part 1 of multipart
'        Payload8Bits = (PayloadByteArraySize + 1) * 8
'v133 .AisPayloadBits is to end of part 1, if part 2 not received
'when all parts have been received .AisPayloadBits will be reduced by fill bits
'If .AisPayloadBits <> Payload8Bits Then Stop    'debugv133
        If .AisMsgPart1Missing = False Then
            .AisMsgType = pLong(1, 6)  '230k/min
            .AisMsgRepeat = pLong(7, 2)
        End If
'get other details when weve got a message type and enough bits
'message type will be missing if we have not received part 1
        If .AisMsgType <> "" Then
            If .AisPayloadBits >= 38 Then      '39
                .AisMsgFromMmsi = Format$(pLong(9, 30), "000000000")
            End If

'MMSI From
            Select Case .AisMsgType
'We have to keep lat/lon and vessel names before
'any input filtering - as the user may change this later
'These are added to to vessels cache - last position
            Case Is = "1", "2", "3", "9"    'SAR is the same
                If IsPayloadFillOK Then
                    Call AddLatLon(90, 27, 62, 28, 4)
               End If
            Case Is = "4", "11" 'Base station
                If IsPayloadFillOK Then
                    Call AddLatLon(108, 27, 80, 28, 4)
               End If
            Case Is = "17"  'GNSS
                If IsPayloadFillOK Then
                    Call AddLatLon(59, 17, 41, 18, 1)
               End If
            Case Is = "18"  'Standard class B
                If IsPayloadFillOK Then
                    Call AddLatLon(86, 27, 58, 28, 4)
               End If
            Case Is = "19"  'Extended class B
                If IsPayloadFillOK = 312 Then
                    Call AddLatLon(86, 27, 58, 28, 4)
               End If
            Case Is = "21"  'AtoN
'                If .AisPayloadBits >= 10 Then     '233
'                    .VesselName = Trim$(p6bit(44, 120))
'                    Call AddVessel(.AisMsgFromMmsi, .VesselName, .NmeaRcvTime)
'                End If
                If IsPayloadFillOK Then
                    Call AddName(44, 120)
                    Call AddLatLon(193, 27, 165, 28, 4)
               End If
            Case Is = "27"  'Long range
                If IsPayloadFillOK Then
                    Call AddLatLon(63, 17, 45, 18, 2)
               End If
'MMSI TO
            Case Is = "7", "10", "12", "13", "15", "16"
                If .AisPayloadBits >= 70 Then      '71
                    .AisMsgToMmsi = Format$(pLong(41, 30), "000000000")
                End If
'Vessel Name
            Case Is = "5"
                If IsPayloadFillOK Then
                    Call AddName(113, 120)  'if payload is complete
                End If
'                If .AisPayloadBits >= 232 Then     '233
'                    .VesselName = Trim$(p6bit(113, 120))
'only add if not a shore station
'                    If ChkSentenceFilter("shore", "") = True Then
'                        Call AddVessel(.AisMsgFromMmsi, .VesselName, .NmeaRcvTime)
'                    End If
'                End If
            Case Is = "24"
                If IsPayloadFillOK Then
                    Select Case pLong(39, 2)
                    Case Is = 0    'part A only
                        .AisMsgFiId = "A"
                        Call AddName(41, 120)
                    Case Is = 1     'part B
                        .AisMsgFiId = "B"
                    End Select
                End If
'                If .AisPayloadBits >= 40 Then  '41
'                    Select Case pLong(39, 2)
'                    Case Is = 0    'part A only
'                        .AisMsgFiId = "A"
'                        If .AisPayloadBits >= 160 Then     'ok
'                            .VesselName = Trim$(p6bit(41, 120))
'only add if not a shore station
'                            If ChkSentenceFilter("shore", "") = True Then
'                                Call AddVessel(.AisMsgFromMmsi, .VesselName, .NmeaRcvTime)
'                            End If
'                        End If
'                    Case Is = 1     'part B
'                        .AisMsgFiId = "B"
'                    End Select
'                End If
'ASM's
            Case Is = "6", "8", "25", "26"
                Select Case .AisMsgType
                Case Is = "6"
                    If .AisPayloadBits >= 70 Then  '71
                        .AisMsgToMmsi = Format$(pLong(41, 30), "000000000")
                        DacFrom = 73
                    End If
                Case Is = "8"
                    DacFrom = 41
                Case Is = "25", "26"
                    If pLong(40, 1) = 1 Then 'DAC/FI in use (structured)
                        If pLong(39, 1) = 0 Then
                            DacFrom = 41   'Broadcast
                        Else
                            DacFrom = 73   'Addressed
                        End If
                    End If
                End Select
                
 'If .AisMsgDac <> "" then the sentence is Structured (use dac/fi)
 
                If DacFrom <> 0 And .AisPayloadBits >= DacFrom + 15 Then
                    .AisMsgDac = pLong(DacFrom, 10)
                    .AisMsgFi = pLong(DacFrom + 10, 6)
 'get FIID if applicable
                    Select Case .AisMsgDac
                    Case Is = "0"   'Zeni
                        Select Case .AisMsgFi
                        Case Is = "0"  'Zeni
                            If .AisPayloadBits >= DacFrom + 16 Then   '17
                                .AisMsgFiId = pLong(DacFrom + 16, 16)
                            End If
                        End Select
                    Case Is = "1"
                        Select Case .AisMsgFi
                        Case Is = "21"  'Weather
                           If .AisPayloadBits >= DacFrom + 16 Then   '17
                                .AisMsgFiId = pLong(DacFrom + 16, 1)
                            End If
                        End Select 'Fi
                    Case Is = "316", "366"
                        Select Case .AisMsgFi
                        Case Is = "1", "2", "32"
                            If .AisPayloadBits >= DacFrom + 23 Then   '24
                                .AisMsgFiId = pLong(DacFrom + 18, 6)
                            End If
                        Case Is = "33"
                            If .AisPayloadBits >= DacFrom + 19 Then   '20
                                .AisMsgFiId = pLong(DacFrom + 16, 4)
                            End If
                        End Select 'Fi
'Northern Lights 6-235-15
                    Case Is = "235", "250"
                        Select Case .AisMsgFi
                        Case Is = "15"
                            If .AisPayloadBits >= DacFrom + 23 Then   '24
                            .AisMsgFiId = pLong(DacFrom + 16, 8)
                            End If
                        End Select 'Fi
                    End Select  'Dac with fiid
'all Ais messages
                    .PayloadReassemblerComments = .PayloadReassemblerComments _
                    & " [" & .AisMsgDac
                    .PayloadReassemblerComments = .PayloadReassemblerComments _
                    & "-" _
                    & .AisMsgFi
'FiId  messages only
                    If .AisMsgFiId <> "" Then
                        .PayloadReassemblerComments = .PayloadReassemblerComments & ":" _
                        & .AisMsgFiId
                    End If
                    .PayloadReassemblerComments = .PayloadReassemblerComments & "]"
'addressed  messages only
                    If .AisMsgType = "6" Then
                        .PayloadReassemblerComments = .PayloadReassemblerComments & " to " _
                        & .AisMsgToMmsi
                    End If
                End If  'End ASM
            End Select
'.AisMsgPartsComplete move from here up and now in PayloadReassembler

'try and get name here as we may will not get is from the scheduled messages
'if we have filtered messages containing the name
            If .VesselName = "" And CacheVessels = True _
            And .AisMsgFromMmsi <> "" Then
                    .VesselName = GetVessel(.AisMsgFromMmsi)
            End If

            If RangeReq(0) = True And .AisMsgFromMmsi <> "" Then
                ThisShip = GetCachedVessel(.AisMsgFromMmsi)
'This causes each sentence to be output even when scheduled
'                Call SetLatLongTagValue(ThisShip.RcvTime, (ThisShip.Lat), CStr(ThisShip.Lon))
'but never outputs static data if CSV all - try this
'Stopjna
                If .VesselLat = 91 Then .VesselLat = ThisShip.Lat
                If .VesselLon = 181 Then .VesselLon = ThisShip.Lon
            End If

        End If      'Ais Sentence with AisMsgType
    End If 'AIS Sentence
End With
'MsgBox "FullSentence=" & .FullSentence & vbCrLf _
'& "CommentBlock=" & clsCb.block & vbCrLf _
'& "CBerrmsg=" & clsCb.errmsg & vbCrLf _
'& "NmeaSentence=" & .NmeaSentence
'    DecodingState = DecodingState Xor 64
 
End Function

Private Function ClearInputSentence()   'v129
Dim i As Long

    Set clsSentence = Nothing   '(also Cleared in SchedulerOutput)
    Set clsCb = Nothing          'V142 Clear comment block
    clsSentence.VesselLon = 181 'Not Available
    clsSentence.VesselLat = 91 'Not Available
    Set clsCb = Nothing
    Erase NmeaWords
    Erase CbWords
End Function

'This must only be called if the sentence is going to be output
'because it will force the sentence to be output (because a Tag value is set)
Private Function SetLatLongTagValue(RcvTime As String, Lat As String, Lon As String)
Dim i As Long
    For i = 1 To UBound(TagArray)
        If TagArray(i, 0) = "lat" Then
            If TagArray(i, 1) = "" Then
                TagArray(i, 1) = Lat
                TagArray(i, 2) = RcvTime
            End If
        End If
        If TagArray(i, 0) = "lon" Then
            If TagArray(i, 1) = "" Then
                TagArray(i, 1) = Lon
                TagArray(i, 2) = RcvTime
            End If
        End If
    Next i
End Function

'This stores the latest positions in the vessels cache
'and traps myship sentences
'clsSentence.AisPositionOK is set if this sentence is OK to use for GIS output
Function AddLatLon(LatFrom As Long, LatBits As Long, LonFrom As Long, LonBits As Long, Precision As Long)
Dim wSi As Single
Dim Lat As Single
Dim Lon As Single
Dim arry() As String
Dim i As Long

'Latitude
        wSi = pSi(LatFrom, LatBits)
        Lat = wSi / (60 * 10 ^ Precision)
        clsSentence.VesselLat = Lat ' clssentence
'Longitude
        wSi = pSi(LonFrom, LonBits)
        Lon = wSi / (60 * 10 ^ Precision)
        clsSentence.VesselLon = Lon
        
'Set this position sentence as OK for GIS output
        If Not (Lon = 91 Or Lat = 181) Then clsSentence.AisPositionOK = True
        
'Update Vessels.dat
        Call AddVessel(clsSentence.AisMsgFromMmsi, "", clsSentence.NmeaRcvTime, Lat, Lon)
        
'Trap Own Ship sentences
'        If clsSentence.IecFormat = "VDO" Then  'VDO – AIS VHF data-link own-vessel repor
        If clsSentence.IecFormat = "VDO" And TreeFilter.Check1(24).Value = vbChecked Then 'VDO – AIS VHF data-link own-vessel repor
'Myship requires changing
            If clsSentence.AisMsgFromMmsi <> MyShip.Mmsi Then
'Need to set the name as the current name for this vessel
'To display on NmeaRcv
                Call ClearMyShip(MyShip)
'Set up MyShip from the vessels cache
                MyShip = GetCachedVessel(clsSentence.AisMsgFromMmsi)
                If MyShip.Mmsi <> "" Then
                    If MyShip.Name <> "" Then
                        arry = Split(MyShip.Name, " ")
                    Else
                        arry = Split(MyShip.Mmsi, " ")
                    End If
                    For i = 0 To UBound(arry)
                        If Len(arry(i)) > 1 Then
                            Mid$(arry(i), 2) = LCase(Mid$(arry(i), 2))
                        End If
                    Next i
                    NmeaRcv.Label4(0) = Join(arry, " ")
                    NmeaRcv.Frame11.Visible = True
                End If
            End If
'Update position of MyShip
            MyShip.Lat = Lat
            MyShip.Lon = Lon
            MyShip.RcvTime = clsSentence.NmeaRcvTime
        End If

End Function

'This stores the latest name in the vessels cache
'Add/update vessel name to vassels.dat
'clsSentence.AisPositionOK is set if this sentence is OK to use for GIS output
Function AddName(From As Long, Bits As Long)
Dim i As Long
    With clsSentence
        .VesselName = Trim$(p6bit(From, Bits))  'leading & trailing spaces
'Remove trailing @ in vessel name data base 'v136
' Debug.Print "AddName" & " "; .VesselName
       For i = Len(.VesselName) To 1 Step -1
            If Right$(.VesselName, 1) = "@" Then
                .VesselName = Left$(.VesselName, i - 1)
            Else
                Exit For
            End If
        Next i
'V 144 Jan 2017 unless $AITAG(Jason time stamp)has been found make CSV output compatible with V129
'Change Vessel@Name@@@@@@@@@ to Vessel Name
    If Left$(GroupSentence.NmeaSentence, 6) = "$AITAG" Then
        .VesselName = Trim$(Replace(.VesselName, "@", " "))
    End If
'Set this position sentence as OK for GIS output
        .AisPositionOK = True
        
'Update Vessels.dat
        Call AddVessel(.AisMsgFromMmsi, .VesselName, .NmeaRcvTime)
'Debug.Print "AddName" & " "; .VesselName
    End With
    
End Function

'Constructs TagList from FieldList
'Constructs FieldArray from FieldList
'Constructs TagArray from TagList
'Constructs FieldArrayFromTo to speed up transfer of field values to tag values
'when TagsFromFields adds values to TagArray

'FieldList contains the source of the value for each field and a tag
'TagList contains the Max & Min permitted value for each Tag
'Both are populated on startup, so neither can be deleted if a reference
'in either exists to the other. Likewise neither can contain a reference
'to the other if it does not exist in the other.

'(A) Fields can be added or deleted from FieldList
'(B) Tags can be deleted from TagList

'(A)    1. Add any Tags in FieldList to TagList if Tag not in TagList
'       2. Delete any Tags in TagList if Tag not in Field List

'(B)    Delete any Fields in FieldList if the Tag does not in TagList
'?      Add any Tags in TagList to FieldList if Tag not in FieldList
'Then Construct TagArray
'TagArray is used because it is faster when outputting the Tags

'FieldList contains the AisMsgKey. Any Incoming sentence with the same Key
'is put in the Schedule Buffer (AisMsgs) collection.
'Newer messages with the same key + partno are replaced in AisMsgs
'When Scheduled messages are retrieved (in MMSI order) from AisMsgs, the value for
'any Fields in FieldList are obtained, and placed in TagArray. Again newer values
'replacing older.

'FieldList should be made into a collection keyed on AisMsgKey, with the Fields
'separate items within each key. This would enable faster recovery of Tag.Value
'from the outputted message. I dont think the individual parts of the message key
'need keeping in colFieldList as clsSentence will contain all the detail required
'to get the value of the Source info.

'If Parent = "FieldList" the Tag is synchronised to FieldList
'If Parent = "TagList" the Field is synchronised to TagList

Public Function ResetTags(Parent As String)
Dim i As Long
Dim j As Long   'TagList counter
Dim arry() As Boolean
Dim Restart As Boolean
Dim FieldKey As String
Dim elem As Long   'Ais msg no (0 to 63, NMEA is 0)
Dim PermTagCount As Long

PermTagCount = 0    'removed because of duplicated vesselnames
'(A)    1. Add any Tags in FieldList to TagList - if not already there
Select Case Parent
Case Is = "FieldList"
    With TreeFilter.FieldList
        If .TextMatrix(1, 0) <> "" Then 'if only 2rows and first blank there are no entries
            For i = 1 To .Rows - 1
                With TreeFilter.TagList
                    For j = 1 To .Rows - 1
'Copy Tag Description (Des) to TagList
                        If .TextMatrix(j, 1) = TreeFilter.FieldList.TextMatrix(i, 12) Then
'Dont overwrite name if already set up
If .TextMatrix(j, 4) = "" Then
    .TextMatrix(j, 4) = TreeFilter.FieldList.TextMatrix(i, 13)
End If
                            Exit For
                        End If
                    Next j
                        If j = .Rows Then   'last row
                            .AddItem vbTab & TreeFilter.FieldList.TextMatrix(i, 12) _
                            & vbTab & vbTab & vbTab & TreeFilter.FieldList.TextMatrix(i, 13)
                            If .TextMatrix(1, 1) = "" Then
                                .RemoveItem (1)
                                j = j - 1
                            End If
                        End If
'add name to tag list (used for csv header)
'.TextMatrix(j, 4) = TreeFilter.FieldList.TextMatrix(i, 13)
                End With
            Next i
        End If
    End With
'(A)    2. Delete any Tags in TagList not in Field List
    Do
    Restart = False
        With TreeFilter.TagList
        If .TextMatrix(1, 1) <> "" Then 'if only 2rows and first blank there are no entries
                For j = 1 To .Rows - 1
                    With TreeFilter.FieldList
                        For i = 1 To .Rows - 1
                            If .TextMatrix(1, 1) = "" Then    'blank Field
                                i = .Rows                    'force delete
                                Exit For
                            End If
                            If .TextMatrix(i, 12) = TreeFilter.TagList.TextMatrix(j, 1) Then
                                Exit For    'found Tag in Field
                            End If
                        Next i
                        If i = .Rows Then   'Tag not found in field
                            If TreeFilter.TagList.Rows = 2 Then   'last Tag being removed
                                TreeFilter.TagList.AddItem ""
                            End If
                            TreeFilter.TagList.RemoveItem j 'delete tag
                            Restart = True
                            Exit For
                        End If
                    End With
                Next j
            End If
        End With
    Loop Until Restart = False  'there's been a tag deleted
Case Is = "TagList"
    Do
    Restart = False
        With TreeFilter.FieldList
        If .TextMatrix(1, 12) <> "" Then 'if only 2rows and first blank there are no entries
                For i = 1 To .Rows - 1
                    With TreeFilter.TagList
                        For j = 1 To .Rows - 1
                            If .TextMatrix(1, 1) = "" Then    'blank Tag
                                j = .Rows                    'force delete
                                Exit For
                            End If
                            If TreeFilter.FieldList.TextMatrix(i, 12) = .TextMatrix(j, 1) Then
                                Exit For    'found Field Tag in TagList
                            End If
                        Next j
                        If j = .Rows Then   'Tag not found in field
                            If TreeFilter.FieldList.Rows = 2 Then   'last Field being removed
                                TreeFilter.FieldList.AddItem ""
                            End If
'put row to be deleted in red helps any debugging
                            If i > 1 Then
                                TreeFilter.FieldList.TopRow = i - 1
                            Else
                                TreeFilter.FieldList.TopRow = i
                            End If
                            TreeFilter.FieldList.Row = i
                            TreeFilter.FieldList.col = 12
'            TreeFilter.FieldList.ColSel = TreeFilter.FieldList.Cols - 1
                            TreeFilter.FieldList.CellBackColor = vbRed
                            TreeFilter.FieldList.RemoveItem i 'delete field
                            Restart = True
                            Exit For
                        End If
                    End With
                If Restart = True Then Exit For
                Next i
            End If
        End With
    Loop Until Restart = False  'there's been a tag deleted
End Select

'Construct TagArray from TagList - VesselName must be added here
With TreeFilter.TagList
    ReDim TagArray(.Rows - 2 + PermTagCount, 5) 'First line & col (Descriptions) not required
'set up any permanent tags (only 1 at the moment) mmsifrom
'these will NOT appear in Treefilter.TagList
    If .TextMatrix(1, 1) <> "" Then     'first line not blank, number 1st col
        j = 1
        Do While j <= .Rows - 1
            .TextMatrix(j, 0) = j
'0=Tag,1=Value,2=RcvTime,3=Min,4=Max,5=name
            TagArray(j - 1, 0) = .TextMatrix(j, 1) 'Tag
            TagArray(j - 1, 3) = .TextMatrix(j, 2) 'Min
            TagArray(j - 1, 4) = .TextMatrix(j, 3) 'Max
            TagArray(j - 1, 5) = .TextMatrix(j, 4) 'Name
            j = j + 1
        Loop
'Output PermTags at the end if decoded CSV output
        If PermTagCount <> 0 Then
        End If
    End If
End With

'SynchroniseTagsToInputFilter must be ticked otherwise not actioned
Call TreeFilter.SynchroniseTags

TreeFilter.FieldList.col = 0    'sort on displayed msg/dac/fi
TreeFilter.FieldList.ColSel = 0
TreeFilter.FieldList.Sort = flexSortStringAscending
'TreeFilter.FieldList.Col = 0    'clear the selection cols
'TreeFilter.FieldList.ColSel = 0
'construct FieldArray array from FieldList
'field list is in sorted order
Erase FieldArrayFromTo
With TreeFilter.FieldList
    ReDim FieldArray(.Rows - 1) 'minimum of 2 of which (1)
'be blank, if there are no fields
    If .TextMatrix(1, 0) <> "" Then 'if only 2rows and first blank there are no entries
        For i = 1 To .Rows - 1
            FieldArray(i).MsgKey = .TextMatrix(i, 0)
            FieldArray(i).Source = .TextMatrix(i, 1)
            FieldArray(i).Member = .TextMatrix(i, 2)
            FieldArray(i).From = .TextMatrix(i, 3)
            FieldArray(i).reqbits = .TextMatrix(i, 4)
            FieldArray(i).Arg = .TextMatrix(i, 5)
            FieldArray(i).Arg1 = .TextMatrix(i, 6)
            FieldArray(i).Column = .TextMatrix(i, 7)
'            FieldArray(i).MsgType = .TextMatrix(i, 8)
'            FieldArray(i).Dac = .TextMatrix(i, 9)
'            FieldArray(i).Fi = .TextMatrix(i, 10)
'            FieldArray(i).FiId = .TextMatrix(i, 11)
            FieldArray(i).Tag = .TextMatrix(i, 12)
            FieldArray(i).Valdes = .TextMatrix(i, 13)
'            FieldArray(i).Value = .TextMatrix(i, 14)
'msgtype will be blank if nmea or 2 spaces
'            If .TextMatrix(i, 8) <> "" And Left$(.TextMatrix(i, 8), 2) <> "  " Then
'                elem = .TextMatrix(i, 8)
'            Else
'                elem = 0    'nmea sentence
'            End If

'Force Field array (0) if not an Ais Message (replacing above code)
'            elem = NullToZero(.TextMatrix(i, 8))

            elem = NullToZero(.TextMatrix(i, 8))


'set up first and last element in FieldArray array for this message type
'note 0 element is present but is always blank
            If FieldArrayFromTo(elem, 0) = "" Then
                FieldArrayFromTo(elem, 0) = i     'first element in FieldArray for this aismsgtype
            End If
            FieldArrayFromTo(elem, 1) = i         'last element in FieldArray
        Next i
    End If
End With
'Debug.Print "Reset-" & FieldArray(1).Tag

Call ClearVesselTagValues

End Function

'Clears TagValues in strTags if TagsArray is out of sync with VesselTagsHeader
'Sets VesselTagsHeader to TagArray
Sub ClearVesselTagValues()
Dim myVessel As New clsVessel

If CacheVessels = True And TagsToHeader <> "" Then
'MsgBox "Old:" & VesselTagsHeader & vbCrLf & "New:" & TagsToHeader
    If VesselTagsHeader <> TagsToHeader Then
        VesselTagsHeader = TagsToHeader
        For Each myVessel In Vessels
            myVessel.strTags = ""
        Next myVessel
    End If
End If

End Sub
'Adds either Position or Name or both to Vessels collection (and vessels.dat on exit)
Sub AddVessel(Mmsi As String, _
Optional VesselName As String, _
Optional RcvTime As String, _
Optional Lat As Single = 91, _
Optional Lon As Single = 181)
'If lat or lon not passed then default is 91 or 181

Dim myVessel As clsVessel       'V142 Vessel changed to myVessel
Dim kb As String

If IsNumeric(Mmsi) = False Then Exit Sub 'V3.4.143   MMSI must not be "" (some routines use mmsi is long)

On Error GoTo Key_NewVessel
kb = clsSentence.AisMsgType
Set myVessel = Vessels(Mmsi)  'see if this ship is in collection (err=5 if not)
                             'if not create new ship in ships collection
'v146 removed. Jason error 457 requires error trap to remain because if 457 returned when in
'Key_NewVessel, AddVessel must be exited
'On Error GoTo 0         'V142

Vessel_Update:
    If VesselName <> "" Then    'Name passed to AddVessel
'Only Named Vessels for stats
        If myVessel.Name = "" Then  'New vessel name will added
            NamedVessels = NamedVessels + 1
'Debug.Print "addVessel(" & NamedVessels & ")" & VesselName
        End If
'Update vessel name as it may have changed
        myVessel.Name = VesselName
        If RcvTime <> "" Then myVessel.RcvTime = RcvTime 'UTC
    End If
'dec separator problem (remove if to end if)
'Stop
    
'If both 0 then assume we do not actually have a position - set defaults =None
    If Lat = 0 And Lon = 0 Then
        Lat = 91
        Lon = 181
        End If
    myVessel.LastLat = Lat
    myVessel.LastLon = Lon
'If myVessel.LastLat = 91 Then Stop  'jna 20161118
    If Lat <> 91 And Lon <> 181 Then
        If RcvTime <> "" Then myVessel.RcvTime = RcvTime 'UTC last position received
    End If
    Set myVessel = Nothing
'testing    Vessels.Remove Mmsi 'This causes the class terminated in gcolDebug
'    Call SaveVessels    'jna test only 20161116
Exit Sub
    
Key_NewVessel:                    'create new ship in ships collection
    Select Case err.Number
    Case Is = 5
'Set up in collection if only a position (will not be saved to vessels.dat)
        If VesselName <> "" Or (Lat <> 91 And Lon <> 181) Then
            Set myVessel = New clsVessel  '3
'testing class terminate Set myVessel = Nothing 'this terminates the class for this mmsi
            myVessel.Mmsi = Mmsi
            Vessels.Add myVessel, Mmsi
'this does not terminate the class because it is still referenced in the collection Set myVessel = Nothing
            Resume Vessel_Update
        End If
        err.Clear
    End Select
End Sub

'if arg blank then Name is returned, if "time" time name recveied returned
Function GetVessel(ByRef Mmsi As String, Optional Arg As String) As String
'returns vessel name
Dim myVessel As clsVessel       'V142 Vessel changed to MyVessel
On Error GoTo No_Name
Set myVessel = Vessels(Mmsi)  'see if this ship is in collection
On Error GoTo 0             'if not create new ship in ships collection
Select Case Arg
Case Is = "cache"
    If DateDiff("n", DecoderStartTime, myVessel.RcvTime) < 0 Then GetVessel = "Cached"
Case Is = "time"
    GetVessel = myVessel.RcvTime
Case Else
    GetVessel = myVessel.Name
End Select
No_Name:
Set myVessel = Nothing
End Function

'if arg blank then Name is returned, if "time" time name recveied returned
Function GetCachedVessel(Mmsi As String) As ShipDef
'returns vessel name
Dim Vessel As clsVessel
Dim CachedVessel As ShipDef
    On Error GoTo No_Name
    Set Vessel = Vessels(Mmsi)  'see if this ship is in collection
    On Error GoTo 0             'if not create new ship in ships collection
'return the MMSI that was actually asked for
    CachedVessel.Mmsi = Mmsi
'dec separator
    CachedVessel.Lat = Vessel.LastLat
    CachedVessel.Lon = Vessel.LastLon
'    CachedVessel.PositionTime = Vessel.PositionTime
    CachedVessel.Name = Vessel.Name
    CachedVessel.RcvTime = Vessel.RcvTime
    GetCachedVessel = CachedVessel
    Set Vessel = Nothing
    Exit Function
No_Name:
    GetCachedVessel.Mmsi = Mmsi
End Function

Public Sub SaveVessels()
Dim Vessel As New clsVessel
Dim arry(5) As String  'dec separator change to (2)
'Dim arry(2) As String
Dim kb As String
Dim ch As Long
Dim Des As String
Dim i As Long
Dim Abort As Boolean
Dim Rejects As Long
Dim arVessels() As String    'Vessels collection as CSV array before sorting and writing to disk
Dim arShipNames() As String 'Used to output ShipNames.txt (ShipPlotter format)
'If save vessels is called by TimeZone_Timer before vessels.dat have been saved to vessels
'vessels.dat will be cleared to 0 vessels
'Note abort still updates VesselsFileSyncTime
kb = Vessels.Count
    If Vessels.Count > 0 Then
        DecimalSeparator = GetDecimalSeparator
 'v3.4.143 On Error GoTo Exit_err    'v143 can exit without copying .tmp to .bak
'This was caused by the MMSI not being set in AddVessel (was "") causing Type mismatch in MmsiFmt
'kb = Vessels.Count
        ReDim arVessels(Vessels.Count - 1)
        i = 0
        For Each Vessel In Vessels
            With Vessel
'only add if not a shore station
'Temp remove shore stations from saved vessels
'they don't broadcast names (but could have been set-up
'by user in vessels.dat)
'If CLng(.Mmsi) >= 100000000 Then
'If .Mmsi = "236407000" Then Stop
'dont write away "wrong"mmsi's
                If IsNumeric(.Mmsi) Then
                    Des = MmsiFmt(.Mmsi, "D")
'Only save vessels to vessels.dat that have a name
                    If Des <> "Reserved for Testing" _
                    And Des <> "Invalid MMSI" _
                    And .Name <> "" Then
                        arry(0) = Format$(.Mmsi, "000000000")   'include leading zeros
                        arry(1) = .Name
                        arry(2) = .RcvTime
'Write European Number format away with "." not ","
                        arry(3) = Replace(.LastLat, DecimalSeparator, ".")
'arry(3) = "-8.996666666E01"
'Convert "-8.996666666E01" to "-89.966667"
                        arry(3) = Format$(arry(3), "##0.0#####")
                        arry(4) = CStr(Replace(.LastLon, DecimalSeparator, "."))
'arry(4) = "-1.796666666E02"
'Convert "-1.796666666E02" to "-179.666667"
                        arry(4) = Format$(arry(4), "##0.0#####")
                        arry(5) = .strTags
'If .Mmsi = "244387000" Then Stop
                        arVessels(i) = Join(arry, ",")
'                        Print #ch, Join(arry, ",")
                        i = i + 1
'Debug.Print "SaveVessels(" & i & ") " & arry(1)
                    Else
                        Rejects = Rejects + 1
'Stop   'debug vessels not saved
                    End If
                Else
                    Rejects = Rejects + 1
'stop debug non numeric mmsi's
                End If
'End If 'remove shore stations
            End With
'        kb = kb & Join(arry, ",") & vbCrLf
        Next
'remove "" entries from end of array
        ReDim Preserve arVessels(Vessels.Count - Rejects - 1)
'i = 0
'Do While i < UBound(arVessels)
'    If arVessels(i) = "" Then Stop
'    i = i + 1
'Loop
        Call QuickSort(arVessels)
'i = 0
'Do While i < UBound(arVessels)
'    If arVessels(i) = "" Then Stop
'    i = i + 1
'Loop
        ch = FreeFile
'v148  hourly backup of vessels.dat causes error 70 (Permission Denied) if running twice
'Also occurs when vessels.dat updated on program exit
'Dont attempt to update
        If IsFileInUse(VesselsFileName) Then Abort = True
        If IsFileInUse(VesselsFileName & ".bak") Then Abort = True
        If Abort = False Then
'rename
            FileCopy VesselsFileName, VesselsFileName & ".bak"
            Open VesselsFileName For Output As #ch
            i = 0
            Do While i < UBound(arVessels)
                Print #ch, arVessels(i)
'If arVessels(i) = "" Then Stop
                i = i + 1
            Loop
        End If
        Close #ch
'Output ShipNames.txt (ShipPlotter format)
        If IsFileInUse(Replace(VesselsFileName, "Vessels.dat", "ShipNames.txt")) = False Then
            ch = FreeFile
            Open Replace(VesselsFileName, "Vessels.dat", "ShipNames.txt") For Output As #ch
                i = 0
                Do While i < UBound(arVessels)
                    arShipNames = Split(arVessels(i), ",", 3)
                    Print #ch, arShipNames(0) & " " & Trim$(Replace(arShipNames(1), "@", " "))
                    i = i + 1
                Loop
            Close ch
        End If

        If IsFileInUse(Replace(VesselsFileName, "Vessels", "VesselTagsHeader")) = False Then
            ch = FreeFile
            Open Replace(VesselsFileName, "Vessels", "VesselTagsHeader") For Output As #ch
            Print #ch, TagsToHeader
            Close #ch
        End If
'Stop    'jna test only 20161116
        Set Vessel = Nothing
'MsgBox kb

        Call SaveTagsHeader
    End If

    Call GetSystemTime(VesselsFileSyncTime)

Exit Sub

Exit_err:
    On Error GoTo 0
    Close #ch
End Sub

Public Sub SaveTagsHeader()
Dim ch As Long

ch = FreeFile
Open Replace(VesselsFileName, "Vessels", "VesselTagsHeader") For Output As #ch
Print #ch, TagsToHeader
Close #ch
End Sub

Public Sub ReadTagsHeader()
Dim ch As Long
        
ch = FreeFile
On Error GoTo No_Header
Open Replace(VesselsFileName, "Vessels", "VesselTagsHeader") For Input As #ch
Line Input #ch, VesselTagsHeader
Close #ch
No_Header:
End Sub

Public Function TagsToHeader() As String
Dim TagNo As Long
Dim arTags() As String          'Tag is the Unique FieldTag (not the user's tag name)
ReDim arTags(UBound(TagArray, 1))
For TagNo = 0 To UBound(TagArray, 1)
    arTags(TagNo) = TagArray(TagNo, 0)
Next TagNo
TagsToHeader = Join(arTags, ",")
End Function

Public Sub ReadVessels()
Dim ch As Long
Dim nextline As String
Dim arry() As String
Dim i As Long
Dim arTags() As String
Dim strTags As String
''Dim myVessel As New clsVessel

    DecimalSeparator = GetDecimalSeparator
    VesselsFileName = FileSelect.SetFileName("VesselsFileName")
    If FileExists(VesselsFileName) Then
                
'Set the TagsHeader to the SavedHeader on disk that will have the format of
'Vessels.strTags when Vessels has been read
        Call ReadTagsHeader
        ch = FreeFile
        Open VesselsFileName For Input As #ch
        Do Until EOF(ch)
''            Set myVessel = New clsVessel        'V142
            Line Input #ch, nextline
            If Date < CDate(#12/6/2015#) Then
                nextline = Replace(nextline, "~", ",")
            End If
            strTags = ""
            arry() = Split(nextline, ",")
            If UBound(arry) < 5 Then ReDim Preserve arry(5) 'discard all data after lat lon
'            If UBound(arry) >= 2 Then  'decseparator
'add all names (even if shore station)
'arry for lat lon can be blank
'                If IsNumeric(arry(3)) = False Then arry(3) = 0
'                If IsNumeric(arry(4)) = False Then arry(4) = 0
'                Call AddVessel(arry(0), arry(1), arry(2), CSng(arry(3)), CSng(arry(4)))
'                Call AddVessel(arry(0), arry(1), arry(2))
                With clsSentence
'From v13, vessel.dat file lat/lon always written away with . as separator
'Convert any lat lon written away with local separator to "."
'All vessels.dat are now use . as decimal separator.
                    arry(3) = Replace(arry(3), DecimalSeparator, ".")
                    arry(4) = Replace(arry(4), DecimalSeparator, ".")
                    .AisMsgFromMmsi = NullToZero(arry(0))   'Force 0 if not numeric
                                                            'user may have edited vessels.dat
                    .VesselName = arry(1)
                    .NmeaRcvTime = arry(2)
'Vessel.dat may not include Lat & Lon
                    If Not IsNumeric(arry(3)) Then arry(3) = 91 'Default not available
'Convert to PC's separator
'If you dont 90.0000 when converted to single is 900000
'not 90,0000 (in local representation)
                    .VesselLat = CSng(Replace(arry(3), ".", DecimalSeparator))
                    If Not IsNumeric(arry(4)) Then arry(4) = 181 'Default not available
                    .VesselLon = CSng(Replace(arry(4), ".", DecimalSeparator))
'                    Call AddVessel(.AisMsgFromMmsi, .VesselName, .NmeaRcvTime, _
'                    .VesselLat, .VesselLon, arry(5))
'V3.4.143 remove from here
'                    If UBound(arry) > myVessel.VesselTagsStart Then
'                        ReDim arTags(UBound(arry) - myVessel.VesselTagsStart)
'                        For i = myVessel.VesselTagsStart To UBound(arry)
'                            arTags(i - myVessel.VesselTagsStart) = arry(i)
'                        Next i
'                        strTags = Join(arTags, ",")
'                    End If
'Call AddVessel(.AisMsgFromMmsi, .VesselName, .NmeaRcvTime,.VesselLat, .VesselLon, strTags)    '3
'v3.4.143 remove to here
'Stop
                    Call AddVessel(.AisMsgFromMmsi, .VesselName, .NmeaRcvTime, _
                    .VesselLat, .VesselLon) 'v3.4.143 strTags must not be saved/restored as tags may have been changed/deleted
'                                           eg with a new profile
                End With
''            Set myVessel = Nothing      'V142
        Loop
        Call ClearInputSentence 'clear clsSentence
        Call NmeaRcv.UpdateStats
        NmeaRcv.Stats.TextMatrix(9, 1) = NamedVessels    'Vessels.Count
        Close #ch
    End If
'Keep the header for the Vessels just loaded
    VesselTagsHeader = TagsToHeader
    Call GetSystemTime(VesselsFileSyncTime)
'V142    Set myVessel = Nothing
End Sub

Sub TrappedMsgInput()
Dim TrappedMsg As New clsAisMsg
Dim TrappedMsgKey As String 'Excludes MMSI

With clsSentence
    TrappedMsgKey = cKey(.AisMsgType, 2) _
    & cKey(.SentencePart, 1) _
    & cKey(.AisMsgDac, 4) _
    & cKey(.AisMsgFi, 2) _
    & cKey(.AisMsgFiId, 2)
End With

On Error GoTo Key_NewTrappedMsg
Set TrappedMsg = TrappedMsgs(TrappedMsgKey)  'see if this AisMsg is in collection
On Error GoTo 0             'if not create new AisMsg in AisMsgs collection
Set TrappedMsg = Nothing
Exit Sub

TrappedMsg_Update:
    TrappedMsg.AisSentence = clsSentence.NmeaSentence
    TrappedMsg.NmeaRcvTime = clsSentence.NmeaRcvTime
    Set TrappedMsg = Nothing
Exit Sub

Key_NewTrappedMsg:                    'create new ship in ships collection
    On Error GoTo 0
    Set TrappedMsg = New clsAisMsg
    TrappedMsg.AisMsgKey = TrappedMsgKey    'dont actually need to keep, but hels debugging
    TrappedMsgs.Add TrappedMsg, TrappedMsgKey
    Resume TrappedMsg_Update
End Sub

'needs writing & testing
Public Sub ReadTrappedMsgs()
Dim ch As Long
Dim nextline As String
Dim arry() As String

'If DecodingState <> 0 Then Exit Sub
'DecodingState = DecodingState Or 16
TrappedMsgsFileName = FileSelect.SetFileName("TrappedMsgsFileName")
If FileExists(TrappedMsgsFileName) Then
    ch = FreeFile
    Open TrappedMsgsFileName For Input As #ch
    Do Until EOF(ch)
        Line Input #ch, nextline
'create clsSentence
        Call DecodeSentence(nextline)
        
        Call NmeaOutBufClear
'load into collection
        Call TrappedMsgInput
    Loop
    Close #ch
End If
'DecodingState = DecodingState Xor 5
End Sub

Public Sub SaveTrappedMsgs()
Dim TrappedMsg As New clsAisMsg
Dim kb As String
Dim ch As Long

ch = FreeFile
On Error GoTo Exit_err
'should already been set when file is read
If TrappedMsgsFileName = "" Then TrappedMsgsFileName = FileSelect.SetFileName("TrappedMsgsFileName")
Open TrappedMsgsFileName & ".tmp" For Output As #ch

    For Each TrappedMsg In TrappedMsgs
        Print #ch, TrappedMsg.AisSentence
'        kb = kb & Join(arry, ",") & vbCrLf
    Next
Close #ch
FileCopy TrappedMsgsFileName & ".tmp", TrappedMsgsFileName
Kill TrappedMsgsFileName & ".tmp"
Set TrappedMsg = Nothing
'MsgBox kb
Exit Sub
Exit_err:
On Error GoTo 0
Close #ch
End Sub

Sub FileToLog(FileName As String)
Dim ch As Long
Dim nextline As String
Call WriteStartUpLog("")
Call WriteStartUpLog("Print of: " & FileName)

ch = FreeFile
On Error GoTo nofil
Open FileName For Input As #ch
Do Until EOF(ch)
    Line Input #ch, nextline
    Call WriteStartUpLog(nextline)
Loop
Close #ch
Call WriteStartUpLog("Finished print of: " & FileName)
Exit Sub
nofil:
On Error GoTo 0
Close #ch
WriteStartUpLog ("File not found " & FileName)
End Sub
Function SchedulerInput() As Boolean  'returns true if output is required
'Static LastTimeOutput As String
Dim AisMsg As New clsAisMsg    'ship is class
Dim AisMsgKey As String
Dim TimeSinceOutput As Long 'seconds

'If NmeaRcv.ScheduledTimer.Enabled = False Then
'    Stop
'End If

 'Debug.Print Time$ & "SchedulerInput Start"
 'Debug.Print NmeaBufNxtIn & " " & NmeaBufNxtOut & " " & NmeaBufUsed & " SchedulerInput"
With clsSentence
    AisMsgKey = cKey(.AisMsgFromMmsi, 9) _
    & cKey(.AisMsgType, 2) _
    & cKey(.SentencePart, 1) _
    & cKey(.AisMsgDac, 4) _
    & cKey(.AisMsgFi, 2) _
    & cKey(.AisMsgFiId, 2)
End With

'Add clsSentence into AisMsgs
On Error GoTo Key_NewAisMsg
Set AisMsg = AisMsgs(AisMsgKey)  'see if this AisMsg is in collection
On Error GoTo 0             'if not create new AisMsg in AisMsgs collection

AisMsg_Update:
'If clsSentence.AisMsgFromMmsi = "" Then Stop
    AisMsg.AisSentence = clsSentence.NmeaSentence
    AisMsg.NmeaRcvTime = clsSentence.NmeaRcvTime

'Then check if scheduled output is required
'may be reading log file which may not have a timestamp
    If IsDate(AisMsg.NmeaRcvTime) = False Then AisMsg.NmeaRcvTime = NowUtc()
    If IsDate(LastTimeOutput) = False Then  'set up defaults
        LastTimeOutput = AisMsg.NmeaRcvTime
    End If
    TimeSinceOutput = DateDiff("s", LastTimeOutput, AisMsg.NmeaRcvTime)
    If TimeSinceOutput >= (UserScheduledSecs + ScheduledFreqAdj) And clsSentence.AisMsgPartsComplete = True Then
'   if so output all scheduled buffer & clear dead entries
 'Debug.Print Time$ & " Last Time Output" & LastTimeOutput
        LastTimeOutput = AisMsg.NmeaRcvTime
        SchedulerInput = True
    End If
Set AisMsg = Nothing
 'Debug.Print NmeaBufNxtIn & " " & NmeaBufNxtOut & " " & NmeaBufUsed & " SchedulerInputExit"
 'Debug.Print Time$ & "SchedulerInput Finish"
Exit Function

Bad_Freq:       'not numeric
Exit Function

Key_NewAisMsg:                    'create new ship in ships collection
    On Error GoTo 0
    Set AisMsg = New clsAisMsg
    AisMsg.AisMsgKey = AisMsgKey    'dont actually need to keep, but hels debugging
    AisMsgs.Add AisMsg, AisMsgKey
    Resume AisMsg_Update
    
End Function

Function SchedulerOutput(Optional Caller As String)

'Only called at the end of ProcessSentence and by ScheduledTimer when scheduled output is due
'Outputing from the Scheduled buffer must be after all processing of the current message
'because current message details (clsSentence, PayloadBytes etc) will be overwritten
Dim TimeToLive As Long  'minutes set by user
Dim Age As Long
Dim AisMsg As New clsAisMsg    'ship is class
Dim Count As Long
Dim CompleteMessages As Long

Dim AisMsgKeyLen As Integer
'Dim LastTimeOutput As String
Dim i As Long
Dim arry() As String     'hold key + sentence
Dim LastMessage As Boolean  'force output of tags on last message
Static LastOutputted As Long    'debug speed only
Dim Channel As Long
Dim NmeaOutputted As Boolean
Dim TagsInputted As Boolean
Dim LastOutput As Long
Dim OutputOK As Boolean
Dim LastMmsi As String
Dim ThisMmsi As String
Dim TagsFound As Boolean    'tags found since last output of tags
Dim StartTime As Single
Dim ScheduledTimeTaken As Single  'seconds (for display)
Dim OutputCount As Long

'    DecodingState = DecodingState Or 2
'spool received nmea sentences while processing scheduled AisMsgs
'NmeaRcv.ScheduledTimer.Enabled = False
'SuspendProcessScheduler = True
'ProcessSuspended = True
    Processing.Suspended = True
    Processing.Scheduler = True

'Set time last output RealTime or Message Time
Select Case Caller
Case Is = "Timer"
    LastTimeOutput = NowUtc()
Case Else
'Set as rcvtime (even if reading from file)
    If IsDate(LastTimeOutput) = False Then  'set up defaults
        LastTimeOutput = clsSentence.NmeaRcvTime
'If no time on file assume current time as we
'must set some time
        If IsDate(LastTimeOutput) = False Then LastTimeOutput = NowUtc()   'locale
    End If
End Select

On Error GoTo Bad_Time:
TimeToLive = TreeFilter.Text1(1).Text
On Error GoTo 0

DecimalSeparator = GetDecimalSeparator
NmeaRcv.cbSpawnGis.Enabled = False
StartTime = Timer

 'Debug.Print Time$() & " Start " & Caller
 'Debug.Print Time$ & " AisMsgs " & AisMsgs.Count
'Screen.MousePointer = 11 ' Hourglass (wait).
'Count = AisMsgs.Count

If AisMsgs.Count > 0 Then
    AisMsgKeyLen = Len(AisMsgs.Item(1).AisMsgKey)

'load messages in schedule buffer(AisMsgs) into sort array (arry)
'removing "old" messages (if timetolive <> 0) otherwise
'messages since last pass will never be output
    ReDim arry(AisMsgs.Count - 1)
    For Each AisMsg In AisMsgs
        Age = DateDiff("n", AisMsg.NmeaRcvTime, LastTimeOutput)
        If Age > TimeToLive And TimeToLive > 0 Then
'delete message
            AisMsgs.Remove AisMsg.AisMsgKey
        Else
            arry(Count) = AisMsg.AisMsgKey & AisMsg.AisSentence
'If Len(Arry(i)) = 0 Then Stop
            Count = Count + 1
        End If
    Next AisMsg
    Set AisMsg = Nothing
 'Debug.Print Time$() & "Start Sort Array " & UBound(arry) + 1
    Call QuickSort(arry)
 'Debug.Print Time$() & "Start Sort  Array " & UBound(Arry) + 1
'Call ShuffleArray(Arry)
End If

'if outputting to list clear and set up existing list
If NmeaRcv.Option1(3).Value = True Then
    With List.NmeaDecodeList
        .Redraw = False    'to slow if sentences display onlist while outputting
'start again at top of list when outputting AisMsgs (buffer)
        Do While .Rows > 20
            .RemoveItem .Rows
        Loop
        .Clear  'blank remaining rows
        .FormatString = "<NMEA|^Sentence|^MMSI|^Message Type|^DAC|^FI|^ID|Vessel Name|Comments"
        .ColWidth(0) = 0    'must be after formatstring
'when displaying Sched AisMsgs
        MaxNmeaDecodeListCount = MAXNMEADECODELISTCOUNT_Sched
    End With
End If

NmeaRcv.StatusBar.Panels(1) = Time$ & "Output Started " & "[" & AisMsgs.Count & " messages]"
        
 'Debug.Print Time$ & " Msgs Alive " & Count
        
'we only come here if some of the list display is scheduled to clear the existing entries
For Channel = 1 To 2
    If ScheduledReq(Channel) And DisplayOutput(Channel) = True And ChannelMethod(Channel) = "file" Then
Outputs(Channel).AutoRedraw = False
        With Outputs(Channel).MSFlexGrid1
'        .Redraw = False
'start again with output display - dont use clear (initial blank rows)
        Do While .Rows > 10 + .FixedRows    '10 displayed initially
            .RemoveItem .Rows
        Loop
        .Clear
        End With
'Modal form is displayed
        On Error Resume Next
        Outputs(Channel).Hide
'if you dont set focus when you hide output it also hides the forms behind it
        NmeaRcv.SetFocus
        On Error GoTo 0
        End If
Next Channel


'now output decode and output the sentences
SchedOut = 0    'for stats
Call ClearInputSentence 'v129
'will be set up again if DecodeSentence is called
'Now got to handle the output if it is 0
If Count > 0 Then   'Is something to output
    For i = LBound(arry) To UBound(arry)
'occasionally a null record is returned from the array. VB6 returns a null
'record length if the record contains a nul' and mid$ will fail
'this may be caused by a timing issue, so ignore this record
        If Len(arry(i)) <> 0 Then
            ThisMmsi = Left$(arry(i), SchedKeyLen)
            If LastMmsi = "" Then  'first one
                LastMmsi = ThisMmsi    'MMSI (dont use clssentence will be "" if GPS)
            End If
'CHECK IF ALL PARTS COMPLETE
'Output each message List Scheduled (3) if required
'dont do here as well as in ProcessSentence else output will be duplicated
            If NmeaRcv.Option1(3).Value = True Then
                Call List.AddToNmeaSummaryList  'This will be the new list
            End If
 'debug only
'List.NmeaDecodeList.Redraw = True
'If clssentence.AisMsgFromMmsi = "259024000" Then Stop
        
            If ThisMmsi <> LastMmsi Then
                Call OutputAll(True)    'True = on any scheduled channel
                LastMmsi = ThisMmsi    'MMSI (dont use clssentence will be "" if GPS)
                TagsFound = False
            End If
'all output must have been done before we set up the next cslSentence, PayloadBytes and Vessels.dat
'creates clsSentence, PayloadBytes() and looksup Vessels.dat
            Call DecodeSentence(Mid$(arry(i), AisMsgKeyLen + 1))
'If clsSentence.AisMsgPartsComplete = False Then Stop
'get the tags once for all channels
            If TagsReq(0) = True And clsSentence.AisMsgPartsComplete = True Then
                    TagsFound = TagsFromFields   'found a tag
            End If
'V123   'all Nmea put into NmeaOutBuf when DecodeSentence is called
'if any raw output put sentence into FEN buffer  , for output later
'            If FenReq(0) = True Then
'                NmeaArrayNo = NmeaArrayNo + 1
'v102            NmeaArray(NmeaArrayNo) = clsSentence.NmeaSentence
'                NmeaArray(NmeaArrayNo) = clsSentence.FullSentence
'Debug.Print "FenInS(" & NmeaArrayNo & ")" & Left$(NmeaArray(NmeaArrayNo), 15)
'            End If
    
        End If
        DoEvents    'get new udp data
    Next i

'Clear all messages since last scheduled if TimetoLive is zero
    If TimeToLive = 0 Then
        For Each AisMsg In AisMsgs
'delete message
            AisMsgs.Remove AisMsg.AisMsgKey
        Next AisMsg
        Set AisMsg = Nothing
    End If

'need to clear again as DecodeSentence will have set it up again
    Call ClearInputSentence 'v129
End If  'are some sentences in array

'output the last mmsi (if any) by channel
Call OutputAll(True)    'True = on any scheduled channel
'Close the Scheduled Output Files
'Stop
For Channel = 1 To 2
    If ScheduledReq(Channel) = True Then
        If ChannelMethod(Channel) = "udp" Then
            Call NmeaRcv.CloseOutputUdp(Channel)
        End If
        If ChannelMethod(Channel) = "file" Then
            Call NmeaRcv.CloseOutputDataFile(Channel)    'write out No Data if not opened
'restart output file
'If output is kmz, CloseOutputFile will zipup KML output with
'Overlay Output into one file. frmFTP.FtpUpload is done
'after scheduler processing has been re-started (later in this
'routine)
        End If
    End If
Next Channel

'All messages are now in list (if required) so display them
If NmeaRcv.Option1(3).Value = True Then     'scheduled list output
    With List.NmeaDecodeList
        .TopRow = 1             'position scroll bar at top
        .ScrollTrack = True     'allow user to move scroll bar
        .Redraw = True          'update display
    End With
'show list as it may not be visible first time, or could have been closed
'note cant show non-modal form when modal form displayed
    If cmdNoWindow = False And FieldInput.Visible = False Then List.Show
End If

'Show Output channel Window (if reqd)
For Channel = 1 To 2
    If ScheduledReq(Channel) Then
'Debug.Print "SchedulerOutput=" & Screen.ActiveForm.Name
        If DisplayOutput(Channel) Then
Outputs(Channel).AutoRedraw = True
            With Outputs(Channel).MSFlexGrid1
'               .TopRow = .FixedRows             'position scroll bar at top (base 0)
                .ScrollTrack = True     'allow user to move scroll bar
                .Redraw = True          'update display
                .TopRow = .Rows - 1     'Position scroll bar at bottom
            
            End With
'modal form is displayed
            On Error Resume Next
            If cmdNoWindow = False Then
                If Outputs(Channel).Visible = False Then
                    Outputs(Channel).Show
                End If
            End If
            On Error GoTo 0
        End If
'Debug.Print "SchedulerOutputEnd=" & Screen.ActiveForm.Name
    End If
Next Channel

LastOutputted = CLng(NmeaRcv.Stats.TextMatrix(4, 1))

'IF FTP'ing rename the output file(s) so that if when the output file
'is re-opened when the next scheduled time is reached
'it will not overwrite the current FTP file(s), if the FTP upload
'has not finished
'IF the output file is a zip archive (.kmz) and there is an overal file
'it will already be included in the .kmz output file
'If were going to FTP the Output file, it is copied to
'OutputFile + .ftp when it is closed. This allows the file to be read
'whilst a new output file is being created.

'MOVED to CloseOutputDataFile
If FtpUploadExecuting = False Then
    For Channel = 1 To 1
        If FtpOutput(Channel) = True And ScheduledReq(Channel) = True Then
'Error if OutputFile has not been created
            On Error GoTo Copy_err
            FileCopy OutputFileName, OutputFileName & ".ftp"
'MsgBox OutputFileName
            On Error GoTo 0
        End If
    Next Channel
End If

'stop spooling received sentences
StartProcessing:
'DecodingState = DecodingState Xor 2
 'Debug.Print Time$() & " End Output"
'Done before FTP transfer
ScheduledTimeTaken = Timer - StartTime
NmeaRcv.StatusBar.Panels(1) = "Finished Scheduled Output in " & Format$(ScheduledTimeTaken, "###0.00") & " seconds"
NmeaRcv.ClearStatusBarTimer.Enabled = True

NmeaRcv.ScheduledTimer.Enabled = True
'SuspendProcessScheduler = False
'Call ResumeProcess
'    If NmeaBufXoff = False Then
'        Call NmeaRcv.ProcessNmeaBuf
'    End If
Call ResumeProcessing("Scheduler")

NmeaRcv.cbSpawnGis.Enabled = SpawnGisOk
Screen.MousePointer = 0 ' Hourglass (wait).

'MOVED to CloseOutputDataFile
'start off the FTP upload (will not happen if previous one not finished)
'Note the file to be uploaded has been copied to .ftp
For Channel = 1 To 1
    If FtpOutput(Channel) = True And ScheduledReq(Channel) = True Then
'MsgBox "FtpUploadExecuting=" & FtpUploadExecuting
    Call frmFTP.FtpUpload   'move to here from above
'    frmFTP.StartTimer.Enabled = True
 'Debug.Print Time$() & " End FTP"
    End If
Next Channel
 'Debug.Print "Due " & ScheduledDue
 'Debug.Print Time$ & " Finished"
Exit Function

Copy_err:
On Error GoTo 0
Resume StartProcessing
Exit Function

Bad_Time:
    Call ResumeProcessing("Scheduler")
'DecodingState = DecodingState Xor 2
End Function

'file will not exist when first started
Public Sub WriteErrorLog(kb As String)
ErrorLogFile = FileSelect.SetFileName("ErrorLogFile")
ErrorLogFileCh = FreeFile
Open ErrorLogFile For Append As #ErrorLogFileCh
Print #ErrorLogFileCh, Now() & vbTab & kb
Close #ErrorLogFileCh
End Sub

'only turn off taggedOutput(0) if neither udp or file tagged output
'Also closes OutputFile if tagged output
Sub WriteTagTail(Channel As Long)
Dim kb As String

'Debug.Print "#WriteTagTail(" & Channel & ")"
    On Error GoTo error_close
'add tail on if using Tag template
    If TagTemplateTail(Channel) <> "" Then
        kb = TagTemplateTail(Channel)
    End If

'Replace the [LINK] filename ,it could have been changed
    If ChannelMethod(Channel) = "file" Then
        If OverlayOutputFileCh <> 0 Then
            kb = Replace(kb, "[LINK]", _
                "<NetworkLink><name>Close Up View</name>" _
                & "<Link><href> " _
                & NameFromFullPath(OverlayOutputFileName) _
                & "</href></Link></NetworkLink>")
        End If
        
        Call OutputToDataFile(Channel, kb)
'always close overlay even if output file is kml
        If OverlayOutputFileCh <> 0 Then
            Print #OverlayOutputFileCh, "</Document></kml>"
 'Debug.Print "Close " & OverlayOutputFileCh
            Close #OverlayOutputFileCh
            OverlayOutputFileCh = 0
        End If
    End If
        
    If ChannelMethod(Channel) = "udp" Then
        Call OutputUdp(Channel, TagTemplateTail(Channel))
    End If
    
    TaggedOutputOn(Channel) = False
        
Exit Sub

error_close:
    MsgBox err.Description, vbExclamation, "Close Tagged Output"

End Sub

Sub WriteTagHead(Channel As Long)
Dim kbtag As String
Dim strNowUtc As String
'Debug.Print "#WriteTagHead(" & Channel & ")"
    kbtag = TagTemplateHead(Channel)
    strNowUtc = NowUtc()
    kbtag = Replace(kbtag, TagChr(0) & "Now" & TagChr(1), strNowUtc)
    If kbtag <> "" Then
        If ChannelMethod(Channel) = "file" Then Call OutputToDataFile(Channel, kbtag)
        If ChannelMethod(Channel) = "udp" Then Call OutputUdp(Channel, kbtag)
    End If
End Sub

Private Sub pDisplayError(ByVal sError As String)
Dim sMsg As String

If Trim$(sError) = "" Then
    sMsg = err.Description
Else
    sMsg = sError & vbCrLf & vbCrLf & err.Description
End If

If err.Number <> 0 Then sMsg = sMsg & " (" & CStr(err.Number) & ")"

'If cmdNoWindow = False Then
'stops decoder continuing
'    MsgBox sMsg, vbCritical
'End If
    
Call WriteErrorLog(sMsg)
End Sub

Sub OutputToDataFile(Channel As Long, kb As String)  'file only
Dim Mess As String

'Exit Sub
'Debug.Print "#OutputToDataFile"

    If MethodOutput(Channel) = True Then    'udp or file, only file comes here
        If OutputDataFile Is Nothing Then
            Call NmeaRcv.OpenOutputDataFile(Channel)
            If Left$(kb, 8) <> "No Data " Then
                If TaggedOutput(Channel) = True Then
                    If TaggedOutputOn(Channel) = False Then  'head not yet written out
                        TaggedOutputOn(Channel) = True  'Write Tail on Close
                        Call WriteTagHead(Channel)  'Re-entrant call to OutputToDataFile
                    End If
                End If
            End If
        End If
'file could be locked
        If Not (OutputDataFile Is Nothing) Then
            OutputDataFile.WriteString kb & vbCrLf
            NoDataOutput(Channel) = False
        Else
                Mess = "[Locked] "  'Only used for Display channel
        End If
    End If
    If DisplayOutput(Channel) = True Then
        Call Output.OutputDisplay(Channel, Mess & kb)
    End If
'Debug.Print "#Exit OutputToDataFile"
End Sub

Function AsciiToXml(ByVal Channel As Long, ByVal Instring As String) As String
'converts ascii to numbered entities
Dim OutString As String
Select Case ChannelEncoding(Channel)
Case Is = "xml", "kml"
    OutString = Instring
    OutString = Replace(OutString, "&", "&amp;")
    OutString = Replace(OutString, "<", "&lt;")
    OutString = Replace(OutString, ">", "&gt;")
    OutString = Replace(OutString, """", "&quot;")
    OutString = Replace(OutString, "'", "&apos;")
    OutString = Replace(OutString, "°", "&#176;")
    OutString = Replace(OutString, "‰", "&#8240;")
    AsciiToXml = OutString
    If ChannelEncoding(Channel) = "kml" Then
        OutString = Replace(OutString, "=", "&#061;")   'GoogleEarth
    End If
    AsciiToXml = OutString
Case Else
    AsciiToXml = Instring
End Select
End Function

Function QuotedString(ByVal kb As String, Optional ByVal Delim As String) As String
Dim Found As Long
If Delim <> "" Then
    Found = InStr(kb, Delim)
Else
    Found = 1
End If
If Found Then
'Replace " with "" if found in field RFC4180
    kb = Replace(kb, """", """""")
    QuotedString = """" & kb & """"
Else
    QuotedString = kb
End If
End Function

Public Function UrlToName(Url)
Dim i As Integer
Dim j As Integer
j = 0
Do
i = j + 1
j = InStr(i, Url, "/")
Loop Until j = 0
UrlToName = Mid$(Url, i, Len(Url) - i + 1)
End Function

Public Function Terminate()
Dim i As Integer
Dim kb As String
Dim myObj As Variant
Dim f As Form
Dim ctrl As Control     'V142
Dim myVessel As clsVessel

Call WriteStartUpLog("Terminating")

If FormLoaded("NmeaRcv") Then   'v3.4.143 stop NmeaRcv being reloaded on exit
    NmeaRcv.StatusBar.Panels(1).Text = "Terminating AisDecoder"

'Stop the all event timers
'kb = ""
    For Each ctrl In NmeaRcv    'V142
        If TypeOf ctrl Is Timer Then
'kb = kb & ctrl.Name & vbCrLf
            ctrl.Enabled = False
        End If
    Next ctrl                                   'V142 to here
End If

Call PrintRegistry

'V142 add from here
For i = LBound(ListFilters) To UBound(ListFilters)
    Set ListFilters(i) = Nothing
Next i
For i = LBound(Outputs) To UBound(Outputs)
    Set Outputs(i) = Nothing
Next i
'V142 to here
'V142 remove from here
'For Each myObj In ListFilters
'    Set myObj = Nothing
'Next
'For Each myObj In Outputs
'    Set myObj = Nothing
'Next
'V142 to here
frmSplash.lblAction = "Unloading vessels"
'frmSplash.Refresh
Set AisMsgs = Nothing
Debug.Print "Vessels = " & Vessels.Count
Debug.Print "Classes open = " & gcolDebug.Count
Set Vessels = Nothing   'this removes all the vessels in
Debug.Print "Vessels = " & Vessels.Count
Debug.Print "Classes open = " & gcolDebug.Count
Set TrappedMsgs = Nothing
Set clsSentence = Nothing
Set clsCb = Nothing          'Clear comment block
Set clsField = Nothing
'treefilter must be unloaded first otherwise if it is visible
'some settings mey be reset when unloaded
frmSplash.lblAction = "Unloading filters"
frmSplash.Refresh
Unload TreeFilter
frmSplash.lblAction = "Terminating"
frmSplash.Refresh


Call UnloadAllForms(NmeaRcv)
'Unload Detail   'It no longer reloads nmearcv
'For Each f In Forms
'kb = f.Name & ":" & f.Caption
'    f.Visible = False
'    If f.Name <> "frmSplash" Then
'        Unload f
'    End If
'Next f
'do a second time as it doesnt unload nmearcv mmsifilter first time
'it seems to reload it - not sure why
'This is due to a QueryUnload reloading a different form
'NmeaRcv must not get reloaded otherwise you are asked twice if you want to quit
'List Filter (only) is still getting re-loaded by something
'kb = Forms.Count
'For Each f In Forms
'kb = f.Name & ":" & f.Caption
'    f.Visible = False
'    If f.Name <> "frmSplash" Then
'        Unload f
'    End If
'Next f


kb = Forms.Count

'If OutputFileCh <> 0 Then Call CloseOutputDataFile
'Cant do this because any output file that has been properly closed
'With Stop or EOF will be changed to No Data
'Call CloseOutputDataFile    'write out No Data if not opened
If HexDump = True Then
    If HexDumpFileCh <> 0 Then Close HexDumpFileCh
End If

Call WriteStartUpLog("Terminating Process " & Now())
Call CloseStartupLogFile  'v140
'Unload frmSplash

If IsAnyFileOpen(True) = True Then
'    MsgBox "File left open"
End If
ReleaseMutex (hMutex)
CloseHandle hMutex
'Must not force END otherwise NmeaRcv is not unloaded properly
'v140 If Forms.Count = 0 Then End   'terminate process (if Terminate not called by NmeaRcv)
'V3.4.143 If terminate has been called by NmeaRcv it must not be called again by .main
If Forms.Count > 0 Then  'v3.4.143
    Call LogForms("Open forms at end of MainControl.Terminate")  'v3.4.143
    Call CloseStartupLogFile       'v3.4.143
End If 'v3.4.143
End Function

Public Function daDat(Ddat As String) As String
daDat = Mid$(Ddat, 7, 2) & "/" & Mid$(Ddat, 5, 2) & "/" & Mid$(Ddat, 1, 4) _
& " " & Mid$(Ddat, 9, 2) & ":" & Mid$(Ddat, 11, 2) & ":" & Mid$(Ddat, 13, 2)
End Function

'when converted to a full tree as below, other arguments will be required
Function ProcessSentenceFilter() As Boolean
Dim Range As Single

    ProcessSentenceFilter = True   'assume a pass because were not checking everything yet
'We cannot reject a sentence if part1 is missing because
'we may output part sentences.
    If clsSentence.AisMsgPart1Missing = True Then
        Exit Function
    End If
'Here starts any AIS sentence checks
    If clsSentence.IsAisSentence = True Then
'because we check MMSI on the first part
        If TreeFilter.Check1(0).Value <> 0 Then
            If ChkSentenceFilter("shore", "") = False Then
                ProcessSentenceFilter = False
                Exit Function
            End If
        End If
        
'Range check is applied to both output channels - should really be individual on Output
'and as an input filter all channels - would require a separate tick box for input
        If RangeReq(0) = True Then
            If ChkSentenceFilter("latlonrange", "") = False Then
                ProcessSentenceFilter = False
                Exit Function
            End If
        End If
    End If
End Function

#If False Then
Sub EnumInputTree_old(n As Node, Level As Long, result As Boolean)
'Result is made true when first path to last node has passed
'all data checks, at this point no further checking is done
'also exits at the first node with DetailOut as the tag, this was
'originally done as fields were set on the input tree
'but the fields have now been moved to the output tree
'Passed node is the note that results in the tree being passed
Dim nC As Node
'PassedOK need keeping separately as it required passing back up the call stack
'so that when the last EnumTree returns Passed is returned to the first
'calling routine
Dim DataPassedOk As Boolean 'true if n.key test on data has passed

    If Level = 0 Then
        result = False    'clear first time
'        Set ExitNode = Nothing
    End If
    
    Level = Level + 1
    DataPassedOk = ChkInputFilter(n.Tag, n.Key)
    If n.Children = 0 And DataPassedOk Then 'end of branch
        result = True
'        Set ExitNode = n
    End If
 'Continue down this branch until no children, Data fails, or result
    If n.Children And DataPassedOk And Not result Then
        Set nC = n.Child
         Do
'trying to see if we can stop after input filter, This is the
'first child with DetailOut as the tag
#If False Then  'with separate output tree go to first leaf that is ok
            If nC.Tag = "DetailOut" Then
                If DataPassedOk Then result = True
'Restart at the child to get the output fields
'                Set ExitNode = n
                Exit Do
            End If
#End If
            Call EnumInputTree(nC, Level, result)
            If nC.Index = n.Child.LastSibling.Index Or result Then
                Exit Do 'exit this branch
            End If
            Set nC = nC.Next    'get next node at same level on this branch
        Loop
    End If
    Level = Level - 1
End Sub
#End If

'Remove level later
'NOTE this uses the NmeaRcv TreeFilter, which only contains
'the Items that are ticked on the Options TreeFilter
Function IsFilterOK(nR As Node) As Boolean
Dim tNode As Node
Dim DataPassedOk As Boolean 'true if n.key test on data has passed

'Debug.Print "Filter " & clsSentence.AisMsgType & "-" & clsSentence.AisMsgDac
'If clsSentence.AisMsgType = 8 Then Stop

'Should not happen
    If nR Is Nothing Then
        Exit Function
    End If

    IsFilterOK = False
'.Tag is the Filter that is being used, .key is the detail
'that the Filter will be using to check
    Set tNode = nR
    Do While Not tNode Is Nothing
'If tNode.Key = "FiList29" Then Stop
        DataPassedOk = ChkInputFilter(tNode.Tag, tNode.Key)
'Debug.Print tNode & " " & DataPassedOk
'Exit filter when the first node has no children and is OK
        If DataPassedOk = True And tNode.Children = 0 Then 'end of branch
                IsFilterOK = True
                Set tNode = Nothing     'exit from While loop
        Else
            If DataPassedOk = True Then
'Continue to the next sibling, or this sibling's child
'                Call NextNode(tNode)
            Else
'Go back to the next parent's sibling
'                Call NextBranch(tNode)
            End If
            Call NextValidNode(tNode, Not DataPassedOk)
        End If
    Loop
    
    Set tNode = Nothing
'Debug.Print "IsFilterOk " & IsFilterOK
'If clsSentence.AisMsgDac = "0" And IsFilterOK = True Then Stop
End Function

'returns next node ignoring any children (or grandchildren) of tNode if ThisNodeFailed=true
Private Sub NextValidNode(ByRef tNode As Node, Optional ThisNodeFailed As Boolean)
    On Error GoTo Node_err
        If ThisNodeFailed = False Then
            If tNode.Children Then  ' has child nodes so move to 1st child
                Set tNode = tNode.Child.FirstSibling
                Exit Sub
            End If
        End If
        
'If The current node failed then go to the next sibling, if none move back up the tree
        If tNode.Next Is Nothing Then 'No sibling, gotta move up level(s)
            Set tNode = tNode.Parent
            Do Until tNode Is Nothing   ' if Nothing, then done
                If Not tNode.Next Is Nothing Then
                    Set tNode = tNode.Next  'next sibling of parent (aunt/uncle) is valid
                    Exit Do                 'return
                End If
                If tNode.Parent Is Nothing Then 'No parent so at the root of the tree
                    Set tNode = Nothing         'exit no more nodes (whole tree traversed)
                    Exit Do
                End If
                Set tNode = tNode.Parent ' move up again
            Loop
        Else    'Has another sibling
            Set tNode = tNode.Next ' move to next sibling
        End If
Exit Sub

Node_err:
    Select Case err.Number
    Case Is = 35603     'invalid key
        err.Clear
    Case Is = 35605      'Control has been deleted
        err.Clear
    Case Else
        MsgBox "Error " & err.Number & vbCrLf & err.Description, , "CheckAll"
    End Select
    Resume Next

End Sub


'.next returns next sibling
Private Sub NextNode(ByRef tNode As Node)
        
        If tNode.Children Then  ' has child nodes so move to 1st child
            Set tNode = tNode.Child.FirstSibling
        ElseIf tNode.Next Is Nothing Then 'No sibling, gotta move up level(s)
            Set tNode = tNode.Parent
            Do Until tNode Is Nothing   ' if Nothing, then done
                If Not tNode.Next Is Nothing Then
                    Set tNode = tNode.Next  'next Child of parent
                    Exit Do
                End If
                If tNode.Parent Is Nothing Then 'No parent so at the root of the tree
                    Set tNode = Nothing         'exit from IsFilterOk
                    Exit Do
                End If
                Set tNode = tNode.Parent ' move up again
            Loop
        Else    'Has another sibling
            Set tNode = tNode.Next ' move to next sibling
        End If
End Sub

Private Sub NextBranch(ByRef tNode As Node)
        If tNode.Next Is Nothing Then ' gotta move up level(s)
            Set tNode = tNode.Parent
            Do Until tNode Is Nothing   ' if Nothing, then done
                If Not tNode.Next Is Nothing Then
                    Set tNode = tNode.Next  'next Child of parent
                    Exit Do
                End If
                If tNode.Parent Is Nothing Then 'No parent so set next node to nothing
                    Set tNode = Nothing         'Exit
                    Exit Do
                End If
                Set tNode = tNode.Parent ' move up again
            Loop
        Else
            Set tNode = tNode.Next ' move to next sibling
        End If

End Sub

'http://www.vbforums.com/showthread.php?t=473677
Public Sub QuickSort(ByRef pvarArray As Variant, Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)
    Dim lngFirst As Long
    Dim lngLast As Long
    Dim varMid As Variant
    Dim varSwap As Variant
    
    If plngRight = 0 Then
        plngLeft = LBound(pvarArray)
        plngRight = UBound(pvarArray)
    End If
    lngFirst = plngLeft
    lngLast = plngRight
    varMid = pvarArray((plngLeft + plngRight) \ 2)
    Do
        Do While pvarArray(lngFirst) < varMid And lngFirst < plngRight
            lngFirst = lngFirst + 1
        Loop
        Do While varMid < pvarArray(lngLast) And lngLast > plngLeft
            lngLast = lngLast - 1
        Loop
        If lngFirst <= lngLast Then
            varSwap = pvarArray(lngFirst)
            pvarArray(lngFirst) = pvarArray(lngLast)
            pvarArray(lngLast) = varSwap
            lngFirst = lngFirst + 1
            lngLast = lngLast - 1
        End If
    Loop Until lngFirst > lngLast
    If plngLeft < lngLast Then QuickSort pvarArray, plngLeft, lngLast
    If lngFirst < plngRight Then QuickSort pvarArray, lngFirst, plngRight
End Sub

'Called by DetailLineOut and TagsFromFields
'Gets the Msg key used to match the field records to the current sentence
Public Function FieldKeyFromSentence(CallingRoutine) As String
'Required to avoid type missmatch error if null string
'REmoved when CINT removed
'MsgBox clsField.CallingRoutine
    
'Because the CommentBlock can occur with any type of sentence
'We must not use the AisMsgType as part of the Field Key
If CallingRoutine = "" Then
MsgBox "Can't find calling routine " & CallingRoutine & " in FieldKeyFromSentence"
'    Stop
End If
'If clsField.CallingRoutine = "NmeaOut" Then Stop
'This affects the filtering
    With clsSentence
        Select Case CallingRoutine
'These are only used if an NMEA AisSentence (MyShiP requires AIVDO)
        Case Is = "SentenceOut"    'Any Sentence
            FieldKeyFromSentence = cKey("  $", 10)
        Case Is = "NmeaAisOut", "MyShipOut" 'Any AIS sentence
            FieldKeyFromSentence = cKey("  ais", 10)
        Case Is = "NmeaOut"                 'Specific NMEA sentence
            FieldKeyFromSentence = cKey("  " & .NmeaSentenceType, 10)
        Case Is = "CommentOut"              'NMEA Comment
            FieldKeyFromSentence = cKey("  \", 10)
        Case Else                           'Assume Specific AIS Sentence
            FieldKeyFromSentence = cKey(.AisMsgType, 2) _
            & cKey(.AisMsgDac, 4) _
            & cKey(.AisMsgFi, 2) _
            & cKey(.AisMsgFiId, 2)
        End Select
    End With
End Function

Public Function SplitFieldKey(ByVal FieldKey As String _
, Optional ByRef Msg As String _
, Optional ByRef Dac As String _
, Optional ByRef Fi As String _
, Optional ByRef Fiid As String _
) As Boolean
Dim Val As String

    If Len(FieldKey) <> 10 Then
        Exit Function
    End If
    If Left$(FieldKey, 2) <> "  " Then
'AIS key
        Val = Trim$(Left$(FieldKey, 2))
        If IsNumeric(Val) Then
            Msg = Val
            SplitFieldKey = True
        End If
        Dac = Trim$(Mid$(FieldKey, 3, 4))
        Fi = Trim$(Mid$(FieldKey, 7, 2))
        Fiid = Trim$(Mid$(FieldKey, 9, 2))
    Else
'NonAis key
        Msg = Trim$(FieldKey)
        SplitFieldKey = True
    End If

End Function
'Used only to display Message Type on Options Fields and Tags
Public Function DisplayFieldMsgType(MsgType As String, CallingRoutine As String) As String
    If NullToZero(MsgType) = 0 Then
        Select Case CallingRoutine
'These are only used if an NMEA AisSentence (MyShiP requires AIVDO)
        Case Is = "NmeaAisOut", "MyShipOut"
            DisplayFieldMsgType = "ais"
        Case Is = "NmeaOut", "SentenceOut"
            DisplayFieldMsgType = "nmea"
        Case Is = "CommentOut"
            DisplayFieldMsgType = "\"
        Case Else
            DisplayFieldMsgType = "??"   'Unkmown
        End Select
    Else
        DisplayFieldMsgType = MsgType 'Remove leading spaces
    End If

End Function

#If False Then
Public Function DisplayFieldKey(FieldKey As String, CallingRoutine As String) As String
Dim AisMsgType As String
Dim Dac As String
Dim Fi As String
Dim Fiid As String
Dim kb As String

    If Left$(FieldKey, 2) = "  " Then
       Call SplitFieldKey(FieldKey, AisMsgType, Dac, Fi, Fiid)
    
    If NullToZero(AisMsgType) = 0 Then
        Select Case CallingRoutine
'These are only used if an NMEA AisSentence (MyShiP requires AIVDO)
        Case Is = "NmeaAisOut", "MyShipOut"
            DisplayFieldKey = "ais"
        Case Is = "NmeaOut", "SentenceOut"
            DisplayFieldKey = "nmea"
        Case Is = "CommentOut"
            DisplayFieldKey = "\"
        Case Else
            DisplayFieldKey = "??"   'Unkmown
        End Select
    Else
        kb = AisMsgType
        If Dac <> "" Then
            kb = kb & " [" & Trim$(Dac)
            If Trim$(Fi) <> "" Then kb = kb & "-" & Trim$(Fi)
            If Trim$(Fiid) <> "" Then kb = kb & ":" & Trim$(Fiid)
            kb = kb & "]"
        End If
            Else
                AisMsgType = Trim$(FieldKey) & "_" & arry(4)
            End If
        DisplayFieldKey = FieldKey 'No leading spaces
    End If

End Function
#End If

'put the latest value for the tag into TagArray from FieldArray
'this can only be cleared when the mmsi changes.
'Saves IconTags
'When NmeaOut is called this is
'not known, so Nmeaout fields cannot be output to csv files
'Last forces an output, this is required to output the last scheduled vessel
'Requires clsSentence & Potentially could return value, if split
'so part of routine could be used to determine if we want to put sentence
'into the Scheduled buffer.

'true if weve found a field tag for this message
'this is because we may (if csv)only output the message if we have found
'a required tag in this message
Function TagsFromFields() As Boolean
Dim i As Long
Dim kb As String
'Dim Tag As clsTag
Dim RcvTime As String
Dim TagNo As Long   'base 0
Dim FieldKey As String
Dim MsgType As Long
Dim TagFound As Boolean

'Dim ThisShip As ShipDef
'if any tagged output or any range check extract the fields
'and do the range check (will be true if no ranges set)
    If FieldArray(1).MsgKey <> "" Then 'if only 2rows and first blank there are no entries
'can be done here because we already have decoded the sentence and know the mmsi in clssentence
'extract field values
'        For i = 1 To UBound(FieldArray)
'if nmea then msgtype = 0
        
'Always check MsgType 0 (Checks All of sentence except AisMsgTypes)
'Check any tags that apply to all NMEA sentences (eg received time)
'This is just a clever way of not overwriting tagfound once true
'While still calling set tag values

        TagFound = TagFound Or SetTagValues(0)
        
        If IsNumeric(clsSentence.AisMsgType) Then
'Check if Tag exists for AIS of any ais msg type (1-27) type
            MsgType = clsSentence.AisMsgType
'clssentence.vesselname may not have been loaded into clssentence
'if detail not being displayed
            If CacheVessels = True Then
                If clsSentence.VesselName = "" Then
                clsSentence.VesselName = GetVessel(clsSentence.AisMsgFromMmsi)
                End If
            End If
            TagFound = TagFound Or SetTagValues(MsgType)
        End If
        

'If there are Nmea messages with fields set up there will be
'some elements set up for message type 0
    End If
    
    If TagFound = True Then
        Call RestoreTagValues(clsSentence.AisMsgFromMmsi)
    End If
    
    TagsFromFields = TagFound
            
End Function

'Replace any blank value with saved value
Sub RestoreTagValues(Mmsi As String)
Dim myVessel As clsVessel
Dim TagNo As Long
Dim arTags() As String

Debug.Print "Restore:" & Mmsi
    On Error GoTo Key_NoVessel
    Set myVessel = Vessels(Mmsi)   'test if already in collection
'    arTags = Split(myVessel.strTags, "~")
    arTags = Split(myVessel.strTags, ",")
'Both arrays should be the same size
    If UBound(arTags) = UBound(TagArray, 1) Then
        For TagNo = 0 To UBound(arTags)
            If TagArray(TagNo, 1) = "" Then 'Dont overwrite new values with saved values
                TagArray(TagNo, 1) = arTags(TagNo)
            End If
        Next TagNo
    End If
    
Mmsi_Update:
Set myVessel = Nothing
Exit Sub
    
Key_NoVessel:                    'No saved value in MmsiTags collection
    Select Case err.Number
    Case Is = 5
        Resume Mmsi_Update
        err.Clear
    End Select

End Sub

'Sets the Value of the Tag from the value from the received sentence
'WRONG The MsgType is passed as a speed up so that only those
'WRONG Fields pertaining to a message type are processed.
'The MsgType must be valid
'If a message contains a comment block or is an NMEA sentence
'(rather than an AIS sentence) this must be called twice
'by TagsFromFields first to get the NMEA/Comment Values
'Then to get the AIS values for this AIS message type
Function SetTagValues(MsgType As Long) As Boolean
Dim i As Long
Dim TagNo As Long
Dim kb As String
Dim RcvTime As String
'Dim DefaultLatLon As Boolean    'true if 181,91
Dim MemberArry() As String
Dim FieldKey As String
Dim TagVal As String
Dim LatOK As Boolean    'Both lat and lon must be OK for Position to be OK
Dim LonOK As Boolean

'==== Start Ais Messages
 Debug.Print "MsgType=", MsgType
 Debug.Print "SetTag-" & FieldArray(1).Tag
        For i = FieldArrayFromTo(MsgType, 0) To FieldArrayFromTo(MsgType, 1)
'if the min elem in FieldArray is 0 then this msgtype
'has no field tags set up for it
'            If i = 0 Then Exit For
'Exit For   'debug speed
'If i = 3 Then Stop
'because treefilter.fieldlist key has been sorted (with ResetTags)
'we can exit as soon as a highet FieldKey (Ais msg no) is received
'AND it is an AIS message were processing
'kb = .TextMatrix(i, 0) & ":" & .TextMatrix(i, 2) 'debug
                        
            If MsgType > 0 Then
'AisSentences
                FieldKey = FieldKeyFromSentence("SetAisMessageTag")
'exit if higher key value
               If FieldArray(i).MsgKey > FieldKey Then
                    Exit For
                End If
            Else
'non-AIS Sentences
'Exit if reached a AisMsgType Key ("  1") or higher
                If Left$(FieldArray(i).MsgKey, 2) <> "  " Then
                    Exit For
                End If
kb = Trim$(FieldArray(i).MsgKey)
                
                Select Case Trim$(FieldArray(i).MsgKey)
                Case Is = "\"
'Output if sentence contains Comment block
                    If clsCb.Block <> "" Then
                        FieldKey = FieldArray(i).MsgKey
                    End If
                Case Is = "$"
'Output for every sentence
                    FieldKey = FieldArray(i).MsgKey
                Case Is = "ais"
'Output any Ais Specific fields
                    If clsSentence.IsAisSentence = True Then
                        FieldKey = FieldArray(i).MsgKey
                    End If
                Case Else
'Output any NMEA specific sentence fields
                    If Trim$(FieldArray(i).MsgKey) = clsSentence.NmeaSentenceType Then
                        FieldKey = FieldArray(i).MsgKey
                    End If
                End Select
            End If
'check if current message has a field we require the value of
'FieldKey is for Current Message
            If FieldKey = FieldArray(i).MsgKey Then
'                If FieldKey = "          " Then
'MemberArry = Split(FieldArray(i).Member, "_")
'                    If clsSentence.NmeaSentenceType <> MemberArry(0) Then
'                        GoTo skip
'                    Else
'                        FieldKey = FieldArray(i).Member
'Stop
'                    End If
'                End If
'recover latest tag for this field
'If clssentence.AisMsgType = "24" Then MsgBox FieldKey
                TagNo = -1
                Do
                    TagNo = TagNo + 1
'this should never happen as tag will not have been setup in the array
                    If TagNo > UBound(TagArray, 1) Then
                        kb = "FieldArray(" & i & ").MsgKey is """ & FieldArray(i).MsgKey & """" & vbCrLf _
                        & "Can't find Tag """ & FieldArray(i).Tag & """" & vbCrLf _
                        & " in TagArray" & vbCrLf
                        Call WriteErrorLog(kb & vbCrLf & clsSentence.NmeaSentence)
 '                       Stop
                    End If
 'If FieldArray(i).Source <> clsField.CallingRoutine Then Stop
                Loop Until TagArray(TagNo, 0) = FieldArray(i).Tag
'If Tagno = 14 And TagArray(14, 1) <> "" Then Stop 'debug Destination only 6 characters
'If FieldArray(i).Tag = "destination" Then Stop
'this should never happen as tag will not have been setup in the array
                If TagNo > UBound(TagArray, 1) Then
                    kb = "FieldArray(" & i & ").MsgKey is """ & FieldArray(i).MsgKey & """" & vbCrLf _
                    & "Can't find Tag """ & FieldArray(i).Tag & """" & vbCrLf _
                    & " in TagArray" & vbCrLf _
                    & "Possibly arry(2) in Flexgrid in .ini file" & vbCrLf _
                    & "Not found" & vbCrLf
                    Call WriteErrorLog(kb & vbCrLf & clsSentence.NmeaSentence)
'Stop
                End If
'kb = TagArray(Tagno, 0) & ":" & .TextMatrix(i, 12)
'if time "" then subcript error, if same time force update
'If clsSentence.AisMsgFromMmsi = "000000172" Then Stop
                If Not IsDate(clsSentence.NmeaRcvTime) Then clsSentence.NmeaRcvTime = NowUtc()   'locale
                If IsDate(TagArray(TagNo, 2)) Then
                    RcvTime = TagArray(TagNo, 2)
                Else
                    RcvTime = clsSentence.NmeaRcvTime
                End If

'If this sentence received later than current Tag Time received
'If existing tav value is present don't overwrite it
'because the same tag may be used on more then one part of the same
'sentence (eg comment and AisWord 7 bith get unixtime
'Update the value of this tag
'kb = DateDiff("s", TagArray(Tagno, 2), clsSentence.NmeaRcvTime)
                If DateDiff("s", RcvTime, clsSentence.NmeaRcvTime) >= 0 _
                And TagArray(TagNo, 1) = "" Then
                    Select Case FieldArray(i).Source
                    Case Is = "DetailOut"
'If .TextMatrix(i, 13) = "Ship Type" Then Stop
'If FieldArray(i).Member = "destination" Then Stop
                        TagArray(TagNo, 1) = Detail.DetailOut( _
                        FieldArray(i).Column, _
                        FieldArray(i).Valdes, _
                        FieldArray(i).Member, _
                        FieldArray(i).From, _
                        FieldArray(i).reqbits, _
                        FieldArray(i).Arg, _
                        FieldArray(i).Arg1)
                        If GisReq(0) Then
                            If TagArray(TagNo, 1) <> "" _
                            And FieldArray(i).Column = 2 Then
'user cannot set member Select on the member not the tag
'and all the GISReq values must be column 2 as user may also have selected
'the description column
                                Select Case FieldArray(i).Member
                                Case Is = "lat"
'may be more than one lat/lon - Maggie M problem - position not set up on ais
'                                    If TagArray(TagNo, 1) = "91.000000" Then DefaultLatLon = True
'Stop
'                                    LatOK = True
                                    If IsNumeric(TagArray(TagNo, 1)) Then
                                        If CSng(TagArray(TagNo, 1)) <> 91 Then
                                            LatOK = True
                                        End If
                                    End If
                                Case Is = "lon"
'                                    If TagArray(TagNo, 1) = "181.000000" Then DefaultLatLon = True
'Stop
'                                    LonOK = True
                                    If IsNumeric(TagArray(TagNo, 1)) Then
                                        If CSng(TagArray(TagNo, 1)) <> 181 Then
                                            LonOK = True
                                        End If
                                    End If
                                Case Is = "heading"
'                                    If TagArray(TagNo, 1) <> "511" Then IconHeading = TagArray(TagNo, 1)
                                    If TagArray(TagNo, 1) <> "511" Then Gis.Heading = TagArray(TagNo, 1)
                                Case Is = "course"
'                                    If TagArray(TagNo, 1) <> "511" Then IconCourse = TagArray(TagNo, 1)
                                    If TagArray(TagNo, 1) <> "511" Then Gis.Course = TagArray(TagNo, 1)
                                Case Is = "clength"
                                    If TagArray(TagNo, 1) <> "0" Then
'If clssentence.AisMsgType = "24" Then Stop
'                                        IconScale = 1 + (CLng(TagArray(TagNo, 1)) - 100) / 200
                                        Gis.IScale = 1 + (CLng(TagArray(TagNo, 1)) - 100) / 200
                                    End If
                                Case Is = "ship_type"
'Color order is BGR html is RGB
        Select Case Int(CInt(TagArray(TagNo, 1)) / 10)
        Case Is = 2, 4
'            IconColor = "ff0000ff"  'yellow hsc
            Gis.Color = "ff0000ff"  'yellow hsc
        Case Is = 3
'            IconColor = "ffff00ff"  'Purple YAcht+Other
            Gis.Color = "ffff00ff"  'Purple YAcht+Other
        Case Is = 5
'            IconColor = "ffffff00"  'Cyan Tug,Pilot
            Gis.Color = "ffffff00"  'Cyan Tug,Pilot
        Case Is = 6
'            IconColor = "ffff0000"  'Blue Passenger
            Gis.Color = "ffff0000"  'Blue Passenger
        Case Is = 7, 9
'            IconColor = "ff00ff00"  'Green Cargo
            Gis.Color = "ff00ff00"  'Green Cargo
        Case Is = 8
'            IconColor = "ff0000ff"  'Tanker Red
            Gis.Color = "ff0000ff"  'Tanker Red
        Case Else
'            IconColor = "ffd3d3d3"  'Grey Other
            Gis.Color = "ffd3d3d3"  'Grey Other
        End Select
                                
                                End Select
                            End If
                        End If
                    Case Is = "MyShipOut"
                        TagArray(TagNo, 1) = Detail.MyShipOut( _
                        FieldArray(i).Column, _
                        FieldArray(i).Valdes, _
                        FieldArray(i).Member)
                  '      FieldArray(i).From, _
                  '      FieldArray(i).reqbits, _
                  '      FieldArray(i).Arg, _
                  '      FieldArray(i).Arg1)
                    Case Is = "NmeaOut"
                        TagArray(TagNo, 1) = Detail.NmeaOut( _
                        CLng(FieldArray(i).Column), _
                        FieldArray(i).Tag, _
                        FieldArray(i).Member, _
                        FieldArray(i).From)
                    Case Is = "SentenceOut"
                        TagArray(TagNo, 1) = Detail.SentenceOut( _
                        CLng(FieldArray(i).Column), _
                        FieldArray(i).Tag, _
                        FieldArray(i).Member, _
                        FieldArray(i).From)
                    Case Is = "NmeaAisOut"
                        TagArray(TagNo, 1) = Detail.NmeaAisOut( _
                        CLng(FieldArray(i).Column), _
                        FieldArray(i).Tag, _
                        FieldArray(i).Member, _
                        FieldArray(i).From)
                    Case Is = "Dac1Out"
                        TagArray(TagNo, 1) = Detail.Dac1Out( _
                        FieldArray(i).Column, _
                        FieldArray(i).Valdes, _
                        FieldArray(i).Member, _
                        FieldArray(i).From, _
                        FieldArray(i).reqbits, _
                        FieldArray(i).Arg, _
                        FieldArray(i).Arg1)
                    Case Is = "Dac200Out"
                        TagArray(TagNo, 1) = Detail.Dac200Out( _
                        FieldArray(i).Column, _
                        FieldArray(i).Valdes, _
                        FieldArray(i).Member, _
                        FieldArray(i).From, _
                        FieldArray(i).reqbits, _
                        FieldArray(i).Arg, _
                        FieldArray(i).Arg1)
                    Case Is = "Dac235Out"
                        TagArray(TagNo, 1) = Detail.Dac235Out( _
                        FieldArray(i).Column, _
                        FieldArray(i).Valdes, _
                        FieldArray(i).Member, _
                        FieldArray(i).From, _
                        FieldArray(i).reqbits, _
                        FieldArray(i).Arg, _
                        FieldArray(i).Arg1)
                    Case Is = "Dac366Out"
                        TagArray(TagNo, 1) = Detail.Dac366Out( _
                        FieldArray(i).Column, _
                        FieldArray(i).Valdes, _
                        FieldArray(i).Member, _
                        FieldArray(i).From, _
                        FieldArray(i).reqbits, _
                        FieldArray(i).Arg, _
                        FieldArray(i).Arg1)
                    Case Is = "CommOut"
                        TagArray(TagNo, 1) = Detail.CommOut( _
                        FieldArray(i).Column, _
                        FieldArray(i).Valdes, _
                        FieldArray(i).Member)
                    Case Is = "CommentOut"
                        TagArray(TagNo, 1) = Detail.CommentOut( _
                        FieldArray(i).Column, _
                        FieldArray(i).Valdes, _
                        FieldArray(i).Member)
                    Case Else
                        MsgBox FieldArray(i).Source & " Requires setting up in TagsFromFields"
                    End Select
'ignore blank value records
                    If TagArray(TagNo, 1) <> "" Then
                        TagArray(TagNo, 2) = clsSentence.NmeaRcvTime
                    End If
'if here weve found an output tag
'                    If Not (clsSentence.VesselLat = 91 And clsSentence.VesselLon = 181) Then
'                        LatOk = True
'                        LonOK = True
'                    End If
                    SetTagValues = True
'                OutputTagFound = True (not in use)
'Reformat any specific output fields
                    Select Case TagArray(TagNo, 0)
'Now uses vessellat,vessellon lat/lon below is for compatibility
                    Case Is = "nmea_lat"
                        If IsNumeric(clsSentence.VesselLat) Then
'kb = LatToNmea(CSng(-5.684358))
                             TagArray(TagNo, 1) = LatToNmea(CSng(clsSentence.VesselLat))   'DDMM.MM,N or S
                        End If
                    Case Is = "nmea_lon"
'kb = LonToNmea(CSng(-170.4225))
                        If IsNumeric(clsSentence.VesselLat) Then
                             TagArray(TagNo, 1) = LonToNmea(CSng(clsSentence.VesselLon))   'DDMM.MM,N or S
                        End If
                    Case Is = "nmea_date"
                        TagArray(TagNo, 1) = Format$(TagArray(TagNo, 1), "ddmmyy")
                    Case Is = "nmea_time"
                        TagArray(TagNo, 1) = Format$(TagArray(TagNo, 1), "hhnnss")
                    Case Is = "range"
'Stop
                    End Select
                Else
'Stop
                End If
            End If
skip:
        Next i  'no tags found (next output sentence for this msg type)
'    IconName = clsSentence.VesselName
    Gis.Name = clsSentence.VesselName
'Both the Lat and Lon tag values must be OK for the position to be valid
    If LatOK = True And LonOK = True Then
        Gis.PositionOK = True
    End If
'at this point all TagArray(Tagno) has been loaded
'with the tag values that are to be output for the current message
'    If DefaultLatLon Then
'        LatOk = False
'        LonOK = False
'    End If
 'Debug.Print "exitSet-" & FieldArray(1).Tag

End Function

'Called by OutputAll
'Check if any tags have a range set, if they have and range not passed
'returns false.
'TagArray(,cols) Cols 0=Tag,1=Value,2=RcvTime,3=Min,4=Max,5-Name(for csvhead
Function RangeCheck() As Boolean
Dim arry() As String
Dim TagNo As Long

'scan TagArray to see if there is anything to output
'Exit Function    'debug speed
    RangeCheck = True
    If (Not TagArray) = True Then  'there are no tags to output (Scheduled output should be stopped ? )
        Exit Function
    End If
        
    For TagNo = 0 To UBound(TagArray, 1)    'value
        
'Debug RangeCheck
'        If TagArray(TagNo, 0) = "mmsi" Then
'            If TagArray(TagNo, 1) <> clsSentence.AisMsgFromMmsi Then
'                MsgBox "Range and Sentence MMSI's differ", , "RangeCheck"
'            End If
'        End If

'if no message received for this tag assume true, therefore
'range will be assumed to pass
        If TagArray(TagNo, 2) <> "" Then    'RcvTime
'Exit For   'debug speed
'check if value is within range to output, if any out of range then don't output
'Check Min
            If TagArray(TagNo, 3) <> "" Then    'Min
                If IsNumeric(TagArray(TagNo, 1)) = True Then    'Value is numeric
'lat/lon is always dot, force lat/lon & range to locale
'because Csng is locale aware
                    If CSng(Replace(TagArray(TagNo, 1), ".", DecimalSeparator)) < CSng(Replace(TagArray(TagNo, 3), ".", DecimalSeparator)) Then
                        RangeCheck = False
                        Exit For
                    End If
                Else            'string tag
                    If TagArray(TagNo, 1) <> TagArray(TagNo, 3) Then    'Value <> Min
                        RangeCheck = False
                        Exit For
                    End If
                End If
            End If
'Check max
            If TagArray(TagNo, 4) <> "" Then    'Max
                If IsNumeric(TagArray(TagNo, 1)) = True Then    'Value is numeric
'If CSng(TagArray(Tagno, 1)) < 30 Then Stop
                    If CSng(Replace(TagArray(TagNo, 1), ".", DecimalSeparator)) > CSng(Replace(TagArray(TagNo, 4), ".", DecimalSeparator)) Then
                        RangeCheck = False
                        Exit For
                    End If
                Else            'String
                    If TagArray(TagNo, 1) <> TagArray(TagNo, 4) Then    'Value <> Max
                        RangeCheck = False
                        Exit For
                    End If
                End If
            End If
'These have been reinserted to block output if a blank tag
'if a range has been set 23/11 Vesseltype no in range passes
'if message 5 not yet received
'This can only be done if not outputting all messages
        
        Else    'No message has been received containing this Tag
'Min or Max range has been set for this Tag
            If (TagArray(TagNo, 3) <> "" Or TagArray(TagNo, 4) <> "") Then 'Min or Max not blank
'Output on mmsi change
                If TreeFilter.Check1(16).Value = vbChecked Then 'was <>  0  'output on MMSI change
                    RangeCheck = False
                    Exit For
                End If
            End If
        End If
    Next TagNo

End Function

'If not Scheduled, called for each complete sentence
'If Scheduled called for each MMSI
'If NmeaOutput indivisual sentences Output
'Else Tagged or CSV for each MMSI
'outputs nmea,csv or tagged to both channels
'If Scheduled=true then only if called from SchedulerOutput
'else only if called from ProcessSentence (not Scheduled)
Sub OutputAll(ByVal Scheduled As Boolean)
Dim i As Long
Dim GisOk As Boolean
Dim RangeOk As Boolean
Dim NmeaOutputted As Boolean
Dim OutputOK As Boolean
Dim OutputRejected As Boolean
Dim Channel As Long
Dim TagNo As Long
'Debug RangeCheck
'If TagArray(0, 1) <> clsSentence.AisMsgFromMmsi Then Stop
'If TagArray(3, 1) <> CStr(clsSentence.VesselLat) Then Stop

'Debug.Print "OutputAll-start" & Outputted
    For Channel = 1 To 2    'reset if channel in use
        If ChannelOutput(Channel) = True Then
            OutputOK = True
'Check MaxMin values on all tags
'I THINK CSVALL may be unecessary
'    If RangeReq(Channel) = True And CsvAll(Channel) = False Then
            If RangeReq(Channel) = True Then
                If RangeCheck = False Then  'V123 changed 3/4/14 because range being applied even if range required not set for
'one channel only
                    OutputOK = False
                Else
'Debug Broos range
If TagArray(0, 1) <> clsSentence.AisMsgFromMmsi Then
'Stop
End If
                End If
            End If
                        
'If clsSentence.AisMsgType = 5 Then Stop
                                               
            If OutputOK = True Then
'This validates output if Nmea being output
                If GisReq(Channel) = True Then
'This checks the current sentence is valid for gis output
'note tag range is also checked as valid
                    If ChkSentenceFilter("gis", "") = False Then
                        OutputOK = False
                    End If
                End If

'Tagged output we need to check the Tag Positions are valid
'This validates output if tags are being used to output data
'We dont output any static data unless we have a position in this batch of data
'THIS IS CHECKED by AisPositionOK (I think)
'                If TagsReq(Channel) = True Then
'                    If Gis.PositionOK = False Then
'                        OutputOK = False
'                    End If
'                End If
            End If
 
            If OutputOK = True Then
'Expression too complex !!!
                If TaggedOutput(Channel) = True Then
                    Call OutputTagged(Channel)
                End If
                
                If CsvOutput(Channel) = True Then Call OutputCsv(Channel)
'Nmea and csvAll require individual sentences outputting
'clsSentence is not NmeaArray(i)
                If NmeaOutput(Channel) = True Then
                    For i = 0 To NmeaOutBufCount - 1
'Debug.Print "NmeaBufOut(" & i & ")" & Left$(NmeaArray(i), 15)
                        Call OutputNmea(Channel, NmeaOutBuf(i))
                    Next i
                End If
                
                If CsvAll(Channel) = True Then
                    For i = 0 To NmeaOutBufCount - 1
'Debug.Print "OutputCsvAll(" & i & ")" & Left$(NmeaArray(i), 15)
'Only output when the last part is reached
                        If i = NmeaOutBufCount - 1 Then
                            Detail.SentenceAndPayloadDetail (4)
                            Call OutputCsvAll(Channel)
'Test to see which nmeasentence is being output
'                Call OutputNmea(Channel, NmeaOutBuf(i))
                        End If
                    Next i
                End If
                
'NmeaOutBufCount is the no of parts in the message + 1
                Outputted = Outputted + 1
'Debug.Print "OutputAll " & Filtered & " " & Outputted

'                If ScheduledReq(Channel) Then SchedOut = SchedOut + 1
            Else
'Stop
'Call WriteErrorLog(kb)
'MsgBox clssentence.nmeasentence 'debug range rejected sentences only
            End If
        End If  'No output on this channel
    Next Channel
    If Scheduled Then
        SchedOut = SchedOut + 1
    End If
    
'Save Tag Values if range OK
'    If OutputOK = True Then (doesnt save static data if range fails
'    End If
    
'Some sentences eg !AITXT do not have a MMSI
    If IsNumeric(clsSentence.AisMsgFromMmsi) Then
       Call SaveMmsiTags(clsSentence.AisMsgFromMmsi)
    Else
'eg !AITXT
'Debug.Print "Non numeric MMSI"
    End If

'must output all channels before we clear the Range & buffer
    Call ClearTagValues
    Call NmeaOutBufClear

End Sub

Sub SaveMmsiTags(Mmsi As String)
Dim TagNo As Long
Dim arTags() As String
Dim myVessel As clsVessel       'V142 Vessel changed to myVessel
Dim Tags As String

On Error GoTo Key_NewVessel
Set myVessel = Vessels(Mmsi)  'see if this ship is in collection (err=5 if not)
                             'if not create new ship in ships collection
'v146 removed. Jason error 457 requires error trap to remain because if 457 returned when in
'Key_NewVessel, AddVessel must be exited
'On Error GoTo 0         'V142

Vessel_Update:
    ReDim arTags(UBound(TagArray, 1))
    For TagNo = 0 To UBound(TagArray, 1)
        arTags(TagNo) = TagArray(TagNo, 1)
    Next TagNo
    Tags = Join(arTags, ",")
    If Tags <> "" Then      'No tags
        myVessel.strTags = Tags
    End If
    Set myVessel = Nothing
Exit Sub
    
Key_NewVessel:                    'create new ship in ships collection
    Select Case err.Number
    Case Is = 5
            Set myVessel = New clsVessel  'even
            myVessel.Mmsi = Mmsi
            Vessels.Add myVessel, Mmsi
            Resume Vessel_Update
        err.Clear
    End Select

End Sub

'Only used to test spoofing ownship
'Return original sentence if any problem
Function VDMtoVDO(Sentence As String) As String
Dim i As Long   'Position of *
Dim VDO As String
Dim crc As String
    i = InStr(1, Sentence, "*")
    If i > 0 Then
        crc = Mid$(Sentence, i, 3)
        VDO = Replace(Sentence, "!AIVDM", "!AIVDO")
        VDO = Replace(VDO, crc, "*[CRC]")
        VDO = Replace(VDO, "[CRC]", CalculateCrc(VDO))
        VDMtoVDO = VDO
    Else
        VDMtoVDO = Sentence
    End If
End Function

'All channels must be output before ClearTagValues is called
'It must be used whenever the MMSI has changed after DecodeSentence in order
'to Reset the NmeaOutBuf
Sub ClearTagValues()
Dim TagNo As Long
'clear tag values/rcvtimes and FEN buffer
'the check below doesnt work properly
'    If (Not TagArray) = False Then 'tags to clear !
    For TagNo = 0 To UBound(TagArray, 1)
        TagArray(TagNo, 1) = ""
        TagArray(TagNo, 2) = ""
    Next TagNo
'    End If
'v123
'If FenReq(0) = True Then
'Debug.Print "ClearFenR"
'    Erase NmeaArray
'    NmeaArrayNo = 0
'End If
'LatOk = False
'LonOK = False
'IconHeading = ""
'IconCourse = ""
'IconScale = "1"
'IconColor = "ffffffff"
'IconName = ""
Gis.Color = "ffffffff"
Gis.Course = ""
Gis.Heading = ""
Gis.IScale = "1"
Gis.Name = ""
Gis.PositionOK = False
End Sub

Sub NmeaOutBufClear()
Dim i As Long

'Debug.Print "NmeaOutBufClear"
    If UBound(NmeaOutBuf) > MINNMEAOUTBUFSIZE Then
        ReDim NmeaOutBuf(MINNMEAOUTBUFSIZE)
    End If
    For i = 0 To NmeaOutBufCount - 1
        NmeaOutBuf(i) = ""
    Next i
    NmeaOutBufCount = 0
End Sub

Sub OutputUdp(Channel As Long, kb As String) 'this can be filtered or tagged
Dim i As Long
Dim j As Long

'if closed try and open it
If MethodOutput(Channel) = True Then
    With NmeaRcv.ServerUDP
        If .State = sckClosed Then Call NmeaRcv.OpenOutputUdp(Channel)
        If .State <> sckOpen Then GoTo Udp_Error
        On Error GoTo Udp_Error
        
        If TaggedOutput(Channel) = True Then
            If TaggedOutputOn(Channel) = False Then  'head not yet written out
                TaggedOutputOn(Channel) = True
                Call WriteTagHead(Channel)
            End If
        End If
        
'MsgBox kb
'        i = 1
'        Do
'            j = InStr(i, kb, vbCrLf)
'May not be a crlf at the end of the last line
'Check we do not have a null length sentence or not termination crlf
'            If j > i Then
'                .SendData Mid$(kb, i, j + i)
'            Else
'                .SendData Mid$(kb, i) & vbCrLf
'            End If
'            i = j + 2       'Skip crlf
'        Loop Until j = 0  'end of all sentences
        
        .SendData kb & vbCrLf
        NoDataOutput(Channel) = False
        On Error GoTo 0
    End With
End If

If DisplayOutput(Channel) = True Then
    Call Output.OutputDisplay(Channel, kb)
End If
Exit Sub

Udp_Error:
    NmeaRcv.ServerUDP.Close
'    TreeFilter.Option1(7).Value = True
    MsgBox "The following error has occurred:" & vbNewLine _
         & "Err # " & err.Number & " - " & err.Description _
         & vbCrLf & " Connection has been closed", _
           vbOK, "UDP Server Error"
    
End Sub

'converts key into consistent format
'AIS key is 2 characters msg no " 1" to "27"
'NMEA key is "  $GPxxx   " total 10 characters
Public Function cKey(Key As String, KeyLen As Long) As String
Dim intg As Long
If IsNumeric(Left$(Key, KeyLen)) Then
'fill right number to size of keylen ie 1 becomes space + 1 if keylen =2
        intg = Left$(Key, KeyLen)
        cKey = Format$(intg, Left$(MYFMT, KeyLen))
Else
        If Key <> "" Then
'! Force left to right fill of placeholders
'The default is to fill from right to left.
'"  $GPZDA  "   (2 spaces + key the spaces to 10 characters
            cKey = Format$(Key, "!" & Left$(MYFMT, KeyLen))
        Else
'fill with keylen spaces
            cKey = Space$(KeyLen)
        End If
End If
End Function

'kb is one nmea sentence, with date stamp
'format & send nmea to either File or UDP, Output.OutputDisplay is handled by the
'called routine
Sub OutputNmea(Channel As Long, kb As String)

'create raw sentence if required
If TimeStampReq(Channel) = False Then
    If IsDate(Mid$(kb, InStrRev(kb, ",") + 1)) Then
        kb = Left$(kb, InStrRev(kb, ",") - 1)
    End If
End If

If ChannelMethod(Channel) = "file" Then Call OutputToDataFile(Channel, kb)
If ChannelMethod(Channel) = "udp" Then Call OutputUdp(Channel, kb)

End Sub

'format & send csv to either File or UDP, Output.OutputDisplay is handled by the
'called routine
Sub OutputCsv(Channel As Long)
Dim TagNo As Long
Dim i As Long
Dim arry() As String
Dim kbCsv As String
Dim NoFieldValues As Boolean
Dim FieldValueExists As Boolean

'create csv head if not existing (there must be tags for csv output)
If CsvHeadOn(Channel) = True And CsvHead(Channel) = "" Then
ReDim arry(UBound(TagArray, 1))
    For TagNo = 0 To UBound(TagArray, 1)
        arry(TagNo) = QuotedString(TagArray(TagNo, 5), CsvDelim(Channel))
        If arry(TagNo) <> "" Then FieldValueExists = True
    Next TagNo
'V101 prefix header with ~ to enable it to be parsed out
'into pairs by NmeaRouter.Formatter (AIVDO)
    kbCsv = "~" & Join(arry, CsvDelim(Channel))
    CsvHead(Channel) = kbCsv
End If
    
If FieldValueExists = True Then
    If ChannelOutput(Channel) = True And ChannelFormat(Channel) = "csv" Then 'File
        If ChannelMethod(Channel) = "file" Then Call OutputToDataFile(Channel, kbCsv)
        If ChannelMethod(Channel) = "udp" Then Call OutputUdp(Channel, kbCsv)
    End If
End If

'output the data
ReDim arry(UBound(TagArray, 1))
If CsvDelim(Channel) <> "" Then   'all channel delimeters are the same
    For TagNo = 0 To UBound(TagArray, 1)
 'Debug.Print "<" &  TagArray(tagno,0) & ">" & TagArray(tagno,1)
        arry(TagNo) = QuotedString(TagArray(TagNo, 1), CsvDelim(Channel))
        If arry(TagNo) <> "" Then FieldValueExists = True
    Next TagNo
    kbCsv = Join(arry, CsvDelim(Channel))
End If
    
If FieldValueExists = True Then
    If ChannelOutput(Channel) = True And ChannelFormat(Channel) = "csv" Then 'File
        If ChannelMethod(Channel) = "file" Then Call OutputToDataFile(Channel, kbCsv)
        If ChannelMethod(Channel) = "udp" Then Call OutputUdp(Channel, kbCsv)
    End If
End If
'Next Channel
End Sub

'format & send csv to either File or UDP, Output.OutputDisplay is handled by the
'called routine
Sub OutputCsvAll(Channel As Long)
Dim i As Long
Dim kbCsv As String
    
If AllFieldsNo > 0 Then
    ReDim Preserve AllFields(AllFieldsNo - 1) 'remove redundant fields
    If CsvDelim(Channel) <> "" Then   'all channel delimeters are the same
'dont (V125 ?) skip time and NMEA sentence
        
        For i = 0 To AllFieldsNo - 1 'UBound(AllFields)
 'Debug.Print "<" &  TagArray(tagno,0) & ">" & TagArray(tagno,1)
            AllFields(i) = QuotedString(AllFields(i), CsvDelim(Channel))
        Next i
        kbCsv = Join(AllFields, CsvDelim(Channel))
    End If
    
    If ChannelOutput(Channel) = True And ChannelFormat(Channel) = "csv" Then 'File
        If ChannelMethod(Channel) = "file" Then Call OutputToDataFile(Channel, kbCsv)
        If ChannelMethod(Channel) = "udp" Then Call OutputUdp(Channel, kbCsv)
    End If
End If
End Sub

'format & send tagged to either File or UDP, Output.OutputDisplay is handled by the
'called routine
Sub OutputTagged(Channel As Long)
Dim TagNo As Long
Dim kbtag As String
Dim Buf As String
Dim Line As String
Dim i As Long
Dim j As Long
Dim arry() As String
Dim IsTagValue As Boolean

kbtag = TagTemplateContent(Channel)
'Blank tag file
If kbtag = "" Then Exit Sub
'MsgBox kbtag

'replace any of the GIS tags that have the latest value received
'that have been extracted from the current TagList
If GisReq(Channel) = True Then
    If Gis.Heading = "" Or Gis.Heading = "511" Then
        If Gis.Course = "" Or Gis.Course = "511" Then
            Gis.Heading = "511"
        Else
            Gis.Heading = CInt(Gis.Course)
        End If
    End If
'    If IconHeading = "" Or IconHeading = "511" Then
'        If IconCourse = "" Or IconCourse = "511" Then
'            IconHeading = "511"
'        Else
'            IconHeading = CInt(IconCourse)
'        End If
'    End If
'    kbtag = Replace(kbtag, TagChr(0) & "IconHeading" & TagChr(1), IconHeading)
'    kbtag = Replace(kbtag, TagChr(0) & "IconScale" & TagChr(1), IconScale)
'    kbtag = Replace(kbtag, TagChr(0) & "IconColor" & TagChr(1), IconColor)
    kbtag = Replace(kbtag, TagChr(0) & "IconHeading" & TagChr(1), Gis.Heading)
    kbtag = Replace(kbtag, TagChr(0) & "IconScale" & TagChr(1), Gis.IScale)
    kbtag = Replace(kbtag, TagChr(0) & "IconColor" & TagChr(1), Gis.Color)
    
End If
    
'any permanent cached data for the MMSI being output
'must have been placed in TagArray
'clssentence is NOT the one being output

'Dim kb As String
'Dim TagMmsi As String
'Dim TagCachedName As String
'Dim TagName As String
'Dim TagHeading As String   'heading if n/a then COG
'Dim TagCourse As String   'last COG
'Dim TagScale As String     '1+-10% for each meter +-100
'Dim TagColor As String
'now replace all the tags that have been collected by the scheduler

'Do Until i * 1023 > Len(kbtag)
'    MsgBox Mid$(kbtag, i * 1023 + 1, (i + 1) * 1023)
'    i = i + 1
'Loop

'If TagArray(14, 1) <> "" Then Stop
For TagNo = 0 To UBound(TagArray, 1)
'v110 check is tag is actually used in THIS template
    If InStr(1, kbtag, TagChr(0) & TagArray(TagNo, 0) & TagChr(1)) > 0 Then
        kbtag = Replace(kbtag, TagChr(0) & TagArray(TagNo, 0) & TagChr(1), AsciiToXml(Channel, TagArray(TagNo, 1)))
'If TagArray(Tagno, 0) = "CachedVesselName" Then TagCachedName = TagArray(Tagno, 1)
'If TagArray(Tagno, 5) = "MMSI" Then TagMmsi = TagArray(Tagno, 1)
'If TagArray(Tagno, 5) = "Vessel Name" Then TagName = TagArray(Tagno, 1)
'v110
        If TagArray(TagNo, 1) <> "" Then
            IsTagValue = True
        End If
    End If
Next TagNo
'MsgBox "Output Tagged" & vbCrLf _
'& "TagName=" & TagName & " " & TagMmsi & vbCrLf _
'& "TagCachedName=" & TagCachedName & vbCrLf _
'& "IconName = " & IconName & vbCrLf _
'& "clsSentenceName=" & clssentence.VesselName & " " & clssentence.AisMsgFromMmsi & vbCrLf

'If TagMmsi <> "" Then
'kb = clssentence.VesselName
'    If GetVessel(TagMmsi) <> TagName Then Stop
'    If GetVessel(TagMmsi) <> IconName Then Stop
'End If

'If no tags have been replaced with a value then don't output anything
'This will happen with incomplete messages or if Tags have
'been set up but are not used in the CURRENT template
'v110
If IsTagValue = False Then
    Exit Sub
End If
If kbtag <> "" Then
    
'For VDO to GPHDT & GPRMC
'Split kbtag into individual lines and replace crc
    If InStr(1, kbtag, "[CRC]") > 0 Then
        i = 1
        Do
            j = InStr(i, kbtag, vbCrLf)
'May not be a crlf at the end of the last line
'            If j = 0 And Len(kbtag) > i Then j = Len(kbtag) + 2
'kb = Right$(kbtag, 2)
'Check we do not have a null length sentence or not termination crlf
                If j > i Then
                    Line = Mid$(kbtag, i, j + i)
                Else
                    Line = Mid$(kbtag, i) ' & vbCrLf
                End If
'Expression too complex !!!
                Line = Replace(Line, "[CRC]", CalculateCrc(Line))
                Buf = Buf + Line
                i = j + 2       'Skip crlf
'                End If          'end of this sentence
        Loop Until j = 0  'end of all sentences
'MsgBox kbtag & ":" & buf
        kbtag = Buf
    End If
    
'    If TaggedOutputOn(Channel) = False Then Call OpenTaggedOutput(Channel)
    If ChannelMethod(Channel) = "file" Then
        Call OutputToDataFile(Channel, kbtag)
'this is for testing the overlay
'Dim Heading As Long
'        If OverlayOutputFileCh <> 0 Then Call OutputOverlay("NorthUp")
'        If OverlayOutputFileCh <> 0 Then Call OutputOverlay("Rotate")
'        If OverlayOutputFileCh <> 0 Then
'            Do Until Heading > 90
'                Call OutputOverlay("Move", _
'                CStr(Heading), CStr(to_bow), CStr(to_stern), _
'                CStr(to_port), CStr(to_starboard))
'            Heading = Heading + 90
'            Loop
'        End If
        If OverlayOutputFileCh <> 0 Then
'            Call OutputOverlay("", "0", "2", "10", "5", "1")  'ne
'            Call OutputOverlay("", "0", "4", "8", "8", "4")  'ne square
'            Call OutputOverlay("", "0", "6", "6", "6", "6")  'centre
'            Call OutputOverlay("", "0", "10", "2", "5", "1")  'se
'            Call OutputOverlay("", "0", "10", "2", "1", "5")  'sw
'            Call OutputOverlay("", "0", "2", "10", "1", "5")  'nw
            Call OutputOverlay(Channel)
'            Call OutputOverlay("", "0")
        End If
    End If
    If ChannelMethod(Channel) = "udp" Then Call OutputUdp(Channel, kbtag)
End If

End Sub

'channel required for xml coding
Sub OutputOverlay(Channel As Long, Optional Arg As String, _
Optional xHeading As String, _
Optional xto_bow As String, _
Optional xto_stern As String, _
Optional xto_port As String, _
Optional xto_starboard As String)
'as per offset.xls + constant M_to_deg
Dim Offset As Single    '0-90 because its tan
Dim Quadrant As Single  '0,90,180 or 270 (to add to offset)
Dim Ais_Distance As Single
Dim AisAngle As Single
Dim dx As Single    'for the centroid offset
Dim dy As Single
Dim dxR As Single   'for the rotation
Dim dyR As Single   'for the rotation
Dim dLon As Double  'adjustment to lat,lon to centroid
Dim dLat As Double
Dim CLon As Double
Dim CLat As Double
Dim M_to_Deg As Double

Dim VesselName As String
Dim Lat As Double
Dim Lon As Double
Dim North As Double
Dim South As Double
Dim East As Double
Dim West As Double
Dim to_bow As Long
Dim to_stern As Long
Dim to_starboard As Long
Dim to_port As Long
Dim Heading As Single
Dim Rotation As Single
Dim TagNo As Long
Dim kb As String
Dim PI As Double
Dim Description As String
Dim Color As String
Dim ratio As Single
Dim Mmsi As String

For TagNo = 0 To UBound(TagArray, 1)
kb = ""
    If TagArray(TagNo, 1) <> "" Then
        Select Case TagArray(TagNo, 0)
        Case Is = "vesselname"
            VesselName = AsciiToXml(Channel, TagArray(TagNo, 1))
        Case Is = "lat"
            Lat = CSng(TagArray(TagNo, 1))
        Case Is = "lon"
            Lon = CSng(TagArray(TagNo, 1))
        Case Is = "to_bow"
            to_bow = CLng(TagArray(TagNo, 1))
        Case Is = "to_stern"
            to_stern = CLng(TagArray(TagNo, 1))
        Case Is = "to_starboard"
            to_starboard = CLng(TagArray(TagNo, 1))
        Case Is = "to_port"
            to_port = CLng(TagArray(TagNo, 1))
        Case Is = "mmsi"
            Mmsi = TagArray(TagNo, 1)
        End Select
    End If
Next TagNo

'If Mmsi <> "244387000" And Mmsi <> "245881000" Then Exit Sub

If VesselName <> "" And Lat <> 0 And Lon <> 0 _
    And to_bow <> 0 And to_stern <> 0 _
    And to_starboard <> 0 And to_port <> 0 Then

    M_to_Deg = 110574.2727
'testing overrides
    If xHeading <> "" Then
        Heading = xHeading
    Else
'        Heading = IconHeading
        Heading = Gis.Heading
    End If
    If xto_bow <> "" Then to_bow = xto_bow
    If xto_stern <> "" Then to_stern = xto_stern
    If xto_port <> "" Then to_port = xto_port
    If xto_starboard <> "" Then to_starboard = xto_starboard
'ratio = (to_bow + to_stern) / (to_port + to_starboard)
'to_port = to_port * ratio
'to_starboard = to_starboard * ratio

'    ratio = (to_bow + to_stern) / 2
'    to_bow = ratio
'    to_stern = ratio
'calculations from offset.xls
    PI = 4 * Atn(1)
    Color = "ffffffff"
'dx,dy is the position of the centroid relative to AIS
'when north up to East is +, to north is +
'dx dy is used here only to get the Ais_distance and the correct angle
    dx = to_starboard - (to_starboard + to_port) / 2
    dy = to_bow - (to_bow + to_stern) / 2
    Ais_Distance = Sqr(dx ^ 2 + dy ^ 2) 'to centroid is 1/2
'AisAngle is Angle between AIS & Centroid from East Anticlockwise
'we need this angle in 360 degrees to establish direction of
'dx and dy
'kb = Atn(1000 / 1) * 180 / Pi
    If dx <> 0 Then
        Offset = Atn(dy / dx) * 180 / PI    'always between 0 and 90
    End If
'express the AIS angle (from Centroid to AIS) between 0 and 360
'this could probably be simplified, but I don't want to break it until it's stable
    If dx <= 0 And dy <= 0 Then Quadrant = 180 '1   North East
    If dx <= 0 And dy >= 0 Then Quadrant = 180  '2  South East
    If dx > 0 And dy >= 0 Then Quadrant = 0     '3  South West
    If dx > 0 And dy <= 0 Then Quadrant = 0  '4 North West
'When dx = 0 (Ais on centre line of ship) dy must be set to 90 or 270 so that
'offset has correct sign. If wrong AIS is outside ship.
    If dx = 0 Then
        If dy > 0 Then
            Quadrant = 90
        Else
            Quadrant = 270
        End If
    End If
'AisAngle is from AIS to centroid from East anticlockwise (0 to 360)
'Cartesian coordinates angle is from x axis (+ is east) to y axis (+ is North)
'intercept on x axis is cos(angle) on y axis is sin(angle)
    AisAngle = Offset + Quadrant
'now dx,dy is the actual x,y vector of the offest
'Calculate dx,dy for the centroid offset, these 6 lines are for testing only
    dx = Ais_Distance * Cos(AisAngle * PI / 180)
    dy = Ais_Distance * Sin(AisAngle * PI / 180)
're-calculate centroid from Ais in Lat/lon
    dLon = dx / M_to_Deg / Cos(Lat * PI / 180)
    dLat = dy / M_to_Deg
    CLon = Lon + dLon
    CLat = Lat + dLat
'rotate
'calculate dxR,dyR rotation offset
        
    dxR = Ais_Distance * Cos((AisAngle - Heading) * PI / 180)
    dyR = Ais_Distance * Sin((AisAngle - Heading) * PI / 180)
        
    dLon = dxR / M_to_Deg / Cos(Lat * PI / 180)
    dLat = dyR / M_to_Deg
    CLon = Lon + dLon
    CLat = Lat + dLat
    Rotation = Heading * -1     'calculated +/- 180
    If Rotation > 180 Then Rotation = Rotation - 360

'testing only
    If Arg = "NorthUp" Then
        Color = "ff00ffff"
        Rotation = 0
        dxR = 0
        dyR = 0
    End If
    
'testing only
    If Arg = "Rotate" Then
        Color = "ffed9564"
        dx = 0
        dy = 0
    End If
    
'Scale to dimensions of ship
    North = CLat + (to_bow + to_stern) / 2 / M_to_Deg
    South = CLat - (to_bow + to_stern) / 2 / M_to_Deg
    East = CLon + (to_port + to_starboard) / 2 / M_to_Deg / Cos(Lat * PI / 180)
    West = CLon - (to_port + to_starboard) / 2 / M_to_Deg / Cos(Lat * PI / 180)
    
'construct the "ballon" - for debugging at the moment
#If False Then  'DONT DELETE may need again for debugging
    Description = _
    "Heading = " & Heading & vbCrLf & _
    "Rotation = " & Rotation & vbCrLf & _
    "to_bow = " & to_bow & vbCrLf & _
    "to_stern = " & to_stern & vbCrLf & _
    "to_port = " & to_port & vbCrLf & _
    "to_starboard = " & to_starboard & vbCrLf & _
    "Quadrant = " & Quadrant & vbCrLf & _
    "Offset = " & Offset & vbCrLf & _
    "Ais_Distance = " & CDec(Ais_Distance) & vbCrLf & _
    "AisAngle = " & CDec(AisAngle) & vbCrLf & _
    "dx = " & dx & vbCrLf & _
    "dy = " & dy & vbCrLf & _
    "Lon = " & CDec(Lon) & vbCrLf & _
    "Lat = " & CDec(Lat) & vbCrLf & _
    "dLon = " & CDec(dLon) & vbCrLf & _
    "dLat = " & CDec(dLat) & vbCrLf & _
    "dxR = " & dxR & vbCrLf & _
    "dyR = " & dyR & vbCrLf & _
    "CLon = " & CDec(CLon) & vbCrLf & _
    "CLat = " & CDec(CLat) & vbCrLf & _
    "North = " & CDec(North) & vbCrLf & _
    "South = " & CDec(South) & vbCrLf & _
    "East = " & CDec(East) & vbCrLf & _
    "West = " & CDec(West) & vbCrLf
#End If

'MsgBox Description
'We require vesselname,lat,lon,Heading,to_bow,to_stern,to_starboard,to_port
    Print #OverlayOutputFileCh, "<GroundOverlay>"
    Print #OverlayOutputFileCh, "  <name>" & VesselName & "</name>"
    Print #OverlayOutputFileCh, "  <description>" & Description & "</description>"
    Print #OverlayOutputFileCh, "<color>" & Color & "</color>"
    
    Print #OverlayOutputFileCh, "  <Icon><href>square.png</href></Icon>"
    Print #OverlayOutputFileCh, "   <LatLonBox>"
    Print #OverlayOutputFileCh, "    <north>" & North & "</north>"
    Print #OverlayOutputFileCh, "    <south>" & South & "</south>"
    Print #OverlayOutputFileCh, "    <west>" & West & "</west>"
    Print #OverlayOutputFileCh, "    <east>" & East & "</east>"
    Print #OverlayOutputFileCh, "    <rotation>" & Rotation & "</rotation> <!-- heading anticlockwise -->"
    Print #OverlayOutputFileCh, "   </LatLonBox>"
    Print #OverlayOutputFileCh, "</GroundOverlay>"
    
'DO NOT REMOVE  - will display the centroid for debugging purposes
'    Print #OverlayOutputFileCh, "<Placemark>" _
'    & "<Style><IconStyle>" _
'    & "<color>" & color & "</color>" _
'    & "<hotSpot x=""20"" y=""4"" xunits=""pixels"" yunits=""pixels""/>" _
'    & "<Icon><href>http://maps.google.com/mapfiles/kml/pushpin/wht-pushpin.png</href></Icon></IconStyle></Style>" _
'    & "<Point><coordinates>" _
'    & CLon & "," & CLat & "</coordinates></Point></Placemark>"

End If

End Sub

'remove crc from word containing crc
Public Function SplitCrc(Word As String) As String
Dim i As Long
'NmeaCrc = ""
i = InStr(1, Word, "*")
If i > 0 Then
'    NmeaCrc = Mid$(Word, i + 1)     'crc
    SplitCrc = Mid$(Word, i + 1, 2)  'crc
    Word = Left$(Word, i - 1)       'remove crc (if any)
End If
End Function

Public Function HttpSpawn(Url As String)
Dim r As Long
Dim Command As String

If Environ("windir") <> "" Then
    r = ShellExecute(0, "open", Url, 0, 0, 1)
Else
'try for linux compatibility
    Command = "winebrowser " & Url & " ""%1"""

    Shell (Command)
End If
End Function
' Return True if a folder exists
Public Function FolderExists(Foldername As String) As Boolean
    FolderExists = False
    On Error GoTo errorhandler
    FolderExists = (GetAttr(Foldername) And vbDirectory) = vbDirectory
'MsgBox Filename & vbCrLf & FileExists
errorhandler:
    err.Clear
    ' if an error occurs, this function returns False
End Function

'return true if they are the same
Public Function FileCompare(File1 As String, File2 As String) As Boolean
Dim Len1 As Long
Dim Ch1 As Long
Dim b1() As Byte
Dim Len2 As Long
Dim fNum As Long
Dim b2() As Byte
Dim i As Long
Dim kb1 As String
Dim kb2 As String
Dim j As Long
Dim LastLf As Long
Dim kb As String

    FileCompare = True
    Call WriteStartUpLog("FileCompare") 'v142
    On Error GoTo Err_NoFile
    Len1 = FileLen(File1)
    Call WriteStartUpLog("File1=" & File1 & " (" & Len1 & " bytes)") 'v147
    Len2 = FileLen(File2)
    Call WriteStartUpLog("File2=" & File2 & " (" & Len2 & " bytes)") 'v147
'MsgBox "File1=" & File1 & " (" & Len1 & " bytes)" & vbCrLf & "File2=" & File2 & " (" & Len2 & " bytes)"
    On Error GoTo 0
'    If Len1 <> Len2 Then   'speed up exit
'        Exit Function
'    End If
    fNum = FreeFile
    ReDim b1(1 To Len1)
    Open File1 For Binary As #fNum
    Get #fNum, 1, b1
    Close fNum
    fNum = FreeFile
    ReDim b2(1 To Len2)
    Open File2 For Binary As #fNum
    Get #fNum, 1, b2
    Close fNum
    For i = 1 To Len1
        If b1(i) = 10 Then LastLf = i
        If b1(i) <> b2(i) Then
 'display first changed line in status bar File1 is the new file
            i = LastLf + 1
            Do Until i > Len1
                If b1(i) = 13 Then Exit Do  'end of line
                kb1 = kb1 & Chr$(b1(i))
                i = i + 1
            Loop
'            NmeaRcv.StatusBar.Panels(1) = kb1
            i = LastLf + 1
            Do Until i > Len2
                If b2(i) = 13 Then Exit Do  'end of line
                kb2 = kb2 & Chr$(b2(i))
                i = i + 1
            Loop
'MsgBox kb2 & vbCrLf & kb1
            Call WriteStartUpLog(vbTab & "Existing  :" & kb2)   'v142
            Call WriteStartUpLog(vbTab & "Changed to:" & kb1)   'v142
            FileCompare = False
'v146 report all differences
'            Exit Function
             Exit Function      'v146 is incorrect (all differences not reported
                                'HUGE file can be output
        End If
    Next i
    If FileCompare = True Then
        Call WriteStartUpLog(vbTab & "Files are the same")  'v142
    End If
    Exit Function

Err_NoFile:
    Call WriteStartUpLog(vbTab & "No File")     'v142
    FileCompare = False
End Function

' Generalized Function to Check if an Instance of Application is running in the machine
'this doesnot work to see if vbe is running
Public Function IsAppRunning(ByVal sAppName) As Boolean
Dim oApp As Object
On Error Resume Next
Set oApp = GetObject(, sAppName)
If Not oApp Is Nothing Then
Set oApp = Nothing
IsAppRunning = True
End If
End Function
'used to split out the domain from the path for FTPing files
'ftp:// is not handled
Function HostFromPath(Path As String) As String
Dim i As Long
i = InStr(Path, "/")
If i > 0 Then
    HostFromPath = Left$(Path, i - 1)
    i = InStr(HostFromPath, ".")
    If i = 0 Then HostFromPath = ""  'must have a dot in url
Else
    HostFromPath = ""
End If
End Function

'used to split out the remote folder from the path for FTP
'ie the domain is removed. ftp:// is not handled
Function FolderFromPath(Path As String) As String
Dim i As Long
i = InStr(Path, "/")
If i > 0 Then
    FolderFromPath = Mid$(Path, i)
'Removed V111 Paul Taylor has . in directory of URL
'    i = InStr(FolderFromPath, ".")
'    If i <> 0 Then FolderFromPath = ""  'must NOT have a dot in url
Else
    FolderFromPath = ""
End If

End Function

Function ExtFromFullName(FullPath As String) As String
Dim i As Long
i = InStrRev(FullPath, ".")
If i > 0 Then
    ExtFromFullName = Mid$(FullPath, i + 1)
Else
    ExtFromFullName = ""
End If
End Function
Function CheckUpdates() As Boolean

Dim WebCurrentVersionNo As String  'on Website in CurrentVersion.csv App.EXEName,3.1.0.57
Dim arry() As String    'to split above
Dim InstalledBuildNo As Long    'this is the app that is currently running
Dim LastUpdatedBuildNo As Long  'last build no updated for this user
Dim WebBuildNo As Long
Dim i As Long

'Check if weve a new version available on the web
'MsgBox "Getting Web Version"
    Call WriteStartUpLog("Checking Web for Updates")
If TreeFilter.Check1(2).Value <> 0 Then
    WebCurrentVersionNo = Download_csv.GetWebCurrentVersion
'    For i = 0 To UBound(WebCurrentVersionNo)
    Call WriteStartUpLog("WebCurrentVersion =" & WebCurrentVersionNo)
'    Next i
    If WebCurrentVersionNo <> "" Then
        arry = Split(WebCurrentVersionNo, ".")
        WebBuildNo = arry(0) * 2 ^ 8 + arry(1) * 2 ^ 4 + arry(3)
        Call WriteStartUpLog("WebBuildNo=" & WebBuildNo)
    End If
'WebBuildNo=0 if no internet access
End If
    
'MsgBox "Checking Web Version"
'If new version on web
InstalledBuildNo = App.Major * 2 ^ 8 + App.Minor * 2 ^ 4 + App.Revision
'testing only WebBuildNo = InstalledBuildNo + 1
If WebBuildNo <> 0 And WebBuildNo <> InstalledBuildNo Then
    NmeaRcv.cbDownload.Enabled = True
    Call WriteStartUpLog("InstalledBuildNo (" & InstalledBuildNo & ") differs - enabling update")
End If
Call GetSystemTime(LastProgramUpdateCheckTime)
End Function

Public Function FormLoaded(FormName As String, Optional FormCaption As String) As Boolean
Dim f As Form
On Error Resume Next
    For Each f In Forms
        If f.Name = FormName Then
            If FormCaption = "" Then
                FormLoaded = True
            Else
                If f.Caption = FormCaption Then FormLoaded = True
            End If
        End If
    Next f
End Function

'V99 changed
Public Function NowUtc() As Variant
'Public Function NowUtc() As String
Dim UtcTime As SYSTEMTIME
Dim NewDate As Variant
'Get Current UTC time
    Call GetSystemTime(UtcTime) 'SYSTEMTIME
    ' Convert it to a Date.
    With UtcTime
'.wYear = 0
'.wMonth = 0
'.wDay = 0
'.wHour = 0
'.wMinute = 0
'.wSecond = 0
'.wMilliseconds = 0
'        the_date = DateSerial(.wYear, .wMonth, .wDay, .wHour, .wMinute, .wSecond)
        NewDate = DateSerial(UtcTime.wYear, UtcTime.wMonth, UtcTime.wDay) _
        + TimeSerial(UtcTime.wHour, UtcTime.wMinute, UtcTime.wSecond)
'NewDate = DateSerial(2013, 10, 27)
'NewDate = NewDate + TimeSerial(UtcTime.wHour, UtcTime.wMinute, UtcTime.wSecond)
    End With
'v99 changed
    NowUtc = Format$(NewDate)
'    NowUtc = Format$(the_date, DateTimeOutputFormat)
'Call strNowUtc
End Function

Public Function NowUnix() As Long
Dim UtcTime As SYSTEMTIME
'Get Current UTC time
    Call GetSystemTime(UtcTime) 'SYSTEMTIME
    NowUnix = SysTimeToUnix(UtcTime)
End Function
'Public Function DateTime(dtDate As Date) As String
'    datetime=date$(dtdate) & " " & time$(dtdate)
'End Function

#If False Then
Public Function NowUtc_old() As String

   Dim tzi As TIME_ZONE_INFORMATION
   Dim gmt As Date
   Dim dwBias As Long
   Dim tmp As String
    
   Select Case GetTimeZoneInformation(tzi)
   Case TIME_ZONE_ID_DAYLIGHT
      dwBias = tzi.Bias + tzi.DaylightBias
   Case Else
      dwBias = tzi.Bias + tzi.StandardBias
   End Select
'Now() returns blank time at midnight so no time is put on log timestamp
   tmp = DateAdd("n", dwBias, Now())
    If DateValue(tmp) = NowUtc Then
        tmp = tmp & " " & TimeValue(NowUtc)
    End If
'    If InStr(1, NowUtc, " ") = 0 Then
'        NowUtc = NowUtc & " 00:00:00"
'    End If
    NowUtc = tmp
End Function
#End If

Public Function LocalTimeZoneName() As String
   Dim LocalTZI As TIME_ZONE_INFORMATION
   Dim kb As String
    Dim lRetVal As Long
    
    lRetVal = GetTimeZoneInformation(LocalTZI)
    Select Case lRetVal
    Case Is = TIME_ZONE_ID_UNKNOWN
        kb = LocalTZI.StandardName    ' "Cannot determine current time zone"
    Case Is = TIME_ZONE_ID_STANDARD
        kb = LocalTZI.StandardName
    Case Is = TIME_ZONE_ID_DAYLIGHT
        kb = LocalTZI.DaylightName
    Case Is = TIME_ZONE_ID_INVALID
        kb = "Invalid"
    End Select
   LocalTimeZoneName = TrimNull(kb)
End Function


Public Function TrimNull(Item As String)
    Dim pos As Integer
      'double check that there is a chr$(0) in the string
    pos = InStr(Item, Chr$(0))
    If pos Then
       TrimNull = Left$(Item, pos - 1)
    Else
       TrimNull = Item
    End If
  
End Function



'This is used for testing form visibility and loading
#If False Then  '(Moved to ModGeneral)
Sub DisplayForms(Title As String)
Dim f As Form
Dim kb As String
Dim Visible As String

For Each f In Forms
    If f.Visible = False Then
        Visible = " (Hidden)"
    Else
        Visible = ""
    End If
    kb = kb & f.Name & ":" & f.Caption & Visible & vbCrLf
Next f
kb = kb & "Total of " & Forms.Count & " forms loaded" & vbCrLf

MsgBox kb, , Title
End Sub
#End If

Public Function DecryptString(kb As String) As String
    Dim Secret As EncryptedData
'kb = "open363dir"
'If Len(kb) < 128 Then
'    DecryptString = Encrypt_Original(kb)
'Else
    If kb = "" Then Exit Function
    Set Secret = New EncryptedData
    Secret.Algorithm.Name = CAPICOM_ENCRYPTION_ALGORITHM_AES
    Secret.Algorithm.KeyLength = CAPICOM_ENCRYPTION_KEY_LENGTH_256_BITS
    Secret.SetSecret "My Secret Encryption Key" & Environ$("ComputerName") & Environ$("username")
'    Secret.Content = "password" ' just so we know that this is being reset by decryption
    On Error Resume Next    'errror if .secret differs
    Secret.Decrypt kb
    On Error GoTo 0
    On Error GoTo Not_Encrypted    'error if string is not encrypted
    DecryptString = Secret.Content
    Set Secret = Nothing
Exit Function

Not_Encrypted:
    Set Secret = Nothing
    kb = EncryptString(kb)
    DecryptString = DecryptString(kb)
'End If
'MsgBox Decrypt
End Function

Public Function EncryptString(kb As String) As String
Dim Secret As EncryptedData
    
    If kb = "" Then Exit Function
    Set Secret = New EncryptedData
    Secret.Algorithm.Name = CAPICOM_ENCRYPTION_ALGORITHM_AES
    Secret.Algorithm.KeyLength = CAPICOM_ENCRYPTION_KEY_LENGTH_256_BITS
    Secret.SetSecret "My Secret Encryption Key" & Environ$("ComputerName") & Environ$("username")
    Secret.Content = kb ' what we want to encrypt
'we must remove the split lines secret.content includes
    EncryptString = Replace(Secret.Encrypt, vbCrLf, "")
'    MsgBox Encrypt
    Set Secret = Nothing
End Function

Public Function Encrypt_Original(kb As String) As String
Dim i As Long
Dim j As Long
Dim strKey As String
Dim strChar1 As String
Dim StrChar2 As String
Dim strTemp As String

strKey = Environ$("ComputerName") & Environ$("username")

    For i = 1 To Len(kb)
        'Get the next character from the text
        strChar1 = Mid(kb, i, 1)
        'Find the current "frame" within the key
        j = ((i - 1) Mod Len(strKey)) + 1
        'Get the next character from the key
        StrChar2 = Mid(strKey, j, 1)
        'Convert the charaters to ASCII, XOR them, and convert to a character again
'add 128 to avoid nul terminating string
        strTemp = strTemp & Chr(Asc(strChar1) Xor Asc(StrChar2) + 128)
    Next i
Encrypt_Original = strTemp
End Function

Public Function SysTimeToUnix(Systime As SYSTEMTIME) As Long
Dim DinY()
Dim Dayno As Long
Dim ret As Long
Dim kb As String

    With Systime
        SysTimeToUnix = DateDiff("s", DateSerial(1970, 1, 1), DateSerial(.wYear, .wMonth, .wDay) + TimeSerial(.wHour, .wMinute, .wSecond))
    End With

End Function

Public Function UnixTimeToDate(ByVal Timestamp As Long) As String

          Dim intDays As Integer, intHours As Integer, intMins As Integer, intSecs As Integer
          intDays = Timestamp \ 86400
          intHours = (Timestamp Mod 86400) \ 3600
          intMins = (Timestamp Mod 3600) \ 60
          intSecs = Timestamp Mod 60
          UnixTimeToDate = DateSerial(1970, 1, intDays + 1) + TimeSerial(intHours, intMins, intSecs)
        If IsDate(UnixTimeToDate) = False Then
            UnixTimeToDate = ""
        End If
'MsgBox Timestamp
'            UnixTimeToDate = Format$(UnixTimeToDate, DateTimeOutputFormat)
'            Val = SysTimeToUnix(MsgTimeSys)
'            ValDes = UnixTimeToDate(Val)
      
End Function

'Converts a DateTime string to Unix Time
Public Function DateToUnixTime(ByVal strDate As String) As Long
    On Error Resume Next
    DateToUnixTime = DateDiff("s", DateSerial(1970, 1, 1), strDate)
    On Error GoTo 0
End Function

'Hide must be variant
Sub FormsLog(LogFile As String, Optional Caller As String, Optional Hide As Variant)
Dim f As Form
Dim kb As String

If Caller <> "" Then kb = "Caller is " & Caller & vbCrLf
    kb = kb & "Total Forms = " & Forms.Count & vbCrLf
    kb = kb & "cmdNoWindow = " & cmdNoWindow & vbCrLf
    For Each f In Forms
        kb = kb & f.Name & ":" & f.Caption & " - Visible = " & f.Visible & vbCrLf
'Force all forms to be visible or invisible
'IsMissing only works with variant
        If IsMissing(Hide) = False Then
            If f.Visible <> Hide Then
                f.Visible = Hide
                kb = kb & "Forced " & f.Name & " to " & Hide & vbCrLf
            End If
        End If
    Next f
    Select Case LogFile
    Case Is = "MsgBox"
        MsgBox kb, , "AisDecoder - Forms Visibility"
    Case Is = "StartUp"
        Call WriteStartUpLog(kb)
    End Select

End Sub

Public Function CtrlToString(Line As String) As String
Dim kb As String
Dim Outbuf As String
Dim i As Long
Dim Chrno As Long
Dim b() As Byte

    b = StrConv(Line, vbFromUnicode)
'    If Not b Is Nothing Then
        For i = 0 To UBound(b)
            If b(i) >= 32 And b(i) <= 127 Then
            Outbuf = Outbuf & Chr$(b(i))
        Else
            Outbuf = Outbuf & "<" & b(i) & ">"
        End If
        Next i
'    Else
'        DpyBuf = "<empty>"
'    End If
    CtrlToString = Outbuf
End Function

Public Function StringToCrlf(Line As String) As String
Dim kb As String
Dim Outbuf As String
Dim i As Long
Dim Chrno As Long
Dim b() As Byte

    Outbuf = Line
    Outbuf = Replace(Outbuf, "<cr>", Chr$(13), , , vbTextCompare)
    Outbuf = Replace(Outbuf, "<lf>", Chr$(10), , , vbTextCompare)
    Outbuf = Replace(Outbuf, "<nul>", Chr$(0), , , vbTextCompare)
    StringToCrlf = Outbuf
End Function

Public Function CrlfToString(Line As String) As String
Dim kb As String
Dim Outbuf As String
Dim i As Long
Dim Chrno As Long
Dim b() As Byte

    Outbuf = Line
    Outbuf = Replace(Outbuf, Chr$(13), "<CR>", , , vbTextCompare)
    Outbuf = Replace(Outbuf, Chr$(10), "<LF>", , , vbTextCompare)
    Outbuf = Replace(Outbuf, Chr$(0), "<NUL>", , , vbTextCompare)
    CrlfToString = Outbuf
End Function

'Not used
Public Function ByteArrayToHexStr(b() As Byte) As String
   Dim n As Long, i As Long
   
   ByteArrayToHexStr = Space$(3 * (UBound(b) - LBound(b)) + 2)
   n = 1
   For i = LBound(b) To UBound(b)
      Mid$(ByteArrayToHexStr, n, 2) = Right$("00" & Hex$(b(i)), 2)
      n = n + 3
   Next
End Function

Function AsciiToHexStr(AsciiStr As String) As String
   Dim i As Long
   Dim kb As String
   
    For i = 1 To Len(AsciiStr)
        If kb <> "" Then kb = kb & " "
        kb = kb & Right$("00" & Hex$(Asc(Mid$(AsciiStr, i, 1))), 2) 'to change "A" to "0A"
    Next i
    AsciiToHexStr = kb
End Function

Public Function IsIniFile(FileName As String) As Boolean
Dim ch As Integer
Dim kb As String

'MsgBox "IsIniFile=" & ExpEnvStr(FileName)
'File may not exist
    If FileExists(FileName) Then
        ch = FreeFile
        Open FileName For Input As #ch
'File may be zero length
        Do Until EOF(ch)
            Line Input #ch, kb
            Close ch
'Exit after first line
            Exit Do
        Loop
        If kb = "[FRAME]" Then
            IsIniFile = True
        End If
    End If
End Function

Function ExpEnvStr(strInput As String) As String
Dim result As Long
Dim strOutput As String

'' Two calls required, one to get expansion buffer length first then do expansion
result = ExpandEnvironmentStrings(strInput, strOutput, result)
strOutput = Space$(result)
'ExpEnvStr = Space$(result)
result = ExpandEnvironmentStrings(strInput, strOutput, result)
ExpEnvStr = strOutput
'result = 0
'result = ExpandEnvironmentStrings(strInput, strOutput, result)
'strOutput = Space$(result)
'result = ExpandEnvironmentStrings(strInput, strOutput, result)
End Function

Public Function GetDecimalSeparator() As String
    GetDecimalSeparator = Mid$(1 / 2, 2, 1)
End Function

'http://www.geodatasource.com/developers/vb
':::    South latitudes are negative, east longitudes are positive           :::
':::    lat1, lon1 = Latitude and Longitude of point 1 (in decimal degrees)  :::
':::    lat2, lon2 = Latitude and Longitude of point 2 (in decimal degrees)  :::
':::    unit = the unit you desire for results                               :::
':::           where: 'M' is statute miles                                   :::
':::                  'K' is kilometers (default)                            :::
':::                  'N' is nautical miles                                  :::

Function RhumbLineDistance(lat1 As Single, lon1 As Single, lat2 As Single, lon2 As Single) As Single
Dim Departure As Single
Dim dLong As Single
Dim dLat As Single
Dim AveLat As Single
Dim wLat As Single
Dim wlong As Single

'Departure = Change of Longitude (in minutes) x Cosine of Latitude
    If lon1 >= lon2 Then
        dLong = lon1 - lon2
    Else
        dLong = lon2 + 360 - lon1
    End If
    If dLong >= 360 Then dLong = dLong - 360
    If lat1 >= lat2 Then
        dLat = lat1 - lat1
    Else
        dLat = lat2 + 180 - lat1
    End If
    If dLat >= 180 Then dLat = dLat - 180
    AveLat = (lat1 + lat2) / 2
    Departure = dLong * Cos(deg2rad(CDbl(AveLat)))
    RhumbLineDistance = Sqr(Departure ^ 2 + dLat ^ 2) * 60
End Function

Function LatLonDistance(lat1 As Single, lon1 As Single, lat2 As Single, lon2 As Single) As Single
Dim theta As Single
Dim dist As Double
    theta = lon1 - lon2
    dist = Sin(deg2rad(lat1)) * Sin(deg2rad(lat2)) + Cos(deg2rad(lat1)) * Cos(deg2rad(lat2)) * Cos(deg2rad(theta))
    dist = acos(dist)
    dist = rad2deg(dist) * 60   '60 nm per degree
LatLonDistance = CSng(dist)
End Function

'This function get the arccos function using arctan function   :::
Function acos(Rad As Double) As Double
    On Error GoTo bad_argument
'abs(double) returns a double data type, which may not be
'exactly 1 !!!! (causes an invalid procedure call error
    If Abs(Rad) <> 1 Then
        acos = PI / 2 - Atn(Rad / Sqr(1 - Rad * Rad))
    ElseIf Rad = -1 Then
        acos = PI
    End If
    Exit Function
bad_argument:
    acos = 0
End Function

'This function converts decimal degrees to radians             :::
Function deg2rad(Deg As Single) As Double
    deg2rad = CDbl(Deg * PI / 180)
End Function

'This function converts radians to decimal degrees             :::
Function rad2deg(Rad As Double) As Double
    rad2deg = CDbl(Rad * 180 / PI)
End Function

Public Function aLatLon(LatLon As Single, LatorLon As String) As String
Dim kb As String

    kb = Int(Abs(LatLon)) & Chr$(176) & " " _
    & Format((Abs(LatLon) - Int(Abs(LatLon))) * 60, "0.0000") & "'"
    Select Case LatorLon
    Case Is = "Lon"
        If LatLon >= 0 Then
            kb = kb & " E"
        Else
            kb = kb & " W"
        End If
    Case Is = "Lat"
        If LatLon >= 0 Then
            kb = kb & " N"
        Else
            kb = kb & " S"
        End If
    End Select
    aLatLon = kb
End Function

'Clears MysShip and display
Public Function ClearMyShip(Ship As ShipDef)
    Ship.Lat = 91   'not available
    Ship.Lon = 181  'not available
    Ship.Mmsi = ""
    Ship.Name = ""
    Ship.RcvTime = ""
'Clear the display
    NmeaRcv.Label4(0) = ""
    NmeaRcv.Frame11.Visible = False
'    Ship.PositionTime = ""
End Function

Public Function RoundIt(ByVal aNumberToRound As Single, _
  Optional ByVal aDecimalPlaces As Long = 0) As Single

On Error GoTo errHandler

Dim nFactor As Double
Dim nTemp As Double

    nFactor = 10 ^ aDecimalPlaces
    nTemp = (aNumberToRound * nFactor) + 0.5
    RoundIt = Int(CDec(nTemp)) / nFactor

'-----------EXIT POINT------------------
ExitPoint:

Exit Function

'-----------ERROR HANDLER---------------
errHandler:

    Select Case err.Number
        Case Else
            ' Your error handling here
            RoundIt = 0
            Resume ExitPoint
    End Select
End Function

'returns ddmm.mmmmmm with leading zero's
Function LatToNmea(Lat As Single) As String
Dim posMins As Long
Dim posDegs As Long
Dim sngMins As Single
Dim kb As String
    posDegs = Int(Abs(Lat)) 'if -ve return +ve
    kb = Format$(posDegs, "00")
    sngMins = (Abs(Lat) - posDegs) * 60
    kb = kb & sngMins
    If Lat >= 0 Then
        kb = kb & ",N"
    Else
        kb = kb & ",S"
    End If
    LatToNmea = kb
End Function

'returns dddmm.mmmmmm with leading zero's
Function LonToNmea(Lon As Single) As String
Dim posMins As Long
Dim posDegs As Long
Dim sngMins As Single
Dim kb As String
    posDegs = Int(Abs(Lon)) 'if -ve return +ve
    kb = Format$(posDegs, "000")
    sngMins = (Abs(Lon) - posDegs) * 60
    kb = kb & sngMins
    If Lon <= 0 Then
        kb = kb & ",W"
    Else
        kb = kb & ",E"
    End If
    LonToNmea = kb
End Function

Function NmeaLatLon(LatLon As String) As String
Dim arry() As String
'Split into 55 12.33 N
    On Error GoTo Convert_Error
    arry = Split(LatLon, " ")
    If UBound(arry) = 2 Then
'Remove degree and minute symbols
        arry(0) = Replace(arry(0), Chr$(176), "")
        arry(1) = Replace(arry(1), "'", "")
'Add leading zero's to degrees 2 for lat 3 for lon
        Select Case arry(2)
        Case Is = "N", "S"
            arry(0) = Format$(CLng(arry(0)), "00")
        Case Is = "E", "W"
            arry(0) = Format$(CLng(arry(0)), "000")
        Case Else
            Exit Function
        End Select
'Add leading zero's to minutes
        arry(1) = Format$(CSng(arry(1)), "00.00")
'recombine back into tag array
        NmeaLatLon = arry(0) & arry(1) & ", " & arry(2)
    End If
Convert_Error:
End Function

Function NmeaCrcChk(ByVal NmeaSentence As String)
Dim i As Long
Dim CheckSum As Byte
Dim lngChkSum As Long
Dim Chr As String
Dim HexChecksum As String
Dim b() As Byte
Dim PassedCrc As String
Dim Offset As Long
Dim wlong As Long
    On Error GoTo error_crc
'check checksum
'Note passed string includes (! $ or \)
    Select Case Left$(NmeaSentence, 1)
    Case Is = "!", "$", "\"
        Offset = 1
    End Select

b = StrConv(NmeaSentence, vbFromUnicode)
If UBound(b) = 0 Then Exit Function '0 or less characters
CheckSum = b(Offset) 'set the first byte to be checked
For i = 1 + Offset To UBound(b) 'Excluces !,$,\ and * in last word
    If b(i) = 42 Then Exit For  '* found
    CheckSum = CheckSum Xor b(i)
Next i
    lngChkSum = CheckSum
'return checksum
'On Error Resume Next
    If CheckSum < 16 Then
        HexChecksum = "0" & Hex(CheckSum)
    Else
        HexChecksum = Hex(CheckSum)
    End If
'Expression too complex !!!
'HexChecksum = Hex$((CheckSum And 240) / 16) & Hex$(CheckSum And 15)
'if [CRC] is the checksum, we are going to replace it with the calculated check sum
If Mid$(NmeaSentence, i + 2, 5) = "[CRC]" Then
    NmeaCrcChk = HexChecksum
Else
    PassedCrc = Mid$(NmeaSentence, i + 2, 2)
    If PassedCrc <> "hh" Then     'test data
        If HexChecksum <> PassedCrc Then
            NmeaCrcChk = "{CRC error " & HexChecksum & "}"
        End If
    End If
End If
Exit Function
error_crc:
    NmeaRcv.StatusBar.Panels(1).Text = "NmeaCrcChk error no " & err.Number & " " & err.Description
    NmeaRcv.ClearStatusBarTimer.Enabled = True
    Call WriteErrorLog(NmeaRcv.StatusBar.Panels(1).Text & vbCrLf & clsSentence.NmeaSentence)
 'Debug.Print "End CC"
    err.Clear
End Function

'Given a sentence gets the CRC for the input string
'between $!or\ and * and returns CRC
Function CalculateCrc(InputStr As String) As String
Dim i As Long
Dim CheckSum As Byte
Dim lngChkSum As Long
Dim Chr As String
Dim HexChecksum As String
Dim b() As Byte
Dim PassedCrc As String
Dim Offset As Long
'check checksum
'Note passed string includes (! $ or \)
'inserted to debug Expression too complex !!!
 'Debug.Print "CC"
    On Error GoTo error_crc
    
    Select Case Left$(InputStr, 1)
    Case Is = "!", "$", "\"
        Offset = 1
    End Select
b = StrConv(InputStr, vbFromUnicode)
 'Debug.Print "CCa"
    If UBound(b) = 0 Then
 'Debug.Print "Exit CC"
    
        Exit Function '0 or less characters
    End If
 'Debug.Print "CCb"
    CheckSum = b(Offset) 'set the first byte to be checked
    For i = 1 + Offset To UBound(b) 'Excluces !,$,\ and * in last word
        If b(i) = 42 Then Exit For  '* found
        CheckSum = CheckSum Xor b(i)
    Next i
 'Debug.Print "CCc"
    lngChkSum = CheckSum
'return checksum
'On Error Resume Next    'Expression too complex !!!
'wlong = CLng(CheckSum) And 240
'If wlong Then wlong = wlong / 16
'HexChecksum = Hex$(wlong)
'HexChecksum = Hex$((CLng(CheckSum) And 240) / 16)
'HexChecksum = HexChecksum & Hex$(CLng(CheckSum) And 15)
'Expression too complex !!!
 'Debug.Print "CCd"
    If CheckSum < 16 Then
 'Debug.Print "CCda"
        HexChecksum = "0" & Hex(CheckSum)
    Else
 'Debug.Print "CCdb"
        HexChecksum = Hex(CheckSum)
    End If
 'Debug.Print "CCdc"
'HexChecksum = Hex((CheckSum And 240) / 16) & Hex(CheckSum And 15)
 'Debug.Print "CCe"
'On Error GoTo 0
'if [CRC] is the checksum, we are going to replace it with the calculated check sum
    If Mid$(InputStr, i + 2, 5) = "[CRC]" Then
 'Debug.Print "CCf"
        CalculateCrc = HexChecksum
    Else
 'Debug.Print "CCg"
      PassedCrc = Mid$(InputStr, i + 2, 2)
     If PassedCrc <> "hh" Then     'test data
            If HexChecksum <> PassedCrc Then
                CalculateCrc = "{CRC error " & HexChecksum & "}"
            End If
        End If
    End If
 'Debug.Print "Exit CC"
Exit Function
error_crc:
    NmeaRcv.StatusBar.Panels(1).Text = "CalculateCrc error no " & err.Number & " " & err.Description
    NmeaRcv.ClearStatusBarTimer.Enabled = True
    Call WriteErrorLog(NmeaRcv.StatusBar.Panels(1).Text & vbCrLf & clsSentence.NmeaSentence)
 'Debug.Print "End CC"
    err.Clear
End Function

'Converts ^hh to ascii code
Public Function ConvEscChrs(ByVal kb As String) As String
Dim i As Long
    i = InStr(1, kb, "^")
    Do Until i = 0
        kb = Replace(kb, Mid$(kb, i, 3), Chr$(CLng("&H" & "21")))
        i = InStr(1, kb, "^")
    Loop
    ConvEscChrs = kb
End Function

Public Function GroupSentenceClear_old()
    If GroupSentence.Encapsulates = False Then
        Set GroupSentence = Nothing
    End If
End Function

'LF only because CRLF can be split when counting
Public Function LFCount(ByVal kb As String) As Long
Dim Count As Long
Dim i As Long

'NmeaRcv.Refresh
    If Len(kb) = 0 Then Exit Function
    i = InStr(i + 1, kb, vbLf)
    Do While i > 0
        Count = Count + 1
        i = InStr(i + 1, kb, vbLf)
    Loop
    LFCount = Count
End Function

'Find the position of the No'th LF
'Return 0 if none or last CRLF if < No or No=0
Public Function CRLFPos(ByVal kb As String, ByVal No As String) As Long
Dim i As Long
Dim Count As Long
'i = 0
    If Len(kb) = 0 Then
        CRLFPos = -1     'No length
        Exit Function
    End If
    If No = 0 Then      '-1   0      1
        CRLFPos = -1    '<cr><LF><1st Character>
        Exit Function
    End If
    i = InStr(i + 1, kb, vbCrLf)
    If i = 0 Then Exit Function
    Do While i > 0
        Count = Count + 1
        CRLFPos = i
        If Count = No Then
            Exit Do
        End If
        i = InStr(i + 1, kb, vbCrLf)
    Loop
End Function

Public Function aInputState() As String
    Select Case InputState
    Case Is = 0
        aInputState = "Stopped"
    Case Is = 1
        aInputState = "Started"
    Case Is = 2
        aInputState = "Paused"
    End Select
End Function

'The string can be formatted ie 1,000,000
Public Function aByte(strBytes As String) As String
Dim ctotalByte As Currency
Dim ctotalKb As Currency
Dim cTotalMb As Currency
Dim ctotalGb As Currency
Dim ctotalTb As Currency

    If Len(strBytes) < 4 Then
        aByte = ""
        Exit Function
    End If
    ctotalByte = CDec(strBytes)
    
    ctotalKb = ctotalByte / 1000@    'to 3 decimal places
'Dont display if < 1k bytes, except when stop (in case very few bytes)
    If ctotalKb < 1 Then
        If InputState = 0 Then
            aByte = CDec(ctotalByte) & " bytes"
        Else
            aByte = ""
        End If
        Exit Function
    End If
    
    cTotalMb = ctotalKb / 1000@    'to 3 decimal places
    If cTotalMb < 1 Then
        aByte = Format$(CDec(ctotalKb), "##0.0\ Kb")
        Exit Function
    End If
    
    ctotalGb = cTotalMb / 1000@    'to 3 decimal places
    If ctotalGb < 1 Then
        aByte = Format$(CDec(cTotalMb), "##0.0\ Mb")
            Exit Function
    End If
    
    ctotalTb = ctotalGb / 1000@    'to 3 decimal places
    If ctotalTb < 1 Then
        aByte = Format$(CDec(ctotalGb), "##0.0\ Gb")
        Exit Function
    End If
    
    aByte = Format$(CDec(ctotalTb), "#,###,##0.0\ Tb")
End Function

#If False Then
Public Function StringToCurrency_notused(Value As String) As Currency
Dim C As MungeCurr
Dim l As Munge2Long
Stop
    C.Value = Value / 10000@
    LSet l = C
    StringToCurrency = C.Value
End Function
#End If

Public Function DisplayAisMsg()
Dim AisMsg As New clsAisMsg    'ship is class
Dim AisMsgKey As String
Dim kb As String

With clsSentence
    AisMsgKey = cKey(.AisMsgFromMmsi, 9) _
    & cKey(.AisMsgType, 2) _
    & cKey(.SentencePart, 1) _
    & cKey(.AisMsgDac, 4) _
    & cKey(.AisMsgFi, 2) _
    & cKey(.AisMsgFiId, 2)
End With

    kb = "AisMsgKey=" & AisMsgKey & vbCrLf
    On Error GoTo keynotfound
    Set AisMsg = AisMsgs(AisMsgKey)  'see if this AisMsg is in collection
    On Error GoTo 0
    kb = kb & "AisMsgSentence=" & AisMsg.AisSentence & vbCrLf
    kb = kb & "NmeaRcvTime=" & AisMsg.NmeaRcvTime & vbCrLf
    MsgBox kb
    Set AisMsg = Nothing
    Exit Function

keynotfound:
    kb = "Key Not Found"
    Resume Next
End Function

'Now only called by InputBuffersToNmeaBuf
'Called by cbStart, ClientUDP_DataArrival, ClientTCP_DataArrival, MSComm1_OnComm
'Must only be called if Xoff = false, because calling
Public Sub ProcessNmeaBuf()
Dim Processed As Long
    
'    If DecodingState <> 0 Or Processing.Suspended = True Then
'        Exit Sub
'    End If
    
    Processing.Suspended = True
    Processing.NmeaBuf = True
'    DecodingState = DecodingState Or 1
'Only allow to be called once - potential stack overflow
'    If ProcessNmeaBufCalls = 0 Then
    
'update on entry
'    Call UpdateStats
'    Refresh
'Pause is checked on each iteration because it may have been set after the call
'to ProcessNmeaBuf
    Do Until NmeaBufUsed = 0    'Empty
        ProcessSentence (NmeaBuf(NmeaBufNxtOut))
        NmeaBufNxtOut = NmeaBufNxtOut + 1
        If NmeaBufNxtOut > NMEABUF_MAX Then NmeaBufNxtOut = 0
        NmeaBufUsed = NmeaBufUsed - 1
        Processed = Processed + 1
        NmeaBufXoff = False
'        If NmeaBufXoff = True Then
'NmeaBuf is 50% full (<= so that it only processes exactly 1/2 (10000 if 20k)
'            If NmeaBufUsed * 2 <= NMEABUF_MAX Then
'                XonNmeaBuf 'Start receiving more data into NmeaBuf
'Allow a Pause to be actioned
'Also allows UDP/TCP DataRcv
'                DoEvents    'process input interupts
'               Exit Do     'exit Processing NmeaBuf
'           End If
'        End If
    
''    If Processed Mod 1000 = 0 Then
'keep user informed
'        Call NmeaRcv.UpdateStats    'NO this will display NMEABuf Wait
'ensure NmeaRcv remains responsive to click events every 1000 sentences
'do input interrupts every 1000 sentences processed
''        DoEvents
'        Exit Do
''    End If
'    If NmeaRcv.cbPause.Caption = "Continue" Then   'Continue to allow routines paused to exit
'        Exit Do
'    End If
'Stop    'should not happen
    Loop
    
'    End If
'update on exit
'    Call UpdateStats
'    Refresh
'Check no more data received (no need to in RcvData or InputBuffersToNmeaBuf)
'If Stop has been pressed during ProcessNmeaBuf - InputFile will now be nothing
'Because ProcessNmeaBuf my have been called in FileDataArrival - FileDataArrival must
'Now stop calling InputBuffersToNmeaBuf (before DoEvents ?)
'Stop
'DoEvents   'With fast UDP input causes UDPDataRcv to be called again (potential stack overflow)

'If Processed = 0 Then Stop  'debugv131Debug.Print "Processed " & Processed & " NmeaBuf " & NmeaBufUsed & " left"
'If NmeaBufStart - NmeaBufUsed = 0 Then Stop
    Call ResumeProcessing("NmeaBuf")
'    DecodingState = DecodingState Xor 1
End Sub

'Called when ProcessNmeaBuf reaches Lower limit
Public Sub XonNmeaBuf()   'Call to turn on getting data (set Xoff=false)
        
'Only take action if not currently turned off, otherwise call to start processing buffer
'could be re-entrant (called from itslf)
    If NmeaBufXoff = True Then
        NmeaBufXoff = False
Debug.Print "Xon (Nmeabuf)"
'Call to start processing buffer here
    End If
End Sub

Public Function ResumeProcessing(Caller As String)
    With Processing
        Select Case Caller
        Case Is = "NmeaBuf"
            .NmeaBuf = False
        Case Is = "Scheduler"
            .Scheduler = False
        Case Is = "InputOptions"
            .InputOptions = False
        Case Is = "Paused"
            .Paused = False
        Case Else
MsgBox "Unexpected Caller argument " & Caller, vbExclamation, "Resume Processing"
        End Select
        If .NmeaBuf = False And .InputOptions = False And .Scheduler = False And .Paused = False Then
            .Suspended = False
        End If
    End With
End Function


'Checks clsSentence to validate Payload
'PayloadBits and FillBits must be within valid range for this AisMsg
Private Function IsPayloadFillOK() As Boolean

    With clsSentence
        If IsNumeric(.AisPayloadFillBits) And IsNumeric(.AisMsgType) Then   'In case of "" in clssentence
            If .AisPayloadFillBits >= 0 And .AisPayloadFillBits <= 5 Then     'must be 0 to 5 (6 bit ascii)
'if -1 variable length payload
                If AisMsgTypeBitsFill(.AisMsgType) < 0 Or .AisPayloadFillBits = AisMsgTypeBitsFill(.AisMsgType) Then
'Fill bits OK, check length
                    If .AisPayloadBits >= AisMsgTypeBitsMin(.AisMsgType) And .AisPayloadBits <= AisMsgTypeBitsMax(.AisMsgType) Then
                        IsPayloadFillOK = True
                    End If
                End If
            If .AisMsgType = 24 And .AisPayloadBits = 160 And .AisPayloadFillBits = 2 Then
                    IsPayloadFillOK = True
            End If
         End If
        End If
        
'this is only used to reject lat/lon before filtering
'        If .AisMsgPartsComplete = True And IsPayloadFillOK = False Then
'            If DisableNmeaFillBitsError = True Then
'                IsPayloadFillOK = True
'            End If
'        End If
    
    End With
End Function

'If not AIS sentence then Error is False (Only AIS sentences can return an error)
'Error if IsPayloadError() <> ""
Public Function IsPayloadError(InputSentence As String) As String
Dim From As Long    'to start position of bits
Dim Payload8Bits As Long
Dim CBFrom As Long  'Start of comment block
Dim CBTo As Long   'end of Comment block = 0 (excl delimeter) if not found
Dim NmeaFrom As Long   'Start of NMEA (incl $ or !)part of sentence = 0 if not found
Dim kb As String

    If InputSentence = "" Then Exit Function
'find start of nmea (IEC spec - ! or $ is always start of NMEA
    NmeaFrom = InStr(1, InputSentence, "!")
    If NmeaFrom = 0 Then Exit Function  'Not AIS
'Comment Block is discarded
    kb = Mid$(InputSentence, NmeaFrom)
'check we don't just have a comment block
    If kb = "" Then Exit Function
'created into Public Variable becuase you cant have a dynamic array in a class module
    NmeaWords = Split(kb, ",")  '350k/min
    If UBound(NmeaWords) < 6 Then
        IsPayloadError = "AIS NMEA sentence incomplete"
        Exit Function
    End If
    
    If NmeaWords(2) = "1" And Len(NmeaWords(5)) < 2 Then
        IsPayloadError = "Payload incomplete"
        Exit Function
    End If
'The first character of the 1st Part
    If NmeaWords(2) = "1" Then
    Select Case Left$(NmeaWords(5), 1)  'Message type
    Case Is = "1", "2", "3", "4", "9", "B"  '4=Base,9=SAR,B=msg18
        If Len(NmeaWords(5)) <> 28 Then
            IsPayloadError = "Payload length incorrect"
            Exit Function
        End If
    End Select
    End If
    
End Function

'Check to see of the argumnt can be converted to Long (required arg for UnixTimeToDate
Function IsLong(kb As Variant) As Boolean
Dim lngLong As Long
    On Error GoTo skip
    lngLong = kb
    IsLong = True
skip:
End Function

'So I can prematurely exit a subroutine for testing closedown problem
Function ExitSub() As Boolean
ExitSub = False
End Function
