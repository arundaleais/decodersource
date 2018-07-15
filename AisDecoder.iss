; -- AisDecoder.iss --

;SourcePath is where the .iss file is located
#pragma message SourcePath
#define MyAppName "AisDecoder.exe" 
#pragma message "MyAppName info: " + MyAppName
#define MyAppFile SourcePath + MyAppName
#pragma message "MyAppFile info: " + MyAppFile
#define MyAppVersion GetFileVersion(MyAppFile)
#pragma message "Detailed version info: " + MyAppVersion
;Remove the 3rd digit (0)
#define MyAppVersion StringChange(MyAppVersion, ".0.", "." )
#pragma message "Stripped version info: " + MyAppVersion

#define public MyFileDateTimeString GetFileDateTimeString(MyAppFile, 'dd/mm/yyyy hh:nn:ss', '-', ':');
#pragma message "File Date info: " + MyFileDateTimeString
#define MyDateTimeString GetDateTimeString('dd/mm/yyyy hh:nn:ss', '-', ':');
#define result Exec('cmd /c xcopy/s/y', '"C:\Users\Admin\My Documents\Ais\DecoderSource" "C:\Users\Admin\My Documents\Ais\DecoderSourceBackup\AisDecoder_' + MyAppVersion + '\"')
;Copy CommonSource files into Common backup as we may have changed a common routine
#define result Exec('cmd /c xcopy/s/y/q', '"C:\Users\Admin\My Documents\Ais\CommonSource" "C:\Users\Admin\My Documents\Ais\CommonSourceBackup\Common_' + MyAppVersion + '\"')

;The location on Win10 (Was CommonAppData  or AllUsers on XP)
#define MyProgramData "C:\ProgramData"
#pragma message "MyProgramData: " + MyProgramData
#define MySys32 "C:\Windows\SysWOW64"
#pragma message "MySys32: " + MySys32

[Setup]
;version explorer displays for setup.exe, recovered with VB6 app.major & app.minor
VersionInfoVersion={#MyAppVersion}
;minimum windows version sofware will run on (0=no Win98, 4.0= nt or 2000,XP upwards)
MinVersion= 0,5.0
AppName=Ais Decoder
AppId=Ais Decoder
;CreateUninstallRegKey=no
;UpdateUninstallLogAppName=no
;On INNO installer "This will install Ais Decoder Version x.x.x.x on your computer"
AppVerName=Ais Decoder
AppPublisher=Neal Arundale
AppPublisherURL=http://arundale.com/docs/ais/sp_map.html
;where the users files are placed
DefaultDirName={pf}\Arundale\Ais Decoder
DefaultGroupName=Ais Decoder
;UsePreviousAppDir=No
;UninstallDisplayIcon=E:\jna\arundale\website\docs\arundale.ico
UninstallDisplayIcon=arundale.ico
;outputdir=E:\jna\Arundale\website\docs\ais\
;outputdir=C:\website\
;outputdir= "C:\Users\Admin\My Documents\DirectNic\Live Parent (ArundaleCom)\docs\ais"
outputdir="C:\Users\Admin\Documents\ais\DecoderSource"
OutputBaseFilename= AisDecoder_setup_{#MyAppVersion}
setuplogging=yes
SetupIconFile=arundale.ico
;required for vbfiles installation
PrivilegesRequired=admin
LicenseFile=license.txt
;FileDateTimeString= (#MyFileDateTimeString)
AppMutex="AisDecoder"

[Dirs]
;only required if creating an empty directory [files] creates the directory
;these get copied to userappdata when AisDecoderns new version
Name: "{commonappdata}\Arundale\Ais Decoder\Settings"
Name: "{commonappdata}\Arundale\Ais Decoder\Logs"
Name: "{commonappdata}\Arundale\Ais Decoder\Output"
Name: "{commonappdata}\Arundale\Ais Decoder\Files"
Name: "{commonappdata}\Arundale\Ais Decoder\Templates"
[Files]
; begin VB system files
; (Note: Scroll to the right to see the full lines!)
Source: "vbfiles\stdole2.tlb";  DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regtypelib
Source: "vbfiles\msvbvm60.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "vbfiles\oleaut32.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "vbfiles\olepro32.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "vbfiles\asycfilt.dll"; DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile
Source: "vbfiles\comcat.dll";   DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "vbfiles\scrrun.dll";   DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
Source: "vbfiles\advpack.dll";   DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile
Source: "vbfiles\capicom.dll";   DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
;Source: "vbfiles\imagehlp.dll";   DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
;Source: "vbfiles\imagehlp.dll";   DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
;Source: "vbfiles\wininet.dll";   DestDir: "{sys}"; OnlyBelowVersion: 0,6; Flags: restartreplace uninsneveruninstall sharedfile regserver
;This should not be redistributed (see Unsafe files in Help)
; end VB system files
;3rd party DLL must be copied for all versions
Source: "vbfiles\vbzip11.dll";   DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile
Source: "vbfiles\activelock3.6.dll";   DestDir: "{sys}"; Flags: restartreplace uninsneveruninstall sharedfile
Source: "{#MyAppName}"; DestDir: "{app}" ;flags: replacesameversion ignoreversion
;Source: "E:\My Documents\Ais\Decoder_v3\{#MyAppVersion}.txt"; DestDir: "{commonappdata}\Arundale\Ais Decoder\Files" ;flags: replacesameversion ignoreversion
Source: "arundale.ico"; DestDir: "{app}"  ;flags: replacesameversion ignoreversion
;have to create first before creating read only otherwise it fails - doesnt work
;Source: "C:\Documents and Settings\All Users\Application Data\Arundale\Ais Decoder\Settings\default.ini"; DestDir: "{commonappdata}\Arundale\Ais Decoder\Settings"  ;flags: replacesameversion ignoreversion
;keep a copy of the default ini file with the application so if deleted in commonappdata we can copy the default file
Source: "C:\Documents and Settings\All Users\Application Data\Arundale\Ais Decoder\Settings\default.ini"; DestDir: "{app}"  ;flags: replacesameversion ignoreversion
;all other files and rhe working copy of default.ini
;go in commonappdata which is All Users profile if user has
;adminitrative priviledges or cuurent user profile if no admin priviledges
;when the set-up program is run
;Initialisation Files
Source: "{#MyProgramData}\Arundale\Ais Decoder\Settings\aspx.ini"; DestDir: "{commonappdata}\Arundale\Ais Decoder\Settings"  ;flags: replacesameversion ignoreversion
Source: "{#MyProgramData}\Application Data\Arundale\Ais Decoder\Settings\aspx.ini"; DestDir: "{commonappdata}\Arundale\Ais Decoder\Settings"  ;flags: replacesameversion ignoreversion
Source: "{#MyProgramData}\Arundale\Ais Decoder\Settings\CsvAll.ini"; DestDir: "{commonappdata}\Arundale\Ais Decoder\Settings"  ;flags: replacesameversion ignoreversion
Source: "{#MyProgramData}\Arundale\Ais Decoder\Settings\default.ini"; DestDir: "{commonappdata}\Arundale\Ais Decoder\Settings"  ;flags: replacesameversion ignoreversion
Source: "{#MyProgramData}\Arundale\Ais Decoder\Settings\Excel.ini"; DestDir: "{commonappdata}\Arundale\Ais Decoder\Settings"  ;flags: replacesameversion ignoreversion
Source: "{#MyProgramData}\Arundale\Ais Decoder\Settings\GoogleEarth.ini"; DestDir: "{commonappdata}\Arundale\Ais Decoder\Settings"  ;flags: replacesameversion ignoreversion
Source: "{#MyProgramData}\Arundale\Ais Decoder\Settings\GoogleEarthOverlay.ini"; DestDir: "{commonappdata}\Arundale\Ais Decoder\Settings"  ;flags: replacesameversion ignoreversion
Source: "{#MyProgramData}\Arundale\Ais Decoder\Settings\GoogleMaps.ini"; DestDir: "{commonappdata}\Arundale\Ais Decoder\Settings"  ;flags: replacesameversion ignoreversion
Source: "{#MyProgramData}\Arundale\Ais Decoder\Settings\Html.ini"; DestDir: "{commonappdata}\Arundale\Ais Decoder\Settings"  ;flags: replacesameversion ignoreversion
Source: "{#MyProgramData}\Arundale\Ais Decoder\Settings\spnmea.ini"; DestDir: "{commonappdata}\Arundale\Ais Decoder\Settings"  ;flags: replacesameversion ignoreversion
Source: "{#MyProgramData}\Arundale\Ais Decoder\Settings\UdpTagsRange.ini" ; DestDir: "{commonappdata}\Arundale\Ais Decoder\Settings" ;flags: replacesameversion ignoreversion
;Templates
Source: "{#MyProgramData}\Arundale\Ais Decoder\Logs\aismsgs.dat"; DestDir: "{commonappdata}\Arundale\Ais Decoder\Logs"  ;flags: replacesameversion ignoreversion
Source: "{#MyProgramData}\Arundale\Ais Decoder\Templates\example.aspx"; DestDir: "{commonappdata}\Arundale\Ais Decoder\Templates"  ;flags: replacesameversion ignoreversion
Source: "{#MyProgramData}\Arundale\Ais Decoder\Templates\example.html"; DestDir: "{commonappdata}\Arundale\Ais Decoder\Templates"  ;flags: replacesameversion ignoreversion
Source: "{#MyProgramData}\Arundale\Ais Decoder\Templates\data.kml"; DestDir: "{commonappdata}\Arundale\Ais Decoder\Templates"  ;flags: replacesameversion ignoreversion
Source: "{#MyProgramData}\Arundale\Ais Decoder\Templates\data.xml"; DestDir: "{commonappdata}\Arundale\Ais Decoder\Templates"  ;flags: replacesameversion ignoreversion
Source: "{#MyProgramData}\Arundale\Ais Decoder\Templates\GoogleEarth.kml"; DestDir: "{commonappdata}\Arundale\Ais Decoder\Templates"  ;flags: replacesameversion ignoreversion
Source: "{#MyProgramData}\Arundale\Ais Decoder\Templates\GoogleEarthLink.kml"; DestDir: "{commonappdata}\Arundale\Ais Decoder\Templates"  ;flags: replacesameversion ignoreversion
Source: "{#MyProgramData}\Arundale\Ais Decoder\Templates\GoogleMaps.kml"; DestDir: "{commonappdata}\Arundale\Ais Decoder\Templates"  ;flags: replacesameversion ignoreversion
;images
Source: "{#MyProgramData}\Arundale\Ais Decoder\Output\ship1.png"; DestDir: "{commonappdata}\Arundale\Ais Decoder\Output"  ;flags: replacesameversion ignoreversion
Source: "{#MyProgramData}\Arundale\Ais Decoder\Output\triangle.png"; DestDir: "{commonappdata}\Arundale\Ais Decoder\Output"  ;flags: replacesameversion ignoreversion
Source: "{#MyProgramData}\Arundale\Ais Decoder\Output\square.png"; DestDir: "{commonappdata}\Arundale\Ais Decoder\Output"  ;flags: replacesameversion ignoreversion
Source: "{#MyProgramData}\Arundale\Ais Decoder\Output\ScreenOverlay.png"; DestDir: "{commonappdata}\Arundale\Ais Decoder\Output"  ;flags: replacesameversion ignoreversion
;help
Source: "Help\AisDecoder.chm"; DestDir: "{app}\Help"  ;flags: replacesameversion ignoreversion
;above need uncommenting
;Source: "C:\Documents and Settings\All Users\\Application Data\Arundale\Ais Decoder\default.ini"; DestDir: "{commonappdata}\Arundale\Ais Decoder"  ;Attribs: readonly ;flags: replacesameversion ignoreversion
;Source: "C:\Documents and Settings\All Users\\Application Data\Arundale\Ais Decoder\blank_ais_messages.log"; DestDir: "{commonappdata}\Arundale\Ais Decoder"  ;Attribs: readonly ;flags: replacesameversion ignoreversion
;Help file Source: "MyProg.chm"; DestDir: "{app}"
Source: "Readme.txt"; DestDir: "{app}"; Flags: isreadme ignoreversion

;These appear not to be distributed in all versions of windows
Source: "{#MySys32}\MSFlxGrd.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "{#MySys32}\mscomctl.OCX"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "{#MySys32}\MSWINSCK.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "{#MySys32}\MSINET.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "{#MySys32}\ComDlg32.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
Source: "{#MySys32}\MSComm32.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver
;Windows 8
Source: "{#MySys32}\richtx32.ocx"; DestDir: "{sys}"; Flags: restartreplace sharedfile regserver

;Config file (One Created in Users) - this is the file when the installation.exe is created
;Source: "C:\Documents and Settings\jna\Application Data\Arundale\Ais Decoder\Users\Setup.cfg"; DestDir: "{userappdata}\Arundale\Ais Decoder"  ;flags: replacesameversion ignoreversion

[Icons]
Name: "{group}\Ais Decoder"; Filename: "{app}\AisDecoder.exe"; IconFilename:"{app}\arundale.ico"
; NOTE: Most apps do not need registry entries to be pre-created. If you
; don't know what the registry is or if you need to use it, then chances are
; you don't need a [Registry] section.
;Name: "{userdocs}\Ais Decoder"; Filename: "{userappdata}\Arundale\Ais Decoder\Output"; Flags: foldershortcut ; IconFilename:"{app}\arundale.ico" ;Comment:"AisDecoder Files"

[InstallDelete]
Type: files; Name: "{app}\AisDecoder.exe"

[Registry]
; Start "Software\My Company\My Program" keys under HKEY_CURRENT_USER
; and HKEY_LOCAL_MACHINE. The flags tell it to always delete the
; "My Program" keys upon uninstall, and delete the "My Company" keys
; if there is nothing left in them.
;Root: HKCU; Subkey: "Software\Arundale"; Flags: uninsdeletekeyifempty
;Root: HKCU; Subkey: "Software\Arundale\AisDecoder"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Arundale"; Flags: uninsdeletekeyifempty
Root: HKLM; Subkey: "Software\Arundale\AisDecoder"; Flags: uninsdeletekey
Root: HKLM; Subkey: "Software\Arundale\AisDecoder\Settings"; ValueType: string; ValueName: "InstalledAppName"; ValueData: {#MyAppName}
Root: HKLM; Subkey: "Software\Arundale\AisDecoder\Settings"; ValueType: string; ValueName: "InstalledAppDateTime"; ValueData: {#MyFileDateTimeString}
Root: HKLM; Subkey: "Software\Arundale\AisDecoder\Settings"; ValueType: string; ValueName: "InstalledAppVersion"; ValueData: {#MyAppVersion}
Root: HKLM; Subkey: "Software\Arundale\AisDecoder\Settings"; ValueType: string; ValueName: "InstallDateTime"; ValueData: {#MyDateTimeString}
;Root: HKLM; Subkey: "Software\Arundale\AisDecoder\Settings"; ValueType: expandsz; ValueName: "AllUsersPath"; ValueData: "%ALLUSERSPROFILE%\Application Data\Arundale\Ais Decoder" ; flags: createvalueifdoesntexist
;Root: HKLM; Subkey: "Software\Arundale\AisDecoder\Settings"; ValueType: expandsz; ValueName: "NewUserInitialisationFile"; ValueData: "%ALLUSERSPROFILE%\Application Data\Arundale\Ais Decoder\Settings\default.ini" ; flags: createvalueifdoesntexist
Root: HKLM; Subkey: "Software\Arundale\AisDecoder\Settings"; ValueType: expandsz; ValueName: "AllUsersPath"; ValueData: "%ALLUSERSPROFILE%\Application Data\Arundale\Ais Decoder"
Root: HKLM; Subkey: "Software\Arundale\AisDecoder\Settings"; ValueType: expandsz; ValueName: "NewUserInitialisationFile"; ValueData: "%ALLUSERSPROFILE%\Application Data\Arundale\Ais Decoder\Settings\default.ini"
;set a key for the FallBack file
; Incorrect wrong if user changes default directory
;Root: HKLM; Subkey: "Software\Arundale\AisDecoder\Settings"; ValueType: expandsz; ValueName: "FallBackInitialisationFile"; ValueData: "%PROGRAMFILES%\Arundale\Ais Decoder\default.ini"
Root: HKLM; Subkey: "Software\Arundale\AisDecoder\Settings"; ValueType: expandsz; ValueName: "FallBackInitialisationFile"; ValueData: "{app}\default.ini"
;set a key for all the reserved files (my templates)
Root: HKLM; Subkey: "Software\Arundale\AisDecoder\Settings\ReservedFiles"; ValueType: string; ValueName: "0"; ValueData: "\Settings\default.ini" ;
Root: HKLM; Subkey: "Software\Arundale\AisDecoder\Settings\ReservedFiles"; ValueType: string; ValueName: "1"; ValueData: "\Settings\CsvAll.ini" ;
Root: HKLM; Subkey: "Software\Arundale\AisDecoder\Settings\ReservedFiles"; ValueType: string; ValueName: "2"; ValueData: "\Templates\example.html" ;
Root: HKLM; Subkey: "Software\Arundale\AisDecoder\Settings\ReservedFiles"; ValueType: string; ValueName: "3"; ValueData: "\Templates\example.aspx" ;
Root: HKLM; Subkey: "Software\Arundale\AisDecoder\Settings\ReservedFiles"; ValueType: string; ValueName: "4"; ValueData: "\Settings\UdpTagsRange.ini" ;
Root: HKLM; Subkey: "Software\Arundale\AisDecoder\Settings\ReservedFiles"; ValueType: string; ValueName: "5"; ValueData: "\Settings\GoogleEarth.ini" ;
Root: HKLM; Subkey: "Software\Arundale\AisDecoder\Settings\ReservedFiles"; ValueType: string; ValueName: "6"; ValueData: "\Templates\GoogleEarth.kml" ;
Root: HKLM; Subkey: "Software\Arundale\AisDecoder\Settings\ReservedFiles"; ValueType: string; ValueName: "7"; ValueData: "\Output\ship1.png" ;
Root: HKLM; Subkey: "Software\Arundale\AisDecoder\Settings\ReservedFiles"; ValueType: string; ValueName: "8"; ValueData: "\Settings\GoogleMaps.ini" ;
Root: HKLM; Subkey: "Software\Arundale\AisDecoder\Settings\ReservedFiles"; ValueType: string; ValueName: "9"; ValueData: "\Templates\GoogleMaps.kml" ;
Root: HKLM; Subkey: "Software\Arundale\AisDecoder\Settings\ReservedFiles"; ValueType: string; ValueName: "10"; ValueData: "\Templates\data.kml" ;
Root: HKLM; Subkey: "Software\Arundale\AisDecoder\Settings\ReservedFiles"; ValueType: string; ValueName: "11"; ValueData: "\Templates\data.xml" ;
Root: HKLM; Subkey: "Software\Arundale\AisDecoder\Settings\ReservedFiles"; ValueType: string; ValueName: "12"; ValueData: "\Logs\aismsgs.dat" ;
Root: HKLM; Subkey: "Software\Arundale\AisDecoder\Settings\ReservedFiles"; ValueType: string; ValueName: "13"; ValueData: "\Settings\Excel.ini" ;
Root: HKLM; Subkey: "Software\Arundale\AisDecoder\Settings\ReservedFiles"; ValueType: string; ValueName: "14"; ValueData: "\Settings\GoogleEarthOverlay.ini" ;
Root: HKLM; Subkey: "Software\Arundale\AisDecoder\Settings\ReservedFiles"; ValueType: string; ValueName: "15"; ValueData: "\Templates\GoogleEarthLink.kml" ;
Root: HKLM; Subkey: "Software\Arundale\AisDecoder\Settings\ReservedFiles"; ValueType: string; ValueName: "16"; ValueData: "\Settings\CsvUdpTags.ini" ;

;setting for NewVersion (2015) 
Root: HKLM; Subkey: "Software\Arundale\AisDecoder\Settings"; ValueType: string; ValueName: "LastVersion"; ValueData: "{#MyAppVersion}"

;set initialisation file for All Users (First time program is run)
Root: HKCU; Subkey: "Software\Arundale\AisDecoder\Settings"; ValueType: string; ValueName: "InstalledIniFileDateTime"; ValueData: {#MyDateTimeString}
Root: HKCU; Subkey: "Software\Arundale\AisDecoder\Settings"; ValueType: expandsz; ValueName: "InitialisationFile"; ValueData: "%APPDATA%\Arundale\Ais Decoder\Settings\default.ini" ; flags: createvalueifdoesntexist
Root: HKCU; Subkey: "Software\Arundale\AisDecoder\Settings"; ValueType: expandsz; ValueName: "CurrentUserPath"; ValueData: "%APPDATA%\Arundale\Ais Decoder" ; flags: createvalueifdoesntexist

[Run]
;Filename: "{app}\license.txt"; Description: "View the README file"; Flags: postinstall shellexec unchecked skipifsilent
Filename: "{app}\AisDecoder.exe"; Description: "Launch application"; Flags: postinstall nowait skipifsilent

