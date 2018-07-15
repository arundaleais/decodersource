Attribute VB_Name = "modSysCalls"
Option Explicit

Private UACValue   As Long  'The Saved UAC Value, used by DisableUAC and RestoreUAC
Private Const MaxCount = 600000000000000#   '600*10^12 = 600TBytes
'http://support.microsoft.com/kb/189323/en-us
'http://www.experts-exchange.com/Programming/Languages/Visual_Basic/Q_27261445.html
Private Const OFFSET_4 = 4294967295#    'Note Microsoft is wrong (4294967296#)
Private Const MAXINT_4 = 2147483647
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

Public Enum mceIDLPaths
    CSIDL_ALTSTARTUP = &H1D                 '    * CSIDL_ALTSTARTUP - File system directory that corresponds to the user's nonlocalized Startup program group. (All Users\Startup?)
    CSIDL_APPDATA = &H1A                    '    * CSIDL_APPDATA - File system directory that serves as a common repository for application-specific data. A common path is C:\WINNT\Profiles\username\Application Data.
    CSIDL_BITBUCKET = &HA                   '    * CSIDL_BITBUCKET - Virtual folder containing the objects in the user's Recycle Bin.
    CSIDL_COMMON_ALTSTARTUP = &H1E          '    * CSIDL_COMMON_ALTSTARTUP - File system directory that corresponds to the nonlocalized Startup program group for all users. Valid only for Windows NT systems.
    CSIDL_COMMON_APPDATA = &H23             '    * CSIDL_COMMON_APPDATA - Version 5.0. Application data for all users. A common path is C:\WINNT\Profiles\All Users\Application Data.
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19    '    * CSIDL_DESKTOPDIRECTORY - File system directory used to physically store file objects on the desktop (not to be confused with the desktop folder itself). A common path is C:\WINNT\Profiles\username\Desktop
    CSIDL_COMMON_DOCUMENTS = &H2E           '    * CSIDL_COMMON_DOCUMENTS - File system directory that contains documents that are common to all users. A common path is C:\WINNT\Profiles\All Users\Documents. Valid only for Windows NT systems.
    CSIDL_COMMON_FAVORITES = &H1F           '    * CSIDL_COMMON_FAVORITES - File system directory that serves as a common repository for all users' favorite items. Valid only for Windows NT systems.
    CSIDL_COMMON_PROGRAMS = &H17            '    * CSIDL_COMMON_PROGRAMS - File system directory that contains the directories for the common program groups that appear on the Start menu for all users. A common path is c:\WINNT\Profiles\All Users\Start Menu\Programs. Valid only for Windows NT systems.
    CSIDL_COMMON_STARTMENU = &H16           '    * CSIDL_COMMON_STARTMENU - File system directory that contains the programs and folders that appear on the Start menu for all users. A common path is C:\WINNT\Profiles\All Users\Start Menu. Valid only for Windows NT systems.
    CSIDL_COMMON_STARTUP = &H18             '    * CSIDL_COMMON_STARTUP - File system directory that contains the programs that appear in the Startup folder for all users. A common path is C:\WINNT\Profiles\All Users\Start Menu\Programs\Startup. Valid only for Windows NT systems.
    CSIDL_COMMON_TEMPLATES = &H2D           '    * CSIDL_COMMON_TEMPLATES - File system directory that contains the templates that are available to all users. A common path is C:\WINNT\Profiles\All Users\Templates. Valid only for Windows NT systems.
    CSIDL_COOKIES = &H21                    '    * CSIDL_COOKIES - File system directory that serves as a common repository for Internet cookies. A common path is C:\WINNT\Profiles\username\Cookies.
    CSIDL_DESKTOPDIRECTORY = &H10           '    * CSIDL_COMMON_DESKTOPDIRECTORY - File system directory that contains files and folders that appear on the desktop for all users. A common path is C:\WINNT\Profiles\All Users\Desktop. Valid only for Windows NT systems.
    CSIDL_FAVORITES = &H6                   '    * CSIDL_FAVORITES - File system directory that serves as a common repository for the user's favorite items. A common path is C:\WINNT\Profiles\username\Favorites.
    CSIDL_FONTS = &H14                      '    * CSIDL_FONTS - Virtual folder containing fonts. A common path is C:\WINNT\Fonts.
    CSIDL_HISTORY = &H22                    '    * CSIDL_HISTORY - File system directory that serves as a common repository for Internet history items.
    CSIDL_INTERNET_CACHE = &H20             '    * CSIDL_INTERNET_CACHE - File system directory that serves as a common repository for temporary Internet files. A common path is C:\WINNT\Profiles\username\Temporary Internet Files.
    CSIDL_LOCAL_APPDATA = &H1C              '    * CSIDL_LOCAL_APPDATA - Version 5.0. File system directory that serves as a data repository for local (non-roaming) applications. A common path is C:\WINNT\Profiles\username\Local Settings\Application Data.
    CSIDL_PROGRAMS = &H2                    '    * CSIDL_PROGRAMS - File system directory that contains the user's program groups (which are also file system directories). A common path is C:\WINNT\Profiles\username\Start Menu\Programs.
    CSIDL_PROGRAM_FILES = &H26              '    * CSIDL_PROGRAM_FILES - Version 5.0. Program Files folder. A common path is C:\Program Files.
    CSIDL_PROGRAM_FILES_COMMON = &H2B       '    * CSIDL_PROGRAM_FILES_COMMON - Version 5.0. A folder for components that are shared across applications. A common path is C:\Program Files\Common. Valid only for Windows NT and Windows® 2000 systems.
    CSIDL_PERSONAL = &H5                    '    * CSIDL_PERSONAL - File system directory that serves as a common repository for documents. A common path is C:\WINNT\Profiles\username\My Documents.
    CSIDL_RECENT = &H8                      '    * CSIDL_RECENT - File system directory that contains the user's most recently used documents. A common path is C:\WINNT\Profiles\username\Recent. To create a shortcut in this folder, use SHAddToRecentDocs. In addition to creating the shortcut, this function updates the shell's list of recent documents and adds the shortcut to the Documents submenu of the Start menu.
    CSIDL_SENDTO = &H9                      '    * CSIDL_SENDTO - File system directory that contains Send To menu items. A common path is c:\WINNT\Profiles\username\SendTo.
    CSIDL_STARTUP = &H7                     '    * CSIDL_STARTUP - File system directory that corresponds to the user's Startup program group. The system starts these programs whenever any user logs onto Windows NT or starts Windows® 95. A common path is C:\WINNT\Profiles\username\Start Menu\Programs\Startup.
    CSIDL_STARTMENU = &HB                   '    * CSIDL_STARTMENU - File system directory containing Start menu items. A common path is c:\WINNT\Profiles\username\Start Menu.
    CSIDL_SYSTEM = &H25                     '    * CSIDL_SYSTEM - Version 5.0. System folder. A common path is C:\WINNT\SYSTEM32.
    CSIDL_TEMPLATES = &H15                  '    * CSIDL_TEMPLATES - File system directory that serves as a common repository for document templates.
    CSIDL_WINDOWS = &H24                    '    * CSIDL_WINDOWS - Version 5.0. Windows directory or SYSROOT. This corresponds to the %windir% or %SYSTEMROOT% environment variables. A common path is C:\WINNT.
End Enum

Public Declare Function SHGetSpecialFolderPath Lib "SHELL32.DLL" Alias "SHGetSpecialFolderPathA" (ByVal hwnd As Long, ByVal lpszPath As String, ByVal nFolder As Integer, ByVal fCreate As Boolean) As Boolean
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwnd As Long) As Long

   Private Type STARTUPINFO
      cb As Long
      lpReserved As String
      lpDesktop As String
      lpTitle As String
      dwX As Long
      dwY As Long
      dwXSize As Long
      dwYSize As Long
      dwXCountChars As Long
      dwYCountChars As Long
      dwFillAttribute As Long
      dwFlags As Long
      wShowWindow As Integer
      cbReserved2 As Integer
      lpReserved2 As Long
      hStdInput As Long
      hStdOutput As Long
      hStdError As Long
   End Type

   Private Type PROCESS_INFORMATION
      hProcess As Long
      hThread As Long
      dwProcessID As Long
      dwThreadID As Long
   End Type

   Private Declare Function WaitForSingleObject Lib "KERNEL32" (ByVal _
      hHandle As Long, ByVal dwMilliseconds As Long) As Long

   Private Declare Function CreateProcessA Lib "KERNEL32" (ByVal _
      lpApplicationName As String, ByVal lpCommandLine As String, ByVal _
      lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
      ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
      ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
      lpStartupInfo As STARTUPINFO, lpProcessInformation As _
      PROCESS_INFORMATION) As Long

   Private Declare Function CloseHandle Lib "KERNEL32" _
      (ByVal hObject As Long) As Long

   Private Declare Function GetExitCodeProcess Lib "KERNEL32" _
      (ByVal hProcess As Long, lpExitCode As Long) As Long

Private Declare Sub CopyMemory Lib "KERNEL32" Alias _
    "RtlMoveMemory" (dest As Any, src As Any, ByVal nbytes _
    As Long)
   
   Private Const NORMAL_PRIORITY_CLASS = &H20&
   Private Const INFINITE = -1&

Declare Function FindWindow32 Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

'This is required as ShellExecute is aynchronous and
'we need a synchronous shelled process
   Public Function ExecCmd(cmdLine$)
      Dim proc As PROCESS_INFORMATION
      Dim start As STARTUPINFO
        Dim Ret As Long
    Dim cmd As String
    Dim ProcessDir As String
'cmdline = "/k " & Environ("PROGRAMFILES") & "\Arundale\NmeaRouter\com0com\com0com_setup_driver.bat"
'cmd = "cmd.exe"
'cmd = vbNullString  'important not vbnull
'Direct = Environ("PROGRAMFILES") & "\Arundale\NmeaRouter\com0com"
'cmdline = "dir"
      ' Initialize the STARTUPINFO structure:
'    frmrouter.
'Declare Function LockWindowUpdate Lib "user32" (ByVal hwnd As Long) As Long
'    ret = LockWindowUpdate(frmRouter.hwnd)
      
'        frmRouter.Refresh
      start.cb = Len(start)

      ' Start the shelled application:
'      ret = CreateProcessA(vbNullString, cmdline$, 0&, 0&, 1&, _
'         NORMAL_PRIORITY_CLASS, 0&, vbNullString, start, proc)
'next ok
'cmdline = "cmd dir"
'      ret = CreateProcessA(vbNullString, cmdline$, 0&, 0&, 1&, _
'         NORMAL_PRIORITY_CLASS, 0&, Direct, start, proc)
'next ok
'cmd = vbNullString  'important not vbnull
'      ret = CreateProcessA(cmd$, cmdline$, 0&, 0&, 1&, _
'         NORMAL_PRIORITY_CLASS, 0&, Environ("PROGRAMFILES") & "\Arundale\NmeaRouter\com0com", start, proc)

'cmdline = "com0com_setup_driver.bat"
'next ok
    cmd = vbNullString

'Command prompt created in Program Files
'/C closes the Console when finished /K doesnt
'below works fine
'cmdline = "cmd.exe /C ""setupc remove 0""" 'on its own creates command prompt"
'cmdline = "cmd.exe /C ""setupc install PortName=- PortName=COM#""" 'on its own creates command prompt"
'cmdline = "cmd.exe /C ""reg add HKLM\Software\Policies\Microsoft\Windows\DeviceInstall\Settings /v SuppressNewHWUI /t REG_DWORD /d 1 /f"""
'cmdline = "cmd.exe /C ""reg add HKLM\Software\Policies\Microsoft\Windows\DeviceInstall\Settings /v SuppressNewHWUI /t REG_DWORD /d 0 /f"""
'cmdline = "cmd.exe /C ""setup_com0com_x86.exe /S /D=%ProgramFiles%\Arundale\NmeaRouter\com0com"""
'    ProcessDir = Environ("PROGRAMFILES") & "\Arundale\NmeaRouter\com0com"
'
'NmeaRouter        ProcessDir = App.path & "\com0com"
'ProcessDir = App.path
ProcessDir = PathFromFullName(ShellFileName)
        Ret = CreateProcessA(cmd$, cmdLine$, 0&, 0&, 1&, _
         NORMAL_PRIORITY_CLASS, 0&, ProcessDir, start, proc)
        If Ret = 0 Then
        MsgBox GetLastSystemError
        End If

      ' Wait for the shelled application to finish:
         Ret = WaitForSingleObject(proc.hProcess, INFINITE)
'MsgBox "Completed"
         Call GetExitCodeProcess(proc.hProcess, Ret&)
         Call CloseHandle(proc.hThread)
         Call CloseHandle(proc.hProcess)
         ExecCmd = Ret
' ret = LockWindowUpdate(0)
'        frmRouter.Refresh
   End Function

Public Function GetSpecialFolderA(ByVal eSpecialFolder As mceIDLPaths) As String
On Error GoTo Hell

Dim Ret As Long
Dim Trash As String: Trash = Space$(260)

    Ret = SHGetSpecialFolderPath(0, Trash, eSpecialFolder, False)
    If Trim$(Trash) <> Chr(0) Then Trash = Left$(Trash, InStr(Trash, Chr(0)) - 1) & "\"
    
    GetSpecialFolderA = Trash
   
Hell:
End Function

Public Function GetLastSystemError() As String

Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Dim sError As String * 500 '\\ Preinitilise a string buffer to put any error message into
Dim lErrNum As Long
Dim lErrMsg As Long

lErrNum = err.LastDllError

lErrMsg = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, ByVal 0&, lErrNum, 0, sError, Len(sError), 0)

GetLastSystemError = Trim(sError)

End Function

Public Function DisableUAC()
Dim kb As String
Dim Ret As Long
'Get initial value of UAC
        UACValue = QueryValue(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Windows\CurrentVersion\Policies\System", "EnableLUA")
        If UACValue = 0 Then
            WriteStartUpLog "UAC is not set"
        Else
            WriteStartUpLog "UAC is set"
'Disable UAC
            Call SetKeyValue(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Windows\CurrentVersion\Policies\System", "EnableLUA", 0, REG_DWORD)
            WriteStartUpLog "UAC has been disabled"
        End If
'Disable New Hardware prompt
        Call SetKeyValue(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Windows\DeviceInstall\Settings", "SuppressNewHWUI", 1, REG_DWORD)
        WriteStartUpLog "New Hardware Prompt has been disabled"
'        kb = "cmd.exe /C ""reg add HKLM\Software\Policies\Microsoft\Windows\CurrentVersion\Policies\System /v EnableLUA /t REG_DWORD /d 0 /f"""
'        ret = ExecCmd(kb)
'        kb = "cmd.exe /C ""reg add HKLM\Software\Policies\Microsoft\Windows\DeviceInstall\Settings /v SuppressNewHWUI /t REG_DWORD /d 1 /f"""
'        ret = ExecCmd(kb)
End Function

Public Function RestoreUAC()
'Re-enable add New Hardware prompt
        If UACValue = 1 Then
            Call SetKeyValue(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Windows\DeviceInstall\Settings", "SuppressNewHWUI", 0, REG_DWORD)
            WriteStartUpLog "UAC is re-enabled"
        End If
'        kb = "cmd.exe /C ""reg add HKLM\Software\Policies\Microsoft\Windows\DeviceInstall\Settings /v SuppressNewHWUI /t REG_DWORD /d 0 /f"""
'        ret = ExecCmd(kb)
're-set UAC to original value UAC
        Call SetKeyValue(HKEY_LOCAL_MACHINE, "Software\Policies\Microsoft\Windows\CurrentVersion\Policies\System", "EnableLUA", UACValue, REG_DWORD)
        WriteStartUpLog "New Hardware Prompt has been re-enabled"
'        kb = "cmd.exe /C ""reg add HKLM\Software\Policies\Microsoft\Windows\CurrentVersion\Policies\System /v EnableLUA /t REG_DWORD /d " & UACValue & " /f"""
'        ret = ExecCmd(kb)
 End Function

'Incr is used to prevent overflow errors or debug overflow
Public Function incr(ByRef var As Variant)  'Variand so it workd for long,single and double
    If IsNumeric(var) Then
        On Error GoTo Reset
        var = var + MESSAGE_COUNT_STEP  'change from 1 to test
    End If
Exit Function
Reset:
    Call WriteErrorLog("Incr Reset " & CStr(var))
    var = 1     'not 0 because we dont want a div/0 error
End Function

Public Function UnsignedToLong(Value As Double) As Long
    If Value < 0 Or Value >= OFFSET_4 Then Error 6 ' Overflow
    If Value <= MAXINT_4 Then
        UnsignedToLong = Value
    Else
        UnsignedToLong = Value - OFFSET_4
    End If
End Function

'http://support.microsoft.com/kb/189323
Public Function LongToUnsigned(Value As Long) As Double
    If Value < 0 Then
        LongToUnsigned = Value + OFFSET_4
    Else
        LongToUnsigned = Value
    End If
End Function

#If False Then
'Must be modified to use double
Public Function ElapsedTickCount_old(StartTick As Long, Optional FinishTick As Long) As Long
Dim ElapsedTick As Long

StartTick = FinishTick + 1
If FinishTick >= StartTick Then
    ElapsedTickCount = FinishTick - StartTick
'60# is to prevent overflow (otherwise I think it assumes an integer)
Else
'    ElapsedTime = ElapsedTime + CSng(60# * 60# * 24#)
    ElapsedTickCount = 2 ^ 16 - 1 - StartTick + FinishTick
End If
End Function
#End If

Public Function ElapsedTickCount(StartTick As Long, Optional FinishTick As Long) As Long
Dim ElapsedTick As Long
Dim dblStartTick As Double
Dim dblFinishTick As Double
Dim dblElapsedTime As Double

    dblStartTick = LongToUnsigned(StartTick)
    dblFinishTick = LongToUnsigned(FinishTick)
    If dblFinishTick >= dblStartTick Then
        dblElapsedTime = dblFinishTick - dblStartTick
    Else    'Rollover after 49 days
        dblElapsedTime = dblStartTick - dblFinishTick
    End If
    ElapsedTickCount = UnsignedToLong(dblElapsedTime)
End Function

'http://www.devx.com/vb2themax/Tip/19007
' The standard deviation of an array of any type
'
' if the second argument is True or omitted,
' it evaluates the standard deviation of a sample,
' if it is False it evaluates the standard deviation of a population
'
' if the third argument is True or omitted, Empty values aren't accounted for

Public Function ArrayStdDev(arr As Variant, Optional SampleStdDev As Boolean = True, _
    Optional IgnoreEmpty As Boolean = True) As Double
    Dim sum As Double
    Dim sumSquare As Double
    Dim Value As Double
    Dim Count As Long
    Dim Index As Long

    ' evaluate sum of values
    ' if arr isn't an array, the following statement raises an error
    For Index = LBound(arr) To UBound(arr)
        Value = arr(Index)
        ' skip over non-numeric values
        If IsNumeric(Value) Then
            ' skip over empty values, if requested
               If Not (IgnoreEmpty And IsEmpty(Value)) Then
                 ' add to the running total
                   Count = Count + 1
                    sum = sum + Value
                    sumSquare = sumSquare + Value * Value
                End If
         End If
    Next

    ' evaluate the result
    ' use (Count-1) if evaluating the standard deviation of a sample
    If SampleStdDev Then
        ArrayStdDev = Sqr((sumSquare - (sum * sum / Count)) / (Count - 1))
    Else
        ArrayStdDev = Sqr((sumSquare - (sum * sum / Count)) / Count)
    End If
If Count Then
ArrayStdDev = sum / Count + 3 * ArrayStdDev
Else
ArrayStdDev = 0
End If
End Function

