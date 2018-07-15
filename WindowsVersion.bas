Attribute VB_Name = "WindowsVersion"
Option Explicit
'http://vbcity.com/forums/p/99944/422558.aspx
Private Declare Function GetVersionExA Lib "KERNEL32" _
               (lpVersionInformation As OSVERSIONINFO) As Integer
            Public Type OSVERSIONINFO
               dwOSVersionInfoSize As Long
               dwMajorVersion As Long
               dwMinorVersion As Long
               dwBuildNumber As Long
               dwPlatformId As Long
               szCSDVersion As String * 128
            End Type
Private Function LPSTRToVBString$(ByVal S$)
   Dim nullpos&
   nullpos& = InStr(S$, Chr$(0))
   If nullpos > 0 Then
      LPSTRToVBString = Left$(S$, nullpos - 1)
   Else
      LPSTRToVBString = ""
   End If
End Function
Public Function GetVersion1() As String
Dim osinfo As OSVERSIONINFO
Dim retvalue As Integer
osinfo.dwOSVersionInfoSize = 148
osinfo.szCSDVersion = Space$(128)
retvalue = GetVersionExA(osinfo)
With osinfo
    Select Case .dwPlatformId
    Case 1
        Select Case .dwMinorVersion
        Case 0
            GetVersion1 = "Windows 95"
        Case 10
            GetVersion1 = "Windows 98"
        Case 90
            GetVersion1 = "Windows Millenium"
        End Select
    Case 2
        Select Case .dwMajorVersion
        Case 3
            GetVersion1 = "Windows NT 3.51"
        Case 4
            GetVersion1 = "Windows NT 4.0"
        Case 5
            If .dwMinorVersion = 0 Then
                GetVersion1 = "Windows 2000"
            Else
                GetVersion1 = "Windows XP"
            End If
        Case 6
            Select Case .dwMinorVersion
            Case Is = 0
                GetVersion1 = "Windows Vista"
            Case Is = 1         'added from v6.2.136
                GetVersion1 = "Windows 7"
            Case Is = 2
                GetVersion1 = "Windows 8"
            Case Is = 3
                GetVersion1 = "Windows 8.1"
            End Select
        Case 10
            Select Case .dwMinorVersion
            Case Is = 0
                GetVersion1 = "Windows 10"
            Case Else
                GetVersion1 = "Windows 10" & .dwMinorVersion
            End Select
        End Select
    Case Else
        GetVersion1 = "Windows " & .dwMajorVersion & "." & .dwMinorVersion
    End Select
End With
End Function

