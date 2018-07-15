Attribute VB_Name = "Locale"
Option Explicit

Private Declare Function GetLocaleInfo Lib "KERNEL32" _
Alias "GetLocaleInfoA" _
(ByVal Locale As Long, _
ByVal LCType As Long, _
ByVal lpLCData As String, _
ByVal cchData As Long) As Long
 
Private Const LOCALE_SDECIMAL = &HE
Private Declare Function GetThreadLocale Lib "KERNEL32" () As Long
Private Declare Function GetSystemDefaultLCID Lib "KERNEL32" () As Long
Private Declare Function GetUserDefaultLCID Lib "KERNEL32" () As Long
 
Public Function GetDecimalSep() As String
Dim LCID As Integer
Dim Data As String
Dim Ret As Integer
'from http://www.codeguru.com/forum/showthread.php?t=351810
Dim DataLen As Long
 
' Get the local decimal seperator
' Find the threads local
'LCID = GetThreadLocale jna does NOT return correct delimiter
LCID = GetUserDefaultLCID() 'jna uses same as in .main
' Find the required size of the output variables
Ret = GetLocaleInfo(LCID, LOCALE_SDECIMAL, Data, DataLen)
 
If Ret <> 0 Then
     ' prepare the output variable
     DataLen = Ret
     Data = Space(DataLen)
     Ret = GetLocaleInfo(LCID, LOCALE_SDECIMAL, Data, DataLen)
Else
     ' Error no data found
     ' enter some good error handling here, using GetLastError()
End If
' Remove the null terminator from the string
GetDecimalSep = Left(Data, DataLen - 1)
'MsgBox GetDecimalSep & "," & "." & "|"
End Function

