VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAisMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Does not output nmea info
'Public Mmsi As String   'can be from or to
'Public AisMsgType As String '0-63   XX
'Public SentencePart as string '0-9 X
'Public Dac As String    '0-1023 XXXX
'Public Fi As String     '0-63   XX
'Public FiId As String   '0-16   XX

' Storage for the debug ID.
Private mlngDebugID As Long
Public AisMsgKey As String
Public AisSentence As String    '
Public NmeaRcvTime As String  'used to determine dead messages
                            'duplicates words(7), but kept separately
                            'to avoid re-parsing the sentence
'example of how stored in collection
'045133333-06-0200-10-00-   'key (no duplicates)
'!AIVDM......,date/time

Private Sub Class_Initialize()
   mlngDebugID = DebugSerial
   ' Add a string entry to the global collection.
   gcolDebug.Add "clsAisMsg; DebugID=" _
   & mlngDebugID, CStr(mlngDebugID)

End Sub

Private Sub Class_Terminate()
   ' Remove the string entry, so you know the object
   ' isn't around any more.
   gcolDebug.Remove CStr(mlngDebugID)

End Sub
