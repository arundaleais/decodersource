VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInputCB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Storage for the debug ID.
Private mlngDebugID As Long

'This contains the decoded data for
'these must be defined public in the class module (if not using properies)
Public Block As String   '\...\ at begining of sentence IEC61162-1
Public errmsg As String   'Comment Block error
Public CRCerrmsg As String  'Crc error message (from NmeaCrcChk)
Public CbCrc As String  'Crc at end of comment block
Public Time As String   'Unix Time
Public Text As String   'Free form text
Public Source As String 'Source (Talker) ID
Public Destination As String
Public Counter As String    '
Public GroupLine As String  'Line No in this Group ID
Public GroupLines As String 'Total Lines in this Group
Public GroupId As String    'Unique Group ID
Public Unknown As String '"|" separated

Private Sub Class_Initialize()
   mlngDebugID = DebugSerial
   ' Add a string entry to the global collection.
   gcolDebug.Add "clsInputCB; DebugID=" _
   & mlngDebugID, CStr(mlngDebugID)

End Sub

Private Sub Class_Terminate()
   ' Remove the string entry, so you know the object
   ' isn't around any more.
   gcolDebug.Remove CStr(mlngDebugID)

End Sub
