VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVessel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Storage for the debug ID.
Private mlngDebugID As Long
'used to keep ship names and positions
'Only Position or Static Data messges update vessels
Public Mmsi As String   'key
Public Name As String   'From Static messages (5,21,24)
Public RcvTime As String    'Rcv Time of Last message to update
'Public NameRcvTime As String
'dec separator problem
Public LastLat As Single    'From position messages (1,2,3,4,9,11,17,18,19,21,27)
Public LastLon As Single
'Public PositionTime As String
Public strTags As String    'tag(Tagno) as csv
Public VesselTagsStart As Long

'Where strTags starts in Vessels.dat csv record
'Public Property Get VesselTagsStart() As Long
'    VesselTagsStart = 5
'End Property

Private Sub Class_Initialize()
   mlngDebugID = DebugSerial
   ' Add a string entry to the global collection.
   gcolDebug.Add "clsVessel; DebugID=" _
   & mlngDebugID, CStr(mlngDebugID)
    
VesselTagsStart = 5

End Sub

Private Sub Class_Terminate()
   ' Remove the string entry, so you know the object
   ' isn't around any more.
   gcolDebug.Remove CStr(mlngDebugID)

End Sub
