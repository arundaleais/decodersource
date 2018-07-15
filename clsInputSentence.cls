VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInputSentence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Storage for the debug ID.
Private mlngDebugID As Long
'This contains the decoded data for
'   1.  Input Filter
'   2.  Output CRC,NMEA and View
'these must be defined public in the class module (if not using properies)
Public FullSentence As String   'Includes Comment Block
Public NmeaPrefix As String     'Part of sentence before ! or $ if no comment block
'Public CommentBlock As String   '/.../ at begining of sentence IEC61162-1
'Public CBerrmsg As String   'Comment Block error
Public NmeaSentence As String
Public NmeaSentenceType As String  'nmea word 0 sentence type eg !AITXT or $AITXT or $GPRMA
Public IecTalkerID As String                                    'eg AI or GP
Public IecFormat As String                                'eg TXT or RMA
Public EncapsulatingNmeaSentence As Boolean    'This sentence encapsulates
'a following AIS sentence
Public NmeaCrcWord As Long  'Word no containing NMEA CRC
Public NmeaCrc As String
Public CRCerrmsg As String
Public AisPayload As String         'Complete re-assembled payload 6-bit characters(word 6)
Public AisPayloadFillBits As Long   'No of unused bits at end of last 6-bit character
                                    'to get true size of 8 bit payload. This will be from
                                    '0-5 because if 6 (or more) to get to an 8-bit boundary
                                    'an additional 6 bit character would be used.
'It is only the final part of a multipart NMEA sentence that will have fill bits, because
'the transmitted sentence is a continuous bit stream across the data link slots.
'The AIS receiver splits it into NMEA sentences (because of the NMEA 80 character limit)
Public AisPayloadBits As Long   'True size of 8-bit payload taking NMEA fill bits into account
Public NmeaRcvTime As String    'nmea word 7 created time
Public PayloadReassemblerComments As String
Public PayloadReassemblerErr As Boolean
Public AisMsgPart1Missing As Boolean    'means we cannot get AisMsgType
Public AisMsgPartsComplete As Boolean   'this is the last part
                                        'and all prev received ok
Public AisPositionOK As Boolean  'True if Msg1,2,3,4,9,18,27 Lat/Lon OK
Public AisPayloadComments As String     'payload too short
Public IsAisSentence As Boolean   'True if AIS !aaVDO or !aaVDM
Public IsIecEncapsulated As Boolean   'True if !
Public IecEncapsulationComments As String

Public AisMsgType As String     'Number 1- 27
Public SentencePart As String   '0-9 nmea word (2)
Public AisMsgFromMmsi As String
Public AisMsgRepeat As String   'because NmeaAisOut same for all MsgTypes
Public AisMsgToMmsi As String
Public AisMsgDac As String
Public AisMsgFi As String
Public AisMsgFiId As String

Public VesselName As String
Public VesselLat As Single      'Obtained for every sentence
Public VesselLon As Single      'ditto
'Public LastKnownLat As Single
'Public LastKnownLon As Single

Private Sub Class_Initialize()
   mlngDebugID = DebugSerial
   ' Add a string entry to the global collection.
   gcolDebug.Add "clsInputSentence; DebugID=" _
   & mlngDebugID, CStr(mlngDebugID)

End Sub

Private Sub Class_Terminate()
   ' Remove the string entry, so you know the object
   ' isn't around any more.
   gcolDebug.Remove CStr(mlngDebugID)

End Sub
