Attribute VB_Name = "DecodeDefs"
Option Explicit
Public NmeaAisWordName(0 To 6) As String
Public NmeaAIVDMWordName(0 To 6) As String
Public NmeaGPZDAWordName(0 To 6) As String
Public NmeaGPGGAWordName(0 To 14) As String
Public NmeaGPRMCWordName(0 To 13) As String
Public NmeaPGHPWordName(0 To 13) As String
Public NmeaAITXTWordName(0 To 4) As String
Public NmeaAIALRWordName(0 To 5) As String
Public NmeaAITAGWordName(0 To 2) As String
Public NmeaAIABMWordName(0 To 8) As String  'Last field excludes CRC
Public NmeaAIBBMWordName(0 To 7) As String
Public NmeaAIBRMWordName(0 To 4) As String
Public NmeaPTHAJWordName(0 To 4) As String
Public AisBroadcastChannelName(0 To 3) As String
Public IecEncapsulatedWordName(0 To 5) As String 'any sentence starting with ! Min 5 elements
Public IecTXTIdentifierName(0 To 99) As String
Public IecALRIdentifierName(0 To 99) As String
Public AisMsgTypeName(0 To 63) As String    '6 bits
Public AisMsgTypeBitsMin(0 To 63) As Long
Public AisMsgTypeBitsMax(0 To 63) As Long
Public AisMsgTypeBitsFill(0 To 63) As Long  '-1 = (0-5), Nmea Fill Bits(before crc)
Public DacName(0 To 1023) As String  'Index is MID if 200 to 799
Public StatusName(0 To 15) As String
Public AccuracyName(0 To 1) As String   '1 bit
Public RaimName(0 To 1) As String
Public EpfdName(0 To 15) As String      '4 bits
Public TxControlName(0 To 1) As String
Public RepeatName(0 To 3) As String
Public Ais_VersionName(0 To 3) As String    '2 bits
Public Ship_TypeNxName(1 To 9) As String    '8 bits
Public Ship_TypexNName(0 To 9) As String    '8 bits
Public Ship_Type3NName(0 To 9) As String    '8 bits
Public Ship_Type5NName(0 To 9) As String    '8 bits
Public AtoNName(0 To 31) As String          '5 bits
Public StationTypeName(0 To 15) As String
Public TxrxName(0 To 15) As String          '4 bits in msg 22
Public IntervalName(0 To 15) As String
Public DteName(0 To 1) As String
Public CsName(0 To 1) As String
Public DisplayName(0 To 1) As String
Public DscName(0 To 1) As String
Public BandName(0 To 1) As String
Public Msg22Name(0 To 1) As String
Public Off_PositionName(0 To 1) As String
Public Virtual_AidName(0 To 1) As String
Public AddressedName(0 To 1) As String
Public StructuredName(0 To 1) As String
Public Band_WidthName(0 To 1) As String
Public PartnoName(0 To 1) As String
Public GnssPositionName(0 To 1) As String
Public RadioModeName(0 To 1) As String
Public SyncStateName(0 To 3) As String
Public SlotAllocName(0 To 7) As String
Public AltitudeName(0 To 1) As String
Public AssignedName(0 To 1) As String
Public PowerName(0 To 1) As String
Public IfiName(0 To 63) As String
Public TendencyName(0 To 3) As String
Public BeaufortName(0 To 15) As String
Public BeaufortEnvironmentalName(0 To 15) As String
Public YesNoName(0 To 3) As String
Public YesNo2Name(0 To 3) As String
Public PrecipitationName(0 To 7) As String
Public TargTypeName(0 To 3) As String
Public ConeName(0 To 7) As String
Public LoadedName(0 To 3) As String
Public QualSpeedName(0 To 1) As String
Public QualHeadName(0 To 1) As String
Public UsaCanFiName(0 To 63) As String
Public UsaCan1IdName(0 To 63) As String
Public UsaCan2IdName(0 To 63) As String
Public UsaCan32IdName(0 To 63) As String
Public ReportType366Name(0 To 15) As String
Public Owner366Name(0 To 15) As String
Public Timeout366Name(0 To 7) As String
Public SensorData366Name(0 To 7) As String
Public AiAvailable1Name(0 To 1) As String
Public AiResponse1Name(0 To 7) As String
Public SignalStatus1Name(0 To 1) As String
Public ItuWmoTendencyName(0 To 15) As String
Public ItuWmoWeatherName(0 To 15) As String
Public UpDown366Name(0 To 1) As String
Public Level366Name(0 To 1) As String
Public DatumA366Name(0 To 3) As String
Public DatumB366Name(0 To 31) As String
Public Tug200Name(0 To 7) As String
Public LockStatus200Name(0 To 3) As String
Public SignalForm200Name(0 To 15) As String
Public SignalImpact200Name(0 To 7) As String
Public SignalStatus200Name(0 To 7) As String
Public Weather200Name(0 To 15) As String
Public Category200Name(0 To 3) As String
Public Direction200Name(0 To 15) As String
Public Racon235Name(0 To 3) As String
Public Light235Name(0 To 3) As String
Public Alarm235Name(0 To 1) As String
Public Biit235Name(0 To 1) As String
Public ExtFld235Name(0 To 1) As String
Public HighLow235Name(0 To 1) As String
Public OnOffName(0 To 1) As String
Public DgUnitsName(0 To 3) As String
Public DgCodeName(0 To 15) As String
Public BcCodeName(0 To 15) As String
Public Marpol1Name(0 To 15) As String
Public Marpol2Name(0 To 7) As String
Public SalinityType366Name(0 To 3) As String
Public Shape1Name(0 To 7) As String
Public AreaType1Name(0 To 127) As String
Public Sender1Name(0 To 7) As String
Public Route1Name(0 To 31) As String
Public SolasEquipmentName(1 To 26) As String
Public IceClassName(0 To 15) As String
Public OperationalName(0 To 3) As String

#If False Then
Public Example999Name(0 To 7) As String
#End If

Public ArgFmt(0 To 4) As String

'set up here because it's called from more than one form
'des is used because the value must be formatted as well as the decription
Public Function MmsiFmt(ByVal Mmsi As Long, Optional ByVal Des As String) As String
Dim Mmsi9 As String 'formatted as 9 digits
Dim MidPos As Long
Dim MidDes As String
Dim MidNo As Long
Dim MmsiDes As String

Mmsi9 = Format$(Mmsi, "000000000")
If Des = "" Then
    MmsiFmt = Mmsi9
Else
        
'Dec 2014 version
    Select Case Left$(Mmsi9, 1)
    Case Is = "0"
        Select Case Left$(Mmsi9, 2)   '2nd digit
        Case Is = "00"       '00MIDxxxxx Coast station
            MidPos = 3
            MmsiDes = "Shore Station"
            If Mmsi9 = "009990000" Then 'Address all VHF coast stations
                MmsiDes = "All Shore Stations (Group)"
            End If
            If Left$(Mmsi9, 5) = "00000" Then
                MmsiDes = "Reserved for Testing"
            End If
        Case "02" To "07"     '0MIDxxxxx Vessel Group Call
            MidPos = 2
            MmsiDes = "Group Call"
        End Select
    Case Is = "1"
        Select Case Left$(Mmsi9, 3)   '2nd & 3rd
        Case Is = "111"      '111MID000 SAR Aircraft
            MidPos = 4
            MmsiDes = "SAR Aircraft"
            If Right$(Mmsi9, 3) = "000" Then    'SAR Aircraft Group Call
                MmsiDes = "SAR Aircraft (Group)"
            End If
        End Select
    Case "2" To "7"     'Mid for vessel
        MidPos = 1
    Case Is = "8"       '8MIDxxxxx  Handheld VHF transceiver with DSC and GNSS
        MidPos = 2
        MmsiDes = "Portable VHF/DSC/GNSS"
    Case Is = "9"
        Select Case Left$(Mmsi9, 2)   '2nd digit
        Case Is = "97"
            Select Case Left$(Mmsi9, 3) 'freeform
            Case Is = "970"  'Sart
                MidPos = -1 'no mid
                MmsiDes = "SART"
            Case Is = "972" 'MOB device
                MidPos = -1
                MmsiDes = "MOB Device"
            Case Is = "974" 'EPIRB
                MidPos = -1
                MmsiDes = "EPIRB"
            End Select
        Case Is = "98"       '98MIDxxxx  Child Craft
            MidPos = 3
            MmsiDes = "Child Craft"
        Case Is = "99"       '99MIDxxxx  AtoN
            MidPos = 3
            MmsiDes = "AtoN"
        End Select
    Case Else
    End Select
    
#If False Then
        Select Case Left$(Mmsi9, 5)
        Case Is = "00000"
            MmsiDes = "Reserved for Testing"
        Case Else
            MidPos = 4  'Default
            Select Case Left$(Mmsi9, 3)
            Case Is = "111"
                MmsiDes = "SAR Aircraft"
            Case Is = "970"
                MmsiDes = "Search & Rescue Transmitter (SART)"
                MidPos = 0  '(970 EE XXXX)
            Case Else
                MidPos = 3
                Select Case Left$(Mmsi9, 2)
                Case Is = "00"
                If Left$(Mmsi9, 5) <> "00999" Then
                    MmsiDes = "Shore Station"
                    Else
                        MmsiDes = "Global Land Station"
                    End If
                Case Is = "98"
                    MmsiDes = "Tender"
                Case Is = "99"
                    MmsiDes = "Aid to Navigation"
                Case Else
                    MidPos = 2
                    Select Case Left$(Mmsi9, 1)
                    Case Is = "0"
                        MmsiDes = "Group Call"
                    Case Is = "8"
                        MmsiDes = "Portable VHF/DSC/GNSS"
                        MidPos = 0  '
                    Case Else
                        MidPos = 1
                    End Select
                End Select
            End Select
        End Select
#End If
        If Len(Mmsi9) > 9 Then MmsiDes = "Invalid MMSI (ITU M.585-6)"
        If MidPos = 0 Then MmsiDes = "Invalid MMSI (ITU M.585-6)"
        Select Case Des
        Case Is = "M"
            If MidPos > 0 Then
                MidNo = Mid$(Mmsi9, MidPos, 3)
                If MidNo <= 200 Or MidNo > 775 Then MidNo = 0
                MmsiFmt = MidNo
            Else
                MmsiFmt = 0     'no mid encoded in mmsi
            End If
        Case Is = "D"
            MmsiFmt = MmsiDes
        Case Is = "V"
        End Select
End If

#If False Then
Select Case Des
Case Is = "V"
    If Left$(Mmsi9, 2) = "00" Or Left$(Mmsi9, 2) = "99" Then
        Mmsi9 = Format$(Mmsi9, "00 000 0000")
    Else
        If Left$(Mmsi9, 1) = "0" Then
            Mmsi9 = Format$(Mmsi9, "0 000 00000")
        Else
           Mmsi9 = Format$(Mmsi9, "000 000 000")
        End If
    End If
Case Is = "M"
    If Left$(Mmsi9, 2) = "00" Or Left$(Mmsi9, 2) = "99" Then
        Mmsi9 = Mid$(Mmsi9, 3, 3)   '"00 000 0000"
    Else
        If Left$(Mmsi9, 1) = "0" Then
            Mmsi9 = Mid$(Mmsi9, 2, 3)   '"0 000 00000"
        Else
           Mmsi9 = Mid$(Mmsi9, 1, 3)    '"000 000 000"
        End If
    End If
Case Is = "D"
Select Case Mmsi
    Case 0 To 9999   '00 mid 0000!  mid is 000
        Mmsi9 = "Reserved for Testing"
    Case 10000 To 9999999   '00 mid 0000!
        Mmsi9 = "Shore Station " '& DacName(Mid$(Mmsi, 1, 3))
    Case 10000000 To 99999999   '0 mid 00000
        Mmsi9 = "Group Call " '& DacName(Mid$(Mmsi, 1, 3))
    Case 990000000 To 999999999 '99 mid 0000
        Mmsi9 = "Navigation Aid " '& DacName(Mid$(Mmsi, 3, 3))
    Case 980000000 To 989999999 '99 mid 0000
        Mmsi9 = "Tender " '& DacName(Mid$(Mmsi, 3, 3))
    Case 970000000 To 979999999 '99 mid 0000
        Mmsi9 = "Search & Rescue Transmitter (SART) " '& DacName(Mid$(Mmsi, 3, 3))
    Case 111000000 To 111999999 '111 mid 000
       Mmsi9 = "SAR Aircraft "  '& DacName(Mid$(Mmsi, 4, 3))
    Case Else
        Mmsi9 = "Vessel " '& DacName(Mid$(Mmsi, 1, 3))
        If Left$(Mmsi, 6) = "366999" Then Mmsi9 = Mmsi9 & " {USCG}"
        If Left$(Mmsi, 6) = "369493" Then Mmsi9 = Mmsi9 & " {USCG}"
    End Select
    If Mmsi = "45133333" Then Mmsi9 = "VTS Drechtsteden"
End Select
MmsiFmt = Mmsi9
#End If
End Function

Sub Initialise()
Dim i As Integer

NmeaAIVDMWordName(0) = "AIS Sentence"
NmeaAIVDMWordName(1) = "Fragments in this message"
NmeaAIVDMWordName(2) = "Fragment No"
NmeaAIVDMWordName(3) = "Sequential Message ID"
NmeaAIVDMWordName(4) = "Radio Channel"
NmeaAIVDMWordName(5) = "Payload"
NmeaAIVDMWordName(6) = "Fill bits"

NmeaAisWordName(0) = "AIS Sentence"
NmeaAisWordName(1) = "Fragments in this message"
NmeaAisWordName(2) = "Fragment No"
NmeaAisWordName(3) = "Sequential Message ID"
NmeaAisWordName(4) = "Radio Channel"
NmeaAisWordName(5) = "Payload"
NmeaAisWordName(6) = "Fill bits"

NmeaGPZDAWordName(0) = "GPS Date and Time"
NmeaGPZDAWordName(1) = "UTC Time"
NmeaGPZDAWordName(2) = "UTC Day"
NmeaGPZDAWordName(3) = "UTC Month"
NmeaGPZDAWordName(4) = "UTC Year"
NmeaGPZDAWordName(5) = "Local Zone Hours"
NmeaGPZDAWordName(6) = "Local Zone Minutes"

NmeaGPGGAWordName(0) = "GPS Fix Data"
NmeaGPGGAWordName(1) = "UTC Time"
NmeaGPGGAWordName(2) = "Latitude"
NmeaGPGGAWordName(3) = "N or S"
NmeaGPGGAWordName(4) = "Longitude"
NmeaGPGGAWordName(5) = "E or W"
NmeaGPGGAWordName(6) = "Fix Quality"
NmeaGPGGAWordName(7) = "No of Satellites"
NmeaGPGGAWordName(8) = "Horizontal Dilution Of Precision"
NmeaGPGGAWordName(9) = "Altitude"
NmeaGPGGAWordName(10) = "Units"
NmeaGPGGAWordName(11) = "Height of Geoid above WGS84 elipsoid"
NmeaGPGGAWordName(12) = "Units"
NmeaGPGGAWordName(13) = "Time since last DGPS update"
NmeaGPGGAWordName(14) = "DGPS Reference Station ID"

NmeaGPRMCWordName(0) = "Recommended Minimum Specific GNSS Data"
NmeaGPRMCWordName(1) = "UTC of position fix"
NmeaGPRMCWordName(2) = "Status"
NmeaGPRMCWordName(3) = "Latitude"
NmeaGPRMCWordName(4) = "N or S"
NmeaGPRMCWordName(5) = "Longitude"
NmeaGPRMCWordName(6) = "E or W"
NmeaGPRMCWordName(7) = "Speed Over Ground"
NmeaGPRMCWordName(8) = "Course Over Ground"
NmeaGPRMCWordName(9) = "Date"
NmeaGPRMCWordName(10) = "Magnetic Variation"
NmeaGPRMCWordName(11) = "E or W"
NmeaGPRMCWordName(12) = "Mode indicator"
NmeaGPRMCWordName(13) = "Navigational Status"    'from NMEA 0183 V4.10

NmeaPGHPWordName(0) = "GH Internal Message Type 1"
NmeaPGHPWordName(1) = "Message Type"
NmeaPGHPWordName(2) = "Year"
NmeaPGHPWordName(3) = "Month"
NmeaPGHPWordName(4) = "Day"
NmeaPGHPWordName(5) = "Hour"
NmeaPGHPWordName(6) = "Minute"
NmeaPGHPWordName(7) = "Second"
NmeaPGHPWordName(8) = "MilliSecond"
NmeaPGHPWordName(9) = "Country"
NmeaPGHPWordName(10) = "Region"
NmeaPGHPWordName(11) = "Transponder"
NmeaPGHPWordName(12) = "Buffered"   '0=buffered, 1=On Line
NmeaPGHPWordName(13) = "Nmea Sentence Checksum"

NmeaAITXTWordName(0) = "Text transmission"
NmeaAITXTWordName(1) = "Total messages"
NmeaAITXTWordName(2) = "Message number"
NmeaAITXTWordName(3) = "Identifier"
NmeaAITXTWordName(4) = "Message"

NmeaAIALRWordName(0) = "Set Alarm State"
NmeaAIALRWordName(1) = "Time of change UTC"
NmeaAIALRWordName(2) = "Source ID"
NmeaAIALRWordName(3) = "Condition"
NmeaAIALRWordName(4) = "Acknowledged"
NmeaAIALRWordName(5) = "Description"

NmeaAITAGWordName(0) = "MarineCom"
NmeaAITAGWordName(1) = "Unix Time"
NmeaAITAGWordName(2) = "Source"

NmeaAIABMWordName(0) = "AIS Addressed Binary Message"
NmeaAIABMWordName(1) = "Fragments in this message"
NmeaAIABMWordName(2) = "Fragment No"
NmeaAIABMWordName(3) = "Sequential Message ID"  '(0-3)
NmeaAIABMWordName(4) = "MMSI of Destination unit"
NmeaAIABMWordName(5) = "Radio Channel" '(0-3)
NmeaAIABMWordName(6) = "AIS Message ID" '(6 or 12)
NmeaAIABMWordName(7) = "Payload"
NmeaAIABMWordName(8) = "Fill bits"

NmeaAIBBMWordName(0) = "AIS Broadcast Binary Message"
NmeaAIBBMWordName(1) = "Fragments in this message"
NmeaAIBBMWordName(2) = "Fragment No"
NmeaAIBBMWordName(3) = "Sequential Message ID"  '(0-3)
NmeaAIBBMWordName(4) = "Radio Channel" '(0-3)
NmeaAIBBMWordName(5) = "AIS Message ID" '(6 or 12)
NmeaAIBBMWordName(6) = "Payload"
NmeaAIBBMWordName(7) = "Fill bits"

NmeaAIBRMWordName(0) = "Base Station Options Reply of Received Messages"    'True Heading
NmeaAIBRMWordName(1) = "Word 1"
NmeaAIBRMWordName(2) = "Word 2"
NmeaAIBRMWordName(3) = "Signal Strength of Previous Received Message"
NmeaAIBRMWordName(4) = "Slot Number"

NmeaPTHAJWordName(0) = "True Heading A message J"
NmeaPTHAJWordName(1) = "AIS Channel"
NmeaPTHAJWordName(2) = "Slot Number"
NmeaPTHAJWordName(3) = "Time measurement sign"
NmeaPTHAJWordName(4) = "Start Time offset to TTS"

AisBroadcastChannelName(0) = "No AIS Channel preference"
AisBroadcastChannelName(1) = "AIS Channel A"
AisBroadcastChannelName(2) = "AIS Channel B"
AisBroadcastChannelName(3) = "AIS Channels A & B"

IecEncapsulatedWordName(0) = "Encapsulated Sentence"
IecEncapsulatedWordName(1) = "Fragments in this message"
IecEncapsulatedWordName(2) = "Fragment No"
IecEncapsulatedWordName(3) = "Sequential Message ID"
'there may be other words in here
IecEncapsulatedWordName(4) = "Binary Data"  'not AIS
IecEncapsulatedWordName(5) = "Fill bits"

IecALRIdentifierName(1) = "Tx malfunction"  'IEC 61993-1 (2002)
IecALRIdentifierName(2) = "Antenna VSWR exceeds limit"  'IEC 61993-1 (2002)
IecALRIdentifierName(3) = "Rx channel 1 malfunction"  'IEC 61993-1 (2002)
IecALRIdentifierName(4) = "Rx channel 2 malfunction"  'IEC 61993-1 (2002)
IecALRIdentifierName(5) = "Rx channel 70 malfunction"  'IEC 61993-1 (2002)
IecALRIdentifierName(6) = "general failure"  'IEC 61993-1 (2002)
IecALRIdentifierName(8) = "MKD connection lost"  'IEC 61993-1 (2002)
IecALRIdentifierName(25) = "external EPFS lost"  'IEC 61993-1 (2002)
IecALRIdentifierName(26) = "no sensor position in use"  'IEC 61993-1 (2002)
IecALRIdentifierName(29) = "no valid SOG information"  'IEC 61993-1 (2002)
IecALRIdentifierName(30) = "no valid COG information"  'IEC 61993-1 (2002)
IecALRIdentifierName(32) = "Heading lost/invalid"  'IEC 61993-1 (2002)
IecALRIdentifierName(35) = "no valid ROT information"  'IEC 61993-1 (2002)
IecALRIdentifierName(37) = "Frame synchronisation failure"  'IALA A124 ed1 2002
IecALRIdentifierName(38) = "DGNSS input failed"  'IALA A124 ed1 2002
IecALRIdentifierName(39) = "DSC Tx malfunction"  'IALA A124 ed1 2002
IecALRIdentifierName(40) = "DSC antenna VSWR exceeds limits"  'IALA A124 ed1 2002

IecTXTIdentifierName(7) = "UTC clock lost"  'IEC 61993-1 (2002)
IecTXTIdentifierName(21) = "external DGNSS in use"  'IEC 61993-1 (2002)
IecTXTIdentifierName(22) = "external GNSS in use"  'IEC 61993-1 (2002)
IecTXTIdentifierName(23) = "internal DGNSS in use (beacon)"  'IEC 61993-1 (2002)
IecTXTIdentifierName(24) = "internal DGNSS in use (message 17)"  'IEC 61993-1 (2002)
IecTXTIdentifierName(25) = "internal GNSS in use"  'IEC 61993-1 (2002)
IecTXTIdentifierName(27) = "external SOG/COG in use"  'IEC 61993-1 (2002)
IecTXTIdentifierName(28) = "internal SOG/COG in use"  'IEC 61993-1 (2002)
IecTXTIdentifierName(31) = "heading valid"  'IEC 61993-1 (2002)
IecTXTIdentifierName(33) = "Rate of Turn indicator in use"  'IEC 61993-1 (2002)
IecTXTIdentifierName(34) = "Other ROT source in use"  'IEC 61993-1 (2002)
IecTXTIdentifierName(36) = "Channel management parameters changed"  'IEC 61993-1 (2002)
IecTXTIdentifierName(91) = "Channel Frequencies (SLR200)"

'IecBIITIdentifierName(41) = "surveyed position in use"  'IALA A124 ed1 2002
'IecBIITdentifierName(42) = "UTC clock OK"  'IALA A124 ed1 2002

AisMsgTypeName(0) = "{Invalid message type [0]}"
AisMsgTypeBitsMin(0) = 0
AisMsgTypeBitsMax(0) = 0
AisMsgTypeBitsFill(0) = -1  '-1 = (0-5), Nmea Fill Bits(before crc)
AisMsgTypeName(1) = "Position Report Class A (Scheduled)"
AisMsgTypeBitsMin(1) = 168
AisMsgTypeBitsMax(1) = 168
AisMsgTypeBitsFill(1) = 0
AisMsgTypeName(2) = "Position Report Class A (Assigned Scheduled)"
AisMsgTypeBitsMin(2) = 168
AisMsgTypeBitsMax(2) = 168
AisMsgTypeBitsFill(2) = 0
AisMsgTypeName(3) = "Position Report Class A (Special)"
AisMsgTypeBitsMin(3) = 168
AisMsgTypeBitsMax(3) = 168
AisMsgTypeBitsFill(3) = 0
AisMsgTypeName(4) = "Base Station Report"
AisMsgTypeBitsMin(4) = 168
AisMsgTypeBitsMax(4) = 168
AisMsgTypeBitsFill(4) = 0
AisMsgTypeName(5) = "Ship and Voyage Report"
AisMsgTypeBitsMin(5) = 424
AisMsgTypeBitsMax(5) = 424
AisMsgTypeBitsFill(5) = 2
AisMsgTypeName(6) = "Addressed Binary Message"
AisMsgTypeBitsMin(6) = 88
AisMsgTypeBitsMax(6) = 1008
AisMsgTypeBitsFill(6) = -1
AisMsgTypeName(7) = "Binary Acknowledge"
AisMsgTypeBitsMin(7) = 72
AisMsgTypeBitsMax(7) = 168
AisMsgTypeBitsFill(7) = -1
AisMsgTypeName(8) = "Binary Broadcast Message"
AisMsgTypeName(9) = "Standard SAR Aircraft Position Report"
AisMsgTypeBitsMin(9) = 168
AisMsgTypeBitsMax(9) = 168
AisMsgTypeBitsFill(9) = 0
AisMsgTypeName(10) = "UTC and Date Inquiry"
AisMsgTypeBitsMin(10) = 72
AisMsgTypeBitsMax(10) = 72
AisMsgTypeBitsFill(10) = 0
AisMsgTypeName(11) = "UTC and Date Response"
AisMsgTypeBitsMin(11) = 168
AisMsgTypeBitsMax(11) = 168
AisMsgTypeBitsFill(11) = 0
AisMsgTypeName(12) = "Addressed Safety Related Message"
AisMsgTypeBitsMin(12) = 72
AisMsgTypeBitsMax(12) = 1008
AisMsgTypeBitsFill(12) = -1
AisMsgTypeName(13) = "Safety Related Acknowledge"
AisMsgTypeBitsMin(13) = 72
AisMsgTypeBitsMax(13) = 168
AisMsgTypeBitsFill(13) = -1
AisMsgTypeName(14) = "Safety Related Broadcast Message"
AisMsgTypeBitsMin(14) = 40
AisMsgTypeBitsMax(14) = 1008
AisMsgTypeBitsFill(14) = -1
AisMsgTypeName(15) = "Interrogation"
AisMsgTypeBitsMin(15) = 88
AisMsgTypeBitsMax(15) = 160
AisMsgTypeBitsFill(15) = -1
AisMsgTypeName(16) = "Assigned Mode Command"
AisMsgTypeBitsMin(16) = 96
AisMsgTypeBitsMax(16) = 114
AisMsgTypeBitsFill(16) = 0
AisMsgTypeName(17) = "GNSS Binary Broadcast Message"
AisMsgTypeBitsMin(17) = 80
AisMsgTypeBitsMax(17) = 816
AisMsgTypeBitsFill(17) = 0
AisMsgTypeName(18) = "Standard Class B CS Position Report"
AisMsgTypeBitsMin(18) = 168
AisMsgTypeBitsMax(18) = 168
AisMsgTypeBitsFill(18) = 0
AisMsgTypeName(19) = "Extended Class B Equipment Position Report"
AisMsgTypeBitsMin(19) = 312
AisMsgTypeBitsMax(19) = 312
AisMsgTypeBitsFill(19) = 0
AisMsgTypeName(20) = "Data Link Management"
AisMsgTypeBitsMin(20) = 72
AisMsgTypeBitsMax(20) = 160
AisMsgTypeBitsFill(20) = -1
AisMsgTypeName(21) = "Aid-to-Navigation Report"
AisMsgTypeBitsMin(21) = 272
AisMsgTypeBitsMax(21) = 350
AisMsgTypeBitsFill(21) = -1
AisMsgTypeName(22) = "Channel Management"
AisMsgTypeBitsMin(22) = 168
AisMsgTypeBitsMax(22) = 168
AisMsgTypeBitsFill(22) = 0
AisMsgTypeName(23) = "Group Assignment Command"
AisMsgTypeBitsMin(23) = 160
AisMsgTypeBitsMax(23) = 160
AisMsgTypeBitsFill(23) = 0
AisMsgTypeName(24) = "Class B CS Static Data Report"
AisMsgTypeBitsMin(24) = 160
AisMsgTypeBitsMax(24) = 168
AisMsgTypeBitsFill(24) = 0
AisMsgTypeName(25) = "Binary Message, Single Slot"
AisMsgTypeBitsMin(25) = 40
AisMsgTypeBitsMax(25) = 168
AisMsgTypeBitsFill(25) = -1
AisMsgTypeName(26) = "Binary Message, Multiple Slot"
AisMsgTypeBitsMin(26) = 64
AisMsgTypeBitsMax(26) = 1064
AisMsgTypeBitsFill(26) = -1
AisMsgTypeName(27) = "Long Range AIS Broadcast"
AisMsgTypeBitsMin(27) = 96
AisMsgTypeBitsMax(27) = 96
AisMsgTypeBitsFill(27) = 0
For i = 28 To UBound(AisMsgTypeName)
    AisMsgTypeName(i) = "{Invalid message type [" & i & "]}"
    AisMsgTypeBitsMin(i) = 0
    AisMsgTypeBitsMax(i) = 0
    AisMsgTypeBitsFill(i) = 0
Next i

For i = 0 To UBound(DacName)
    DacName(i) = "not in use"
Next i
DacName(0) = "reserved for testing"
For i = 1 To 9
    DacName(i) = "International"
Next i
DacName(200) = "Inland Waterways"
DacName(201) = "Albania (Republic of)"
DacName(202) = "Andorra (Principality of)"
DacName(203) = "Austria"
DacName(204) = "Azores"
DacName(205) = "Belgium"
DacName(206) = "Belarus (Republic of)"
DacName(207) = "Bulgaria (Republic of)"
DacName(208) = "Vatican City State"
DacName(209) = "Cyprus (Republic of)"
DacName(210) = "Cyprus (Republic of)"
DacName(211) = "Germany (Federal Republic of)"
DacName(212) = "Cyprus (Republic of)"
DacName(213) = "Georgia"
DacName(214) = "Moldova (Republic of)"
DacName(215) = "Malta"
DacName(216) = "Armenia (Republic of)"
DacName(218) = "Germany (Federal Republic of)"
DacName(219) = "Denmark"
DacName(220) = "Denmark"
DacName(224) = "Spain"
DacName(225) = "Spain"
DacName(226) = "France"
DacName(227) = "France"
DacName(228) = "France"
DacName(230) = "Finland"
DacName(231) = "Faroe Islands"
DacName(232) = "United Kingdom of Great Britain and Northern Ireland"
DacName(233) = "United Kingdom of Great Britain and Northern Ireland"
DacName(234) = "United Kingdom of Great Britain and Northern Ireland"
DacName(235) = "United Kingdom of Great Britain and Northern Ireland"
DacName(236) = "Gibraltar"
DacName(237) = "Greece"
DacName(238) = "Croatia (Republic of)"
DacName(239) = "Greece"
DacName(240) = "Greece"
DacName(241) = "Greece"
DacName(242) = "Morocco (Kingdom of)"
DacName(243) = "Hungary (Republic of)"
DacName(244) = "Netherlands (Kingdom of the)"
DacName(245) = "Netherlands (Kingdom of the)"
DacName(246) = "Netherlands (Kingdom of the)"
DacName(247) = "Italy"
DacName(248) = "Malta"
DacName(249) = "Malta"
DacName(250) = "Ireland"
DacName(251) = "Iceland"
DacName(252) = "Liechtenstein (Principality of)"
DacName(253) = "Luxembourg"
DacName(254) = "Monaco (Principality of)"
DacName(255) = "Madeira"
DacName(256) = "Malta"
DacName(257) = "Norway"
DacName(258) = "Norway"
DacName(259) = "Norway"
DacName(261) = "Poland (Republic of)"
DacName(262) = "Montenegro (Republic of)"
DacName(263) = "Portugal"
DacName(264) = "Romania"
DacName(265) = "Sweden"
DacName(266) = "Sweden"
DacName(267) = "Slovak Republic"
DacName(268) = "San Marino (Republic of)"
DacName(269) = "Switzerland (Confederation of)"
DacName(270) = "Czech Republic"
DacName(271) = "Turkey"
DacName(272) = "Ukraine"
DacName(273) = "Russian Federation"
DacName(274) = "The Former Yugoslav Republic of Macedonia"
DacName(275) = "Latvia (Republic of)"
DacName(276) = "Estonia (Republic of)"
DacName(277) = "Lithuania (Republic of)"
DacName(278) = "Slovenia (Republic of)"
DacName(279) = "Serbia (Republic of)"
DacName(301) = "Anguilla"
DacName(303) = "Alaska (State of)"
DacName(304) = "Antigua and Barbuda"
DacName(305) = "Antigua and Barbuda"
DacName(306) = "Netherlands Antilles"
DacName(307) = "Aruba"
DacName(308) = "Bahamas (Commonwealth of the)"
DacName(309) = "Bahamas (Commonwealth of the)"
DacName(310) = "Bermuda"
DacName(311) = "Bahamas (Commonwealth of the)"
DacName(312) = "Belize"
DacName(314) = "Barbados"
DacName(316) = "Canada"
DacName(319) = "Cayman Islands"
DacName(321) = "Costa Rica"
DacName(323) = "Cuba"
DacName(325) = "Dominica (Commonwealth of)"
DacName(327) = "Dominican Republic"
DacName(329) = "Guadeloupe (French Department of)"
DacName(330) = "Grenada"
DacName(331) = "Greenland"
DacName(332) = "Guatemala (Republic of)"
DacName(334) = "Honduras (Republic of)"
DacName(336) = "Haiti (Republic of)"
DacName(338) = "United States of America"
DacName(339) = "Jamaica"
DacName(341) = "Saint Kitts and Nevis (Federation of)"
DacName(343) = "Saint Lucia"
DacName(345) = "Mexico"
DacName(347) = "Martinique (French Department of)"
DacName(348) = "Montserrat"
DacName(350) = "Nicaragua"
DacName(351) = "Panama (Republic of)"
DacName(352) = "Panama (Republic of)"
DacName(353) = "Panama (Republic of)"
DacName(354) = "Panama (Republic of)"
DacName(355) = "Panama (Republic of)"
DacName(356) = "Panama (Republic of)"
DacName(357) = "Panama (Republic of)"
DacName(358) = "Puerto Rico"
DacName(359) = "El Salvador (Republic of)"
DacName(361) = "Saint Pierre and Miquelon (Territorial Collectivity of)"
DacName(362) = "Trinidad and Tobago"
DacName(364) = "Turks and Caicos Islands"
DacName(366) = "United States of America"
DacName(367) = "United States of America"
DacName(368) = "United States of America"
DacName(369) = "United States of America"
DacName(370) = "Panama (Republic of)"
DacName(371) = "Panama (Republic of)"
DacName(372) = "Panama (Republic of)"
DacName(373) = "Panama (Republic of)"
DacName(375) = "Saint Vincent and the Grenadines"
DacName(376) = "Saint Vincent and the Grenadines"
DacName(377) = "Saint Vincent and the Grenadines"
DacName(378) = "British Virgin Islands"
DacName(379) = "United States Virgin Islands"
DacName(401) = "Afghanistan"
DacName(403) = "Saudi Arabia (Kingdom of)"
DacName(405) = "Bangladesh (People's Republic of)"
DacName(408) = "Bahrain (Kingdom of)"
DacName(410) = "Bhutan (Kingdom of)"
DacName(412) = "China (People's Republic of)"
DacName(413) = "China (People's Republic of)"
DacName(414) = "China (People's Republic of)"
DacName(416) = "Taiwan (Province of China)"
DacName(417) = "Sri Lanka (Democratic Socialist Republic of)"
DacName(419) = "India (Republic of)"
DacName(422) = "Iran (Islamic Republic of)"
DacName(423) = "Azerbaijani Republic"
DacName(425) = "Iraq (Republic of)"
DacName(428) = "Israel (State of)"
DacName(431) = "Japan"
DacName(432) = "Japan"
DacName(434) = "Turkmenistan"
DacName(436) = "Kazakhstan (Republic of)"
DacName(437) = "Uzbekistan (Republic of)"
DacName(438) = "Jordan (Hashemite Kingdom of)"
DacName(440) = "Korea (Republic of)"
DacName(441) = "Korea (Republic of)"
DacName(443) = "Palestine (In accordance with Resolution 99 Rev. Antalya, 2006)"
DacName(445) = "Democratic People's Republic of Korea"
DacName(447) = "Kuwait (State of)"
DacName(450) = "Lebanon"
DacName(451) = "Kyrgyz Republic"
DacName(453) = "Macao (Special Administrative Region of China)"
DacName(455) = "Maldives (Republic of)"
DacName(457) = "Mongolia"
DacName(459) = "Nepal (Republic of)"
DacName(461) = "Oman (Sultanate of)"
DacName(463) = "Pakistan (Islamic Republic of)"
DacName(466) = "Qatar (State of)"
DacName(468) = "Syrian Arab Republic"
DacName(470) = "United Arab Emirates"
DacName(473) = "Yemen (Republic of)"
DacName(475) = "Yemen (Republic of)"
DacName(477) = "Hong Kong (Special Administrative Region of China)"
DacName(478) = "Bosnia and Herzegovina"
DacName(501) = "Adelie Land"
DacName(503) = "Australia"
DacName(506) = "Myanmar (Union of)"
DacName(508) = "Brunei Darussalam"
DacName(510) = "Micronesia (Federated States of)"
DacName(511) = "Palau (Republic of)"
DacName(512) = "New Zealand"
DacName(514) = "Cambodia (Kingdom of)"
DacName(515) = "Cambodia (Kingdom of)"
DacName(516) = "Christmas Island (Indian Ocean)"
DacName(518) = "Cook Islands"
DacName(520) = "Fiji (Republic of)"
DacName(523) = "Cocos (Keeling) Islands"
DacName(525) = "Indonesia (Republic of)"
DacName(529) = "Kiribati (Republic of)"
DacName(531) = "Lao People's Democratic Republic"
DacName(533) = "Malaysia"
DacName(536) = "Northern Mariana Islands (Commonwealth of the)"
DacName(538) = "Marshall Islands (Republic of the)"
DacName(540) = "New Caledonia"
DacName(542) = "Niue"
DacName(544) = "Nauru (Republic of)"
DacName(546) = "French Polynesia"
DacName(548) = "Philippines (Republic of the)"
DacName(553) = "Papua New Guinea"
DacName(555) = "Pitcairn Island"
DacName(557) = "Solomon Islands"
DacName(559) = "American Samoa"
DacName(561) = "Samoa (Independent State of)"
DacName(563) = "Singapore (Republic of)"
DacName(564) = "Singapore (Republic of)"
DacName(565) = "Singapore (Republic of)"
DacName(566) = "Singapore (Republic of)"
DacName(567) = "Thailand"
DacName(570) = "Tonga (Kingdom of)"
DacName(572) = "Tuvalu"
DacName(574) = "Viet Nam (Socialist Republic of)"
DacName(576) = "Vanuatu (Republic of)"
DacName(577) = "Vanuatu (Republic of)"
DacName(578) = "Wallis and Futuna Islands"
DacName(601) = "South Africa (Republic of)"
DacName(603) = "Angola (Republic of)"
DacName(605) = "Algeria (People's Democratic Republic of)"
DacName(607) = "Saint Paul and Amsterdam Islands"
DacName(608) = "Ascension Island"
DacName(609) = "Burundi (Republic of)"
DacName(610) = "Benin (Republic of)"
DacName(611) = "Botswana (Republic of)"
DacName(612) = "Central African Republic"
DacName(613) = "Cameroon (Republic of)"
DacName(615) = "Congo (Republic of the)"
DacName(616) = "Comoros (Union of the)"
DacName(617) = "Cape Verde (Republic of)"
DacName(618) = "Crozet Archipelago"
DacName(619) = "Côte d'Ivoire (Republic of)"
DacName(620) = "Comoros (Union of the)"
DacName(621) = "Djibouti (Republic of)"
DacName(622) = "Egypt (Arab Republic of)"
DacName(624) = "Ethiopia (Federal Democratic Republic of)"
DacName(625) = "Eritrea"
DacName(626) = "Gabonese Republic"
DacName(627) = "Ghana"
DacName(629) = "Gambia (Republic of the)"
DacName(630) = "Guinea-Bissau (Republic of)"
DacName(631) = "Equatorial Guinea (Republic of)"
DacName(632) = "Guinea (Republic of)"
DacName(633) = "Burkina Faso"
DacName(634) = "Kenya (Republic of)"
DacName(635) = "Kerguelen Islands"
DacName(636) = "Liberia (Republic of)"
DacName(637) = "Liberia (Republic of)"
DacName(638) = "South Sudan (Republic of)"
DacName(642) = "Socialist People's Libyan Arab Jamahiriya"
DacName(644) = "Lesotho (Kingdom of)"
DacName(645) = "Mauritius (Republic of)"
DacName(647) = "Madagascar (Republic of)"
DacName(649) = "Mali (Republic of)"
DacName(650) = "Mozambique (Republic of)"
DacName(654) = "Mauritania (Islamic Republic of)"
DacName(655) = "Malawi"
DacName(656) = "Niger (Republic of the)"
DacName(657) = "Nigeria (Federal Republic of)"
DacName(659) = "Namibia (Republic of)"
DacName(660) = "Reunion (French Department of)"
DacName(661) = "Rwanda (Republic of)"
DacName(662) = "Sudan (Republic of the)"
DacName(663) = "Senegal (Republic of)"
DacName(664) = "Seychelles (Republic of)"
DacName(665) = "Saint Helena"
DacName(666) = "Somali Democratic Republic"
DacName(667) = "Sierra Leone"
DacName(668) = "Sao Tome and Principe (Democratic Republic of)"
DacName(669) = "Swaziland (Kingdom of)"
DacName(670) = "Chad (Republic of)"
DacName(671) = "Togolese Republic"
DacName(672) = "Tunisia"
DacName(674) = "Tanzania (United Republic of)"
DacName(675) = "Uganda (Republic of)"
DacName(676) = "Democratic Republic of the Congo"
DacName(677) = "Tanzania (United Republic of)"
DacName(678) = "Zambia (Republic of)"
DacName(679) = "Zimbabwe (Republic of)"
DacName(701) = "Argentine Republic"
DacName(710) = "Brazil (Federative Republic of)"
DacName(720) = "Bolivia (Republic of)"
DacName(725) = "Chile"
DacName(730) = "Colombia (Republic of)"
DacName(735) = "Ecuador"
DacName(740) = "Falkland Islands (Malvinas)"
DacName(745) = "Guiana (French Department of)"
DacName(750) = "Guyana"
DacName(755) = "Paraguay (Republic of)"
DacName(760) = "Peru"
DacName(765) = "Suriname (Republic of)"
DacName(770) = "Uruguay (Eastern Republic of)"
DacName(775) = "Venezuela (Bolivarian Republic of)"

For i = 1 To UBound(AreaType1Name)
    AreaType1Name(i) = "not in use"
Next i
AreaType1Name(0) = "Caution Area: Marine mammals habitat"
AreaType1Name(1) = "Caution Area: Marine mammals in area - Reduce Speed"
AreaType1Name(2) = "Caution Area: Marine mammals in area - Stay Clear"
AreaType1Name(3) = "Caution Area: Marine mammals in area - Report Sightings"
AreaType1Name(4) = "Caution Area: Protected Habitat - Reduce Speed"
AreaType1Name(5) = "Caution Area: Protected Habitat - Stay Clear"
AreaType1Name(6) = "Caution Area: Protected Habitat - No fishing or anchoring"
AreaType1Name(7) = "Caution Area: Derelicts (drifting objects)"
AreaType1Name(8) = "Caution Area: Traffic congestion"
AreaType1Name(9) = "Caution Area: Marine event"
AreaType1Name(10) = "Caution Area: Divers down"
AreaType1Name(11) = "Caution Area: Swim area"
AreaType1Name(12) = "Caution Area: Dredge operations"
AreaType1Name(13) = "Caution Area: Survey operations"
AreaType1Name(14) = "Caution Area: Underwater operation"
AreaType1Name(15) = "Caution Area: Seaplane operations"
AreaType1Name(16) = "Caution Area: Fishery - nets in water"
AreaType1Name(17) = "Caution Area: Cluster of fishing vessels"
AreaType1Name(18) = "Caution Area: Fairway closed"
AreaType1Name(19) = "Caution Area: Harbor closed"
AreaType1Name(20) = "Caution Area: Risk (defined in free text field)"
AreaType1Name(21) = "Caution Area: Underwater vehicle operation"
AreaType1Name(22) = "reserved for future use"
AreaType1Name(23) = "Storm front (line squall)"
AreaType1Name(24) = "Env. Caution Area: Hazardous sea ice"
AreaType1Name(25) = "Env. Caution Area: Storm warning (storm cell or line of storms)"
AreaType1Name(26) = "Env. Caution Area: High wind"
AreaType1Name(27) = "Env. Caution Area: High waves"
AreaType1Name(28) = "Env. Caution Area: Restricted visibility (fog, rain, etc)"
AreaType1Name(29) = "Env. Caution Area: Strong currents"
AreaType1Name(30) = "Env. Caution Area: Heavy icing"
AreaType1Name(31) = "reserved for future use"
AreaType1Name(32) = "Restricted Area: Fishing prohibited"
AreaType1Name(33) = "Restricted Area: No anchoring."
AreaType1Name(34) = "Restricted Area: Entry approval required prior to transit"
AreaType1Name(35) = "Restricted Area: Entry prohibited"
AreaType1Name(36) = "Restricted Area: Active military OPAREA"
AreaType1Name(37) = "Restricted Area: Firing - danger area."
AreaType1Name(38) = "Restricted Area: Drifting Mines"
AreaType1Name(39) = "reserved for future use"
AreaType1Name(40) = "Anchorage Area: Anchorage open"
AreaType1Name(41) = "Anchorage Area: Anchorage closed"
AreaType1Name(42) = "Anchorage Area: Anchoring prohibited"
AreaType1Name(43) = "Anchorage Area: Deep draught anchorage"
AreaType1Name(44) = "Anchorage Area: Shallow draught anchorage"
AreaType1Name(45) = "Anchorage Area: Vessel transfer operations"
AreaType1Name(46) = "reserved for future use"
AreaType1Name(47) = "reserved for future use"
AreaType1Name(48) = "reserved for future use"
AreaType1Name(49) = "reserved for future use"
AreaType1Name(50) = "reserved for future use"
AreaType1Name(51) = "reserved for future use"
AreaType1Name(52) = "reserved for future use"
AreaType1Name(53) = "reserved for future use"
AreaType1Name(54) = "reserved for future use"
AreaType1Name(55) = "reserved for future use"
AreaType1Name(56) = "Security Alert - Level 1"
AreaType1Name(57) = "Security Alert - Level 2"
AreaType1Name(58) = "Security Alert - Level 3"
AreaType1Name(59) = "reserved for future use"
AreaType1Name(60) = "reserved for future use"
AreaType1Name(61) = "reserved for future use"
AreaType1Name(62) = "reserved for future use"
AreaType1Name(63) = "reserved for future use"
AreaType1Name(64) = "Distress Area: Vessel disabled and adrift"
AreaType1Name(65) = "Distress Area: Vessel sinking"
AreaType1Name(66) = "Distress Area: Vessel abandoning ship"
AreaType1Name(67) = "Distress Area: Vessel requests medical assistance"
AreaType1Name(68) = "Distress Area: Vessel flooding"
AreaType1Name(69) = "Distress Area: Vessel fire/explosion"
AreaType1Name(70) = "Distress Area: Vessel grounding"
AreaType1Name(71) = "Distress Area: Vessel collision"
AreaType1Name(72) = "Distress Area: Vessel listing/capsizing"
AreaType1Name(73) = "Distress Area: Vessel under assault"
AreaType1Name(74) = "Distress Area: Person overboard"
AreaType1Name(75) = "Distress Area: SAR area"
AreaType1Name(76) = "Distress Area: Pollution response area"
AreaType1Name(77) = "reserved for future use"
AreaType1Name(78) = "reserved for future use"
AreaType1Name(79) = "reserved for future use"
AreaType1Name(80) = "Instruction: Contact VTS at this point/juncture"
AreaType1Name(81) = "Instruction: Contact Port Administration at this point/juncture"
AreaType1Name(82) = "Instruction: Do not proceed beyond this point/juncture"
AreaType1Name(83) = "Instruction: Await instructions prior to proceeding beyond this point/juncture"
AreaType1Name(84) = "Proceed to this location - await instructions"
AreaType1Name(85) = "Clearance granted - proceed to berth"
AreaType1Name(86) = "reserved for future use"
AreaType1Name(87) = "reserved for future use"
AreaType1Name(88) = "Information: Pilot boarding position"
AreaType1Name(89) = "Information: Icebreaker waiting area"
AreaType1Name(90) = "Information: Places of refuge"
AreaType1Name(91) = "Information: Position of icebreakers"
AreaType1Name(92) = "Information: Location of response units"
AreaType1Name(93) = "VTS active target"
AreaType1Name(94) = "Rouge or suspicious vessel"
AreaType1Name(95) = "Vessel requesting non-distess assistance"
AreaType1Name(96) = "Chart Feature: Sunken vessel"
AreaType1Name(97) = "Chart Feature: Submerged object"
AreaType1Name(98) = "Chart Feature: Semi-submerged object"
AreaType1Name(99) = "Chart Feature: Shoal area"
AreaType1Name(100) = "Chart Feature: Shoal area due North"
AreaType1Name(101) = "Chart Feature: Shoal area due East"
AreaType1Name(102) = "Chart Feature: Shoal area due South"
AreaType1Name(103) = "Chart Feature: Shoal area due West"
AreaType1Name(104) = "Chart Feature: Channel obstruction"
AreaType1Name(105) = "Chart Feature: Reduced vertical clearance"
AreaType1Name(106) = "Chart Feature: Bridge closed"
AreaType1Name(107) = "Chart Feature: Bridge partially open"
AreaType1Name(108) = "Chart Feature: Bridge fully open"
AreaType1Name(109) = "reserved for future use"
AreaType1Name(110) = "reserved for future use"
AreaType1Name(111) = "reserved for future use"
AreaType1Name(112) = "Report from ship: Icing info"
AreaType1Name(113) = "reserved for future use"
AreaType1Name(114) = "Report from ship: Miscellaneous information   define in free text field"
AreaType1Name(115) = "reserved for future use"
AreaType1Name(116) = "reserved for future use"
AreaType1Name(117) = "reserved for future use"
AreaType1Name(118) = "reserved for future use"
AreaType1Name(119) = "reserved for future use"
AreaType1Name(120) = "Route: Recommended route"
AreaType1Name(121) = "Route: Altenative route"
AreaType1Name(122) = "Route: Recommended route through ice"
AreaType1Name(123) = "reserved for future use"
AreaType1Name(124) = "reserved for future use"
AreaType1Name(125) = "Other Define in free text field"
AreaType1Name(126) = "Cancellation - cancel area as identified by Message Linkage ID"
AreaType1Name(127) = "Undefined (default)"

For i = 0 To 9  'set up defaults
    IfiName(i) = "Reserved for future system applications"
Next i
For i = i To UBound(IfiName) 'note I to I
    IfiName(i) = "Reserved for International Operational Applications"
Next i
IfiName(0) = "Text Telegram"
IfiName(1) = "Application Acknowledgement"
IfiName(2) = "Interrogation for a Specific IFM"
IfiName(3) = "Capability Interrogation"
IfiName(4) = "Capability Interrogation Reply"
IfiName(5) = "Application Acknowlegement to an Addressed Binary Message"
IfiName(11) = "Meteorological and Hydrographic"
IfiName(12) = "Dangerous Cargo Indication (legacy)"
IfiName(13) = "Fairway Closed (legacy)"
IfiName(14) = "Tidal Window (legacy)"
IfiName(15) = "Extended Ship Static and Voyage Related Data (legacy)"
IfiName(16) = "Persons, VTS Targets (legacy)"
IfiName(17) = "VTS-generated/synthetic Targets"
IfiName(18) = "Clearance time, Advice of VTS Waypoints (legacy)"
IfiName(19) = "Marine Traffic Signal, Extended Ship And Voyage related Data (legacy)"
IfiName(20) = "Berthing Data"
IfiName(21) = "Weather Observation Report"
IfiName(22) = "Area Notice - broadcast"
IfiName(23) = "Area Notive - addressed"
IfiName(24) = "Extended ship static and voyage related data"
IfiName(25) = "Dangerous Cargo Indication"
IfiName(26) = "Environmental"
IfiName(27) = "Route information - broadcast"
IfiName(28) = "Route information - addressed"
IfiName(29) = "Text description - broadcast"
IfiName(30) = "Text description - addressed"
IfiName(31) = "Met/Hydrographic"
IfiName(32) = "Tidal Window"
IfiName(40) = "Number of Persons on Board (legacy)"

RepeatName(0) = "Repeatable"
RepeatName(1) = "Repeated once"
RepeatName(2) = "Repeated twice"
RepeatName(3) = "Do not repeat"

StatusName(0) = "Under way using engine (Rule 23(a) or Rule 25(e))"
StatusName(1) = "At anchor (Rule 30(a-c))"
StatusName(2) = "Not under command (Rule 27(a))"
StatusName(3) = "Restricted Manoeuverability (Rule 27(b-h))"
StatusName(4) = "Constrained by her draught (Rule 28)"
StatusName(5) = "Moored"
StatusName(6) = "Aground (Rule 30(d))"
StatusName(7) = "Engaged in Fishing (Rule 26)"
StatusName(8) = "Under way sailing (Rule 25)"
StatusName(9) = "Reserved for future amendment"
StatusName(10) = "Reserved for future amendment"
StatusName(11) = "Reserved for future use"
StatusName(12) = "Reserved for futute use"
StatusName(13) = "Reserved for future use"
StatusName(14) = "AIS-SART (Rule 36)"
StatusName(15) = "Not defined (default) (also SART under test)"

AccuracyName(0) = "low (>10m)(default)"
AccuracyName(1) = "high (<=10m)"

RaimName(0) = "RAIM not in use"
RaimName(1) = "RAIM in use"

EpfdName(0) = "Undefined (default)"
EpfdName(1) = "GPS"
EpfdName(2) = "GNSS (GLONASS)"
EpfdName(3) = "Combined GPS/GLONASS"
EpfdName(4) = "Loran-C"
EpfdName(5) = "Chayka"
EpfdName(6) = "Integrated Navigation System"
EpfdName(7) = "Surveyed"
EpfdName(8) = "Galileo"
For i = 9 To 14
    EpfdName(i) = "{Invalid [" & i & "]}"
Next i
EpfdName(15) = "internal GNSS"

TxControlName(0) = "Stop transmit in base station coverage area (default)"
TxControlName(1) = "Transmit in base station coverage area (default)"

Ais_VersionName(0) = "Compliant with ITU-R M.1371-1"
Ais_VersionName(1) = "Compliant with ITU-R M.1371-3"
Ais_VersionName(2) = "future editions"
Ais_VersionName(3) = "future editions"

Ship_TypeNxName(1) = "Reserved for future use"
Ship_TypeNxName(2) = "WIG"
Ship_TypeNxName(3) = "Refer to type3xName"
Ship_TypeNxName(4) = "HSC"
Ship_TypeNxName(5) = "Refer to type5xName"
Ship_TypeNxName(6) = "Passenger"
Ship_TypeNxName(7) = "Cargo"
Ship_TypeNxName(8) = "Tanker"
Ship_TypeNxName(9) = "Other type of ship"

Ship_TypexNName(0) = "-all ships of this type"
Ship_TypexNName(1) = "-carrying DG,HS,MP,IMO haz or pollutant X"
Ship_TypexNName(2) = "-carrying DG,HS,MP,IMO haz or pollutant Y"
Ship_TypexNName(3) = "-carrying DG,HS,MP,IMO haz or pollutant Z"
Ship_TypexNName(4) = "-carrying DG,HS,MP,IMO haz or pollutant OS"
Ship_TypexNName(5) = "-reserved for future use"
Ship_TypexNName(6) = "-reserved for future use"
Ship_TypexNName(7) = "-reserved for future use"
Ship_TypexNName(8) = "-reserved for future use"
Ship_TypexNName(9) = "-no additional information"

Ship_Type3NName(0) = "=Fishing"
Ship_Type3NName(1) = "-Towing"
Ship_Type3NName(2) = "-Towing, Tow Length > 200m or breadth > 25m"
Ship_Type3NName(3) = "-Engaged in dredging or underwater operations"
Ship_Type3NName(4) = "-Engaged in diving operations"
Ship_Type3NName(5) = "-Engaged in military operations"
Ship_Type3NName(6) = "-Sailing"
Ship_Type3NName(7) = "-Pleasure craft"
Ship_Type3NName(8) = "-Reserved for future use"
Ship_Type3NName(9) = "-Reserved for future use"

Ship_Type5NName(0) = "Pilot vessel"
Ship_Type5NName(1) = "Search and rescue vessel"
Ship_Type5NName(2) = "Tug"
Ship_Type5NName(3) = "Port tender"
Ship_Type5NName(4) = "Anti-pollution vessel"
Ship_Type5NName(5) = "Law enforcement vessel"
Ship_Type5NName(6) = "Spare - for local use"
Ship_Type5NName(7) = "Spare - for local use"
Ship_Type5NName(8) = "Medical transport"
Ship_Type5NName(9) = "Sip/aircraft not party to an armed conflict"

AtoNName(0) = "Type of Aid to Navigation not specified (default)"
AtoNName(1) = "Reference point"
AtoNName(2) = "RACON"
AtoNName(3) = "Fixed structure off shore"
AtoNName(4) = "Spare, reserved for future use"
AtoNName(5) = "Light, without sectors"
AtoNName(6) = "Light, with sectors"
AtoNName(7) = "Leading Light Front"
AtoNName(8) = "Leading Light Rear"
AtoNName(9) = "Beacon, Cardinal N"
AtoNName(10) = "Beacon, Cardinal E"
AtoNName(11) = "Beacon, Cardinal S"
AtoNName(12) = "Beacon, Cardinal W"
AtoNName(13) = "Beacon, Starboard hand"
AtoNName(14) = "Beacon, Port hand"
AtoNName(15) = "Beacon, Prferred Channel port hand"
AtoNName(16) = "Beacon, Preferred Cahnnel starboard hand"
AtoNName(17) = "Beacon, Isolated danger"
AtoNName(18) = "Beacon, Safe water"
AtoNName(19) = "Beacon, Special mark"
AtoNName(20) = "Cardinal Mark N"
AtoNName(21) = "Cardinal Mark E"
AtoNName(22) = "Cardinal Mark S"
AtoNName(23) = "Cardinal Mark W"
AtoNName(24) = "Port hand Mark"
AtoNName(25) = "Starboard hand Mark"
AtoNName(26) = "Preferred Channel Port Hand"
AtoNName(27) = "Preferred Channel Starboard Hand"
AtoNName(28) = "Isolated danger"
AtoNName(29) = "Safe Water"
AtoNName(30) = "Special Mark"
AtoNName(31) = "Light Vessel/LANBY/Rigs"

StationTypeName(0) = "All types of mobiles (default)"
StationTypeName(1) = "Reserved for future use"
StationTypeName(2) = "All types of Class B mobile stations"
StationTypeName(3) = "SAR mobile airborne mobile station"
StationTypeName(4) = "Aid to Navigation station"
StationTypeName(5) = "Class B (CS) shipbourne mobile only"
StationTypeName(6) = "Inland waterways"
StationTypeName(7) = "Regional use"
StationTypeName(8) = "Regional use"
StationTypeName(9) = "Regional use"
StationTypeName(10) = "Base Station Coverage Area"  'Added M.1371-5
StationTypeName(11) = "Reserved for future use"
StationTypeName(12) = "Reserved for future use"
StationTypeName(13) = "Reserved for future use"
StationTypeName(14) = "Reserved for future use"
StationTypeName(15) = "Reserved for future use"

TxrxName(0) = "TxA/TxB, RxA/RxB (default)"
TxrxName(1) = "TxA, RxA/RxB"
TxrxName(2) = "TxB, RxA/RxB"
TxrxName(3) = "Reserved for future use"
For i = 4 To UBound(TxrxName)
    TxrxName(i) = "{Invalid [" & i & "]}"
Next i

IntervalName(0) = "As given by the autonomous mode"
IntervalName(1) = "10 Minutes"
IntervalName(2) = "6 Minutes"
IntervalName(3) = "3 Minutes"
IntervalName(4) = "1 Minutes"
IntervalName(5) = "30 Seconds"
IntervalName(6) = "15 Seconds"
IntervalName(7) = "10 Seconds"
IntervalName(8) = "5 Seconds"
IntervalName(9) = "2 Seconds (not applicable to Class B ""CS"")"
IntervalName(10) = "Next Shorter Reporting Interval"
IntervalName(11) = "Next Longer Reporting Interval"
IntervalName(12) = "Reserved for future use"
IntervalName(13) = "Reserved for future use"
IntervalName(14) = "Reserved for future use"
IntervalName(15) = "Reserved for future use"
           
DteName(0) = "Data Terminal Ready"
DteName(1) = "Not ready (default)"

CsName(0) = "Class B SOTDMA unit"
CsName(1) = "Class B CS (Carrier Sense) unit"

DisplayName(0) = "No visual display (cannot display msg 12 & 14)"
DisplayName(1) = "Has visual display (can display msg 12 & 14)"

DscName(0) = "Not equipped with DSC function"
DscName(1) = "Equipped with DSC function"

BandName(0) = "Can operate over top 525kHz of marine band"
BandName(1) = "Can operate over the whole marine band"

Msg22Name(0) = "No Frequency management via msg22 (AIS1 & AIS2 only)"
Msg22Name(1) = "Unit can accept channel assignment via Message Type 22"
 
AssignedName(0) = "Station operating in autonomous mode (default)"
AssignedName(1) = "Station operating in assigned mode"

Off_PositionName(0) = "on position"
Off_PositionName(1) = "off position"
 
Virtual_AidName(0) = "real aid"
Virtual_AidName(1) = "simulated aid"

GnssPositionName(0) = "Position is current GNSS"
GnssPositionName(1) = "Position is not current GNSS (default)"

AddressedName(0) = "Broadcast"
AddressedName(1) = "Addressed"
                
Band_WidthName(0) = "default (as specified by channel number)"
Band_WidthName(1) = "Spare (formerly 12.5 kHz in M.1371-1)"

StructuredName(0) = "Application Identifier not used"
StructuredName(1) = "Application Identifier used"
        
PartnoName(0) = "Type 24 Part A"
PartnoName(1) = "Type 24 Part B"

AltitudeName(0) = "GNSS"
AltitudeName(1) = "Barometric"

RadioModeName(0) = "SOTDMA"
RadioModeName(1) = "ITDMA"

SyncStateName(0) = "UTC Direct"
SyncStateName(1) = "UTC Indirect"
SyncStateName(2) = "Synchronised to a Base Station"
SyncStateName(3) = "Synchronised to another Station"

SlotAllocName(0) = "1 slot"
SlotAllocName(1) = "2 slots"
SlotAllocName(2) = "3 slots"
SlotAllocName(3) = "4 slots"
SlotAllocName(4) = "5 slots"
SlotAllocName(5) = "1 slot; offset=Slot incr + 8192"
SlotAllocName(6) = "2 slots; offset=Slot incr + 8192"
SlotAllocName(7) = "3 slots; offset=Slot incr + 8192"

PowerName(0) = "High (default)"
PowerName(1) = "Low"

TendencyName(0) = "Steady"
TendencyName(1) = "Decreasing"
TendencyName(2) = "Increasing"
TendencyName(3) = "Not available"

BeaufortName(0) = "Calm"
BeaufortName(1) = "Light Air"
BeaufortName(2) = "Light Breeze"
BeaufortName(3) = "Gentle Breeze"
BeaufortName(4) = "Moderate Breeze"
BeaufortName(5) = "Fresh Breeze"
BeaufortName(6) = "Strong Breeze"
BeaufortName(7) = "Near Gale"
BeaufortName(8) = "Gale"
BeaufortName(9) = "Severe Gale"
BeaufortName(10) = "Storm"
BeaufortName(11) = "Violent Storm"
BeaufortName(12) = "Hurricane"
BeaufortName(13) = "Not Available (default)"    'ASM 1-31 otherwise invalid
BeaufortName(14) = "Reserved for futute use"
BeaufortName(15) = "Reserved for futute use"

BeaufortEnvironmentalName(0) = "Calm"
BeaufortEnvironmentalName(1) = "Light Air"
BeaufortEnvironmentalName(2) = "Light Breeze"
BeaufortEnvironmentalName(3) = "Gentle Breeze"
BeaufortEnvironmentalName(4) = "Moderate Breeze"
BeaufortEnvironmentalName(5) = "Fresh Breeze"
BeaufortEnvironmentalName(6) = "Strong Breeze"
BeaufortEnvironmentalName(7) = "Near Gale"
BeaufortEnvironmentalName(8) = "Gale"
BeaufortEnvironmentalName(9) = "Severe Gale"
BeaufortEnvironmentalName(10) = "Storm"
BeaufortEnvironmentalName(11) = "Violent Storm"
BeaufortEnvironmentalName(12) = "Hurricane"
BeaufortEnvironmentalName(13) = "Invalid"    'ASM 1-31 otherwise invalid
BeaufortEnvironmentalName(14) = "Invalid"
BeaufortEnvironmentalName(15) = "Not available"

YesNoName(0) = "No"
YesNoName(1) = "Yes"
YesNoName(2) = "Invalid"
YesNoName(3) = "Not available (default)"

YesNo2Name(0) = "Not available (default)"
YesNo2Name(1) = "No"
YesNo2Name(2) = "Yes"
YesNo2Name(3) = "Not in use"

'WMO 306 Code table 4.201
PrecipitationName(0) = "Reserved"
PrecipitationName(1) = "Rain"
PrecipitationName(2) = "Thunderstorm"
PrecipitationName(3) = "Freezing Rain"
PrecipitationName(4) = "Mixed Ice"
PrecipitationName(5) = "Snow"
PrecipitationName(6) = "Reserved"
PrecipitationName(7) = "Not available (default)"

TargTypeName(0) = "MMSI Number"
TargTypeName(1) = "IMO Number"
TargTypeName(2) = "Call Sign"
TargTypeName(3) = "Other (default)"

ConeName(0) = "0 Cones"
ConeName(1) = "1 Cone"
ConeName(2) = "2 Cones"
ConeName(3) = "3 Cones"
ConeName(4) = "B-Flag"
ConeName(5) = "Unknown (default)"
ConeName(6) = "invalid"
ConeName(7) = "invalid"

LoadedName(0) = "Not available (default)"
LoadedName(1) = "Loaded"
LoadedName(2) = "Unloaded/Ballast"
LoadedName(3) = "invalid"

QualSpeedName(0) = "low/GNSS (default)"
QualSpeedName(1) = "High"

QualHeadName(0) = "Low (default)"
QualHeadName(1) = "High"

For i = 0 To UBound(UsaCanFiName)
    UsaCanFiName(i) = "Not known"
Next i
UsaCanFiName(1) = "Metrological and Hydrological"
UsaCanFiName(2) = "Vessel/Lock Scheduling"
UsaCanFiName(32) = "Seaway Specific Messages"

For i = 0 To UBound(UsaCan1IdName)
    UsaCan1IdName(i) = "not known"
    UsaCan2IdName(i) = "not known"
    UsaCan32IdName(i) = "not known"
Next i
UsaCan1IdName(1) = "Weather Station Report"
UsaCan1IdName(2) = "Wind Information"
UsaCan1IdName(3) = "Water Level"
UsaCan1IdName(6) = "Water Flow"
    
UsaCan2IdName(1) = "Lockage Order"
UsaCan2IdName(2) = "Estimated Lock Times"

UsaCan32IdName(1) = "Version"

ReportType366Name(0) = "Site Location"
ReportType366Name(1) = "Station ID"
ReportType366Name(2) = "Wind"
ReportType366Name(3) = "Water Level"
ReportType366Name(4) = "Current Flow 2D (x,y) {Awaiting verification)"
ReportType366Name(5) = "Current Flow 3D (x,y,z) {Awaiting verification)"
ReportType366Name(6) = "Horizontal Current Flow {Awaiting verification)"
ReportType366Name(7) = "Sea State {Awaiting verification)"
ReportType366Name(8) = "Salinity {Awaiting verification)"
ReportType366Name(9) = "Weather {Awaiting verification)"
ReportType366Name(10) = "Air gap / Air draught {23Mar15 v3)"
For i = 11 To UBound(ReportType366Name)
ReportType366Name(i) = "Reserved for futute use"
Next i

Owner366Name(0) = "Coastal Directorate (USCG)"
Owner366Name(1) = "Hydrographic Office (NOAA)"
Owner366Name(2) = "Inland Waterway Authority (ACE)"
Owner366Name(3) = "Port Authority"
Owner366Name(4) = "Meteorological Service"
For i = 5 To UBound(Owner366Name)
    Owner366Name(i) = "reserved for future use"
Next i
Owner366Name(15) = "unknown"

Timeout366Name(0) = "Never (default)"
Timeout366Name(1) = "10 min"
Timeout366Name(2) = "1 hr"
Timeout366Name(3) = "6 hrs"
Timeout366Name(4) = "12 hrs"
Timeout366Name(5) = "24 hrs"
Timeout366Name(6) = "reserved for future use"
Timeout366Name(7) = "reserved for future use"

SensorData366Name(0) = "No data (default)"
SensorData366Name(1) = "raw real time"
SensorData366Name(2) = "real time with Quality Control"
SensorData366Name(3) = "predicted"
SensorData366Name(4) = "Forecast"
SensorData366Name(5) = "Nowcast"
SensorData366Name(6) = "reserved for future use"
SensorData366Name(7) = "Sensor not available"

AiAvailable1Name(0) = "Default (no sequence number)"
AiAvailable1Name(1) = "AI Available"

AiResponse1Name(0) = "unable to respond"
AiResponse1Name(1) = "reception acknowledged"
AiResponse1Name(2) = "response to follow"
AiResponse1Name(3) = "able to respond but currently inhibited"
AiResponse1Name(4) = "spare for future use"
AiResponse1Name(5) = "spare for future use"
AiResponse1Name(6) = "spare for future use"
AiResponse1Name(7) = "spare for future use"

SignalStatus1Name(0) = "In regular service"
SignalStatus1Name(1) = "In irregular service"

ItuWmoTendencyName(0) = "Increasing, then decreasing"
ItuWmoTendencyName(1) = "Increasing, then steady"
ItuWmoTendencyName(2) = "Increasing steadily"
ItuWmoTendencyName(3) = "Decreasing or steady"
ItuWmoTendencyName(4) = "Steady"
ItuWmoTendencyName(5) = "Decreasing, then increasing"
ItuWmoTendencyName(6) = "Decreasing, then steady"
ItuWmoTendencyName(7) = "Decreasing steadily"
ItuWmoTendencyName(8) = "Increasing or steady, then decreasing"
For i = 9 To UBound(ItuWmoTendencyName)
    ItuWmoTendencyName(i) = "{Invalid [" & i & "]}"
Next i

ItuWmoWeatherName(0) = "clear (no clouds at any level)"
ItuWmoWeatherName(1) = "cloudy"
ItuWmoWeatherName(2) = "rain"
ItuWmoWeatherName(3) = "fog"
ItuWmoWeatherName(4) = "snow"
ItuWmoWeatherName(5) = "typhoon/hurricane"
ItuWmoWeatherName(6) = "monsoon"
ItuWmoWeatherName(7) = "thunderstorm"
ItuWmoWeatherName(8) = "not available (default)"
For i = 9 To UBound(ItuWmoWeatherName)
    ItuWmoWeatherName(i) = "{reserved for future use [" & i & "]}"
Next i

UpDown366Name(0) = "Down bound"
UpDown366Name(1) = "Up bound"

Level366Name(0) = "Relative to reference datum"
Level366Name(1) = "Water Depth"

DatumA366Name(0) = "MLLW"
DatumA366Name(1) = "IGLD-85"
DatumA366Name(2) = "Reserved for future use"
DatumA366Name(3) = "Reserved for future use"

DatumB366Name(0) = "MLLW"
DatumB366Name(1) = "IGLD-85"
DatumB366Name(2) = "Water Depth"
DatumB366Name(3) = "STND"
DatumB366Name(4) = "MHHW"
DatumB366Name(5) = "MHW"
DatumB366Name(6) = "MSL"
DatumB366Name(7) = "MLW"
DatumB366Name(8) = "NGVD"
DatumB366Name(9) = "NAVD"
DatumB366Name(10) = "WGS-84"
DatumB366Name(11) = "LAT"
DatumB366Name(12) = "Pool"
DatumB366Name(13) = "Gauge"
DatumB366Name(14) = "Local river datum"
For i = 15 To UBound(DatumB366Name)
    DatumB366Name(i) = "Reserved for future use"
Next i
DatumB366Name(31) = "Unknown/Unavailable (default)"

Tug200Name(0) = "None"
Tug200Name(1) = "One"
Tug200Name(2) = "Two"
Tug200Name(3) = "Three"
Tug200Name(4) = "Four"
Tug200Name(5) = "Five"
Tug200Name(6) = "Six"
Tug200Name(7) = "Unknown (default)"

LockStatus200Name(0) = "Operational"
LockStatus200Name(1) = "Limited Operation"
LockStatus200Name(2) = "Out of Order"
LockStatus200Name(3) = "Not Available"

SignalForm200Name(0) = "Unknown (default)"
SignalForm200Name(1) = "One"
SignalForm200Name(2) = "Two Horizontal"
SignalForm200Name(3) = "Three Horizontal"
SignalForm200Name(4) = "Two Vertical"
SignalForm200Name(5) = "One above Two"
SignalForm200Name(6) = "One above Three"
SignalForm200Name(7) = "Two above Two"
SignalForm200Name(8) = "Three Vertical"
SignalForm200Name(9) = "Three above Three"
SignalForm200Name(10) = "Three above Three above Three"
SignalForm200Name(11) = "Two above Four"
SignalForm200Name(12) = "One + Right Arrow"
SignalForm200Name(13) = "Left Arrow + One"
SignalForm200Name(14) = "Up Arrow above One"
SignalForm200Name(15) = "Not used"

SignalStatus200Name(0) = "Unknown (default)"
SignalStatus200Name(1) = "No Light"
SignalStatus200Name(2) = "White"
SignalStatus200Name(3) = "Yellow"
SignalStatus200Name(4) = "Green"
SignalStatus200Name(5) = "Red"
SignalStatus200Name(6) = "White Flashing"
SignalStatus200Name(7) = "Yellow Flashing"

Weather200Name(0) = "Unknown (default)"
Weather200Name(1) = "Wind"
Weather200Name(2) = "Rain"
Weather200Name(3) = "Snow and Ice"
Weather200Name(4) = "Thunderstorm"
Weather200Name(5) = "Fog"
Weather200Name(6) = "Low Temperature"
Weather200Name(7) = "High Temperature"
Weather200Name(8) = "Flood"
Weather200Name(9) = "Fire in the Forests"
For i = 10 To UBound(Weather200Name)
    Weather200Name(i) = "Invalid"
Next i

Category200Name(0) = "Unknown (default)"
Category200Name(1) = "Slight"
Category200Name(2) = "Medium"
Category200Name(3) = "Strong, Heavy"

Direction200Name(0) = "Unknown (default)"
Direction200Name(1) = "North"
Direction200Name(2) = "North East"
Direction200Name(3) = "East"
Direction200Name(4) = "South East"
Direction200Name(5) = "South"
Direction200Name(6) = "South West"
Direction200Name(7) = "West"
Direction200Name(8) = "North West"
For i = 10 To UBound(Direction200Name)
    Direction200Name(i) = "Invalid"
Next i

Racon235Name(0) = "No Racon Installed"
Racon235Name(1) = "Racon not Monitored"
Racon235Name(2) = "Racon Operational"
Racon235Name(3) = "Racon Error"

Light235Name(0) = "No Light of No Monitoring"
Light235Name(1) = "Light On"
Light235Name(2) = "Light Off"
Light235Name(3) = "Light Error"

Alarm235Name(0) = "Good Health"
Alarm235Name(1) = "Alarm"

Biit235Name(0) = "Normal"
Biit235Name(1) = "Failure"

ExtFld235Name(0) = "Not attached"
ExtFld235Name(1) = "Attached"

HighLow235Name(0) = "High"
HighLow235Name(1) = "Low"

OnOffName(0) = "Off"
OnOffName(1) = "On"

DgUnitsName(0) = "Not available (default)"
DgUnitsName(1) = "in kg"
DgUnitsName(2) = "in tonnes (1,000 kg)"
DgUnitsName(3) = "in 1,000 tonnes (1,000,000 kg)"

DgCodeName(0) = "Not available (default)"
DgCodeName(1) = "IMDG Code (in packed form)"
DgCodeName(2) = "IGC Code"
DgCodeName(3) = "BC Code (from 1.1.2011 IMSBC)"
DgCodeName(4) = "MARPOL Annex I list of oild (Appendix 1)"
DgCodeName(5) = "MARPOL Annex II IBC Code"
For i = 6 To UBound(DgCodeName)
    DgCodeName(i) = "reserved for future use"
Next i

BcCodeName(0) = "Not available (default)"
BcCodeName(1) = "A"
BcCodeName(2) = "B"
BcCodeName(3) = "C"
BcCodeName(4) = "MHB - Material Hazardous in Bulk"
For i = 5 To UBound(BcCodeName)
    BcCodeName(i) = "reserved for future use"
Next i

Marpol1Name(0) = "Not available (default)"
Marpol1Name(1) = "Ashphalt solutions"
Marpol1Name(2) = "Oils"
Marpol1Name(3) = "Distillates"
Marpol1Name(4) = "Gas Oil"
Marpol1Name(5) = "Gasoline blending products"
Marpol1Name(6) = "Gasoline"
Marpol1Name(7) = "Jet Fuels"
Marpol1Name(8) = "Naptha"
For i = 9 To UBound(Marpol1Name)
    Marpol1Name(i) = "reserved for future use"
Next i

Marpol2Name(0) = "Not available (default)"
Marpol2Name(1) = "Category X"
Marpol2Name(2) = "Category Y"
Marpol2Name(3) = "Category Z"
Marpol2Name(4) = "Other substances"
For i = 5 To UBound(Marpol2Name)
    Marpol2Name(0) = "reserved for future use"
Next i

SalinityType366Name(0) = "measured"
SalinityType366Name(1) = "calculated using PSS-78"
SalinityType366Name(2) = "calculated using other method"
SalinityType366Name(3) = "Reserved for future use"

Shape1Name(0) = "Circle or Point"
Shape1Name(1) = "Rectangle"
Shape1Name(2) = "Sector"
Shape1Name(3) = "Polyline"
Shape1Name(4) = "Polygon"
Shape1Name(5) = "Free Text"
For i = 6 To UBound(Shape1Name)
    Shape1Name(i) = "reserved"
Next i

Sender1Name(0) = "Ship (default)"
Sender1Name(1) = "Authority"
For i = 2 To UBound(Sender1Name)
    Sender1Name(i) = "reserved for future use"
Next i

Route1Name(0) = "Not Available (default)"
Route1Name(1) = "Mandatory route"
Route1Name(2) = "Recommended route"
Route1Name(3) = "Alternative route"
Route1Name(4) = "Recommended route through ice"
Route1Name(5) = "Ship Route Plan"
For i = 6 To UBound(Route1Name)
    Route1Name(i) = "reserved for future use"
Next i

SolasEquipmentName(1) = "AIS Class A"
SolasEquipmentName(2) = "ATA (Automatic Tracking Aid)"
SolasEquipmentName(3) = "BNWAS (Bridge Navigation Watch Alarm System)"
SolasEquipmentName(4) = "ECDIS Back-up"
SolasEquipmentName(5) = "ECDIS/Paper Nautical Chart"
SolasEquipmentName(6) = "echo sounder"
SolasEquipmentName(7) = "electronic plotting aid"
SolasEquipmentName(8) = "emergency steering gear"
SolasEquipmentName(9) = "navigation system (GPS,Loran,GLONASS)"
SolasEquipmentName(10) = "gyro compass"
SolasEquipmentName(11) = "LRIT"
SolasEquipmentName(12) = "magnetic compass"
SolasEquipmentName(13) = "NAVTEX"
SolasEquipmentName(14) = "radar (ARPA)"
SolasEquipmentName(15) = "radar (S-band)"
SolasEquipmentName(16) = "radar (X-band)"
SolasEquipmentName(17) = "radio HF"
SolasEquipmentName(18) = "radio INMARSAT"
SolasEquipmentName(19) = "radio MF"
SolasEquipmentName(20) = "radio VHF"
SolasEquipmentName(21) = "speed Log (over ground)"
SolasEquipmentName(22) = "speed Log (through water)"
SolasEquipmentName(23) = "THD (Transmitting Heading Device)"
SolasEquipmentName(24) = "track control system"
SolasEquipmentName(25) = "VDR/S-VDR"
SolasEquipmentName(26) = "(reserved for future use)"

IceClassName(0) = "not classified"
IceClassName(1) = "IACS PC1"
IceClassName(2) = "IACS PC2"
IceClassName(3) = "IACS PC3"
IceClassName(4) = "IACS PC4"
IceClassName(5) = "IACS PC5"
IceClassName(6) = "IACS PC6/FSICR IA Super/RS Arc5"
IceClassName(7) = "IACS PC7/FSICR IA/RS Arc4"
IceClassName(8) = "FSICR IB/RS Ice3"
IceClassName(9) = "FSICR IC/RS Ice2"
IceClassName(10) = "RS Ice1"
For i = 11 To 14
    IceClassName(i) = "(reserved for future use)"
Next i
IceClassName(5) = "not available (default)"

OperationalName(0) = "not available (default)"
OperationalName(1) = "Operational"
OperationalName(2) = "Not Operational"
OperationalName(3) = "no data (status unknown)"

#If False Then
Example999Name(0) = ""
#End If

ArgFmt(0) = "0"
ArgFmt(1) = "0.0"
ArgFmt(2) = "0.00"
ArgFmt(3) = "0.000"
ArgFmt(4) = "0.0000"

End Sub
Public Function EniName(Eni As String)
Dim CountryCode As Integer
On Error GoTo EniName_Error    'non-numeric eni
CountryCode = CInt(Mid$(Eni, 1, 3))
On Error GoTo 0
Select Case CountryCode
Case 190 To 199
    EniName = "reserved"
Case 20 To 39
    EniName = "Netherlands"
Case 40 To 59
    EniName = "Germany"
Case 60 To 69
    EniName = "Belgium"
Case 70 To 79
    EniName = "Switzerland"
Case 80 To 99
    EniName = "reserved for vessels from countries that are not party to the Mannheim Convention and for which a Rhine Vessel certificate has been issued before 01.04.2007"
Case 100 To 119
    EniName = "Norway"
Case 120 To 139
    EniName = "Denmark"
Case 140 To 159
    EniName = "United Kingdom"
Case 160 To 169
    EniName = "Iceland"
Case 170 To 179
    EniName = "Ireland"
Case 180 To 189
    EniName = "Portugal"
Case 190 To 199
    EniName = "reserved"
Case 200 To 219
    EniName = "Luxembourg"
Case 220 To 239
    EniName = "Finland"
Case 240 To 259
    EniName = "Poland"
Case 260 To 269
    EniName = "Estonia"
Case 270 To 279
    EniName = "Lithuania"
Case 280 To 289
    EniName = "Latvia"
Case 290 To 299
    EniName = "reserved"
Case 300 To 309
    EniName = "Austria"
Case 310 To 319
    EniName = "Liechtenstein"
Case 320 To 329
    EniName = "Czech Republic"
Case 330 To 339
    EniName = "Slovakia"
Case 340 To 349
    EniName = "Hungary"
Case 350 To 359
    EniName = "Croatia"
Case 360 To 369
    EniName = "Serbia"
Case 370 To 379
    EniName = "Bosnia and Herzegovina"
Case 380 To 399
    EniName = "reserved"
Case 400 To 419
    EniName = "Russian Federation"
Case 420 To 439
    EniName = "Ukraine"
Case 440 To 449
    EniName = "Belarus"
Case 450 To 459
    EniName = "Republic of Moldova"
Case 460 To 469
    EniName = "Romania"
Case 470 To 479
    EniName = "Bulgaria"
Case 480 To 489
    EniName = "Georgia"
Case 490 To 499
    EniName = "reserved"
Case 500 To 519
    EniName = "Turkey"
Case 520 To 539
    EniName = "Greece"
Case 540 To 549
    EniName = "Cyprus"
Case 550 To 559
    EniName = "Albania"
Case 560 To 569
    EniName = "The Former Yugoslav Republic of Macedonia"
Case 570 To 579
    EniName = "Slovenia"
Case 580 To 589
    EniName = "Montenegro"
Case 590 To 599
    EniName = "reserved"
Case 600 To 619
    EniName = "Italy"
Case 620 To 639
    EniName = "Spain"
Case 640 To 649
    EniName = "Andorra"
Case 650 To 659
    EniName = "Malta"
Case 660 To 669
    EniName = "Monaco"
Case 670 To 679
    EniName = "San Marino"
Case 680 To 699
    EniName = "reserved"
Case 700 To 719
    EniName = "Sweden"
Case 720 To 739
    EniName = "Canada"
Case 740 To 759
    EniName = "United States of America"
Case 760 To 769
    EniName = "Israel"
Case 770 To 799
    EniName = "reserved"
Case 800 To 809
    EniName = "Azerbaijan"
Case 810 To 819
    EniName = "Kazakhstan"
Case 820 To 829
    EniName = "Kyrgyzstan"
Case 830 To 839
    EniName = "Tajikistan"
Case 840 To 849
    EniName = "Turkmenistan"
Case 850 To 859
    EniName = "Uzbekistan"
Case 860 To 869
    EniName = "Iran"
Case 870 To 899
    EniName = "reserved"

Case Else
    EniName = "{not known}"
End Select
EniName_Exit:
    Exit Function
EniName_Error:
    EniName = "Invalid ENI number " & Right$(Eni, Len(Eni) - 1)
    GoTo EniName_Exit
End Function


Public Function EriName(kb As Integer) As String
Select Case kb
Case Is = "8000"
    EriName = "Vessel type unknown"
Case Is = "8010"
    EriName = "Motor freighter"
Case Is = "8020"
    EriName = "Motor tanker"
Case Is = "8021"
    EriName = "Motor tanker, liquid cargo, type N"
Case Is = "8022"
    EriName = "Motor tanker, liquid cargo, type C"
Case Is = "8023"
    EriName = "Motor tanker, dry cargo as if liquid"
Case Is = "8030"
    EriName = "Container vessel"
Case Is = "8040"
    EriName = "Gas tanker"
Case Is = "8050"
    EriName = "Motor freighter, tug"
Case Is = "8060"
    EriName = "Motor tanker, tug"
Case Is = "8070"
    EriName = "Motor freighter with one of more ships alongside"
Case Is = "8080"
    EriName = "Motor freighter with tanker"
Case Is = "8090"
    EriName = "Motor freighter pushing one or more freighters"
Case Is = "8100"
    EriName = "Motor freighte pushing at least one tank-ship"
Case Is = "8110"
    EriName = "Tug, freighter"
Case Is = "8120"
    EriName = "Tug, tanker"
Case Is = "8130"
    EriName = "Tug freighter, coupled"
Case Is = "8140"
    EriName = "Tug, freighter/tanker, coupled"
Case Is = "8150"
    EriName = "Freightbarge"
Case Is = "8160"
    EriName = "Tankbarge"
Case Is = "8161"
    EriName = "Tankbarge, liquid cargo, type N"
Case Is = "8162"
    EriName = "Tankbarge, liquid cargo, type C"
Case Is = "8163"
    EriName = "Tankbarge, dry cargos as if liquid"
Case Is = "8170"
    EriName = "Freightbarge with containers"
Case Is = "8180"
    EriName = "Tankbarge, gas"
Case Is = "8210"
    EriName = "Pushtow, one cargo barge"
Case Is = "8220"
    EriName = "Pushtow, two cargo barge"
Case Is = "8230"
    EriName = "Pushtow, three cargo barge"
Case Is = "8240"
    EriName = "Pushtow, four cargo barge"
Case Is = "8250"
    EriName = "Pushtow, five cargo barge"
Case Is = "8260"
    EriName = "Pushtow, six cargo barge"
Case Is = "8270"
    EriName = "Pushtow, seven cargo barge"
Case Is = "8280"
    EriName = "Pushtow, eight cargo barge"
Case Is = "8290"
    EriName = "Pushtow, nine cargo barge"
Case Is = "8310"
    EriName = "Pushtow, one tank/gas barge"
Case Is = "8320"
    EriName = "Pushtow, two barges at least one tanker or gas barge"
Case Is = "8330"
    EriName = "Pushtow, three barges at least one tanker or gas barge"
Case Is = "8340"
    EriName = "Pushtow, four barges at least one tanker or gas barge"
Case Is = "8350"
    EriName = "Pushtow, five barges at least one tanker or gas barge"
Case Is = "8360"
    EriName = "Pushtow, six barges at least one tanker or gas barge"
Case Is = "8370"
    EriName = "Pushtow, seven barges at least one tanker or gas barge"
Case Is = "8380"
    EriName = "Pushtow, eight barges at least one tanker or gas barge"
Case Is = "8390"
    EriName = "Pushtow, nine barges at least one tanker or gas barge"
Case Is = "8400"
    EriName = "Tug, single"
Case Is = "8410"
    EriName = "Tug, one or more tows"
Case Is = "8420"
    EriName = "Tug, assisting a vessel or linked combination"
Case Is = "8430"
    EriName = "Pushboat, single"
Case Is = "8440"
    EriName = "Passenger ship,ferry,cruise ship,red cross ship"
Case Is = "8441"
    EriName = "Ferry"
Case Is = "8442"
    EriName = "Red cross ship"
Case Is = "8443"
    EriName = "Cruise ship"
Case Is = "8444"
    EriName = "Passenger ship without accomodation"
Case Is = "8450"
    EriName = "Service vessel, police patrol, port service"
Case Is = "8460"
    EriName = "Vessel, work maintenence craft,floating derrick,cable ship,buoy ship,dredge"
Case Is = "8470"
    EriName = "Object towed, not otherwise specified"
Case Is = "8480"
    EriName = "Fishing boat"
Case Is = "8490"
    EriName = "Bunker ship"
Case Is = "8500"
    EriName = "Barge, tanker, chemical"
Case Is = "8510"
    EriName = "Object. not otherwise specified"
Case Is = "1500"
    EriName = "General cargo Vessel maritime"
Case Is = "1510"
    EriName = "Unit carrier maritime"
Case Is = "1520"
    EriName = "Bulk carrier maritime"
Case Is = "1530"
    EriName = "Tanker"
Case Is = "1540"
    EriName = "Liquified gas tanker"
Case Is = "1850"
    EriName = "Pleasure craft longer than 20 meters"
Case Is = "1900"
    EriName = "Fast ship"
Case Is = "1910"
    EriName = "Hydofoil"
Case Else
    EriName = "{not known}"
End Select
End Function

Function SecondDes(Val As String) As String
SecondDes = "Second of UTC timestamp"
If CInt(Val) = 60 Then SecondDes = "not available (default)"
If CInt(Val) = 61 Then SecondDes = "Positioning system is in manual input mode"
If CInt(Val) = 62 Then SecondDes = "Electronic Positioning Fixing System operates in estimated (dead reckoning) mode"
If CInt(Val) = 63 Then SecondDes = "Positioning system is inoperative"
End Function

Public Function ParameterDes(ParameterCode) As String
    Select Case ParameterCode
    Case Is = "c"
        ParameterDes = "Unix Time"
    Case Is = "i"
        ParameterDes = "Information"
    Case Is = "s"
        ParameterDes = "Source"
    Case Is = "d"
        ParameterDes = "Destination"
    Case Is = "x"
        ParameterDes = "Counter"
    Case Is = "G", "g"
        ParameterDes = "Group"
    Case Else
        ParameterDes = "Unknown Parameter"
    End Select
End Function

Public Function IecFormatDes(IecFormat As String) As String
    Select Case clsSentence.IecFormat
'Proprietary sentences
    Case Is = "GHP": IecFormatDes = "GateHouse internal message type P"
    Case Is = "THAJ": IecFormatDes = "Data Slot & Jitter (True Heading)"
    Case Is = "THAR": IecFormatDes = "Data Link HDLC CRC Error (True Heading)"
    Case Is = "TAG": IecFormatDes = "MarineCom internal message"
    Case Is = "ASHR": IecFormatDes = "Attitude Sensor, INS"
    Case Is = "GSSRV_": IecFormatDes = "Vehicle Attitude"
    Case Is = "RDID": IecFormatDes = "RDI Proprietary Heading, Pitch, Roll"
'NMEA Semtemces
Case Is = "AAM": IecFormatDes = "Waypoint arrival alarm"
Case Is = "ABK": IecFormatDes = "AIS addressed and binary broadcast acknowledgement"
Case Is = "ABM": IecFormatDes = "AIS Addressed binary and safety related message"
Case Is = "ACA": IecFormatDes = "AIS channel assignment message"
Case Is = "ACK": IecFormatDes = "Acknowledge alarm"
Case Is = "AIR": IecFormatDes = "AIS Interrogation request"
Case Is = "AKD": IecFormatDes = "Acknowledge detail alarm condition"
Case Is = "ALA": IecFormatDes = "Report detailed alarm condition"
Case Is = "ALR": IecFormatDes = "Set alarm state"
Case Is = "APB": IecFormatDes = "Heading/track controller (autopilot) sentence B"
Case Is = "BBM": IecFormatDes = "AIS Broadcast binary message"
Case Is = "BEC": IecFormatDes = "Bearing and distance to waypoint - dead reckoning"
Case Is = "BOD": IecFormatDes = "Bearing origin to destination"
Case Is = "BRM": IecFormatDes = "Base Station Options Reply of Received Messages (True Heading only ?)"
Case Is = "BWC": IecFormatDes = "Bearing and distance to waypoint - great circle"
Case Is = "BWR": IecFormatDes = "Bearing and distance to waypoint - rhumb line"
Case Is = "BWW": IecFormatDes = "Bearing waypoint to waypoint"
Case Is = "CBR": IecFormatDes = "Configure Broadcast Rates for AIS AtoN Station Message Command"
Case Is = "CUR": IecFormatDes = "Water current layer - Multi-layer water current data"
Case Is = "DBS": IecFormatDes = "Depth below surface"
Case Is = "DBT": IecFormatDes = "Depth below transducer"
Case Is = "DDC": IecFormatDes = "Display Dimming Control"
Case Is = "DOR": IecFormatDes = "Door status detection"
Case Is = "DPT": IecFormatDes = "Depth"
Case Is = "DSC": IecFormatDes = "Digital selective calling information"
Case Is = "DSE": IecFormatDes = "Expanded digital selective calling"
Case Is = "DTM": IecFormatDes = "Datum reference"
Case Is = "ETL": IecFormatDes = "Engine telegraph operation status"
Case Is = "EVE": IecFormatDes = "General event message"
Case Is = "FIR": IecFormatDes = "Fire detection"
Case Is = "FSI": IecFormatDes = "Frequency set information"
Case Is = "GBS": IecFormatDes = "GNSS satellite fault detection"
Case Is = "GEN": IecFormatDes = "Generic binary information"
Case Is = "GFA": IecFormatDes = "GNSS fix accuracy and integrity"
Case Is = "GGA": IecFormatDes = "Global positioning system (GPS) fix data"
Case Is = "GLL": IecFormatDes = "Geographic position - latitude/longitude"
Case Is = "GNS": IecFormatDes = "GNSS fix dataGRS - GNSS range residuals"
Case Is = "GSA": IecFormatDes = "GNSS DOP and active satellites"
Case Is = "GST": IecFormatDes = "GNSS pseudorange noise statistics"
Case Is = "GSV": IecFormatDes = "GNSS satellites in view"
Case Is = "HBT": IecFormatDes = "Heartbeat supervision sentence"
Case Is = "HDG": IecFormatDes = "Heading, deviation and variation"
Case Is = "HDT": IecFormatDes = "Heading true"
Case Is = "HMR": IecFormatDes = "Heading monitor receive"
Case Is = "HMS": IecFormatDes = "Heading monitor set"
Case Is = "HSC": IecFormatDes = "Heading steering command"
Case Is = "HSS": IecFormatDes = "Hull stress surveillance systems"
Case Is = "HTC": IecFormatDes = "Heading/track control command"
Case Is = "HTC": IecFormatDes = "Heading/track control command"
Case Is = "HTD": IecFormatDes = "Heading /track control data"
Case Is = "LR1": IecFormatDes = "AIS long-range reply sentence 1"
Case Is = "LR2": IecFormatDes = "AIS long-range reply sentence 2"
Case Is = "LR3": IecFormatDes = "AIS long-range reply sentence 3"
Case Is = "LRF": IecFormatDes = "AIS long-range function"
Case Is = "LRI": IecFormatDes = "AIS long-range interrogation"
Case Is = "M01": IecFormatDes = "Zeno Met Sensors"
Case Is = "MEB": IecFormatDes = "Message input for broadcast command"
Case Is = "MSK": IecFormatDes = "MSK receiver interface"
Case Is = "MSS": IecFormatDes = "MSK receiver signal status"
Case Is = "MTW": IecFormatDes = "Water temperature"
Case Is = "MWD": IecFormatDes = "Wind direction and speed"
Case Is = "MWV": IecFormatDes = "Wind speed and angle"
Case Is = "NRM": IecFormatDes = "NAVTEX receiver mask"
Case Is = "NRX": IecFormatDes = "NAVTEX received message"
Case Is = "OSD": IecFormatDes = "Own ship data"
Case Is = "POS": IecFormatDes = "Device position and ship dimensions report or configuration command"
Case Is = "PRC": IecFormatDes = "Propulsion remote control status"
Case Is = "RMA": IecFormatDes = "Recommended minimum specific LORAN-C data"
Case Is = "RMB": IecFormatDes = "Recommended minimum navigation information"
Case Is = "RMC": IecFormatDes = "Recommended minimum specific GNSS data"
Case Is = "ROT": IecFormatDes = "Rate of Turn"
Case Is = "ROR": IecFormatDes = "Rudder order status"
Case Is = "RPM": IecFormatDes = "Revolutions"
Case Is = "RSA": IecFormatDes = "Rudder sensor angle"
Case Is = "RSD": IecFormatDes = "Radar system data"
Case Is = "RTE": IecFormatDes = "Routes"
Case Is = "SFI": IecFormatDes = "Scanning frequency information"
Case Is = "SSD": IecFormatDes = "AIS ship static data"
Case Is = "STN": IecFormatDes = "Multiple data ID"
Case Is = "THS": IecFormatDes = "True heading and status"
Case Is = "TLB": IecFormatDes = "Target label"
Case Is = "TLL": IecFormatDes = "Target latitude and longitude"
Case Is = "TRC": IecFormatDes = "Thruster control data"
Case Is = "TRD": IecFormatDes = "Thruster response data"
Case Is = "TTD": IecFormatDes = "Tracked Target Data"
Case Is = "TTM": IecFormatDes = "Tracked target message"
Case Is = "TUT": IecFormatDes = "Transmission of multi-language text"
Case Is = "TXT": IecFormatDes = "Text transmission"
Case Is = "VBW": IecFormatDes = "Dual ground/water speed"
Case Is = "VDM": IecFormatDes = "AIS VHF data-link message"
Case Is = "VDO": IecFormatDes = "AIS VHF data-link own-vessel report"
Case Is = "VDR": IecFormatDes = "Set and drift"
Case Is = "VER": IecFormatDes = "Version"
Case Is = "VHW": IecFormatDes = "Water speed and heading"
Case Is = "VLW": IecFormatDes = "Dual ground/water distance"
Case Is = "VPw": IecFormatDes = "Speed measured parallel to wind"
Case Is = "VSD": IecFormatDes = "AIS voyage static data"
Case Is = "VTG": IecFormatDes = "Course over ground and ground speed"
Case Is = "WAT": IecFormatDes = "Water level detection"
Case Is = "WCV": IecFormatDes = "Waypoint closure velocity"
Case Is = "WNC": IecFormatDes = "Distance waypoint to waypoint"
Case Is = "WPL": IecFormatDes = "Waypoint location"
Case Is = "XDR": IecFormatDes = "Transducer measurements"
Case Is = "XTE": IecFormatDes = "Cross-track error, measured"
Case Is = "XTR": IecFormatDes = "Cross-track error, dead reckoning"
Case Is = "ZDA": IecFormatDes = "Time and date"
Case Is = "ZDL": IecFormatDes = "Time and distance to variable point"
Case Is = "ZFO": IecFormatDes = "UTC and time from origin waypoint"
Case Is = "ZTG": IecFormatDes = "UTC and time to destination waypoint"
    Case Else
        IecFormatDes = "Undefined IEC format"
    End Select
End Function

Public Function TalkerDes(IecTalkerID As String) As String
    Select Case clsSentence.IecTalkerID
    Case Is = "AI"
        TalkerDes = "Mobile class A or B"
    Case Is = "AN"
        TalkerDes = "Aids to Navigation"
    Case Is = "AB"
        TalkerDes = "AIS Base Station (NMEA v4.10)"
    Case Is = "BS"
        TalkerDes = "AIS Base Station"
    Case Is = "AL"
        TalkerDes = "Limited Base Station (IEC 62320-1)"
    Case Is = "AS"
        TalkerDes = "Simplex(IEC)/Limited(NMEA) Base Sation"
    Case Is = "AD"
        TalkerDes = "Duplex Repeater Station (IEC61162-1) "
    Case Is = "AR"
        TalkerDes = "Receiving Station (IEC 62320-1)"
    Case Is = "AT"
        TalkerDes = "Transmitting Station (NMEA v4.10)"
    Case Is = "AX"
        TalkerDes = "Simplex Repeater Station (NMEA v4.10)"
    Case Is = "GP"
        TalkerDes = "Global positioning system (GPS)"
    Case Is = "EP"
        TalkerDes = "Emergency position indicating radio beacon (EPIRB)"
    Case Is = "P": TalkerDes = "Proprietary Sentence"
    Case Is = "ZA"
        TalkerDes = "Timekeeper, time/date: atomic clock"
    Case Is = "ZC"
        TalkerDes = "Timekeeper, time/date: chronometer"
    Case Is = "ZQ"
        TalkerDes = "Timekeeper, time/date: quartz"
    Case Is = "ZV"
        TalkerDes = "Timekeeper, time/date: radio update"
Case Is = "AG": TalkerDes = "Heading/track controller (autopilot) general "
Case Is = "AP": TalkerDes = "Heading/track controller (autopilot)magnetic "
Case Is = "AI": TalkerDes = "Automatic identification system "
Case Is = "AN": TalkerDes = "Automatic identification system AtoN station "
Case Is = "AR": TalkerDes = "Automatic identification system receiving station "
Case Is = "AS": TalkerDes = "Automatic identification system station (ITU-R M.1371 (limited base station)) "
Case Is = "AT": TalkerDes = "Automatic identification system station transmitting station "
Case Is = "AX": TalkerDes = "Automatic identification system simplex repeater station "
Case Is = "BI": TalkerDes = "Bilge system "
Case Is = "BN": TalkerDes = "Bridge Navigational Watch Alarm System "
Case Is = "CD": TalkerDes = "Communications: digital selective calling (DSC) "
Case Is = "CR": TalkerDes = "Communications: data receiver "
Case Is = "CS": TalkerDes = "Communications: satellite "
Case Is = "CT": TalkerDes = "Communications: radio-telephone (MF/HF) "
Case Is = "CV": TalkerDes = "radio-telephone (VHF) "
Case Is = "CX": TalkerDes = "Communications: scanning receiver "
Case Is = "DE": TalkerDes = "DECCA navigator "
Case Is = "DF": TalkerDes = "Direction finder "
Case Is = "DU": TalkerDes = "Duplex repeater station "
Case Is = "EC": TalkerDes = "Electronic chart system (ECS) "
Case Is = "EI": TalkerDes = "Electronic chart display and information system (ECDIS) "
Case Is = "EP": TalkerDes = "Emergency position indicating radio beacon (EPIRB) "
Case Is = "ER": TalkerDes = "Engine room monitoring system "
Case Is = "FD": TalkerDes = "Fire door controller/monitoring system "
Case Is = "FE": TalkerDes = "Fire extinguisher system "
Case Is = "FR": TalkerDes = "Fire detection system "
Case Is = "FS": TalkerDes = "Fire sprinkler system "
Case Is = "GA": TalkerDes = "Galileo positioning system "
Case Is = "GP": TalkerDes = "Global positioning system (GPS) "
Case Is = "GL": TalkerDes = "GLONASS positioning system "
Case Is = "GN": TalkerDes = "Global navigation satellite system (GNSS) "
Case Is = "HC": TalkerDes = "Heading sensors: compass magnetic "
Case Is = "HE": TalkerDes = "Heading sensors: gyro north seeking "
Case Is = "HF": TalkerDes = "Heading sensors: fluxgate "
Case Is = "HN": TalkerDes = "Heading sensors: gyro non-north seeking "
Case Is = "HD": TalkerDes = "Hull door controller/monitoring system "
Case Is = "HS": TalkerDes = "Hull stress monitoring "
Case Is = "II": TalkerDes = "Integrated instrumentation "
Case Is = "IN": TalkerDes = "Integrated navigation "
Case Is = "LC": TalkerDes = "LORAN: LORAN-C "
Case Is = "NL": TalkerDes = "Navigation Light Controller "
Case Is = "P": TalkerDes = "Proprietary code "
Case Is = "RA": TalkerDes = "Radar and/or radar plotting "
Case Is = "RC": TalkerDes = "Propulsion machinery including remote control "
Case Is = "SD": TalkerDes = "Sounder depth "
Case Is = "SG": TalkerDes = "Steering gear/steering engine "
Case Is = "SN": TalkerDes = "Electronic positioning system other/general "
Case Is = "SS": TalkerDes = "Sounder scanning "
Case Is = "TI": TalkerDes = "Turn rate indicator "
Case Is = "UP": TalkerDes = "Microprocessor controller "
Case Is = "U0": TalkerDes = "User configured Talker Identifier"
Case Is = "U1": TalkerDes = "User configured Talker Identifier"
Case Is = "U2": TalkerDes = "User configured Talker Identifier"
Case Is = "U3": TalkerDes = "User configured Talker Identifier"
Case Is = "U4": TalkerDes = "User configured Talker Identifier"
Case Is = "U5": TalkerDes = "User configured Talker Identifier"
Case Is = "U6": TalkerDes = "User configured Talker Identifier"
Case Is = "U7": TalkerDes = "User configured Talker Identifier"
Case Is = "U8": TalkerDes = "User configured Talker Identifier"
Case Is = "U9": TalkerDes = "User configured Talker Identifier"
Case Is = "VD": TalkerDes = "Velocity sensors: Doppler other/general "
Case Is = "VM": TalkerDes = "Velocity sensors: speed log water magnetic "
Case Is = "VW": TalkerDes = "Velocity sensors: speed log water mechanical "
Case Is = "VR": TalkerDes = "Voyage data recorder "
Case Is = "WD": TalkerDes = "Watertight door controller/monitoring system "
Case Is = "WL": TalkerDes = "Water level detection system "
Case Is = "YX": TalkerDes = "Transducer "
Case Is = "ZA": TalkerDes = "Timekeeper time/date: atomic clock "
Case Is = "ZC": TalkerDes = "Timekeeper time/date: chronometer "
Case Is = "ZQ": TalkerDes = "Timekeeper time/date: quartz "
Case Is = "ZV": TalkerDes = "Timekeeper time/date: radio update "
Case Is = "WI": TalkerDes = "Weather instrument"
    Case Else
        TalkerDes = "Awaiting defining"
    End Select

End Function
