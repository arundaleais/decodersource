VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form Detail 
   Caption         =   "Detail"
   ClientHeight    =   9660
   ClientLeft      =   168
   ClientTop       =   552
   ClientWidth     =   7812
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Detail.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9660
   ScaleWidth      =   7812
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid DetailDisplay 
      Height          =   9615
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7815
      _ExtentX        =   13780
      _ExtentY        =   16955
      _Version        =   393216
      Rows            =   50
      Cols            =   4
      FixedCols       =   0
      Redraw          =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   2
      ScrollBars      =   2
      AllowUserResizing=   1
      BorderStyle     =   0
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup Menu"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuCreateTag 
         Caption         =   "Create Tag"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuCopyToClipBoard 
         Caption         =   "Copy All to Clip Board as CSV"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSeparator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
         Enabled         =   0   'False
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
   End
End
Attribute VB_Name = "Detail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Des As String   '1st Col
Dim Val As String   '2nd Col
Dim Valdes As String    '3rd Col
Public UserFieldTagName As String

Private Sub Form_Load()
With DetailDisplay
    .FormatString = "<Source|<Description|<Value|<Value Description"
    .Cols = 4
    .ColWidth(0) = 0
#If jnasetup = True Then
    .ColWidth(0) = 5000
#End If
    .ColWidth(1) = 2300
    .ColWidth(2) = 2000
    .ColWidth(3) = 5000
    .Rows = 40
End With
Detail.Hide     'Always hide on initial load
End Sub

'Retvalcol is the Detail column for which we want the data returning
'if 0 the all columns are returned (I think CSV ?)
'if 4 then no output is generated
Sub SentenceAndPayloadDetail(Retvalcol As Long)
'Gets All NMEA Details from clsSentence
''Dim NmeaWords() As String
'Dim err As Boolean
'Dim LastNmeaWords() As String   'PayloadReassembler
'Dim LastNmeaSentence As String  'PayloadReassembler
'Dim Comments As String          'PayloadReassembler
Dim i As Long

'Dim LastRow As Long             'PayloadReassembler
'Dim Payload6Bits As Long    'used to calculate 8 bits
'Dim Payload8Bits As Long
'Dim Val As String
'Dim Des As String
'Dim Valdes As String
'Dim CreatedTime As String
Dim Simple As Boolean   'make true to remove first 2 cols for clarity
Dim ParameterCode As String
Dim CbWordNo As Long
Dim WordNames() As String 'Array of WordNames
Dim LastWordToOutput As Long
Dim WordNo As Long              'Current word
Dim LastWordBeforeCRC As Long
Dim GroupWords() As String
Dim SaveNmeaWords() As String
Dim WordsToAdd As Long
Dim DetailOK As Boolean

'set simple=true to not output all cols for
'construction of help/aisdecoder.html
    Simple = True
    
    If Retvalcol = 4 Then
        AllFieldsNo = 0
        ReDim AllFields(1)
    Else
'Show form if required when data received (UDP & Serial)
'Hide ONLY when option clicked or in QueryUnload
        If cmdNoWindow = False And Detail.Visible = False And NmeaRcv.Option1(4).Value = False Then
            Detail.Show
        End If

#If jnasetup = False Then
        Detail.DetailDisplay.Redraw = False
#End If
        Clear_DetailDisplay
    End If

'dont load created and sentence as it confuses the output
    If Simple = True And Retvalcol <> 4 Then
        Call SentenceOut(Retvalcol, "Creation Time Local", "created")
'Added V117 to give a unix time alternative
        Call SentenceOut(Retvalcol, "Creation Time Unix UTC", "createdunix")
    End If
    
'Output any offset to sentence (may be proprietrary time stamp)
    If clsSentence.NmeaPrefix <> "" Then
'V 144 Jan 2017 if $AITAG(Jason time stamp)has been found make CSV output compatible with V129
        If Left$(GroupSentence.NmeaSentence, 6) <> "$AITAG" Then
            Call SentenceOut(Retvalcol, "Nmea Prefix String", "nmeaprefix")
        End If
'If the prefix looks like a time stamp
'The Prefix (if date) is now copied into .NmeaRcvTime in DecodeSentence v136
'        If IsDate(clsSentence.NmeaPrefix) Then
'           Call SentenceOut(Retvalcol, "Received Time UTC", "receivedtime")
'        End If
    
    End If
        
'Output any comments
    If clsCb.Block <> "" Then
        If Simple = True And Retvalcol <> 4 Then
            Call CommentOut(Retvalcol, "IEC 61162 Comments", "commentblock")
        End If
'Output each word in CB if no errors
        If clsCb.errmsg = "" Then
            For CbWordNo = 0 To UBound(CbWords)
                i = InStr(1, CbWords(CbWordNo), ":")
                If i > 1 Then ParameterCode = Left$(CbWords(CbWordNo), i - 1)
                i = InStr(1, ParameterCode, "G")
                If i >= 1 Then ParameterCode = "G"
'xGy where x & y are integers
                Call CommentOut(Retvalcol, ParameterDes(ParameterCode), ParameterCode)
            
            Next CbWordNo
        End If
        
''        If Simple = True And Retvalcol <> 4 Then    'Suppress if csvall
            Call CommentOut(Retvalcol, "CRC Check", "cbcrc")
''        End If
    End If
 'End of Comments
 
    If clsSentence.NmeaSentence <> "" Then
 'Now Output NMEA sentence details
        If Simple = True And Retvalcol <> 4 Then
            Call SentenceOut(Retvalcol, "Nmea Sentence", "nmea")
        End If
                
        If clsSentence.NmeaRcvTime <> "" Then
            Call SentenceOut(Retvalcol, "Received Time UTC-Unix", "receivedtime")
        End If
            
'Check CRC is valid first
       If clsSentence.CRCerrmsg = "" Then
            
            If clsSentence.IecTalkerID <> "" Then
               Call SentenceOut(Retvalcol, "Talker", "talker")
            End If
        
        
            If clsSentence.IecFormat <> "" Then
            Call SentenceOut(Retvalcol, "Sentence", "iecformat")
            End If
        End If
                 
'Set up the WordNames up to the CRC WordNo
        Select Case clsSentence.IecFormat
        Case Is = "VDM", "VDO", "THAR"
            WordNames = NmeaAIVDMWordName
        Case Is = "TXT"
            WordNames = NmeaAITXTWordName
'Memmber (arg 3) is only used for the default user tag name
        Case Is = "GGA"
            WordNames = NmeaGPGGAWordName
        Case Is = "RMC"
            Call NmeaOut(Retvalcol, "Unix Time (calculated)", "RMC_unixtime")
            WordNames = NmeaGPRMCWordName
        Case Is = "ZDA"
            WordNames = NmeaGPZDAWordName
        Case Is = "ALR"
            WordNames = NmeaAIALRWordName
        Case Is = "GHP"
'+ 1 is calculated unix time
            Call NmeaOut(Retvalcol, "Unix Time (calculated)", "GHP_unixtime")
            WordNames = NmeaPGHPWordName
        Case Is = "TAG"
            WordNames = NmeaAITAGWordName
'For ABM & BBM there are too many problems trying to decode the Payload
'Compared with AIVDM so I do not try and decode the binary data
'The Word positions differ
'The Multipart ID differ (0-3) and 0 is valid if single part
        Case Is = "ABM"
            WordNames = NmeaAIABMWordName
        Case Is = "BBM"
            WordNames = NmeaAIBBMWordName
        Case Is = "BRM" 'Base Station Message
            WordNames = NmeaAIBRMWordName
        Case Is = "THAJ" 'Base Station Message
            WordNames = NmeaPTHAJWordName
        Case Else

'normally we dont output nmeasentence if AllCsv
'An Undecoded NMEA sentence is output as all one starting
'V124 changed to output undecoded NMEA sentence as CSV all fields rather than the whole sentence
'as literal (within ",,,,")
'                    If Simple = True And Retvalcol = 4 Then
'                        Call SentenceOut(Retvalcol, "Nmea Sentence", "nmea")
'                    Else
            Call NmeaOut(Retvalcol, "Sentence Type", "notdecoded")
'These must be called by NmeaOut not SentenceOut otherwise
'When the FieldKey is generated by FieldKeyFromSentence
'If the user picks a Field, The key will apply to All sentences and
'not just this Nmea sentence $xxxxx
                        
'Fake spec for format nmeaword, with names Word 0,Word1 ...etc
'This enables a common routine to be used to parse the NmeaSentence
            ReDim WordNames(clsSentence.NmeaCrcWord)
            For WordNo = 0 To UBound(WordNames)
                WordNames(WordNo) = "Word " & WordNo
            Next WordNo
            clsSentence.IecFormat = "nmeaword"
        End Select
 'Detail.DetailDisplay.Redraw = True
'Output to limit of defined field in Nmea spec
'No CRC, spoof position of crc check
        If clsSentence.NmeaCrc <> "" Then
            LastWordBeforeCRC = clsSentence.NmeaCrcWord
        Else
            LastWordBeforeCRC = UBound(WordNames)
        End If
                
'Output up to CRC
        For WordNo = 0 To UBound(WordNames)
            If WordNo <= LastWordBeforeCRC Then
                Call NmeaOut(Retvalcol, WordNames(WordNo), clsSentence.IecFormat, WordNo)
            Else
'Less words in sentence than defined in spec
                Call NmeaOut(Retvalcol, WordNames(WordNo), "missing")
            End If
        Next WordNo
                
'More Words in sentence before CRC than in spec
        For WordNo = WordNo To LastWordBeforeCRC
            Call NmeaOut(Retvalcol, "Word " & WordNo, "extra", WordNo)
        Next WordNo
'Output CRC (if any) '$AITAG does not have CRC
        If clsSentence.NmeaCrc <> "" Then
            Call SentenceOut(Retvalcol, "CRC check", "nmeacrc")
        End If

'Output from CRC to end of sentence, missing time stamp
        For WordNo = LastWordBeforeCRC + 1 To UBound(NmeaWords)
            If Not IsDate(NmeaWords(WordNo)) Then
'If no CRC dont assume a date
                If clsSentence.NmeaCrc <> "" Then
'Dont output the time stamp as it's already been output at the beginning
                    Call NmeaOut(Retvalcol, "Word " & WordNo, "nmeaword", WordNo)
                End If
            End If
        Next WordNo

'If AIS sentence
        If clsSentence.IsAisSentence = True Then
            If clsSentence.AisMsgPartsComplete Then 'PayloadReassembler payload is complete
'Output Payload
                Call PayloadDetail(Retvalcol)   '(NmeaWords(5)) 'extract Payload details
'check if weve actually got enough bits in the payload
                Call NmeaAisOut(Retvalcol, "Payload Bit Check", "payloadbits")
'Moved from PayloadDetail Jan15
'V 144 Jan 2017 if $AITAG(Jason time stamp)has been found make CSV output compatible with V129
'v129 checks MMSI not Rcvtime
                If Left$(GroupSentence.NmeaSentence, 6) <> "$AITAG" Then
                    If MyShip.RcvTime <> "" Then
                        MyShipDetail (Retvalcol)
                    End If
                Else
                    If MyShip.Mmsi <> "" Then
                        MyShipDetail (Retvalcol)
                    End If
                End If
                
                If GroupSentence.NmeaSentence <> "" Then
'MsgBox clsSentence.NmeaSentence & vbCrLf & GroupSentence.NmeaSentence
                    ReDim SaveNmeaWords(UBound(NmeaWords))
                    SaveNmeaWords = NmeaWords       'Must have been redimensioned
                    NmeaWords = Split(ConvEscChrs(GroupSentence.NmeaSentence), ",")  '350k/min
                    Select Case NmeaWords(0)
                    Case Is = "$AITAG"
                        WordNames = NmeaAITAGWordName
'The encapsulating sentence has a time stamp added
'$AITAG 1402561127,2272,18/06/2014 13:11:03
'Tag has all spaces replaced with comma
'$AITAG,1402561127,2272,18/06/2014,13:11:03
'ONLY output up to the no of fields defined for TAG
                        For WordNo = 0 To UBound(WordNames)
                            Call NmeaOut(Retvalcol, WordNames(WordNo), "TAG", WordNo)
                        Next WordNo
                    End Select
                    ReDim NmeaWords(UBound(SaveNmeaWords))
                    NmeaWords = SaveNmeaWords
                End If
            Else    'Incomplete re-assembled payload
                Call NmeaAisOut(Retvalcol, "AIS Payload", "aispayload")
'Stop
            End If
        End If  'end AIS sentence
                
'Detail.DetailDisplay.Redraw = True
'Stop
    Else    'Blank sentence
        Call SentenceOut(Retvalcol, "Input Sentence", "fullsentence")
    End If  'NMEA sentence not blank
    
    If Retvalcol <> 4 Then
        Fill_DetailDisplay
        Detail.DetailDisplay.Redraw = True
'stops display freezing
        Detail.DetailDisplay.Refresh
    End If
'Detail.Show
End Sub

'Called by SentenceAndPayloadDetail, SetTagValues
'This Outputs any data that is in NmeaWords()
'It is NMEA sentence specific
'Member is not used to Output Fields but IS used to Display the Default
'User Tag (in form FieldInput)
'Member + WordNo is used to format Valdes
Function NmeaOut(Retvalcol As Long, _
ByVal Des As String, _
Member As String, _
Optional WordNo As Long, _
Optional reqChrs As Long) As String   'returned value is (1,des,2=val,3=valdes)
'Des is byVal as it must not be changed in calling program
Dim Valdes As String
Dim Val As String
'Dim Valdes As String
Dim Bold As Boolean
Dim i As Long
Dim kb As String
Dim arry() As String
Dim MsgTimeSys As SYSTEMTIME
Dim Payload6Bits As Long    'used to calculate 8 bits
Dim Payload8Bits As Long
Dim Payload8Words As Long
Dim ByteBoundary As Long

'Check sentence not too short
    If WordNo <= UBound(NmeaWords) Then
        Val = NmeaWords(WordNo)
        If Val = "" Then
            Select Case Member
            Case Is = "extra"    'Any field not specifically formatted and with no value
                Val = NmeaWords(WordNo)
                Valdes = "{not in specification}"
            Case Else
                Valdes = "(blank)"
            End Select
'For NmeaOut Member will be A Specifically decoded sentence
'or "nmeaword"
        Else
            Select Case Member
            Case Is = "notdecoded"  'not aivdm or aivdo (or any recognised nmea sentence)
                Valdes = "{not decoded}"
            Case Is = "missing"  'Field in Specification but Missing (not null ie ,,) in sentence
                Val = ""    'Val will be for the wrong word for this field
                Valdes = "{Missing}"
            Case Is = "extra"    'Any field not specifically formatted
                Val = NmeaWords(WordNo)
                Valdes = "{not in specification}"
            Case Is = "nmeaword"    'Any field not specifically formatted
                Valdes = "{not decoded}"
            Case Is = "VDM", "VDO", "THAR"
                Select Case WordNo
                Case Is = 0
'Debug.Print clsSentence.FullSentence
                    Valdes = TalkerDes(clsSentence.IecTalkerID)
'Check if word(5) exists and is valid - potential subscript error
'    : FullSentence : "!AIVDM,1mJfN18c?d0D00,2*62,03/07/2014 18:49:03" : String : MainControl.ProcessSentence
'                    If UBound(NmeaWords) >= 6 Then
'                        If IsNumeric(NmeaWords(6)) Then
'                            Payload6Bits = Len(NmeaWords(5)) * 6
'                            Payload8Bits = Int(Payload6Bits / 8) * 8
'                            Valdes = Valdes & ", " & Payload8Bits & " bits (" & Len(NmeaWords(5)) & " 6-bit words)"
'                            Payload8Bits = Len(NmeaWords(5)) * 6 - NmeaWords(6)
'                            Valdes = Valdes & ", " & Payload8Bits & " bits (" & Payload8Bits / 8 & ") 8-bit words)"
'                        End If
'                    End If
                Case Is = 1, 2, 3
                    Valdes = AisWordCheck(WordNo)
                Case Is = 4     'Radio Channel
'               Case is = 5
'                    ByteBoundary = Len(NmeaWords(5)) * 6 - NullToZero(NmeaWords(6))
'                    Valdes = ByteBoundary & " bits (" & Len(NmeaWords(5)) & " 6-bit words)"
                Case Is = clsSentence.NmeaCrcWord - 1   'Part Payload
'From v139 - if Encapsulated data is not the 6th word then the sentence is not a corectly
'formatted NMEA AIS sentence (field nos are incorrect)
                    ByteBoundary = Len(NmeaWords(clsSentence.NmeaCrcWord - 1)) * 6 - NullToZero(NmeaWords(clsSentence.NmeaCrcWord))
                    Valdes = ByteBoundary & " bits (" & Len(NmeaWords(clsSentence.NmeaCrcWord - 1)) & " 6-bit words)"
                    If clsSentence.IsAisSentence = False Then Valdes = Valdes & " (not AIS)"
#If jnasetup = True Then
Des = Des & "(" & Len(Val) & " chrs)"
#End If
'                    Valdes = Len(NmeaWords(5)) * 6 - NullToZero(NmeaWords(6))
'            Payload8Bits = Int(Payload6Bits / 8) * 8
'Subscript error if Payload bytes not set up
'                    Payload8Bits = (PayloadByteArraySize + 1) * 8
'                    Valdes = Payload8Bits & " bits (" & Payload8Bits / 8 & " 8-bit words)"
'                    Valdes = clsSentence.AisPayloadBits & " bits (" & clsSentence.AisPayloadBits / 8 & " 8-bit words)"
'Des = Des & "-" & Len(Val)
'                    If clsSentence.AisMsgPartsComplete = False Then Valdes = Valdes & ", incomplete"
                Case Is = clsSentence.NmeaCrcWord - 1   'Fill bits should be word 6
'but if extra word before crc could be 7 (Hung)
'multi-part and NOT the last part
'Fill bits are only used to check complete payload
                    If NullToZero(Val) > 0 Then
                        Valdes = Valdes & "Bits are " & pbits(clsSentence.AisPayloadBits + 1, clsSentence.AisPayloadFillBits)

                    End If
                    If NmeaWords(1) <> NmeaWords(2) Then
                        If NmeaWords(6) <> "0" Then
                            Valdes = Valdes & "NMEA fill bits should be 0"
'                err = True
                        End If
                    Else    'Last part
'We dont know if the fill bits are wrong or the payload characters
'The fill bits must be from 0 to 5, otherwise an additional payload character
                        ByteBoundary = Len(NmeaWords(5)) * 6 - NullToZero(NmeaWords(6))
                        If ByteBoundary Mod 8 <> 0 Then
                            Valdes = Valdes & "Not filled to byte boundary"
                        End If
'If clsSentence.AisPayloadFillBits <> ChrnoToFillBits(Len(clsSentence.AisPayload)) Then
'    Valdes = "Payload fill bits incorrect (" & ChrnoToFillBits(Len(clsSentence.AisPayload)) & ")"
'Stop
'End If
                    End If
                Case Is > clsSentence.NmeaCrcWord   'Words 7 onwards
                    If IsNumeric(Val) Then  'Assume unix time
                        Valdes = Format$(UnixTimeToDate(Val), DateTimeOutputFormat)
                    Else
                        Valdes = "Local Extension " & WordNo
                    End If
                End Select
''                For WordNo = clsSentence.NmeaCrcWord + 1 To UBound(NmeaWords)
''                    Call NmeaAisOut(Retvalcol, "Word " & WordNo, "aisword", WordNo)
''                Next WordNo
'Display detail of incomplete sentences, because Dac316-1,2 (Canadian St Lawrence
'Seaway messages do not appear to transmit all parts !
            Case Is = "TAG"
                Select Case WordNo
                Case Is = 1
                    If IsNumeric(Val) Then  'Assume unix time
                        Valdes = Format$(UnixTimeToDate(Val), DateTimeOutputFormat)
                    Else
                        Valdes = "Local Extension " & WordNo
                    End If
                End Select
            Case Is = "GGA"
                Select Case WordNo
                Case Is = 1   'Time hhmmss.ss
                    If IsNumeric(Val) And Val <> "" Then 'May be blank
                    Valdes = Format$(Int(Val), "00:00:00") & " UTC"
                    Else
                        Valdes = "Invalid Time"
                    End If
                Case Is = 2   'Lat llll.ll First 4 digits are fixed(2 for deg and 2 for mins) .ll is optional
                    If IsNumeric(Val) And Val <> "" Then 'May be blank
                        Valdes = Left$(Val, 2) & " " & Mid$(Val, 3)
                    Else
                        Valdes = "Invalid"
                    End If
                Case Is = 3   'N or S
                Case Is = 4   'Lon 00024.8994 First 5 digits are fixed(3 for deg and 2 for mins) .yy is optional
                    If IsNumeric(Val) And Val <> "" Then 'May be blank
                        Valdes = Left$(Val, 3) & " " & Mid$(Val, 4)
                    Else
                        Valdes = "Invalid"
                    End If
                Case Is = 5   'E or W
                End Select
            Case Is = "RMC"
                Select Case WordNo
                Case Is = 1   'Time hhmmss.ss first 6 digits are fixed .ss is optional
                    If IsNumeric(Val) And Val <> "" Then 'May be blank
                    Valdes = Format$(Int(Val), "00:00:00") & " UTC"
                    Else
                        Valdes = "Invalid Time"
                    End If
                Case Is = 2     'Status A=Data Valid or V=Nav Rcvr Warning
                    Select Case Val
                    Case Is = "A": Valdes = "Data Valid"
                    Case Is = "V": Valdes = "Navigational recever warning"
                    Case Else: Valdes = "Invalid"
                    End Select
                Case Is = 3   'Lat llll.ll First 4 digits are fixed(2 for deg and 2 for mins) .ll is optional
                    If IsNumeric(Val) And Val <> "" Then 'May be blank
                        Valdes = Left$(Val, 2) & " " & Mid$(Val, 3)
                    Else
                        Valdes = "Invalid"
                    End If
                Case Is = 4   'N or S
                Case Is = 5   'Lon yyyyy.yy First 5 digits are fixed(3 for deg and 2 for mins) .yy is optional
                    If IsNumeric(Val) And Val <> "" Then 'May be blank
                        Valdes = Left$(Val, 3) & " " & Mid$(Val, 4)
                    Else
                        Valdes = "Invalid"
                    End If
                Case Is = 6   'E or W
                Case Is = 7     'SOG x.x    'Variable length integer part, Optional Leading/trailing zeros
                                        'Decimal fraction part optional
                Case Is = 8     'COG x.x
                Case Is = 9     'Date ddmmyy
                    Valdes = CDate(Left$(Val, 2) & "/" & Mid$(Val, 3, 2) & "/" & Right$(Val, 2))
                Case Is = 10    'Mag Variation x.x
                Case Is = 11    'E or W
                Case Is = 12    'Mode Alpha String
                    Select Case Val
                    Case Is = "A": Valdes = "Autonomous"
                    Case Is = "D": Valdes = "Differential"
                    Case Is = "E": Valdes = "Estimated (dead reckoning)"
                    Case Is = "F": Valdes = "Float RTK"
                    Case Is = "M": Valdes = "Minual input"
                    Case Is = "N": Valdes = "No fix"
                    Case Is = "P": Valdes = "Precise"
                    Case Is = "R": Valdes = "Real time kinematic"
                    Case Is = "S": Valdes = "Simulator"
                    Case Else: Valdes = "Invalid"
                    End Select
                Case Is = 13    'Nav Status Alpha string
                    Select Case Val
                    Case Is = "S": Valdes = "Safe"
                    Case Is = "C": Valdes = "Caution"
                    Case Is = "U": Valdes = "Unsafe"
                    Case Is = "V": Valdes = "Navigational status is not valid"
                    Case Else: Valdes = "Invalid"
                    End Select
                End Select
            Case Is = "GHP"       'key is sentencename $PGHP GH Internal type 1
                Select Case WordNo
                Case Is <= 8    '0-8    'Just output Des & Val
                Case Is = 9   'Country of origin (MID)
                    If Val <> 0 Then
                        Valdes = DacName(Val)
                    Else
                        Valdes = "Not defined"
                    End If
                Case Is = 10  'MMSI of region
                    Valdes = "MMSI !"
                Case Is = 11  'MMSI of transponder
                    Valdes = "MMSI !"
                Case Is = 12  'Buffered=0, Online=1
                    Select Case Val
                    Case Is = "0"
                        Valdes = "Buffered"
                    Case Is = "1"
                        Valdes = "OnLine"
                    Case Else
                        Valdes = "Invalid"
                    End Select
                Case Is = 13  'Nmea Sentence CheckSum
                End Select
'because this is outside the permitted range of NmeaWords for this
'sentence it must be define separately
            Case Is = "GHP_unixtime"    'unixtime
                MsgTimeSys.wYear = NmeaWords(2)
                MsgTimeSys.wMonth = NmeaWords(3)
                MsgTimeSys.wDay = NmeaWords(4)
                MsgTimeSys.wHour = NmeaWords(5)
                MsgTimeSys.wMinute = NmeaWords(6)
                MsgTimeSys.wSecond = NmeaWords(7)
                MsgTimeSys.wMilliseconds = NmeaWords(8)
                Val = SysTimeToUnix(MsgTimeSys)
                Valdes = UnixTimeToDate(Val)
            Case Is = "RMC_unixtime"    'unixtime
                MsgTimeSys.wYear = 2000 + Right(NmeaWords(9), 2)
                MsgTimeSys.wMonth = Mid(NmeaWords(9), 3, 2)
                MsgTimeSys.wDay = Left(NmeaWords(9), 2)
                MsgTimeSys.wHour = Left(NmeaWords(1), 2)
                MsgTimeSys.wMinute = Mid(NmeaWords(1), 3, 2)
                MsgTimeSys.wSecond = Mid(NmeaWords(1), 5, 2)
                MsgTimeSys.wMilliseconds = 0
                Val = SysTimeToUnix(MsgTimeSys)
                Valdes = UnixTimeToDate(Val)
            Case Is = "TXT"
'With Freq Output from SLR200 the message is "Freq,2087,2088" so the extra words parsed out
'require adding to the message
                If WordNo = UBound(NmeaAITXTWordName) Then
                    For i = WordNo + 1 To UBound(NmeaWords)
                        Val = Val & "," & NmeaWords(i)
                    Next i
                End If
                Select Case WordNo
                Case Is = 3
                    If Val <> "" Then
                        Valdes = IecTXTIdentifierName(Val)
                    End If
                End Select
            Case Is = "ALR"
'Stop
                Select Case WordNo
                Case Is = 2
                    If Val <> "" Then
                        Valdes = IecALRIdentifierName(Val)
                    End If
                Case Is = 3
                    Select Case Val
                    Case Is = "A"
                        Valdes = "threshold exceeded"
                    Case Is = "V"
                        Valdes = "not exceeded"
                    Case Else
                        Valdes = "undefined"
                    End Select
                Case Is = 4
                    Select Case Val
                    Case Is = "A"
                        Valdes = "acknowledged"
                    Case Is = "V"
                        Valdes = "unacknowledged"
                    Case Else
                        Valdes = "undefined"
                    End Select
                End Select
            Case Is = "ABM"
                Select Case WordNo
                Case Is = 7     'Payload
                    ByteBoundary = Len(NmeaWords(WordNo)) * 6 - NullToZero(NmeaWords(WordNo + 1))
                    Valdes = ByteBoundary & " bits (" & Len(NmeaWords(WordNo)) & " 6-bit words)"
#If jnasetup = True Then
Des = Des & "(" & Len(Val) & " chrs)"
#End If
                    clsSentence.AisPayloadBits = ByteBoundary
                End Select
            Case Is = "BBM"
                Select Case WordNo
                Case Is = 6     'Payload
                    ByteBoundary = Len(NmeaWords(WordNo)) * 6 - NullToZero(NmeaWords(WordNo + 1))
                    Valdes = ByteBoundary & " bits (" & Len(NmeaWords(WordNo)) & " 6-bit words)"
#If jnasetup = True Then
Des = Des & "(" & Len(Val) & " chrs)"
#End If
                    clsSentence.AisPayloadBits = ByteBoundary
                End Select
            Case Is = "BRM"
                Select Case WordNo
                Case Is = 3
                    Valdes = Val & " dBm"
                End Select
            Case Is = "THAJ"
                Select Case WordNo
                Case Is = 1
                    Select Case Val
                    Case Is = "A", "B"
                    Case Else
                        Valdes = "invalid"
                    End Select
                Case Is = 2
                    Select Case Val
                    Case Is <= 0, Is > 2249
                        Valdes = "invalid"
                    Case Else
                    End Select
                Case Is = 3
                    Select Case Val
                    Case Is = "N"
                        Valdes = "Negative"
                    Case Is = "P"
                        Valdes = "Positive"
                    Case Else
                        Valdes = "invalid"
                    End Select
                Case Is = 4
                    If IsNumeric(Val) Then
                        If CSng(Val) < 0 Or CSng(Val) > 13333.3 Then       'Out of range
                            Valdes = "invalid"
                        Else
                            Valdes = "Microseconds"
                        End If
                    Else    'not numeric
                        Valdes = "invalid"
                    End If
                End Select
            Case Else            'not a decoded NMEA sentence
                Valdes = "[NmeaOut member (" & Member & ") not found]"
            End Select
        End If
    Else        'More words in sentence description
'Word no requested > words in this sentence
'        Des = "Word " & WordNo
        Valdes = "Too few words (" & UBound(NmeaWords) & ") in sentence"
    End If
'required by detaillineout to construct source etc
    clsField.CallingRoutine = "NmeaOut"
    clsField.Des = Des
    clsField.Member = Member
    clsField.From = WordNo
    clsField.reqbits = reqChrs  'not used
    Select Case Retvalcol
        Case Is = 0
            Call DetailLineOut(Des, Val, Valdes, Bold) 'main output
        Case Is = 1
            NmeaOut = Des
        Case Is = 2
            NmeaOut = Val
        Case Is = 3
            NmeaOut = Valdes
        Case Is = 4
            If AllFieldsNo > UBound(AllFields) Then ReDim Preserve AllFields(AllFieldsNo)
            AllFields(AllFieldsNo) = Val
            AllFieldsNo = AllFieldsNo + 1
    End Select
End Function

'This processes the Current Time + Any required Values in clsSentence
Function SentenceOut(Retvalcol As Long, _
Des As String, _
Member As String, _
Optional WordNo As Long, _
Optional reqChrs As Long) As String   'returned value is (1,des,2=val,3=valdes)

Dim Valdes As String
Dim Val As String
Dim Bold As Boolean
Dim i As Long

    Select Case Member
    Case Is = "nmeaword"
        Val = NmeaWords(WordNo)
    Case Is = "created"
        Val = Format$(Now(), DateTimeOutputFormat)
    Case Is = "createdunix"
'Changed to Unix UTC V117
        Val = NowUnix()
        Valdes = Format$(UnixTimeToDate(Val), DateTimeOutputFormat)
    Case Is = "nmea"
        Val = clsSentence.NmeaSentence
        If clsSentence.CRCerrmsg <> "" Then
            Valdes = clsSentence.CRCerrmsg
        End If
        i = InStr(1, Val, "*")
        If i > 78 Then '*hh<cr><lf> 82 max inc crlf
            Valdes = Valdes & "Nmea Sentence too long (" & i + 2 & ")"
        End If
        If clsSentence.IecEncapsulationComments <> "" Then
            Valdes = Valdes & clsSentence.IecEncapsulationComments
        End If
        Bold = True
    Case Is = "talker"    'not currently used
        Val = clsSentence.IecTalkerID
        Valdes = TalkerDes(Val)
    Case Is = "iecformat"    'not currently used
        Val = clsSentence.IecFormat
        Valdes = IecFormatDes(Val)
    Case Is = "nmeacrc"
        Val = clsSentence.NmeaCrc
        If Val = "hh" Then Valdes = "Test Data (CRC invalid)"
    Case Is = "receivedtime"
        Val = Format$(clsSentence.NmeaRcvTime, DateTimeOutputFormat)
        Valdes = DateToUnixTime(clsSentence.NmeaRcvTime)
    Case Is = "nmeaprefix"
        Val = clsSentence.NmeaPrefix
    Case Is = "notdecoded"  'not aivdm or aivdo (or any recognised nmea sentence)
        Val = clsSentence.NmeaSentenceType
        Valdes = "{not decoded}"
    Case Is = "fullsentence"  'no NMEA delimiter
        Val = clsSentence.FullSentence
        Valdes = "{No NMEA * delimeter)"
    Case Else
        Valdes = "[SentenceOut member (" & Member & ") not found]"
    End Select

'required by detaillineout to construct source etc
clsField.CallingRoutine = "SentenceOut"
clsField.Des = Des
clsField.Member = Member
clsField.From = WordNo
clsField.reqbits = reqChrs  'not used
'Call DetailLineOut(Des, Val, ValDes, Bold) 'main output
Select Case Retvalcol
    Case Is = 0
        Call DetailLineOut(Des, Val, Valdes, Bold) 'main output
    Case Is = 1
        SentenceOut = Des
    Case Is = 2
        SentenceOut = Val
    Case Is = 3
        SentenceOut = Valdes
    Case Is = 4
        If AllFieldsNo > UBound(AllFields) Then ReDim Preserve AllFields(AllFieldsNo)
        AllFields(AllFieldsNo) = Val
        AllFieldsNo = AllFieldsNo + 1
End Select
End Function


'nDetail is the DetailOut node in the output tree - not now used
'the payload must exist in PayloadBytes, which will have been created by PayloadReassembler
'Cycles through the whole of PayloadBytes, outputing all fields in order
'by calling DetailOut to extract the individual fields.
'PayloadBytes will have been previously created by calling PayloadReassembler
'for the current sentence

Sub PayloadDetail(Retvalcol As Long) 'Payload As String)
'Static AisMsgType As String 'keep last type when called
'Static Mmsi As String
Dim wlong As Long
Dim RotAis As Long
Dim Val As String
Dim Valdes As String
Dim Minus As Boolean
Dim kb As String
Dim NextBit As Long 'would be the next bit (if any more to extract)
                    'used to calculate no of bits required to stuff
                    'to 8 bit boundary
Dim Last As Integer 'Last bit used
Dim RadioMode As Long
Dim PayloadExcessBits As Long   'Actual Payload bits - No of Bits to Output (per spec)
Dim Offset As Long
Dim AiStart As Long
Dim AiBits As Long
Dim BinaryDataStart As Long
Dim BinaryBits As Long
Dim SpareBitsToBoundary As Long 'Spares added to end of Message Spec to fill to 8 bit boundary
                                'Used on some variable length Messages
                                'Msg 20

'Output complete cumulative payload (parts 1 + n...)
'V 144 Jan 2017 if $AITAG(Jason time stamp)has been found make CSV output compatible with V129
    If Left$(GroupSentence.NmeaSentence, 6) <> "$AITAG" Then
        Call NmeaAisOut(Retvalcol, "AIS Payload", "aispayload")
    End If
    Call NmeaAisOut(Retvalcol, "Vessel Name", "vesselname")  'as in clssentence
'payload bytes is public created by PayloadReassembler

    
'    clssentence.aispayloadbits = (PayloadByteArraySize + 1) * 8
'    kb = (PayloadByteArraySize + 1) * 8
'   kb = Len(clsSentence.FullSentence)
'If clssentence.aispayloadbits <> (PayloadByteArraySize + 1) * 8 Then Stop

'jnadebug (All payload bits)
'Call DetailOut(RetValCol, "Bits", "bits", 1, clssentence.aispayloadbits)
'clssentence.aispayloadbits = ChrnoToBits(Len(Payload)) 14/2
    
'    Call DetailOut(Retvalcol, "AIS Message Type", "aismsgtype", 1, 6)
    Call NmeaAisOut(Retvalcol, "AIS Message Type", "aismsgtype")
'check to see if weve got it right
'If AisMsgType <> clsSentence.AisMsgType Then Stop 'debug check only
    
'    Call DetailOut(Retvalcol, "Repeat Indicator", "repeat", 7, 2)
    Call NmeaAisOut(Retvalcol, "Repeat Indicator", "repeat")
'    Call DetailOut(Retvalcol, "MMSI", "mmsi", 9, 30)
    Call NmeaAisOut(Retvalcol, "MMSI", "mmsi")
'    Call DetailOut(retvalcol,"MID", "mid", 9, 30)
    Call NmeaAisOut(Retvalcol, "MID", "mid")
'    Mmsi = clsSentence.AisMsgFromMmsi
'    Call UpdateShip(Mmsi, "created", CreatedTime)
    Select Case clsSentence.AisMsgType
      Case Is = "1", "2", "3"
      Call DetailOut(Retvalcol, "Navigation Status", "status", 39, 4)
      Call DetailOut(Retvalcol, "Rate of Turn (ROT)", "turn", 43, 8)
      Call DetailOut(Retvalcol, "Speed Over Ground (SOG)", "speed", 51, 10, 1)
      Call DetailOut(Retvalcol, "Position Accuracy", "accuracy", 61, 1)
      Call DetailOut(Retvalcol, "Longitude", "lon", 62, 28)
      Call DetailOut(Retvalcol, "Latitude", "lat", 90, 27)
      Call DetailOut(Retvalcol, "Course Over Ground (COG)", "course", 117, 12, 1)
      Call DetailOut(Retvalcol, "True Heading (HDG)", "heading", 129, 9, 0)
      Call DetailOut(Retvalcol, "Time Stamp", "second", 138, 6)
      Call DetailOut(Retvalcol, "Manoeuvre Indicator", "manoeuvre", 144, 2)
      Call DetailOut(Retvalcol, "Spare", "spare", 146, 3)
      Call DetailOut(Retvalcol, "RAIM Flag", "raim", 149, 1)
        Call CommDetail(Retvalcol)
        NextBit = 169
  Case Is = "4", "11"
      Call DetailOut(Retvalcol, "Year", "year", 39, 14)
      Call DetailOut(Retvalcol, "Month", "month", 53, 4)
      Call DetailOut(Retvalcol, "Day", "day", 57, 5)
      Call DetailOut(Retvalcol, "Hour", "hour", 62, 5)
      Call DetailOut(Retvalcol, "Minute", "minute", 67, 6)
      Call DetailOut(Retvalcol, "Second", "second", 73, 6)
      Call DetailOut(Retvalcol, "Fix quality", "accuracy", 79, 1)
      Call DetailOut(Retvalcol, "Longitude", "lon", 80, 28)
      Call DetailOut(Retvalcol, "Latitude", "lat", 108, 27)
      Call DetailOut(Retvalcol, "Type of EPFD", "epfd", 135, 4)
      Call DetailOut(Retvalcol, "Transmission Control", "txcontrol", 139, 1)
      Call DetailOut(Retvalcol, "Spare", "spare", 140, 9)
      Call DetailOut(Retvalcol, "RAIM flag", "raim", 149, 1)
        Call CommDetail(Retvalcol)
        NextBit = 169
  Case Is = "5"
      Call DetailOut(Retvalcol, "AIS Version", "ais_version", 39, 2)
      Call DetailOut(Retvalcol, "IMO Number", "imo", 41, 30)
      Call DetailOut(Retvalcol, "Call Sign", "callsign", 71, 42)
      Call DetailOut(Retvalcol, "Vessel Name", "vesselname", 113, 120)
      Call DetailOut(Retvalcol, "Ship Type", "ship_type", 233, 8)
      Call DetailOut(Retvalcol, "Dimension to Bow", "to_bow", 241, 9)
      Call DetailOut(Retvalcol, "Dimension to Stern", "to_stern", 250, 9)
      Call DetailOut(Retvalcol, "Length", "clength", 241, 9)
      Call DetailOut(Retvalcol, "Dimension to Port", "to_port", 259, 6)
      Call DetailOut(Retvalcol, "Dimension to Starboard", "to_starboard", 265, 6)
      Call DetailOut(Retvalcol, "Beam", "cbeam", 259, 6)
      Call DetailOut(Retvalcol, "Position Type Fix", "epfd", 271, 4)
      Call DetailOut(Retvalcol, "ETA month", "month", 275, 4)
      Call DetailOut(Retvalcol, "ETA day", "day", 279, 5)
      Call DetailOut(Retvalcol, "ETA hour", "hour", 284, 5)
      Call DetailOut(Retvalcol, "ETA minute", "minute", 289, 6)
      Call DetailOut(Retvalcol, "Draught", "draught", 295, 8, 1)
      Call DetailOut(Retvalcol, "Destination", "destination", 303, 120)
      Call DetailOut(Retvalcol, "DTE", "dte", 423, 1)
      Call DetailOut(Retvalcol, "Spare", "spare", 424, 1)
          NextBit = 425
Case Is = "6"
      Call DetailOut(Retvalcol, "Sequence Number", "seqno", 39, 2)
      Call DetailOut(Retvalcol, "Destination MMSI", "mmsi", 41, 30, , "V")
      Call DetailOut(Retvalcol, "Retransmit Flag", "retransmit", 71, 1)
      Call DetailOut(Retvalcol, "Spare", "spare", 72, 1)
'      Call DetailOut(retvalcol,"Application ID", "app_id", 73, 16)
'Always structured so use aidata
      Call DetailOut(Retvalcol, "Binary Data", "aidata", 73, clsSentence.AisPayloadBits - 73 + 1)
    Case Is = "7", "13"
        Call DetailOut(Retvalcol, "Spare", "spare", 39, 2)
        Call DetailOut(Retvalcol, "Destination MMSI ID1", "mmsi", 41, 30, , "V")
        Call DetailOut(Retvalcol, "Sequence no for ID1", "seqno", 71, 2)
        NextBit = 73
        If clsSentence.AisPayloadBits > 72 Then
'        If pLong(73, 30) <> 0 Then
            Call DetailOut(Retvalcol, "Destination MMSI ID2", "mmsi", 73, 30, , "V")
            Call DetailOut(Retvalcol, "Sequence no for ID2", "seqno", 103, 2)
            NextBit = 105
        End If
        If clsSentence.AisPayloadBits > 104 Then
'        If pLong(105, 30) <> 0 Then
            Call DetailOut(Retvalcol, "Destination MMSI ID3", "mmsi", 105, 30, , "V")
            Call DetailOut(Retvalcol, "Sequence no for ID3", "seqno", 135, 2)
            NextBit = 137
        End If
'        If pLong(137, 30) <> 0 Then
        If clsSentence.AisPayloadBits > 136 Then
            Call DetailOut(Retvalcol, "Destination MMSI ID4", "mmsi", 137, 30, , "V")
            Call DetailOut(Retvalcol, "Sequence no for ID4", "seqno", 167, 2)
            NextBit = 169
        End If
  Case Is = "8"
      Call DetailOut(Retvalcol, "Spare", "spare", 39, 2)
'      Call DetailOut(retvalcol,"Application ID", "app_id", 41, 16)
      Call DetailOut(Retvalcol, "Binary Data", "aidata", 41, clsSentence.AisPayloadBits - 41 + 1)
' If clssentence.aispayloadbits - 41 <= 0 Then Stop
 
  Case Is = "9"
      Call DetailOut(Retvalcol, "Altitude", "alt", 39, 12)
      Call DetailOut(Retvalcol, "SOG", "speed", 51, 10, 0)
      Call DetailOut(Retvalcol, "Position Accuracy", "accuracy", 61, 1)
      Call DetailOut(Retvalcol, "Longitude", "lon", 62, 28)
      Call DetailOut(Retvalcol, "Latitude", "lat", 90, 27)
      Call DetailOut(Retvalcol, "Course Over Ground", "course", 117, 12, 1)
      Call DetailOut(Retvalcol, "Time Stamp", "second", 129, 6)
      Call DetailOut(Retvalcol, "Altitude Sensor", "altitudesensor", 135, 1)
      Call DetailOut(Retvalcol, "Reserved for future use", "reserved", 136, 7)
      Call DetailOut(Retvalcol, "DTE", "dte", 143, 1)
      Call DetailOut(Retvalcol, "Spare", "spare", 144, 3)
      Call DetailOut(Retvalcol, "Assigned", "assigned", 147, 1)
      Call DetailOut(Retvalcol, "RAIM flag", "raim", 148, 1)
'      Call DetailOut(RetValCol, "Communication mode", "radiomode", 149, 1)
        Call CommDetail(Retvalcol)
        NextBit = 169
  Case Is = "10"
      Call DetailOut(Retvalcol, "Spare", "spare", 39, 2)
      Call DetailOut(Retvalcol, "Destination MMSI", "mmsi", 41, 30, , "V")
      Call DetailOut(Retvalcol, "Spare", "spare", 71, 2)
        NextBit = 73
'  Case Is = "11"   see 4
  Case Is = "12"        'Addressed Safety Related message
      Call DetailOut(Retvalcol, "Sequence Number", "seqno", 39, 2)
      Call DetailOut(Retvalcol, "Destination MMSI", "mmsi", 41, 30, , "V")
      Call DetailOut(Retvalcol, "Retransmit Flag", "retransmit", 71, 1)
      Call DetailOut(Retvalcol, "Spare", "spare", 72, 1)
'v136      Call DetailOut(Retvalcol, "Text", "text", 73, 936)    'Original
'      Call DetailOut(Retvalcol, "Binary Data", "data", 73, clsSentence.AisPayloadBits - 73 + 1)
      Call DetailOut(Retvalcol, "Text", "text", 73, clsSentence.AisPayloadBits - 73 + 1)    'v136
        NextBit = clsSentence.AisPayloadBits + 1

'Try code from AtoN
'        NextBit = 73
'Debug.Print Len(clsSentence.AisPayload)
'        If clsSentence.AisPayloadBits > (NextBit - 1) Then
'            SpareBitsToBoundary = clsSentence.AisPayloadBits - (NextBit - 1) - Int((clsSentence.AisPayloadBits - (NextBit - 1)) / 6) * 6
'            Call DetailOut(Retvalcol, "Text", "text", NextBit, clsSentence.AisPayloadBits - (NextBit - 1) - SpareBitsToBoundary)
'            If SpareBitsToBoundary > 0 Then
'              Call DetailOut(Retvalcol, "Byte boundary filler", "spare", clsSentence.AisPayloadBits - SpareBitsToBoundary + 1, SpareBitsToBoundary)
'            End If
'            NextBit = clsSentence.AisPayloadBits + 1
'        End If
'end of try

'    Case Is = "13" see 7
  Case Is = "14"        'Broadcast Safety Related Message
      Call DetailOut(Retvalcol, "Spare", "spare", 39, 2)
'v136      Call DetailOut(Retvalcol, "Text", "text", 41, 968)
      Call DetailOut(Retvalcol, "Text", "text", 41, clsSentence.AisPayloadBits - 41 + 1) 'v136
'      Call DetailOut(Retvalcol, "Text", "text", 41, clsSentence.AisPayloadBits - 41 + 1)
        NextBit = clsSentence.AisPayloadBits + 1
  
'Try code from AtoN
'        NextBit = 41
'Debug.Print Len(clsSentence.AisPayload)
'        If clsSentence.AisPayloadBits > (NextBit - 1) Then
'            SpareBitsToBoundary = clsSentence.AisPayloadBits - (NextBit - 1) - Int((clsSentence.AisPayloadBits - (NextBit - 1)) / 6) * 6
'            Call DetailOut(Retvalcol, "Text", "text", NextBit, clsSentence.AisPayloadBits - (NextBit - 1) - SpareBitsToBoundary)
'            If SpareBitsToBoundary > 0 Then
'              Call DetailOut(Retvalcol, "Byte boundary filler", "spare", clsSentence.AisPayloadBits - SpareBitsToBoundary + 1, SpareBitsToBoundary)
'            End If
'            NextBit = clsSentence.AisPayloadBits + 1
'        End If
'end of try
  
  
  
  Case Is = "15"
      Call DetailOut(Retvalcol, "Spare", "spare", 39, 2)
      Call DetailOut(Retvalcol, "Interrogated MMSI", "mmsi", 41, 30, , "V")
      Call DetailOut(Retvalcol, "First message type", "aismsgtype", 71, 6)
      Call DetailOut(Retvalcol, "First slot offset", "offset", 77, 12)
        NextBit = 89    'total bits will be 88
      If pLong(91, 6) <> 0 Then
            Call DetailOut(Retvalcol, "Spare", "spare", 89, 2)
            Call DetailOut(Retvalcol, "Second message type", "aismsgtype", 91, 6)
            Call DetailOut(Retvalcol, "Second slot offset", "offset", 97, 12)
            Call DetailOut(Retvalcol, "Spare", "spare", 109, 2)
        NextBit = 111   'Total bits will be 112
      End If
      If pLong(111, 30) <> 0 Then
          Call DetailOut(Retvalcol, "Interrogated MMSI", "mmsi", 111, 30, , "V")
          Call DetailOut(Retvalcol, "First message type", "aismsgtype", 141, 6)
          Call DetailOut(Retvalcol, "First slot offset", "offset", 147, 12)
        NextBit = 159   'Total bits will be 160
      End If
      If NextBit > 89 Then
        Call DetailOut(Retvalcol, "Spare", "spare", NextBit, 2)
        NextBit = NextBit + 2
      End If
  Case Is = "16"
      Call DetailOut(Retvalcol, "Spare", "spare", 39, 2)
      Call DetailOut(Retvalcol, "Destination A MMSI", "mmsi", 41, 30, , "V")
      Call DetailOut(Retvalcol, "Offset A", "offset", 71, 12)
      Call DetailOut(Retvalcol, "Increment A", "increment", 83, 10)
      If pLong(93, 30) = 0 Then
          Call DetailOut(Retvalcol, "Spare", "spare", 93, 4)
          NextBit = 97
      Else
        Call DetailOut(Retvalcol, "Destination B MMSI", "mmsi", 93, 30, , "V")
        Call DetailOut(Retvalcol, "Offset B", "offset", 123, 12)
        Call DetailOut(Retvalcol, "Increment B", "increment", 135, 10)
        NextBit = 145
      End If
  Case Is = "17"
      Call DetailOut(Retvalcol, "Spare", "spare", 39, 2)
      Call DetailOut(Retvalcol, "Longitude", "lon", 41, 18)
      Call DetailOut(Retvalcol, "Latitude", "lat", 59, 17)
      Call DetailOut(Retvalcol, "Spare", "spare", 76, 5)
      Call DetailOut(Retvalcol, "DGNSS Correction Data", "dgnssdata", 81, clsSentence.AisPayloadBits - 81 + 1)
  Case Is = "18"
      Call DetailOut(Retvalcol, "Regional reserved", "reserved", 39, 8)
      Call DetailOut(Retvalcol, "Speed Over Ground (SOG)", "speed", 47, 10, 1)
      Call DetailOut(Retvalcol, "Position Accuracy", "accuracy", 57, 1)
      Call DetailOut(Retvalcol, "Longitude", "lon", 58, 28)
      Call DetailOut(Retvalcol, "Latitude", "lat", 86, 27)
      Call DetailOut(Retvalcol, "Course Over Ground (COG)", "course", 113, 12, 1)
      Call DetailOut(Retvalcol, "True Heading (HDG)", "heading", 125, 9, 0)
      Call DetailOut(Retvalcol, "Time Stamp", "second", 134, 6)
      Call DetailOut(Retvalcol, "Regional reserved", "reserved", 140, 2)
      Call DetailOut(Retvalcol, "CS Unit", "cs", 142, 1)
      Call DetailOut(Retvalcol, "Display flag", "display", 143, 1)
      Call DetailOut(Retvalcol, "DSC flag", "dsc", 144, 1)
      Call DetailOut(Retvalcol, "Band flag", "band", 145, 1)
      Call DetailOut(Retvalcol, "Message 22 flag", "msg22", 146, 1)
      Call DetailOut(Retvalcol, "Assigned", "assigned", 147, 1)
      Call DetailOut(Retvalcol, "RAIM Flag", "raim", 148, 1)
        Call DetailOut(Retvalcol, "Communication Mode", "radiomode", 149, 1)
        Call CommDetail(Retvalcol)
        NextBit = 169
  Case Is = "19"
      Call DetailOut(Retvalcol, "Regional reserved", "reserved", 39, 8)
      Call DetailOut(Retvalcol, "Speed Over Ground (SOG)", "speed", 47, 10, 1)
      Call DetailOut(Retvalcol, "Position Accuracy", "accuracy", 57, 1)
      Call DetailOut(Retvalcol, "Longitude", "lon", 58, 28)
      Call DetailOut(Retvalcol, "Latitude", "lat", 86, 27)
      Call DetailOut(Retvalcol, "Course Over Ground (COG)", "course", 113, 12, 1)
      Call DetailOut(Retvalcol, "True Heading (HDG)", "heading", 125, 9, 0)
      Call DetailOut(Retvalcol, "Time Stamp", "second", 134, 6)
      Call DetailOut(Retvalcol, "Spare", "spare", 140, 4)
      Call DetailOut(Retvalcol, "Vessel Name", "vesselname", 144, 120)
      Call DetailOut(Retvalcol, "Ship Type", "ship_type", 264, 8)
      Call DetailOut(Retvalcol, "Dimension to Bow", "to_bow", 272, 9)
      Call DetailOut(Retvalcol, "Dimension to Stern", "to_stern", 281, 9)
      Call DetailOut(Retvalcol, "Length", "clength", 272, 9)
      Call DetailOut(Retvalcol, "Dimension to Port", "to_port", 290, 6)
      Call DetailOut(Retvalcol, "Dimension to Starboard", "to_starboard", 296, 6)
      Call DetailOut(Retvalcol, "Beam", "cbeam", 290, 6)
      Call DetailOut(Retvalcol, "Position Type Fix", "epfd", 302, 4)
      Call DetailOut(Retvalcol, "RAIM Flag", "raim", 306, 1)
      Call DetailOut(Retvalcol, "DTE", "dte", 307, 1)
      Call DetailOut(Retvalcol, "Assigned", "assigned", 308, 1)
      Call DetailOut(Retvalcol, "Spare", "spare", 309, 4)
        NextBit = 313
  Case Is = "20"
      Call DetailOut(Retvalcol, "Spare", "spare", 39, 2)
'output slots upto the payload length
'nextbit is used to output fill bits at end of PayloadDetail

        If pLong(41, 30) = 0 Then
            Call DetailOut(Retvalcol, "Offset number 1", "offset", 41, 12, " no data link management information available")
            NextBit = 71
        Else
            Call DetailOut(Retvalcol, "Offset number 1", "offset", 41, 12)
            Call DetailOut(Retvalcol, "Reserved slots", "number", 53, 4)
            Call DetailOut(Retvalcol, "Time-out", "minute", 57, 3)
            Call DetailOut(Retvalcol, "Increment", "increment", 60, 11)
            NextBit = 71
            SpareBitsToBoundary = 2
            If clsSentence.AisPayloadBits > 70 + 2 Then '2 fill bits
                Call DetailOut(Retvalcol, "Offset number 2", "offset", 71, 12)
                Call DetailOut(Retvalcol, "Reserved slots", "number", 83, 4)
                Call DetailOut(Retvalcol, "Time-out", "minute", 87, 3)
                Call DetailOut(Retvalcol, "Increment", "increment", 90, 11)
                NextBit = 101
                SpareBitsToBoundary = 4
            End If
            If clsSentence.AisPayloadBits > 104 + 4 Then '4 fill bits
                Call DetailOut(Retvalcol, "Offset number 3", "offset", 101, 12)
                Call DetailOut(Retvalcol, "Reserved slots", "number", 113, 4)
                Call DetailOut(Retvalcol, "Time-out", "minute", 117, 3)
                Call DetailOut(Retvalcol, "Increment", "increment", 120, 11)
                NextBit = 131
                SpareBitsToBoundary = 6
            End If
            If clsSentence.AisPayloadBits > 136 + 6 Then '6 fill bits
                Call DetailOut(Retvalcol, "Offset number 4", "offset", 131, 12)
                Call DetailOut(Retvalcol, "Reserved slots", "number", 143, 4)
                Call DetailOut(Retvalcol, "Time-out", "minute", 147, 3)
                Call DetailOut(Retvalcol, "Increment", "increment", 150, 11)
                NextBit = 161
                SpareBitsToBoundary = 0
            End If
            If SpareBitsToBoundary > 0 Then
              Call DetailOut(Retvalcol, "Spare", "spare", NextBit, SpareBitsToBoundary)
                NextBit = NextBit + SpareBitsToBoundary
            End If
        End If  'data link management data is available
  Case Is = "21"
      Call DetailOut(Retvalcol, "Aid type", "aid_type", 39, 5)
      Call DetailOut(Retvalcol, "Name", "name", 44, 120)
      Call DetailOut(Retvalcol, "Position Accuracy", "accuracy", 164, 1)
      Call DetailOut(Retvalcol, "Longitude", "lon", 165, 28)
      Call DetailOut(Retvalcol, "Latitude", "lat", 193, 27)
      Call DetailOut(Retvalcol, "Dimension to Bow", "to_bow", 220, 9)
      Call DetailOut(Retvalcol, "Dimension to Stern", "to_stern", 229, 9)
      Call DetailOut(Retvalcol, "Length", "clength", 220, 9)
      Call DetailOut(Retvalcol, "Dimension to Port", "to_port", 238, 6)
      Call DetailOut(Retvalcol, "Dimension to Starboard", "to_starboard", 244, 6)
      Call DetailOut(Retvalcol, "Beam", "cbeam", 238, 6)
      Call DetailOut(Retvalcol, "Type of EPFD", "epfd", 250, 4)
      Call DetailOut(Retvalcol, "UTC Second", "second", 254, 6)
      Call DetailOut(Retvalcol, "Off-Position Indicator", "off_position", 260, 1)
      Call DetailOut(Retvalcol, "Regional reserved", "reserved", 261, 8)
      Call DetailOut(Retvalcol, "RAIM Flag", "raim", 269, 1)
      Call DetailOut(Retvalcol, "Virtual-aid flag", "virtual_aid", 270, 1)
      Call DetailOut(Retvalcol, "Assigned", "assigned", 271, 1)
      Call DetailOut(Retvalcol, "Spare", "spare", 272, 1)
      
'v136      Call DetailOut(Retvalcol, "Text", "text", 273, 88)
      Call DetailOut(Retvalcol, "Text", "text", 273, clsSentence.AisPayloadBits - 273 + 1) 'v136
        NextBit = clsSentence.AisPayloadBits + 1

'Debug.Print Len(clsSentence.AisPayload)
'Variable length text
'        If clsSentence.AisPayloadBits > 272 Then
'            SpareBitsToBoundary = clsSentence.AisPayloadBits - 272 - Int((clsSentence.AisPayloadBits - 272) / 6) * 6
'            Call DetailOut(Retvalcol, "Name Extension", "name", 273, clsSentence.AisPayloadBits - 272 - SpareBitsToBoundary)
'            If SpareBitsToBoundary > 0 Then
'              Call DetailOut(Retvalcol, "Byte boundary filler", "spare", clsSentence.AisPayloadBits - SpareBitsToBoundary + 1, SpareBitsToBoundary)
'            End If
'            NextBit = clsSentence.AisPayloadBits + 1
'        End If
'Try code from AtoN
'Debug.Print Len(clsSentence.AisPayload)
'        If clsSentence.AisPayloadBits > (NextBit - 1) Then  'Variable length > 0
'            SpareBitsToBoundary = clsSentence.AisPayloadBits - (NextBit - 1) - Int((clsSentence.AisPayloadBits - (NextBit - 1)) / 6) * 6
'            Call DetailOut(Retvalcol, "Text", "text", NextBit, clsSentence.AisPayloadBits - (NextBit - 1) - SpareBitsToBoundary)
'            If SpareBitsToBoundary > 0 Then
'              Call DetailOut(Retvalcol, "Byte boundary filler", "spare", clsSentence.AisPayloadBits - SpareBitsToBoundary + 1, SpareBitsToBoundary)
'            End If
'            NextBit = clsSentence.AisPayloadBits + 1
'        End If
'end of try
        
'        NextBit = 273   'Start of variable length
'        If clsSentence.AisPayloadBits > (NextBit - 1) Then  'Variable length > 0
'            Call DetailOut(Retvalcol, "Text", "text", NextBit, 0)   'To end of data
'            NextBit = clsSentence.AisPayloadBits + 1
'        End If
'Try setting variable no of bits in DetailOut
  
  
  Case Is = "22"
      Call DetailOut(Retvalcol, "Spare", "spare", 39, 2)
      Call DetailOut(Retvalcol, "Channel A", "channel", 41, 12)
      Call DetailOut(Retvalcol, "Channel B", "channel", 53, 12)
      Call DetailOut(Retvalcol, "Tx Rx mode", "txrx", 65, 4)
      Call DetailOut(Retvalcol, "Power", "power", 69, 1)
        If pLong(140, 1) = 0 Then   'broadcast
            Call DetailOut(Retvalcol, "NE Longitude", "lon", 70, 18, 1)
            Call DetailOut(Retvalcol, "NE Latitude", "lat", 88, 17, 1)
            Call DetailOut(Retvalcol, "SW Longitude", "lon", 105, 18, 1)
            Call DetailOut(Retvalcol, "SW Latitude", "lat", 123, 17, 1)
        Else                        'addressed
            If pLong(70, 30) <> 0 Then
                Call DetailOut(Retvalcol, "MMSI ID1", "mmsi", 70, 30, , "V")
                Call DetailOut(Retvalcol, "Spare", "spare", 100, 5)
            End If
            If pLong(105, 30) <> 0 Then
                Call DetailOut(Retvalcol, "MMSI ID2", "mmsi", 105, 30, , "V")
                Call DetailOut(Retvalcol, "Spare", "spare", 135, 5)
            End If
        End If
      Call DetailOut(Retvalcol, "Addressed", "addressed", 140, 1)
      Call DetailOut(Retvalcol, "Channel A Bandwidth", "band_width", 141, 1)
      Call DetailOut(Retvalcol, "Channel B Bandwidth", "band_width", 142, 1)
      Call DetailOut(Retvalcol, "Zone size", "zonesize", 143, 3)
      Call DetailOut(Retvalcol, "Spare", "spare", 146, 23)
        NextBit = 169
  Case Is = "23"
      Call DetailOut(Retvalcol, "Spare", "spare", 39, 2)
      Call DetailOut(Retvalcol, "NE Longitude", "lon", 41, 18, 1)
      Call DetailOut(Retvalcol, "NE Latitude", "lat", 59, 17, 1)
      Call DetailOut(Retvalcol, "SW Longitude", "lon", 76, 18, 1)
      Call DetailOut(Retvalcol, "SW Latitude", "lat", 94, 17, 1)
      Call DetailOut(Retvalcol, "Station Type", "stationtype", 111, 4)
      Call DetailOut(Retvalcol, "Ship Type", "ship_type", 115, 8)
      Call DetailOut(Retvalcol, "Spare", "spare", 123, 22)
      Call DetailOut(Retvalcol, "Tx Rx mode", "txrx", 145, 2)
      Call DetailOut(Retvalcol, "Report Interval", "interval", 147, 4)
      Call DetailOut(Retvalcol, "Quiet Time", "quiet", 151, 4)
      Call DetailOut(Retvalcol, "Spare", "spare", 155, 6)
        NextBit = 161
   Case Is = "24"
        Select Case DetailOut(Retvalcol, "Part no", "partno", 39, 2)
            Case Is = "0"
                Call DetailOut(Retvalcol, "Vessel Name", "vesselname", 41, 120)
                NextBit = 161
'Removed in ITU M.1371-3
'                Call DetailOut(RetValCol, "Spare", "spare", 161, 8)
            Case Is = "1"
                Call DetailOut(Retvalcol, "Ship Type", "ship_type", 41, 8)
                Call DetailOut(Retvalcol, "Vendor ID (to v3)", "vendorid", 49, 42)
                Call DetailOut(Retvalcol, "Vendor ID (from v4)", "vendorid", 49, 18)
                Call DetailOut(Retvalcol, "Unit Model Code (from v4)", "numeric", 67, 4)
                Call DetailOut(Retvalcol, "Unit Serial No (from v4)", "numeric", 71, 20)
                Call DetailOut(Retvalcol, "Call Sign", "callsign", 91, 42)
'                kb = pLong(9, 30) 'mmsi
                If Left$(clsSentence.AisMsgFromMmsi, 2) <> "98" Then
                    Call DetailOut(Retvalcol, "Dimensions to Bow", "to_bow", 133, 9)
                    Call DetailOut(Retvalcol, "Dimensions to Stern", "to_stern", 142, 9)
                    Call DetailOut(Retvalcol, "Length", "clength", 133, 9)
                    Call DetailOut(Retvalcol, "Dimensions to Port", "to_port", 151, 6)
                    Call DetailOut(Retvalcol, "Dimensions to Starboard", "to_starboard", 157, 6)
                    Call DetailOut(Retvalcol, "Beam", "cbeam", 151, 6)
                Else
                    Call DetailOut(Retvalcol, "Mothership MMSI", "mmsi", 133, 30, , "V")
                End If
                Call DetailOut(Retvalcol, "Type of EPFD", "epfd", 163, 4)   'Added M.1371-5
                Call DetailOut(Retvalcol, "Spare", "spare", 167, 2)         'Changed M.1371-5
                NextBit = 169
        End Select
   Case Is = "25"   'Broadcast or Addressed Binary Message
'Call DetailOut(Retvalcol, "Bits", "bits", 1, clsSentence.AisPayloadBits) 'test Charlie
        
         If DetailOut(Retvalcol, "Addressed", "addressed", 39, 1) = "0" Then
'Broadcast
            If DetailOut(Retvalcol, "Binary Data Flag", "structured", 40, 1) = "1" Then  'Binary Data Flag
                AiStart = 41        'Broadcast
            Else
                BinaryDataStart = 41    'No DacFi
            End If
        Else
'Addressed
            Call DetailOut(Retvalcol, "Destination MMSI", "mmsi", 41, 30, , "V")
            Call DetailOut(Retvalcol, "Spare", "spare", 71, 2)
            If DetailOut(Retvalcol, "Binary Data Flag", "structured", 40, 1) = "1" Then   'Binary Data Flag
'Structured (use Aidata)
                AiStart = 73        'Addressed
            Else
'Unstructured (Output binary as HEX because we have no idea of the format)
                BinaryDataStart = 73
            End If
        End If
        If AiStart > 0 Then
'Structured data (use Aidata)
            BinaryBits = clsSentence.AisPayloadBits - (AiStart - 1)
            Call DetailOut(Retvalcol, "Binary Data", "aidata", AiStart, BinaryBits)
        Else    'No DacFi
            BinaryBits = clsSentence.AisPayloadBits - (BinaryDataStart - 1)    '=128 if max of 168 payload bits
'Output Binary data as hex
            Call DetailOut(Retvalcol, "Binary Data", "data", BinaryDataStart, BinaryBits)

#If jnasetup = True Then
'Try text
            Call DetailOut(Retvalcol, "Try (6-bit Ascii) Text", "text", BinaryDataStart, BinaryBits)
#End If

'NextBit = 112-168  Addressed, 80-168 Broadcast

'Call DetailOut(Retvalcol, "Binary Data (6-bit text)", "text", BinaryDataStart, BinaryBits)
'Call DetailOut(Retvalcol, "Test For Text", "testfortext", BinaryDataStart, BinaryBits)
        
        End If
#If jnasetup = True Then    'Output bit positions
    Call DetailOut(Retvalcol, "Bits", "bits", 1, clsSentence.AisPayloadBits) 'test
#End If
   Case Is = "26"    'Broadcast or Addressed Binary Message

'Call DetailOut(Retvalcol, "Bits", "bits", 1, clsSentence.AisPayloadBits) 'test Charlie
        
        If DetailOut(Retvalcol, "Addressed", "addressed", 39, 1) = "0" Then
'Broadcast
            If DetailOut(Retvalcol, "Binary Data Flag", "structured", 40, 1) = "1" Then  'Binary Data Flag
                AiStart = 41        'Broadcast
            Else
                BinaryDataStart = 41    'No DacFi
            End If
        Else
'Addressed
            Call DetailOut(Retvalcol, "Destination MMSI", "mmsi", 41, 30, , "V")
            Call DetailOut(Retvalcol, "Spare", "spare", 71, 2)
            If DetailOut(Retvalcol, "Binary Data Flag", "structured", 40, 1) = "1" Then   'Binary Data Flag
'Structured (use Aidata)
                AiStart = 73        'Addressed
            Else
'Unstructured (Output binary as HEX because we have no idea of the format)
                BinaryDataStart = 73
            End If
        End If
        
        If AiStart > 0 Then
'Structured data (use Aidata)
            BinaryBits = clsSentence.AisPayloadBits - (AiStart - 1) - 20 - 4
            Call DetailOut(Retvalcol, "Binary Data", "aidata", AiStart, BinaryBits)
            Call DetailOut(Retvalcol, "Spare", "spare", AiStart + BinaryBits, 4)
Call DetailOut(Retvalcol, "Communication Mode", "radiomode", AiStart + BinaryBits + 4, 1)
        Else    'No DacFi
            BinaryBits = clsSentence.AisPayloadBits - (BinaryDataStart - 1) - 20 - 4 '=128 if max of 168 payload bits
'Output Binary data as hex
            Call DetailOut(Retvalcol, "Binary Data", "data", BinaryDataStart, BinaryBits)
#If jnasetup = True Then
'Try text
            Call DetailOut(Retvalcol, "Try (6-bit Ascii) Text", "text", BinaryDataStart, BinaryBits)
#End If
            Call DetailOut(Retvalcol, "Spare", "spare", BinaryDataStart + BinaryBits, 4)
Call DetailOut(Retvalcol, "Communication Mode", "radiomode", BinaryDataStart + BinaryBits + 4, 1)
        End If
        
        Call CommDetail(Retvalcol)  'Always 19 bits
    Case Is = "27"
      Call DetailOut(Retvalcol, "Position Accuracy", "accuracy", 39, 1)
      Call DetailOut(Retvalcol, "RAIM Flag", "raim", 40, 1)
      Call DetailOut(Retvalcol, "Navigation Status", "status", 41, 4)
      Call DetailOut(Retvalcol, "Longitude", "lon", 45, 18)
      Call DetailOut(Retvalcol, "Latitude", "lat", 63, 17)
      Call DetailOut(Retvalcol, "Speed Over Ground (SOG)", "speed", 80, 6, 0)
      Call DetailOut(Retvalcol, "Course Over Ground (COG)", "course", 86, 9, 0)
      Call DetailOut(Retvalcol, "Status of Current GNSS Position", "gnssposition", 95, 1)
      Call DetailOut(Retvalcol, "Spare", "spare", 96, 1)
        NextBit = 97
    End Select
'check if any bit stuffing (Msg 20 only at the moment)

    If NextBit <> 0 Then        'Check Payload length
        PayloadExcessBits = clsSentence.AisPayloadBits - (NextBit - 1)
        If PayloadExcessBits <> 0 Then
'V 144 Jan 2017 if $AITAG(Jason time stamp)has been found make CSV output compatible with V129
            If Left$(GroupSentence.NmeaSentence, 6) <> "$AITAG" Then
                Call DetailOut(Retvalcol, "Payload Size Check", "excessbits", NextBit, PayloadExcessBits)
            End If
        End If
    End If
        
'Move to Sentenceandpayloaddetail Jan 15
'    If MyShip.PositionTime <> "" Then
'        MyShipDetail (Retvalcol)
'    End If
End Sub

Function MyShipDetail(Retvalcol As Long)
    Call MyShipOut(Retvalcol, "MyShip", "myship")
    Call MyShipOut(Retvalcol, "Latitude", "myshiplat")
    Call MyShipOut(Retvalcol, "Longitude", "myshiplon")
    Call MyShipOut(Retvalcol, "Range", "myshiprange")
    Call MyShipOut(Retvalcol, "Position Age Difference", "myshiprangeage")

End Function

'called by MyShipDetail, RetValCol not is required
'Called by TagsFromFields, RetValCol is required
'MyShipOut MUST calculate field positions
Function MyShipOut(Retvalcol As Long, Des As String, _
Member As String) As String   '(1,des,2=val,3=valdes)
Dim wlong As Long
Dim Bold As Boolean
Dim Val As String
Dim Valdes As String
Dim wSi As Single
Dim ThisShip As ShipDef
Dim RhumbLine As Single
Dim GreatCircle As Single

    Select Case Member
    Case Is = "myship"
'update myship (including vessel name)
        MyShip = GetCachedVessel(MyShip.Mmsi)
        Val = MyShip.Mmsi
        Valdes = MyShip.Name
        Bold = True
    Case Is = "myshiplat"
        wSi = MyShip.Lat
        If wSi = 91 Then
            Val = ""
            Valdes = "not available"
        Else
            Val = Replace(Format$(wSi, "0.000000"), ",", ".")
            Valdes = aLatLon(wSi, "Lat")
        End If
'V 144 Jan 2017 unless $AITAG(Jason time stamp)has been found make CSV output compatible with V129
        If Left$(GroupSentence.NmeaSentence, 6) = "$AITAG" And wSi = 0 Then
            Val = ""
        End If
    Case Is = "myshiplon"
        wSi = MyShip.Lon
        If wSi = 181 Then
            Val = ""
            Valdes = "not available"
        Else
            Val = Replace(Format$(wSi, "0.000000"), ",", ".")
            Valdes = aLatLon(wSi, "Lon")
        End If
'V 144 Jan 2017 unless $AITAG(Jason time stamp)has been found make CSV output compatible with V129
        If Left$(GroupSentence.NmeaSentence, 6) = "$AITAG" And wSi = 0 Then
            Val = ""
        End If
    Case Is = "myshiprange"
        ThisShip.Mmsi = Format$(pLong(9, 30), "000000000")
        ThisShip = GetCachedVessel(ThisShip.Mmsi)
        If MyShip.Lat <> 91 And MyShip.Lon <> 181 _
        And ThisShip.Lat <> 0 And ThisShip.Lon <> 0 Then
            GreatCircle = LatLonDistance(MyShip.Lat, MyShip.Lon, ThisShip.Lat, ThisShip.Lon)
'            RhumbLine = RhumbLineDistance(MyShip.Lat, MyShip.Lon, ThisShip.Lat, ThisShip.Lon)
'RhumbLine = RhumbLineDistance(1, -1, 1, 1)
'If GreatCircle - RhumbLine > 2 Or GreatCircle - RhumbLine < -2 Then Stop
            Val = GreatCircle
'this is to allow range to be set to Min= 0.001
            If Val > 0.01 Then
                Val = Format$(Val, "0.00")
            Else
                Val = Format$(Val, "0.000")
            End If
            Valdes = Val & " nm"
        Else
            Val = ""
            Valdes = "not available"
        End If
'V 144 Jan 2017 unless $AITAG(Jason time stamp)has been found make CSV output compatible with V129
        If Left$(GroupSentence.NmeaSentence, 6) = "$AITAG" Then
            If MyShip.Lat = 0 And MyShip.Lon = 0 Then
                 Val = ""
            End If
        End If
    Case Is = "myshiprangeage"
        ThisShip.Mmsi = Format$(pLong(9, 30), "000000000")
        ThisShip = GetCachedVessel(ThisShip.Mmsi)
        If MyShip.RcvTime <> "" And ThisShip.RcvTime <> "" Then
            Val = DateDiff("s", MyShip.RcvTime, ThisShip.RcvTime)
            Valdes = Val & " seconds"
        Else
            Val = ""
            Valdes = "not available"
'V 144 Jan 2017 unless $AITAG(Jason time stamp)has been found make CSV output compatible with V129
            If Left$(GroupSentence.NmeaSentence, 6) = "$AITAG" Then
                 Val = "0"
                 If ThisShip.Lat = 91 And ThisShip.Lon = 181 Then
                    Val = ""
                End If
            End If
        End If
    Case Else
        MsgBox "MyShipOut member (" & Member & ") not found"
    End Select
'required by detaillineout to construct source etc
    clsField.CallingRoutine = "MyShipOut"
    clsField.Des = Des
    clsField.Member = Member
'    clsField.from = CommBase
'    clsField.reqbits = reqbits
'    clsField.Arg = Arg
'   If RetValCol <> 0 Then Call DetailLineOut(Des, Val, ValDes)
'this will be 0 even if arg is not passed, note cant then return col 0
    Select Case Retvalcol
        Case Is = 0
            Call DetailLineOut(Des, Val, Valdes, Bold) 'main output
            MyShipOut = Val
        Case Is = 1
            MyShipOut = Des
        Case Is = 2
            MyShipOut = Val
        Case Is = 3
            MyShipOut = Valdes
        Case Is = 4
            If AllFieldsNo > UBound(AllFields) Then ReDim Preserve AllFields(AllFieldsNo)
            AllFields(AllFieldsNo) = Val
            AllFieldsNo = AllFieldsNo + 1
            MyShipOut = Val
    End Select
End Function

Function CommDetail(Retvalcol As Long)
    Call CommOut(Retvalcol, "Communication", "comm")
    If CommState = 0 Then   'SOTDMA
        Call CommOut(Retvalcol, "Sync State", "syncstate")
        Call CommOut(Retvalcol, "Slot Time-out", "slottimeout")
        Select Case CommSubType
        Case Is = 1
'must be CommOut as if detail out is called in isolation it will
'not know the SlotTimeout
            Call CommOut(Retvalcol, "Received Stations", "rcvstations")
        Case Is = 2
            Call CommOut(Retvalcol, "This Slot Number", "thisslot")
        Case Is = 3
            Call CommOut(Retvalcol, "Hour UTC", "hour")
            Call CommOut(Retvalcol, "Minute UTC", "minute")
        Case Is = 4
            Call CommOut(Retvalcol, "Next Slot Offset", "slotnumber")
        End Select
    Else                'ITDMA
        Call CommOut(Retvalcol, "Sync State", "syncstate")
        Call CommOut(Retvalcol, "Next Slot Offset", "slotincr")
        Call CommOut(Retvalcol, "Slots to Allocate", "slotalloc")
        Call CommOut(Retvalcol, "Keep Flag", "slotkeep")
    End If

#If jnasetup = True Then    'Output bit positions
    Call DetailOut(Retvalcol, "Bits", "bits", 1, clsSentence.AisPayloadBits) 'test
#End If
End Function
'called by CommDetail, RetValCol not is required
'Called by TagsFromFields, RetValCol is required
'CommOut MUST calculate field positions
Function CommOut(Retvalcol As Long, Des As String, _
Member As String) As String   '(1,des,2=val,3=valdes)
Dim wlong As Long
Dim Bold As Boolean
Dim Val As String
Dim Valdes As String
Dim kb As String

'Dim Payloadbits As Long

    Select Case Member
    Case Is = "comm"
        Val = CommState
        Valdes = RadioModeName(CommState) '0=sotdma,1=itdma
        Bold = True
kb = "(" & CommBase & "-" & CommBase + 18 & ")"
    Case Is = "syncstate"       'tdma
        Val = pLong(CommBase, 2)
        Valdes = SyncStateName(Val)
kb = "(" & CommBase & "-" & CommBase + 1 & ")"
    Case Is = "slotincr"        'itdma
        Val = pLong(CommBase + 2, 13)
        Valdes = "Slots"
        CommOut = Val
kb = "(" & CommBase + 2 & "-" & CommBase + 2 + 12 & ")"
    Case Is = "slotalloc"   'itdma
        Val = pLong(CommBase + 15, 3)
        Valdes = SlotAllocName(Val)
kb = "(" & CommBase + 15 & "-" & CommBase + 15 + 2 & ")"
    Case Is = "slotkeep"    'itdma
        Val = pLong(CommBase + 18, 1)
        If Val = 1 Then Valdes = "Keep slot allocated for 1 extra frame"
kb = "(" & CommBase + 18 & "-" & CommBase + 18 & ")"
    Case Is = "slottimeout"      'sotdma
        Val = pLong(CommBase + 2, 3)
        If CInt(Val) = 0 Then
            Valdes = "Last Frame"
        Else
            Valdes = "Slots Left"
        End If
        CommOut = Val
kb = "(" & CommBase + 2 & "-" & CommBase + 2 + 2 & ")"
    Case Is = "rcvstations"     'sotdma
        If CommSubType = 1 Then
            Val = pLong(CommBase + 5, 14)
kb = "(" & CommBase + 5 & "-" & CommBase + 5 + 13 & ")"
        End If
    Case Is = "slotnumber"
        If CommSubType = 4 Then
            Val = pLong(CommBase + 5, 14)
kb = "(" & CommBase + 5 & "-" & CommBase + 5 + 13 & ")"
        End If
    Case Is = "thisslot"
        If CommSubType = 2 Then
            wlong = pLong(CommBase + 5, 14)
            Val = wlong
kb = "(" & CommBase + 5 & "-" & CommBase + 5 + 13 & ")"
            If wlong > 2249 Then Valdes = "Invalid"
        End If
    Case Is = "hour"
        If CommSubType = 3 Then
            Val = Format$(pLong(CommBase + 5, 5), "00")
            If Val = 24 Then Valdes = "not available (default)"
kb = "(" & CommBase + 5 & "-" & CommBase + 5 + 4 & ")"
        End If
    Case Is = "minute"
        If CommSubType = 3 Then
            Val = Format$(pLong(CommBase + 10, 7), "00")
            If Val = 60 Then Valdes = "not available (default)"
kb = "(" & CommBase + 10 & "-" & CommBase + 10 + 6 & ")"
        End If
    Case Else
        MsgBox "Commout member (" & Member & ") not found"
    End Select

#If jnasetup = True Then    'Output bit positions
    Des = Des & kb
#End If

'required by detaillineout to construct source etc
    clsField.CallingRoutine = "CommOut"
    clsField.Des = Des
    clsField.Member = Member
'    clsField.from = CommBase
'    clsField.reqbits = reqbits
'    clsField.Arg = Arg
'   If RetValCol <> 0 Then Call DetailLineOut(Des, Val, ValDes)
'this will be 0 even if arg is not passed, note cant then return col 0
    Select Case Retvalcol
        Case Is = 0
            Call DetailLineOut(Des, Val, Valdes, Bold) 'main output
            CommOut = Val
        Case Is = 1
            CommOut = Des
        Case Is = 2
            CommOut = Val
        Case Is = 3
            CommOut = Valdes
        Case Is = 4
            If AllFieldsNo > UBound(AllFields) Then ReDim Preserve AllFields(AllFieldsNo)
            AllFields(AllFieldsNo) = Val
            AllFieldsNo = AllFieldsNo + 1
            CommOut = Val
    End Select
End Function

'Extracts the des,val and valdes for the Parameter code from cbWords
'If called by SentenceandPayloadDetail & RetValCol = 0
'If called by SetTagValues, RetValCol is required
Function CommentOut(Retvalcol As Long, Des As String, _
Member As String) As String   '(1,des,2=val,3=valdes)
Dim wlong As Long
Dim Bold As Boolean
Dim Val As String
Dim Valdes As String
'Dim Payloadbits As Long

    Select Case Member
    Case Is = "commentblock"
        Val = clsCb.Block
        Valdes = clsCb.errmsg
        Bold = True
    Case Is = "c"   'Unix Time
        Val = clsCb.Time
        If IsNumeric(Val) Then Valdes = UnixTimeToDate(Val) & " UTC"
    Case Is = "i"   'Lowercase i (freeform text)
        Val = clsCb.Text
    Case Is = "s"   'Source
        Val = clsCb.Source
    Case Is = "d"   'Destination
        Val = clsCb.Destination
    Case Is = "x"   'Counter
        Val = clsCb.Counter
    Case Is = "G", "g" 'Group
        Val = clsCb.GroupId
        Valdes = "Line " & clsCb.GroupLine & " of " & clsCb.GroupLines & " lines"
    Case Is = "cbcrc"
        Val = clsCb.CbCrc
        If Val = "hh" Then
            Valdes = "Test Data (CRC invalid)"
        Else
            Valdes = clsCb.CRCerrmsg
        End If
    Case Else
        Val = clsCb.Unknown
        Valdes = "Parameter-Code " & Member & ": is unknown"
    End Select

'required by detaillineout to construct source etc
    clsField.CallingRoutine = "CommentOut"
    clsField.Des = Des
    clsField.Member = Member
'   If RetValCol <> 0 Then Call DetailLineOut(Des, Val, ValDes)
'this will be 0 even if arg is not passed, note cant then return col 0
    Select Case Retvalcol
        Case Is = 0
            Call DetailLineOut(Des, Val, Valdes, Bold) 'main output
            CommentOut = Val
        Case Is = 1
            CommentOut = Des
        Case Is = 2
            CommentOut = Val
        Case Is = 3
            CommentOut = Valdes
        Case Is = 4
            If AllFieldsNo > UBound(AllFields) Then ReDim Preserve AllFields(AllFieldsNo)
            AllFields(AllFieldsNo) = Val
            AllFieldsNo = AllFieldsNo + 1
            CommentOut = Val
    End Select
End Function

'Extracts the des,val and valdes for the NMEA part of any AIS sentence
'(!aaVDO or !aaVDM) from clsAisSentence. IsAisSentence will be true
'If called by SentenceandPayloadDetail & RetValCol = 0
'If called by SetTagValues, RetValCol is required
Function NmeaAisOut(Retvalcol As Long, Des As String, _
Member As String, Optional WordNo As Long) As String  '(1,des,2=val,3=valdes)
Dim wlong As Long
Dim Bold As Boolean
Dim Val As String
Dim Valdes As String
Dim Payload6Bits As Long    'used to calculate 8 bits
Dim Payload8Bits As Long
Dim i As Long
Dim arry() As String
'Dim Payloadbits As Long
    Select Case Member
    Case Is = "vesselname"
        Val = clsSentence.VesselName
        If Val = "" Then
            Valdes = "Not yet received"
        End If
        Bold = True
    Case Is = "aispayload"  'This is the re-assembled payload
        Val = clsSentence.AisPayload
        Valdes = clsSentence.AisPayloadBits & " bits (" & clsSentence.AisPayloadBits / 8 & " 8-bit words)"
#If jnasetup = True Then
Des = Des & "(" & Len(Val) & " chrs)"
#End If
If clsSentence.AisMsgPartsComplete = False Then Valdes = Valdes & ", incomplete"
    Case Is = "aismsgtype"  '?? this is NOT necessarily the received message type
        If IsNumeric(clsSentence.AisMsgType) Then
            Val = clsSentence.AisMsgType
            If Val <= 0 Or Val > 27 Then
                Valdes = "{Invalid message type}"
            Else
                Valdes = AisMsgTypeName(Val)
            End If
        End If
    Case Is = "mmsi"
        If IsNumeric(clsSentence.AisMsgFromMmsi) Then
            Val = clsSentence.AisMsgFromMmsi
            Valdes = MmsiFmt(Val, "D")
        End If
    Case Is = "mid"    'this is NOT necessarily the FromMMSI
        If IsNumeric(clsSentence.AisMsgFromMmsi) Then
            Val = MmsiFmt(clsSentence.AisMsgFromMmsi, "M")
            If Val <> 0 Then
                If Val = 45133333 Then
                    Valdes = "VTS Drechtsteden"
                Else
                    Valdes = DacName(Val)
                End If
            Else
                Valdes = "Not defined"
           End If
        End If
    Case Is = "repeat"
            If IsNumeric(clsSentence.AisMsgRepeat) Then
                Val = clsSentence.AisMsgRepeat
                Valdes = RepeatName(Val)
            End If
    Case Is = "vessellat"
        Val = clsSentence.VesselLat
        If IsNumeric(Val) Then
            Valdes = LatToNmea(CSng(Val))   'DDMM.MM,N or S
        Else
            Valdes = "invalid," ', because N or S is null
        End If
    Case Is = "vessellon"
        Val = clsSentence.VesselLon
        If IsNumeric(Val) Then
            Valdes = LonToNmea(CSng(Val))   'DDD.MM,W or E
        Else
            Valdes = "invalid," ', because W or E is null
        End If
    Case Is = "aisword" 'Called from some tags in .ini files Dec14
                        'Should now use
'Can have less words in sentence than a word that is Tagged in the field list
'testv129   Debug.Print clsSentence.FullSentence
        
        If WordNo <= UBound(NmeaWords) Then
            Val = NmeaWords(WordNo)
            Select Case WordNo
            Case Is = 0
                Valdes = TalkerDes(clsSentence.IecTalkerID)
                Payload6Bits = Len(NmeaWords(5)) * 6
                Payload8Bits = Int(Payload6Bits / 8) * 8
                Valdes = Valdes & ", " & Payload8Bits & " bits (" & Len(NmeaWords(5)) & " 6-bit words)"
            Case Is = 1, 2, 3
                Valdes = AisWordCheck(WordNo)
            Case Is = 4     'Radio Channel
            Case Is = 5
'            Payload8Bits = Int(Payload6Bits / 8) * 8
'Subscript error if Payload bytes not set up
'                Payload8Bits = (PayloadByteArraySize + 1) * 8
'                Valdes = Payload8Bits & " bits (" & Payload8Bits / 8 & " 8-bit words)"
                    Valdes = clsSentence.AisPayloadBits & " bits (" & clsSentence.AisPayloadBits / 8 & " 8-bit words)"
Des = Des & "-" & Len(Val)
                If clsSentence.AisMsgPartsComplete = False Then Valdes = Valdes & ", incomplete"
            Case Is = 6     'Fill bits
'multi-part and NOT the last part
                If NmeaWords(1) <> NmeaWords(2) Then
                    If NmeaWords(6) <> "0" Then
                        Valdes = Valdes & "NMEA fill bits should be 0"
                    End If
                End If
            Case Else   'Words 7 onwards
                If IsNumeric(Val) Then  'Assume unix time
                    Valdes = Format$(UnixTimeToDate(Val), DateTimeOutputFormat)
                Else
                    Valdes = "Local Extension " & WordNo
                End If
            End Select
        End If
   Case Is = "payloadbits"
'For i = 0 To 20
'i = 78
'Debug.Print "CharNo6=" & i & ",Word8=" & ChrnoToWords(i)
'Debug.Print "Bits6=" & i * 6 & ",Bits8=" & ChrnoToWords(i) * 8 & ",Fill=" & ChrnoToFillBits(i); ""
'Next i

'Debug.Print ChrnoToBytes(Len(clsSentence.AisPayload)) * 8 & ":" & ChrnoToWords(Len(clsSentence.AisPayload)) * 8 & ":" & ChrnoToFillBits(Len(clsSentence.AisPayload))
'        Payload6Bits = Len(clsSentence.AisPayload) * 6
'MsgBox Len(clssentence.AisPayload)
'        Payload8Bits = (PayloadByteArraySize + 1) * 8
'        Payload8Bits = (UBound(PayloadBytes) * 8) + 6
'Changed back Raymond spec incorrect for Msg 24A
'+ 6 because it only needs filling to a 6 bit boundary
'eq message 24A where we need 160 bits, 6 bit is 162 & 6 bit is 168
        
'        If Payload6Bits <> Payload8Bits - clsSentence.AisPayloadFillBits Then
'            Val = Payload6Bits - Payload8Bits - clsSentence.AisPayloadFillBits
'            Valdes = "Payload not filled to byte boundary"
'            Bold = True
'        Else
'            Exit Function
'        End If
        Val = clsSentence.AisPayloadBits Mod 8
        If Val <> 0 Then
            Valdes = "Payload not filled to byte boundary"
        Else
            Exit Function   'Dont output anything
        End If
'V 144 Jan 2017 if $AITAG(Jason time stamp)has been found make CSV output compatible with V129
        If Left$(GroupSentence.NmeaSentence, 6) = "$AITAG" Then
            Exit Function
        End If
    Case Else
'this is because the RcvTime will be the last word (if it is a time)
'but there could be other words inserted as a local extension before the rcvtime
    Des = Member
    Val = NmeaWords(WordNo)
    Valdes = "NmeaAisOut member (" & Member & ") not found"
    End Select

'required by detaillineout to construct source etc
    clsField.CallingRoutine = "NmeaAisOut"
    clsField.Des = Des
    clsField.Member = Member
    clsField.From = WordNo
'   If RetValCol <> 0 Then Call DetailLineOut(Des, Val, ValDes)
'this will be 0 even if arg is not passed, note cant then return col 0
    Select Case Retvalcol
        Case Is = 0
            Call DetailLineOut(Des, Val, Valdes, Bold) 'main output
            NmeaAisOut = Val
        Case Is = 1
            NmeaAisOut = Des
        Case Is = 2
            NmeaAisOut = Val
        Case Is = 3
            NmeaAisOut = Valdes
        Case Is = 4
            If AllFieldsNo > UBound(AllFields) Then ReDim Preserve AllFields(AllFieldsNo)
            AllFields(AllFieldsNo) = Val
            AllFieldsNo = AllFieldsNo + 1
            NmeaAisOut = Val
    End Select
End Function

'Jan15 assumes AisPayloadBits agrees with
'get start position of where to decode the communication state
Function CommBase() As Long
Dim PayloadBits As Long
'If clsSentence.AisMsgType <> pLong(1, 6) Then Stop
'If clsSentence.AisPayloadBits <> (PayloadByteArraySize + 1) * 8 Then Stop

    Select Case pLong(1, 6)
    Case Is = 1, 2, 3, 4, 9, 11, 18
        CommBase = 150  'fixed 168-19+1
    Case Is = 26
'AisPayloadBits is the true payload size (6-bit payload characters adjusted by fill bits)
'    PayloadBits = (PayloadByteArraySize + 1) * 8
'    CommBase = PayloadBits - 18
        CommBase = clsSentence.AisPayloadBits - 19 + 1
    Case Else   'This is the first bit after the data (excludes comm state - if any)
        CommBase = clsSentence.AisPayloadBits + 1
    End Select
End Function

Function CommState() As Long
If clsSentence.AisMsgType <> pLong(1, 6) Then
    Call WriteErrorLog("CommState " & clsSentence.AisMsgType & ":" & pLong(1, 6) & " differ" & vbCrLf & clsSentence.NmeaSentence)
End If
Select Case pLong(1, 6)     'msg type clssentence may be incorrect at this point
Case Is = 1, 2, 4, 11
    CommState = 0
Case Is = 3 'always ITDMA
    CommState = 1
Case Is = 9, 18, 26     'Class B (18) should always by SOTDMA
    CommState = pLong(CommBase - 1, 1)  'get the state (previous bit)
Case Else
    MsgBox "No Radio Mode for AIS Message Type " & pLong(1, 6)
End Select
End Function

Function CommSubType() As Long
Dim SlotTimeout As Long

If CommState = 0 Then
    SlotTimeout = pLong(CommBase + 2, 3)
    Select Case SlotTimeout
    Case Is = 3, 5, 7
        CommSubType = 1
    Case Is = 2, 4, 6
        CommSubType = 2
    Case Is = 1
        CommSubType = 3
    Case Is = 0
        CommSubType = 4
    End Select
Else
    CommSubType = 0
End If
End Function

Function Application(Retvalcol As Long, From As Long, DataBits As Long)
'output's details of binary part of the message given
'the bit positions of the binary message including header(DAC & FI)

Dim Dac As String
Dim Fi As String
Dim Last  As Integer 'repeated sections
Dim i As Integer
Dim TargType As Integer
Dim ShapeType As Integer
Dim LatLonPrecision As Long
Dim Multiplier As String    'passed back to detailout as "scale"
Dim kb As String
'Dim Payloadbits As Long

'check weve enough bits in the payload to decode the Dac & fi
'we must have enough bits for the Dac+fi
'     Payloadbits = (UBound(PayloadBytes) + 1) * 8
     If DataBits < 16 Then
    kb = "Application Identifier invalid (" _
                    & 16 - DataBits & " bits too short)"
    Call WriteErrorLog(kb & vbCrLf & clsSentence.NmeaSentence)
        Exit Function
     End If
Debug.Print "from=" & From
Dac = DetailOut(Retvalcol, "DAC", "dac", From, 10)
'Check if Dac has been mapped to another dac
If Dac <> clsSentence.AisMsgDac Then
    clsSentence.AisMsgDac = Dac
End If
Fi = pLong(From + 10, 6)
If Fi <> clsSentence.AisMsgFi Then
    kb = "Stop Encountered in Application" & vbCrLf _
    & "Fi=" & Fi & vbCrLf _
    & "clssentence.AisMsgFi=" & clsSentence.AisMsgFi & vbCrLf
'    MsgBox kb
    Call WriteErrorLog(kb & vbCrLf & clsSentence.NmeaSentence)
'    Stop     'check
End If
Fi = clsSentence.AisMsgFi

'fi descriptions got from name file for international messages
'because we have name for capability reply
'If Dac <> 1 Then Fi = DetailOut(retvalcol,"FI", "fi", from + 10, 6)

Select Case clsSentence.AisMsgDac
Case Is = "0"       'IALA
    Select Case clsSentence.AisMsgFi
    Case Is = "0"   'Zeni
        Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, , "Information required")
        Select Case clsSentence.AisMsgFiId
        Case Is = "1"
        Case Else
            Call DetailOut(Retvalcol, "Message ID", "fiid", From + 16, 16, "Information required")
        End Select
    Case Else
        Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, , "Information required")
    End Select
Case Is = "1"  'international
    Select Case clsSentence.AisMsgFi
    Case Is = "0"
'Stop
        Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6)
        Select Case clsSentence.AisMsgType
        Case Is = "6", "8"
            Call DetailOut(Retvalcol, "Acknowledgement Required", "ack", From + 16, 1)
            Call DetailOut(Retvalcol, "Sequence Number", "seqno", From + 17, 11)
            Call DetailOut(Retvalcol, "Text", "text", From + 28, DataBits - 28)
        Case Else   'No Ack required
            Call DetailOut(Retvalcol, "Sequence Number", "seqno", From + 16, 11)
            Call DetailOut(Retvalcol, "Text", "text", From + 27, DataBits - 27)
        End Select
    Case Is = "1"
        Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6, , "Replaced by 5")
        Call DetailOut(Retvalcol, "DAC code of received FM", "dac", From + 16, 10)
        Call DetailOut(Retvalcol, "FI code of received FM", "ifi", From + 26, 6)
        Call DetailOut(Retvalcol, "Text Sequence number", "seqno", From + 32, 11)
        Call Dac1Out(Retvalcol, "AI Available", "aiavailable", From + 43, 1)
        Call Dac1Out(Retvalcol, "AI Response", "airesponse", From + 44, 3)
        Call DetailOut(Retvalcol, "Spare", "spare", From + 47, 49)
    Case Is = "2"
        Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6)
        Call DetailOut(Retvalcol, "Requested DAC code", "dac", From + 16, 10)
        Call DetailOut(Retvalcol, "Requested FI code", "ifi", From + 26, 6)
        Call DetailOut(Retvalcol, "Spare", "spare", From + 32, 64)
    Case Is = "3"
        Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6)
        Call DetailOut(Retvalcol, "Requested DAC code", "dac", From + 16, 10)
        Call DetailOut(Retvalcol, "Spare", "spare", From + 32, 70)
    Case Is = "4"
'        Call DetailOut(retvalcol,"DAC", "dac", from + 16, 10)
'       next line for debugging
'        Call DetailOut(retvalcol,"FI Availability Bits", "bits", from + 26, 128)
        Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6)
        Call DetailOut(Retvalcol, "FI Availability", "fiavailability", From + 26, 128)
        Call DetailOut(Retvalcol, "Spare", "spare", From + 154, 126)
    Case Is = "5"
        Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6, , "Awaiting Verification")
        Call DetailOut(Retvalcol, "DAC code of received FM", "dac", From + 16, 10)
        Call DetailOut(Retvalcol, "FI code of received FM", "ifi", From + 26, 6)
        Call DetailOut(Retvalcol, "Text Sequence number", "seqno", From + 32, 11)
        Call Dac1Out(Retvalcol, "AI Available", "aiavailable", From + 43, 1)
        Call Dac1Out(Retvalcol, "AI Response", "airesponse", From + 44, 3)
        Call DetailOut(Retvalcol, "Spare", "spare", From + 47, 49)
    Case Is = "11"
        Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6, , "Deprecated")
        Call DetailOut(Retvalcol, "Latitude", "lat", From + 16, 24)
        Call DetailOut(Retvalcol, "Longitude", "lon", From + 40, 25)
        Call DetailOut(Retvalcol, "Day UTC", "day", From + 65, 5)
        Call DetailOut(Retvalcol, "Hour UTC", "hour", From + 70, 5)
        Call DetailOut(Retvalcol, "Minute UTC", "minute", From + 75, 6)
        Call DetailOut(Retvalcol, "Average Wind Speed", "speed", From + 81, 7, 0, "Average for last 10 minutes") '.1Kt
        Call DetailOut(Retvalcol, "Wind Gust", "speed", From + 88, 7, 0, "Average for last 10 minutes") '.1Kt
        Call DetailOut(Retvalcol, "Wind Direction", "heading", From + 95, 9, 0)
        Call DetailOut(Retvalcol, "Wind Gust Direction", "heading", From + 104, 9, 0)
'1-11
        Call DetailOut(Retvalcol, "Air Temperature", "temperature", From + 113, 11, 1, "-600")
        Call DetailOut(Retvalcol, "Relative Humidity", "humidity", From + 124, 7, 0)
        Call DetailOut(Retvalcol, "Dew Point", "temperature", From + 131, 10, 1, "-200")
        Call DetailOut(Retvalcol, "Air Pressure", "pressure", From + 141, 9)
        Call DetailOut(Retvalcol, "Air Pressure tendency", "tendency", From + 150, 2)
'        Call DetailOut(RetValCol, "Horizontal visibility", "visibility", From + 152, 8, 1)
'1-11 does not use msb format
        Call DetailOut(Retvalcol, "Horizontal visibility", "miles", From + 152, 8, 1)
        Call DetailOut(Retvalcol, "Water Level (incl tide)", "chartdatum", From + 160, 9, 1)
        Call DetailOut(Retvalcol, "Water level trend", "tendency", From + 169, 2)
        Call DetailOut(Retvalcol, "Surface current speed", "speed", From + 171, 8, 1)
        Call DetailOut(Retvalcol, "Surface current direction", "heading", From + 179, 9, 0)
        Call DetailOut(Retvalcol, "Current speed, #2", "speed", From + 187, 8, 1)
        Call DetailOut(Retvalcol, "Current direction, #2", "heading", From + 196, 9, 0)
        Call DetailOut(Retvalcol, "Current measuring level, #2", "meters", From + 204, 5, 1)
        Call DetailOut(Retvalcol, "Current speed, #3", "speed", From + 209, 8, 1)
        Call DetailOut(Retvalcol, "Current direction, #3", "heading", From + 218, 9, 0)
        Call DetailOut(Retvalcol, "Current measuring level, #3", "meters", From + 227, 5, 1)
        Call DetailOut(Retvalcol, "Significant wave height", "meters", From + 232, 8, 1)
        Call DetailOut(Retvalcol, "Wave period", "second", From + 240, 6)
        Call DetailOut(Retvalcol, "Wave Direction", "heading", From + 246, 9, 0)
        Call DetailOut(Retvalcol, "Swell height", "meters", From + 255, 8, 1)
        Call DetailOut(Retvalcol, "Swell period", "second", From + 263, 6)
        Call DetailOut(Retvalcol, "Swell direction", "heading", From + 269, 9, 0)
        Call DetailOut(Retvalcol, "Sea State", "beaufort", From + 278, 4)
        Call DetailOut(Retvalcol, "Water temperature", "temperature", From + 282, 10, 1, "-100")
        Call DetailOut(Retvalcol, "Precipitation (type)", "precipitation", From + 292, 3)
        Call DetailOut(Retvalcol, "Salinity", "salinity", From + 295, 9, 1)
        Call DetailOut(Retvalcol, "Ice", "yesno", From + 304, 2)
        Call DetailOut(Retvalcol, "Spare", "spare", From + 306, 6)
    Case Is = "12"
        Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6, , "replaced by 25")
    Case Is = "13"
        Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6, , "replaced by 22")
    Case Is = "14"
        Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6, , "Deprecated")
        Call DetailOut(Retvalcol, "Month UTC", "month", From + 16, 4)
        Call DetailOut(Retvalcol, "Day UTC", "day", From + 20, 5)
        Last = From + 25 - 1
        Do Until Last >= DataBits
            i = i + 1
'            Call Rowout("Position #" & i, "", "", True)
'            Call DetailLineOut("Position # & i", "", "", True)
            Call DetailOut(Retvalcol, "Position #", "literal", , , CStr(i))
            Call DetailOut(Retvalcol, "Latitude", "lat", Last + 1, 27)
            Call DetailOut(Retvalcol, "Longitude", "lon", Last + 28, 28)
            Call DetailOut(Retvalcol, "From UTC hour", "hour", Last + 55, 5)
            Call DetailOut(Retvalcol, "From UTC minute", "minute", Last + 60, 6)
'next 2 fields on position 2 missing on spec
            Call DetailOut(Retvalcol, "To UTC hour", "hour", Last + 66, 5)
            Call DetailOut(Retvalcol, "To UTC minute", "minute", Last + 71, 6)
            Call DetailOut(Retvalcol, "Current Direction predicted", "heading", Last + 77, 9, 0)
            Call DetailOut(Retvalcol, "Current Speed predicted", "speed", Last + 86, 97, 1)
            Last = Last + 112
         Loop
    Case Is = "15"
        Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6, , "Deprecated")
        Call DetailOut(Retvalcol, "Height above Keel", "meters", From + 16, 11, 1)
        Call DetailOut(Retvalcol, "Spare", "spare", From + 27, 5)
    Case Is = "16"
        If DataBits = 32 Then
            Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6, "No of Persons on Board")
            Call DetailOut(Retvalcol, "Persons", "persons", From + 16, 13)
            Call DetailOut(Retvalcol, "Spare", "spare", From + 29, 3)
        Else
            Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6, "VTS Targets", "legacy")
            Last = From + 16 - 1
            Do Until Last >= DataBits
                i = i + 1
'                Call Rowout("VTS Target " & i, "", "", True)
'               Call DetailLineOut("VTS Target " & i, "", "", True)
                Call DetailOut(Retvalcol, "VTS Target", "literal", , , CStr(i))
                TargType = DetailOut(Retvalcol, "Target Identifier Type", "targtype", Last + 1, 2)
                Select Case TargType
                Case Is = 0
                    Call DetailOut(Retvalcol, "Spare", "spare", Last + 3, 12)
                    Call DetailOut(Retvalcol, "MMSI Number", "mmsi", Last + 15, 30)
                Case Is = 1
                    Call DetailOut(Retvalcol, "Spare", "spare", Last + 3, 12)
                    Call DetailOut(Retvalcol, "IMO Number", "imo", Last + 15, 30)
                Case Is = 2
                    Call DetailOut(Retvalcol, "Callsign", "callsign", Last + 3, 42)
                Case Is = 3
                    Call DetailOut(Retvalcol, "Other", "text", Last + 3, 42)
                End Select  'Target type
                Call DetailOut(Retvalcol, "Spare", "spare", Last + 45, 4)
                Call DetailOut(Retvalcol, "Latitude", "lat", Last + 49, 24)
                Call DetailOut(Retvalcol, "Longitude", "lon", Last + 73, 25)
                Call DetailOut(Retvalcol, "COG", "course", Last + 98, 9, 0)
                Call DetailOut(Retvalcol, "Time Stamp", "second", Last + 107, 6)
                Call DetailOut(Retvalcol, "SOG", "speed", Last + 113, 8, 0)
                Last = Last + 120 '120 is size of repeated section
            Loop
        End If
    Case Is = "17"  'see above
        Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6, , "Awaiting verification")
        Last = From + 16 - 1
        Do Until Last >= DataBits
            i = i + 1
'            Call Rowout("VTS Target " & i, "", "", True)
'            Call DetailLineOut("VTS Target " & i, "", "", True)
            Call DetailOut(Retvalcol, "VTS Target", "literal", , , CStr(i))
            TargType = DetailOut(Retvalcol, "Target Identifier Type", "targtype", Last + 1, 2)
            Select Case TargType
            Case Is = 0
                Call DetailOut(Retvalcol, "Spare", "spare", Last + 3, 12)
                Call DetailOut(Retvalcol, "MMSI Number", "mmsi", Last + 15, 30)
            Case Is = 1
                Call DetailOut(Retvalcol, "Spare", "spare", Last + 3, 12)
                Call DetailOut(Retvalcol, "IMO Number", "imo", Last + 15, 30)
            Case Is = 2
                Call DetailOut(Retvalcol, "Callsign", "callsign", Last + 3, 42)
            Case Is = 3
                Call DetailOut(Retvalcol, "Other", "text", Last + 3, 42)
            End Select  'Target type
            Call DetailOut(Retvalcol, "Spare", "spare", Last + 45, 4)
            Call DetailOut(Retvalcol, "Latitude", "lat", Last + 49, 24)
            Call DetailOut(Retvalcol, "Longitude", "lon", Last + 73, 25)
            Call DetailOut(Retvalcol, "COG", "course", Last + 98, 9, 0)
            Call DetailOut(Retvalcol, "Time Stamp", "second", Last + 107, 6)
            Call DetailOut(Retvalcol, "SOG", "speed", Last + 113, 8, 0)
            Last = Last + 120 '120 is size of repeated section
        Loop
    Case Is = "18"
        Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6, , "Awaiting Verification")
        Call DetailOut(Retvalcol, "Message Linkage ID", "seqno", From + 16, 10)
        Call DetailOut(Retvalcol, "Month", "month", From + 26, 4)
        Call DetailOut(Retvalcol, "Day", "day", From + 30, 5)
        Call DetailOut(Retvalcol, "Hour", "hour", From + 35, 5)
        Call DetailOut(Retvalcol, "Minute", "minute", From + 40, 6)
        Call DetailOut(Retvalcol, "Port and Berth", "text", From + 46, 120)
        Call DetailOut(Retvalcol, "Destination", "text", From + 166, 30, , " UN LOCODE")
        Call DetailOut(Retvalcol, "Position, Longitude", "lon", From + 196, 25, 3)
        Call DetailOut(Retvalcol, "Position, Latitude", "lat", From + 221, 24, 3)
        Call DetailOut(Retvalcol, "Spare", "spare", From + 245, 43)
'total size doesnt add up on spec
    Case Is = "19"
        If DataBits = 32 Then
            Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6, "Extended Ship And Voyage related Data", "legacy")
            Call DetailOut(Retvalcol, "Height above Keel", "meters", From + 16, 11, 1)
            Call DetailOut(Retvalcol, "Spare", "spare", From + 27, 5)
        Else
            Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6, , "Awaiting Verification")
            Call DetailOut(Retvalcol, "Message Linkage ID", "seqno", From + 16, 10)
            Call DetailOut(Retvalcol, "Name of Signal Station", "text", From + 26, 120)
            Call DetailOut(Retvalcol, "Position, Longitude", "lon", From + 146, 25, 3)
            Call DetailOut(Retvalcol, "Position, Latitude", "lat", From + 171, 24, 3)
            Call Dac1Out(Retvalcol, "Status of Signal", "signalstatus", From + 195, 1)
'format of signal not defined
            Call Dac1Out(Retvalcol, "Signal in Service", "signalservice", From + 196, 5)
            Call DetailOut(Retvalcol, "UTC Hour of next signal shift", "hour", From + 201, 5)
            Call DetailOut(Retvalcol, "UTC Minute of next signal shift", "minute", From + 206, 6)
            Call Dac1Out(Retvalcol, "Expected Next Signal", "signalservice", From + 212, 5)
            Call DetailOut(Retvalcol, "Spare", "spare", From + 217, 103)
        End If
    Case Is = "20"
        Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6, , "Awaiting Verification")
        Call DetailOut(Retvalcol, "Message Linkage ID", "seqno", From + 16, 10)
        Call DetailOut(Retvalcol, "Name of Berth", "text", From + 26, 120)
        Call DetailOut(Retvalcol, "Position of Berth, Longitude", "lon", From + 146, 25, 3)
        Call DetailOut(Retvalcol, "Position of Berth, Latitude", "lat", From + 171, 24, 3)
        Call DetailOut(Retvalcol, "Spare", "spare", From + 195, 93)
     Case Is = "21" 'Weather Observation from Ship
'xxx
        Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6, , "Awaiting Verification")
        Select Case clsSentence.AisMsgFiId
            Case Is = "0"
            Call DetailOut(Retvalcol, "Type of Weather report", "fiid", From + 16, 1, , "Weather Observation report from ship")
            Call DetailOut(Retvalcol, "Geographic Location", "text", From + 17, 120)
            Call DetailOut(Retvalcol, "Position of Observation Longitude", "lon", From + 137, 25, 3)
            Call DetailOut(Retvalcol, "Position of Observation Latitude", "lat", From + 162, 24, 3)
            Call DetailOut(Retvalcol, "Day of Observation", "day", From + 186, 5)
            Call DetailOut(Retvalcol, "Hour UTC of Observation", "hour", From + 191, 5)
            Call DetailOut(Retvalcol, "Minute UTC of Observation", "minute", From + 196, 6)
            Call DetailOut(Retvalcol, "Present Weather", "ituwmoweather", From + 202, 4)
'msb format
            Call DetailOut(Retvalcol, "Visibility", "visibility", From + 206, 8, 1)
            Call DetailOut(Retvalcol, "Relative Humidity", "humidity", From + 214, 7, 0)
            Call DetailOut(Retvalcol, "Average Wind Speed", "speed", From + 221, 7, 0, "Average for last 10 minutes") '.1Kt
            Call DetailOut(Retvalcol, "Wind Direction", "heading", From + 228, 9, 0)
            Call DetailOut(Retvalcol, "Air Pressure", "pressure", From + 237, 9, 0)
            Call DetailOut(Retvalcol, "Air Pressure tendency", "ituwmotendency", From + 246, 4)
            Call DetailOut(Retvalcol, "Air Temperature", "temperature", From + 250, 11, "air")
            Call DetailOut(Retvalcol, "Water Temperature", "temperature", From + 261, 10, 1, "water")
            Call DetailOut(Retvalcol, "Wave period", "wmosecond", From + 271, 6)
            Call DetailOut(Retvalcol, "Wave height", "wmometers", From + 277, 8, 1)
            Call DetailOut(Retvalcol, "Wave Direction", "heading", From + 285, 9, 0)
            Call DetailOut(Retvalcol, "Swell height", "wmometers", From + 294, 8, 1)
            Call DetailOut(Retvalcol, "Swell direction", "heading", From + 302, 9, 0)
            Call DetailOut(Retvalcol, "Swell period", "wmosecond", From + 311, 6)
            Call DetailOut(Retvalcol, "Spare", "spare", From + 317, 3)
       Case Is = "1"
            Call DetailOut(Retvalcol, "Type of Weather report", "fiid", From + 16, 1, , "WMO Weather Observation report from ship")
       End Select   'fi
    Case Is = "22", "23"    'area notice
        Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6, , "Area Notice")
        Call DetailOut(Retvalcol, "Message Linkage ID", "seqno", From + 16, 10)
        Call Dac1Out(Retvalcol, "Area Type", "areatype", From + 26, 7)
        Call DetailOut(Retvalcol, "Month", "month", From + 33, 4)
        Call DetailOut(Retvalcol, "Day", "day", From + 37, 5)
        Call DetailOut(Retvalcol, "Hour", "hour", From + 42, 5)
        Call DetailOut(Retvalcol, "Minute", "minute", From + 47, 6)
        Call DetailOut(Retvalcol, "Duration", "number", From + 53, 18, , "Minutes")
'Sub-areas
        Last = From + 71 - 1
'Last = Last + 87
        Do Until Last >= DataBits
            i = i + 1
'            Call Rowout("Sub-area " & i, "", "", True)
'            Call DetailLineOut("Sub-area " & i, "", "", True)
            Call DetailOut(Retvalcol, "Sub-area", "literal", , , CStr(i))
'Last = Last - 1
'Call DetailOut(Retvalcol, "Sub-area bits", "bits", Last + 1, 87, Last + 1 & " to " & Last + 88 & " (base 1)")
            ShapeType = Dac1Out(Retvalcol, "Area Shape", "shape", Last + 1, 3)
            Select Case ShapeType
            Case 0 'Circle
                Multiplier = DetailOut(Retvalcol, "Scale Factor", "scale", Last + 4, 2)
                LatLonPrecision = pLong(Last + 55, 3)
                Call DetailOut(Retvalcol, "Longitude", "lon", Last + 6, 25, CStr(LatLonPrecision))
                Call DetailOut(Retvalcol, "Latitude", "lat", Last + 31, 24, CStr(LatLonPrecision))
                Call DetailOut(Retvalcol, "Precision", "precision", Last + 55, 3, , " decimal places")
                Call DetailOut(Retvalcol, "Radius", "scalenumber", Last + 58, 12, Multiplier, " meters")
                Call DetailOut(Retvalcol, "Spare", "spare", Last + 70, 18)
            Case 1
                Multiplier = DetailOut(Retvalcol, "Scale Factor", "scale", Last + 4, 2)
                LatLonPrecision = pLong(Last + 55, 3)
                Call DetailOut(Retvalcol, "Longitude", "lon", Last + 6, 25, CStr(LatLonPrecision))
                Call DetailOut(Retvalcol, "Latitude", "lat", Last + 31, 24, CStr(LatLonPrecision))
                Call DetailOut(Retvalcol, "Precision", "precision", Last + 55, 3, , " decimal places")
                Call DetailOut(Retvalcol, "E dimension", "scalenumber", Last + 58, 8, Multiplier, " meters")
                Call DetailOut(Retvalcol, "N dimension", "scalenumber", Last + 66, 8, Multiplier, " meters")
                Call DetailOut(Retvalcol, "Orientation", "heading", Last + 74, 9, 0)
                Call DetailOut(Retvalcol, "Spare", "spare", Last + 83, 5)
            Case 2
                Multiplier = DetailOut(Retvalcol, "Scale Factor", "scale", Last + 4, 2)
                LatLonPrecision = pLong(Last + 55, 3)
                Call DetailOut(Retvalcol, "Longitude", "lon", Last + 6, 25, CStr(LatLonPrecision))
                Call DetailOut(Retvalcol, "Latitude", "lat", Last + 31, 24, CStr(LatLonPrecision))
                Call DetailOut(Retvalcol, "Precision", "precision", Last + 55, 3, , " decimal places")
                Call DetailOut(Retvalcol, "Radius", "scalenumber", Last + 58, 12, Multiplier, " meters")
                Call DetailOut(Retvalcol, "Left Boundary", "heading", Last + 70, 9, 0)
                Call DetailOut(Retvalcol, "right Boundary", "heading", Last + 79, 9, 0)
                'no spare
            Case 3, 4
                Multiplier = DetailOut(Retvalcol, "Scale Factor", "scale", Last + 4, 2)
                LatLonPrecision = pLong(Last + 55, 3)
                Call DetailOut(Retvalcol, "Point #1 angle", "heading", Last + 6, 10, 1)
                Call DetailOut(Retvalcol, "Point #1 distance", "scalenumber", Last + 16, 10, Multiplier, " meters")
                Call DetailOut(Retvalcol, "Point #2 angle", "heading", Last + 26, 10, 1)
                Call DetailOut(Retvalcol, "Point #2 distance", "scalenumber", Last + 36, 10, Multiplier, " meters")
                Call DetailOut(Retvalcol, "Point #3 angle", "heading", Last + 46, 10, 1)
                Call DetailOut(Retvalcol, "Point #3 distance", "scalenumber", Last + 56, 10, Multiplier, " meters")
                Call DetailOut(Retvalcol, "Point #4 angle", "heading", Last + 66, 10, 1)
                Call DetailOut(Retvalcol, "Point #5 distance", "scalenumber", Last + 76, 10, Multiplier, " meters")
                Call DetailOut(Retvalcol, "Spare", "spare", Last + 86, 2)
            Case 5
                Call DetailOut(Retvalcol, "Text", "text", Last + 4, 84)
                'no spare
            End Select
            Last = Last + 87
         Loop
    Case Is = "24"
        Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6, , "Awaiting Verification")
        Call DetailOut(Retvalcol, "Message Linkage ID", "seqno", From + 16, 10)
        Call DetailOut(Retvalcol, "Air Draught", "airdraught", From + 26, 13, 2)
        Call DetailOut(Retvalcol, "Last port of call", "text", From + 39, 30, 0, "UN Locode")
        Call DetailOut(Retvalcol, "Next port of call", "text", From + 69, 30, 0, "UN Locode")
        Call DetailOut(Retvalcol, "Second port of call", "text", From + 99, 30, 0, "UN Locode")
        For i = 1 To 26
            Call DetailOut(Retvalcol, SolasEquipmentName(i) & " Status", "solasequipment", From + 129 + (2 * (i - 1)), 2)
        Next i
        Call DetailOut(Retvalcol, "Ice Class", "iceclass", From + 181, 4)
                
        Call DetailOut(Retvalcol, "Shaft Horse power", "number", From + 185, 18, 0, "Horsepower")
        Call DetailOut(Retvalcol, "VHF working channel", "vhfchannel", From + 203, 12)
        Call DetailOut(Retvalcol, "Lloyd's Ship Type", "text", From + 215, 42, 0, "Lloyd's statcode5")
        Call DetailOut(Retvalcol, "Gross tonnage", "number", From + 257, 18, 0, "Tonnes")
        Call DetailOut(Retvalcol, "Laden/Ballast", "loaded", From + 275, 2)
        Call DetailOut(Retvalcol, "Heavy Fuel Oil Bunker", "yesno2", From + 277, 2)
        Call DetailOut(Retvalcol, "Light Fuel Oil Bunker", "yesno2", From + 279, 2)
        Call DetailOut(Retvalcol, "Diesel Oil Bunker", "yesno2", From + 281, 2)
        Call DetailOut(Retvalcol, "Total amount of bunker oil", "number", From + 283, 14, 0, "Tonnes")
        Call DetailOut(Retvalcol, "No of persons on board", "number", From + 297, 13, 0, "including crew")
        Call DetailOut(Retvalcol, "Spare", "spare", From + 310, 10)
    Case Is = "25"
        Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6, , "Awaiting verification")
        Call DetailOut(Retvalcol, "Unit of quantity", "dgunits", From + 16, 2)
        Call DetailOut(Retvalcol, "Total quantity of dangerous cargo", "number", From + 18, 10)
        
        Last = From + 28 - 1
        Do Until Last >= DataBits + From - 1
            i = i + 1
            Select Case DetailOut(Retvalcol, "Cargo " & i, "dgcode", Last + 1, 4)
                Case Is = "1"
                    Call DetailOut(Retvalcol, "IMDG Class or division", "imdgcode", Last + 5, 7)
                    Call DetailOut(Retvalcol, "Spare", "spare", Last + 12, 6)
                Case Is = "2"
                    Call DetailOut(Retvalcol, "UN number", "igccode", Last + 5, 13)
                Case Is = "3"
                    Call DetailOut(Retvalcol, "BC Code", "bccode", Last + 5, 3)
                    Call DetailOut(Retvalcol, "IMDG Class or division", "imdgcode", Last + 8, 7)
                    Call DetailOut(Retvalcol, "Spare", "spare", Last + 15, 3)
                Case Is = "4"
                    Call DetailOut(Retvalcol, "Marpol I List of Oils", "marpol1", Last + 5, 4)
                    Call DetailOut(Retvalcol, "Spare", "spare", Last + 9, 9)
                Case Is = "5"
                    Call DetailOut(Retvalcol, "Marpol II IBC code", "marpol2", Last + 5, 3)
                    Call DetailOut(Retvalcol, "Spare", "spare", Last + 8, 10)
            End Select
            Last = Last + 17
        Loop
    Case Is = "26"
        Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6, , "Awaiting Decoding")
'        Call Environmental(RetValCol, From, DataBits)
    Case Is = "27", "28"
        Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6, , "Awaiting verification")
        Call DetailOut(Retvalcol, "Message Linkage ID", "seqno", From + 16, 10)
        Call Dac1Out(Retvalcol, "Sender Classification", "sender", From + 26, 3)
        Call Dac1Out(Retvalcol, "Route Type", "route", From + 29, 5)
        Call DetailOut(Retvalcol, "Month", "month", From + 34, 4)
        Call DetailOut(Retvalcol, "Day", "day", From + 38, 5)
        Call DetailOut(Retvalcol, "Hour", "hour", From + 43, 5)
        Call DetailOut(Retvalcol, "Minute", "minute", From + 48, 6)
        Call DetailOut(Retvalcol, "Duration", "number", From + 54, 18, , "Minutes")
        Call DetailOut(Retvalcol, "Number of Waypoints", "number", From + 72, 5)
'no of waypoints
        Last = From + 77 - 1
'not atlough there is provision fo 10 * 3 lights it appears only a max of 9 are used
        Do Until Last >= DataBits
            i = i + 1
'            Call Rowout("Waypoint " & i, "", "", True)
'            Call DetailLineOut("Waypoint " & i, "", "", True)
            Call DetailOut(Retvalcol, "Waypoint", "literal", , , CStr(i))
            Call DetailOut(Retvalcol, "Longitude", "lon", Last + 1, 28)
            Call DetailOut(Retvalcol, "Latitude", "lat", Last + 29, 27)
            Last = Last + 55
         Loop
    Case Is = "29"
        Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6, , "Awaiting verification")
        Call DetailOut(Retvalcol, "Message Linkage ID", "seqno", From + 16, 10)
'Calculate max bits available in payload (for payload bit check)
'        Last = 26 'max=966 DataBits
        Call DetailOut(Retvalcol, "Text string", "text", From + 26, DataBits - 26)
'        Call DetailOut(Retvalcol, "Text string", "text", From + 26, 966)
    Case Is = "30"
        Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6, , "Awaiting verification")
        Call DetailOut(Retvalcol, "Message Linkage ID", "seqno", From + 16, 10)
'v136        Call DetailOut(Retvalcol, "Text string", "text", From + 26, 930)
        Call DetailOut(Retvalcol, "Text string", "text", From + 26, DataBits - 26)
    Case Is = "31"  'met/hydro
        Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6, , "Awaiting verification")
        Call DetailOut(Retvalcol, "Longitude", "lon", From + 16, 25)
        Call DetailOut(Retvalcol, "Latitude", "lat", From + 41, 24)
        Call DetailOut(Retvalcol, "Position Accuracy", "accuracy", 65, 1)
        Call DetailOut(Retvalcol, "Day UTC", "day", From + 66, 5)
        Call DetailOut(Retvalcol, "Hour UTC", "hour", From + 71, 5)
        Call DetailOut(Retvalcol, "Minute UTC", "minute", From + 76, 6)
        Call DetailOut(Retvalcol, "Average Wind Speed", "speed", From + 82, 7, 0, "Average for last 10 minutes") '.1Kt
        Call DetailOut(Retvalcol, "Wind Gust", "speed", From + 89, 7, 0, "Average for last 10 minutes") '.1Kt
        Call DetailOut(Retvalcol, "Wind Direction", "heading", From + 96, 9, 0)
        Call DetailOut(Retvalcol, "Wind Gust Direction", "heading", From + 105, 9, 0)
        Call DetailOut(Retvalcol, "Air Temperature", "temperature", From + 114, 11, 1, "air")
        Call DetailOut(Retvalcol, "Relative Humidity", "humidity", From + 125, 7, 0)
        Call DetailOut(Retvalcol, "Dew Point", "temperature", From + 132, 10, 1, "dew")
        Call DetailOut(Retvalcol, "Air Pressure", "pressure", From + 142, 9)
        Call DetailOut(Retvalcol, "Air Pressure tendency", "tendency", From + 151, 2)
'msb format
        Call DetailOut(Retvalcol, "Horizontal visibility", "visibility", From + 153, 8, 1)
        Call DetailOut(Retvalcol, "Water Level (incl tide)", "chartdatum", From + 161, 12, 2, -100)
        Call DetailOut(Retvalcol, "Water level trend", "tendency", From + 173, 2)
        Call DetailOut(Retvalcol, "Surface current speed", "speed", From + 175, 8, 1, "251")
        Call DetailOut(Retvalcol, "Surface current direction", "heading", From + 183, 9, 0)
        Call DetailOut(Retvalcol, "Current speed, #2", "speed", From + 192, 8, 1, "251")
        Call DetailOut(Retvalcol, "Current direction, #2", "heading", From + 200, 9, 0)
        Call DetailOut(Retvalcol, "Current measuring level, #2", "meters", From + 209, 5, 1)
        Call DetailOut(Retvalcol, "Current speed, #3", "speed", From + 214, 8, 1, "251")
        Call DetailOut(Retvalcol, "Current direction, #3", "heading", From + 222, 9, 0)
        Call DetailOut(Retvalcol, "Current measuring level, #3", "meters", From + 231, 5, 1)
        Call DetailOut(Retvalcol, "Significant wave height", "meters", From + 236, 8, 1, "251")
        Call DetailOut(Retvalcol, "Wave period", "second", From + 244, 6)
        Call DetailOut(Retvalcol, "Wave Direction", "heading", From + 250, 9, 0)
        Call DetailOut(Retvalcol, "Swell height", "meters", From + 259, 8, 1, "251")
        Call DetailOut(Retvalcol, "Swell period", "second", From + 267, 6)
        Call DetailOut(Retvalcol, "Swell direction", "heading", From + 273, 9, 0)
        Call DetailOut(Retvalcol, "Sea State", "beaufort", From + 282, 4)
        Call DetailOut(Retvalcol, "Water temperature", "temperature", From + 286, 10, 1, "water")
        Call DetailOut(Retvalcol, "Precipitation (type)", "precipitation", From + 296, 3)
        Call DetailOut(Retvalcol, "Salinity", "salinity", From + 299, 9, 1, "501")
        Call DetailOut(Retvalcol, "Ice", "yesno", From + 308, 2)
        Call DetailOut(Retvalcol, "Spare", "spare", From + 310, 10)
    Case Is = "32"  'tidal window
        Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6, , "Awaiting verification")
        Call DetailOut(Retvalcol, "Month UTC", "month", From + 16, 4)
        Call DetailOut(Retvalcol, "Day UTC", "day", From + 20, 5)
        
        Call DetailOut(Retvalcol, "Position #1 Longitude", "lon", From + 25, 25, 3)
        Call DetailOut(Retvalcol, "Position #1 Latitude", "lat", From + 50, 24, 3)
        Call DetailOut(Retvalcol, "From Hour UTC", "hour", From + 74, 5)
        Call DetailOut(Retvalcol, "From Minute UTC", "minute", From + 79, 6)
        Call DetailOut(Retvalcol, "To Hour UTC", "hour", From + 85, 5)
        Call DetailOut(Retvalcol, "To Minute UTC", "minute", From + 90, 6)
        Call DetailOut(Retvalcol, "Current direction predicted #1", "heading", From + 96, 9, 0)
        Call DetailOut(Retvalcol, "Current speed predicted #1", "speed", From + 105, 8, 1, "251")
        
        Call DetailOut(Retvalcol, "Position #2 Longitude", "lon", From + 113, 25, 3)
        Call DetailOut(Retvalcol, "Position #2 Latitude", "lat", From + 138, 24, 3)
        Call DetailOut(Retvalcol, "From Hour UTC", "hour", From + 162, 5)
        Call DetailOut(Retvalcol, "From Minute UTC", "minute", From + 167, 6)
'no to Hour/Min in spec !! SN.1/Circ 289
        Call DetailOut(Retvalcol, "Current direction predicted #2", "heading", From + 173, 9, 0)
        Call DetailOut(Retvalcol, "Current speed predicted #2", "speed", From + 182, 8, 1, "251")
        
        Call DetailOut(Retvalcol, "Position #3 Longitude", "lon", From + 190, 25, 3)
        Call DetailOut(Retvalcol, "Position #3 Latitude", "lat", From + 215, 24, 3)
        Call DetailOut(Retvalcol, "From Hour UTC", "hour", From + 239, 5)
        Call DetailOut(Retvalcol, "From Minute UTC", "minute", From + 244, 6)
        Call DetailOut(Retvalcol, "To Hour UTC", "hour", From + 250, 5)
        Call DetailOut(Retvalcol, "To Minute UTC", "minute", From + 255, 6)
        Call DetailOut(Retvalcol, "Current direction predicted #3", "heading", From + 261, 9, 0)
        Call DetailOut(Retvalcol, "Current speed predicted #3", "speed", From + 270, 8, 1, "251")
        
    Case Is = "40"
        Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6)
        Call DetailOut(Retvalcol, "Persons", "persons", From + 16, 13)
        Call DetailOut(Retvalcol, "Spare", "spare", From + 29, 3)
    Case Else
        Call DetailOut(Retvalcol, "FI", "ifi", From + 10, 6, , "Information required")
    End Select          'international fi
Case Is = "103"
    Select Case Fi
    Case Else
        Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, , "Information required")
    End Select
Case Is = "200"      'inland
    Select Case Fi
    Case Is = "10", "8", "4" '4,8 appears to be an old spec
        If Fi = "10" Then
            Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, "Inland Ship Static and Voyage Related Data")
        Else
            Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, "Inland Ship Static and Voyage Related Data", "Appears an ""old"" spec")
        End If
        Call DetailOut(Retvalcol, "ENI", "eni", From + 16, 48)
        Call DetailOut(Retvalcol, "Length", "meters", From + 64, 13, 1)
        Call DetailOut(Retvalcol, "Beam", "meters", From + 77, 10, 1)
        Call DetailOut(Retvalcol, "Ship or Combination Type", "eri", From + 87, 14)
        Call DetailOut(Retvalcol, "Hazardous Cargo", "cone", From + 101, 3)
        Call DetailOut(Retvalcol, "Draught", "meters", From + 104, 11, 2)
        Call DetailOut(Retvalcol, "Loaded/unloaded", "loaded", From + 115, 2)
        Call DetailOut(Retvalcol, "Quality of Speed information", "qualspeed", From + 117, 1)
        Call DetailOut(Retvalcol, "Quality of Course information", "qualspeed", From + 118, 1)
        Call DetailOut(Retvalcol, "Quality of Heading information", "qualhead", From + 119, 1)
        Call DetailOut(Retvalcol, "Spare", "spare", From + 120, 8)
    Case Is = "21"
        Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, "ETA at lock/bridge/terminal", "Awaiting Verification")
        Call DetailOut(Retvalcol, "UN Country Code", "text", From + 16, 12)
        Call DetailOut(Retvalcol, "UN/LOCODE Code", "text", From + 28, 18)
        Call DetailOut(Retvalcol, "Fairway Section Number", "text", From + 46, 30)
        Call DetailOut(Retvalcol, "Terminal Code", "text", From + 76, 30)
        Call DetailOut(Retvalcol, "Fairway Hectometer", "text", From + 106, 30)
        Call DetailOut(Retvalcol, "ETA UTC Month", "month", From + 136, 4)
        Call DetailOut(Retvalcol, "ETA UTC Day", "day", From + 140, 5)
        Call DetailOut(Retvalcol, "ETA UTC Hour", "hour", From + 145, 5)
        Call DetailOut(Retvalcol, "ETA UTC Minute", "minute", From + 150, 6)
        Call Dac200Out(Retvalcol, "No of Tugboats Assisting", "tug", From + 156, 3)
        Call Dac200Out(Retvalcol, "Air Draught", "airdraught", From + 159, 12, 2)
        Call DetailOut(Retvalcol, "Spare", "spare", From + 171, 5)
    Case Is = "22"
        Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, "RTA at lock/bridge/terminal", "Awaiting Verification")
        Call DetailOut(Retvalcol, "UN Country Code", "text", From + 16, 12)
        Call DetailOut(Retvalcol, "UN/LOCODE Code", "text", From + 28, 18)
        Call DetailOut(Retvalcol, "Fairway Section Number", "text", From + 46, 30)
        Call DetailOut(Retvalcol, "Terminal Code", "text", From + 76, 30)
        Call DetailOut(Retvalcol, "Fairway Hectometer", "text", From + 106, 30)
        Call DetailOut(Retvalcol, "RTA UTC Month", "month", From + 136, 4)
        Call DetailOut(Retvalcol, "RTA UTC Day", "day", From + 140, 5)
        Call DetailOut(Retvalcol, "RTA UTC Hour", "hour", From + 145, 5)
        Call DetailOut(Retvalcol, "RTA UTC Minute", "minute", From + 150, 6)
        Call Dac200Out(Retvalcol, "Lock/bridge/terminal Status", "lockstatus", From + 156, 2)
        Call DetailOut(Retvalcol, "Spare", "spare", From + 158, 2)
    Case Is = "23"
        Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, "EMMA warning", "Awaiting Verification")
        Call DetailOut(Retvalcol, "Start UTC Year", "year", From + 16, 8)
        Call DetailOut(Retvalcol, "Start UTC Month", "month", From + 24, 4)
        Call DetailOut(Retvalcol, "Start UTC Day", "day", From + 28, 5)
        Call DetailOut(Retvalcol, "End UTC Year", "year", From + 33, 8)
        Call DetailOut(Retvalcol, "End UTC Month", "month", From + 41, 4)
        Call DetailOut(Retvalcol, "End UTC Day", "day", From + 45, 5)
        Call DetailOut(Retvalcol, "Start UTC Hour", "hour", From + 50, 5)
        Call DetailOut(Retvalcol, "Start UTC Minute", "minute", From + 55, 6)
        Call DetailOut(Retvalcol, "End UTC Hour", "hour", From + 61, 5)
        Call DetailOut(Retvalcol, "End UTC Minute", "minute", From + 66, 6)
        Call DetailOut(Retvalcol, "Start Longitude", "lon", From + 72, 28)
        Call DetailOut(Retvalcol, "Start Latitude", "lat", From + 100, 27)
        Call DetailOut(Retvalcol, "End Longitude", "lon", From + 127, 28)
        Call DetailOut(Retvalcol, "End Latitude", "lat", From + 155, 27)
        Call Dac200Out(Retvalcol, "Type", "weather", From + 182, 4)
        Call Dac200Out(Retvalcol, "Min Value", "negative", From + 186, 9, 0)
        Call Dac200Out(Retvalcol, "Max Value", "negative", From + 195, 9, 0)
        Call Dac200Out(Retvalcol, "Classification", "category", From + 204, 3)
        Call Dac200Out(Retvalcol, "Wind Direction", "direction", From + 206, 4)
        Call DetailOut(Retvalcol, "Spare", "spare", From + 210, 6)
    Case Is = "24"
        Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, "Water level", "Awaiting Verification")
        Call DetailOut(Retvalcol, "UN Country Code", "text", From + 16, 12)
        Last = From + 28 - 1
        Do Until Last >= DataBits
            i = i + 1
            Call Dac200Out(Retvalcol, "Gauge (" & i & ")ID", "number", Last + 1, 11)
            Call Dac200Out(Retvalcol, "Water Level", "negative", Last + 12, 14, 2, " Meters")
            Last = Last + 25
         Loop
     Case Is = "40"
        Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, "Signal status", "Awaiting Verification")
        Call DetailOut(Retvalcol, "Signal Position Longitude", "lon", From + 16, 28)
        Call DetailOut(Retvalcol, "Signal Position Latitude", "lat", From + 44, 27)
        Call Dac200Out(Retvalcol, "Signal Form", "signalform", From + 71, 4)
        Call DetailOut(Retvalcol, "Orientation of Signal", "heading", From + 75, 9, 0)
        Call Dac200Out(Retvalcol, "Direction of Impact", "signalimpact", From + 84, 3)
        Last = From + 87 - 1
'not atlough there is provision fo 10 * 3 lights it appears only a max of 9 are used
        Do Until Last >= DataBits
            i = i + 1
            Call Dac200Out(Retvalcol, "Light (" & i & ") Status", "signalstatus", Last + 1, 3)
            Last = Last + 3
         Loop
        Call DetailOut(Retvalcol, "Spare", "spare", From + 117, 11)
    Case Is = "55"
        Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, "No of persons on board")
        Call Dac200Out(Retvalcol, "No of Crew Members", "number", From + 16, 8, "", " persons")
        Call Dac200Out(Retvalcol, "No of Passengers", "number", From + 24, 13, "", " persons")
        Call Dac200Out(Retvalcol, "No of Shipboard Personnel", "number", From + 37, 8, "", " persons")
        Call DetailOut(Retvalcol, "Spare", "spare", From + 45, 51)
    Case Else
        Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, , "Information required")
    End Select   'inland fi's
Case Is = "210"     'Cyprus
    Select Case Fi
    Case Is = "0"
        Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, "Text Telegram")
'v136        Call DetailOut(Retvalcol, "Text", "text", From + 16, 906)
      Call DetailOut(Retvalcol, "Text", "text", From + 16, clsSentence.AisPayloadBits - (From + 16) + 1) 'v136
    Case Else
        Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, , "Information required")
    End Select
Case Is = "232" 'uk (PLA)
    Select Case Fi
    Case Is = "1"
        Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, "Used by Port of London Authority (PLA)", "Not decoded")
    Case Else
        Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, , "Information required")
    End Select
Case Is = "235", "250"  'uk,ROI
    Select Case Fi
    Case Is = "10"
        Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, "AtoN Monitoring")
        Call Dac235Out(Retvalcol, "Analogue (internal)", "volts", From + 16, 10, 2)
        Call Dac235Out(Retvalcol, "Analogue (external No 1)", "volts", From + 26, 10, 2)
        Call Dac235Out(Retvalcol, "Analogue (external No 2)", "volts", From + 36, 10, 2)
        Call Dac235Out(Retvalcol, "Racon Status", "racon", From + 46, 2)
        Call Dac235Out(Retvalcol, "Light Status", "light", From + 48, 2)
        Call Dac235Out(Retvalcol, "Alarm", "alarm", From + 50, 1)
        Last = From + 51 - 1
        i = 7
        Do Until i < 0
            Call DetailOut(Retvalcol, "Digital Input " & i, "onoff", Last + 1, 1)
            i = i - 1
            Last = Last + 1
        Loop
        Call DetailOut(Retvalcol, "Off Position Status", "virtual_aid", From + 59, 1)
        Call DetailOut(Retvalcol, "Spare", "spare", From + 60, 4)
    Case Is = "15"
        Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, "Analogue & Digital I/O status")
        Select Case clsSentence.AisMsgFiId
        Case "48"
            Call DetailOut(Retvalcol, "Payload bits", "bits", From, 88, "from payload bit " & From)
            Call DetailOut(Retvalcol, "Function ID", "fiid", From + 16, 8, "Sub App")
            Call Dac235Out(Retvalcol, "Message ID", "msgid", From + 24, 8, "AtoN - I/O status message")
            Call Dac235Out(Retvalcol, "Unit ID", "unitid", From + 32, 6)
            Call Dac235Out(Retvalcol, "BIIT", "biit", From + 38, 1)
            Call Dac235Out(Retvalcol, "Extended Fields", "extfld", From + 39, 1)
            Call Dac235Out(Retvalcol, "Digital Input 1", "highlow", From + 40, 1)
            Call Dac235Out(Retvalcol, "Digital Input 2", "highlow", From + 41, 1)
            Call Dac235Out(Retvalcol, "Digital Input 3", "highlow", From + 42, 1)
            Call Dac235Out(Retvalcol, "Digital Input 4", "highlow", From + 43, 1)
            Call Dac235Out(Retvalcol, "Digital Output 1", "highlow", From + 44, 1)
            Call Dac235Out(Retvalcol, "Digital Output 2", "highlow", From + 45, 1)
            Call Dac235Out(Retvalcol, "Digital Output 3", "highlow", From + 46, 1)
            Call Dac235Out(Retvalcol, "Digital Output 4", "highlow", From + 47, 1)
            Call Dac235Out(Retvalcol, "Analogue Channel 1", "volts", From + 48, 10, 2)
            Call Dac235Out(Retvalcol, "Analogue Channel 2", "volts", From + 58, 10, 2)
            Call Dac235Out(Retvalcol, "Analogue Channel 3", "volts", From + 68, 10, 2)
            Call Dac235Out(Retvalcol, "Analogue Channel 4", "volts", From + 78, 10, 2)
            Call Dac235Out(Retvalcol, "Content-Control", "contentcontrol", From + 88, 8)
        Case Else
            Call DetailOut(Retvalcol, "Message ID", "fiid", From + 18, 6, "Information required")
        End Select
    Case Else
        Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, , "Information required")
    End Select
Case Is = "248"     'Malta
    Select Case Fi
    Case Is = "0"
        Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, "Text Telegram")
'v136        Call DetailOut(Retvalcol, "Text", "text", From + 16, 906)
      Call DetailOut(Retvalcol, "Text", "text", From + 16, clsSentence.AisPayloadBits - (From + 16) + 1) 'v136
    Case Else
        Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, , "Information required")
    End Select
Case Is = "316"
    Select Case Fi
    Case Is = "1", "2", "32"
        Call FiUsaCanada(Retvalcol, From, DataBits)
    Case Else
        Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, , "Information required")
    End Select
Case Is = "353"
    Select Case Fi
    Case Is = "0"
        Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, "Text Telegram")
'v136        Call DetailOut(Retvalcol, "Text", "text", From + 16, 906)
      Call DetailOut(Retvalcol, "Text", "text", From + 16, clsSentence.AisPayloadBits - (From + 16) + 1) 'v136
    Case Else
        Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, , "Information required")
    End Select
Case Is = "366"
    Select Case Fi
    Case Is = "1", "2", "32"
        Call FiUsaCanada(Retvalcol, From, DataBits)
    Case Is = "33"
        Call Environmental(Retvalcol, From, DataBits)
    Case Is = "34"
        Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, "Whale Notice ?", "Not decoded")
    Case Is = "56", "57"
        Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, "USCG {encrypted}", "No details of this FI")
    Case Is = "63"
        Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, "Water level ?")
    Case Else
        Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, , "No details of this FI")
    End Select  'usa fi's
Case Is = "367"
    Select Case Fi
    Case Is = "33"
        Call Environmental367(Retvalcol, From, DataBits)
    Case Else
        Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, , "No details of this FI")
    End Select
Case Else

'        Call DetailOut(Retvalcol, "DAC", "dac", From, 10, , "No details of FI's for this DAC")
        Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, , "No details of this ASM")
        If DataBits > 16 Then
            Call DetailOut(Retvalcol, "Data", "data", From + 16, DataBits - 16)
        End If
1
End Select  'dac

End Function

Function Environmental(Retvalcol As Long, From As Long, DataBits As Long)
Dim Last As Long
Dim i As Long
Dim ReportType As Integer
Dim Precision As String

Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, "Environmental")
Last = From + 16 - 1
Do Until Last >= DataBits
    i = i + 1
'    Call Rowout("Sensor no " & i, "", "", True)
'    Call DetailLineOut("Sensor no " & i, "", "", True)
    Call DetailOut(Retvalcol, "Sensor Report No", "literal", , , CStr(i))
    ReportType = Dac366Out(Retvalcol, "Report Type", "reporttype", Last + 1, 4)
    Call DetailOut(Retvalcol, "UTC Day", "day", Last + 5, 5)
    Call DetailOut(Retvalcol, "UTC Hour", "hour", Last + 10, 5)
    Call DetailOut(Retvalcol, "UTC Minute", "minute", Last + 15, 6)
    Call DetailOut(Retvalcol, "Site ID", "number", Last + 21, 7)
    Select Case ReportType
    Case Is = 0 'Location
'The field is after the field where we wish to use it
'so we need it before, to keep the detail display order the same as the field order
        Precision = pLong(Last + 83, 3)
        Call DetailOut(Retvalcol, "Longitude", "lon", Last + 28, 28)
        Call DetailOut(Retvalcol, "Latitude", "lat", Last + 56, 27)
        Call DetailOut(Retvalcol, "Precision", "number", Last + 83, 3)
        Call DetailOut(Retvalcol, "Altitude", "meters", Last + 86, 11, 1, "2001")
        Call Dac366Out(Retvalcol, "Sensor Owner", "owner", Last + 97, 4)
        Call Dac366Out(Retvalcol, "Data Timeout", "timeout", Last + 101, 3)
        Call DetailOut(Retvalcol, "Spare", "spare", Last + 104, 9)
    Case Is = 1 'Station ID
        Call DetailOut(Retvalcol, "Name", "text", Last + 28, 84)
        Call DetailOut(Retvalcol, "Spare", "spare", Last + 112, 1)
    Case Is = 2 'wind
        Call Dac366Out(Retvalcol, "Average Wind Speed", "windspeed", Last + 28, 7, 0, "Average for last 10 minutes") '1Kt
        Call Dac366Out(Retvalcol, "Wind Gust", "windspeed", Last + 35, 7, 0, "Average for last 10 minutes") '1Kt
        Call Dac366Out(Retvalcol, "Wind Direction", "winddirection", Last + 42, 9, 0)
        Call Dac366Out(Retvalcol, "Wind Gust Direction", "winddirection", Last + 51, 9, 0)
        Call Dac366Out(Retvalcol, "Sensor Data Description", "sensordata", Last + 60, 3)
        Call Dac366Out(Retvalcol, "Forecast Wind Speed", "windspeed", Last + 63, 7, 0) '1Kt
        Call Dac366Out(Retvalcol, "Forecast Gust", "windspeed", Last + 70, 7, 0) '1Kt
        Call Dac366Out(Retvalcol, "Forecast Direction", "winddirection", Last + 77, 9, 0)
        Call DetailOut(Retvalcol, "Valid Day of Forecast", "day", Last + 86, 5)
        Call DetailOut(Retvalcol, "Valid Hour of Forecast", "hour", Last + 91, 5)
        Call DetailOut(Retvalcol, "Valid Minute of Forecast", "minute", Last + 96, 6)
        Call Dac366Out(Retvalcol, "Duration of Forecast", "duration", Last + 102, 8)
        Call DetailOut(Retvalcol, "Spare", "spare", Last + 110, 3)
    Case Is = 3     'water level
        Call Dac366Out(Retvalcol, "Water Level Type", "level", Last + 28, 1)
        Call DetailOut(Retvalcol, "Water Level", "simeters", Last + 29, 16, 2)
        Call DetailOut(Retvalcol, "Trend", "tendency", Last + 45, 2)
        Call Dac366Out(Retvalcol, "Reference Datum", "datum", Last + 47, 5)
        Call Dac366Out(Retvalcol, "Sensor Data Description", "sensordata", Last + 52, 3)
        Call Dac366Out(Retvalcol, "Forecast Water Level Type", "level", Last + 55, 1)
        Call DetailOut(Retvalcol, "Forecast Water Level", "simeters", Last + 56, 16, 2)
        Call DetailOut(Retvalcol, "Valid Day of Forecast", "day", Last + 72, 5)
        Call DetailOut(Retvalcol, "Valid Hour of Forecast", "hour", Last + 77, 5)
        Call DetailOut(Retvalcol, "Valid Minute of Forecast", "minute", Last + 82, 6)
        Call Dac366Out(Retvalcol, "Duration", "duration", Last + 88, 8)
        Call DetailOut(Retvalcol, "Spare", "spare", Last + 96, 17)
    Case Is = 4      'vertical current profile (2D) not verified
        Call Dac366Out(Retvalcol, "Current speed, #1", "currentspeed", Last + 28, 8, 1)
        Call Dac366Out(Retvalcol, "Current direction, #1", "currentdirection", Last + 36, 9, 0)
        Call Dac366Out(Retvalcol, "Current measuring level, #1", "currentlevel", Last + 45, 9, 0)
        Call Dac366Out(Retvalcol, "Current speed, #2", "currentspeed", Last + 54, 8, 1)
        Call Dac366Out(Retvalcol, "Current direction, #2", "currentdirection", Last + 62, 9, 0)
        Call Dac366Out(Retvalcol, "Current measuring level, #2", "currentlevel", Last + 71, 9, 0)
        Call Dac366Out(Retvalcol, "Current speed, #3", "currentspeed", Last + 80, 8, 1)
        Call Dac366Out(Retvalcol, "Current direction, #3", "currentdirection", Last + 88, 9, 0)
        Call Dac366Out(Retvalcol, "Current measuring level, #3", "currentlevel", Last + 97, 9, 0)
        Call Dac366Out(Retvalcol, "Sensor Data Description", "sensordata", Last + 106, 3)
        Call DetailOut(Retvalcol, "Spare", "spare", Last + 109, 4)
    Case Is = 5      'vertical current profile (3D) not verified
        Call Dac366Out(Retvalcol, "Current Vector North, #1", "currentvector", Last + 28, 9, 1)
        Call Dac366Out(Retvalcol, "Current Vector East, #1", "currentvector", Last + 37, 9, 1)
        Call Dac366Out(Retvalcol, "Current Vector Up, #1", "currentvector", Last + 46, 9, 1)
        Call Dac366Out(Retvalcol, "Current measuring level, #1", "currentlevel", Last + 55, 9, 0)
        Call Dac366Out(Retvalcol, "Current Vector North, #2", "currentvector", Last + 64, 9, 1)
        Call Dac366Out(Retvalcol, "Current Vector East, #2", "currentvector", Last + 73, 9, 1)
        Call Dac366Out(Retvalcol, "Current Vector Up, #2", "currentvector", Last + 82, 9, 1)
        Call Dac366Out(Retvalcol, "Current measuring level, #2", "currentlevel", Last + 91, 9, 0)
        Call Dac366Out(Retvalcol, "Sensor Data Description", "sensordata", Last + 100, 3)
        Call DetailOut(Retvalcol, "Spare", "spare", Last + 103, 10)
    Case Is = 6      'horizontal current profile not verified
        Call Dac366Out(Retvalcol, "Current Reading Bearing", "currentdirection", Last + 28, 9, 0)
        Call Dac366Out(Retvalcol, "Vertical Reference Datum", "datum", Last + 37, 5, 0)
        Call Dac366Out(Retvalcol, "Current Reading #1 Distance", "currentlevel", Last + 42, 9, 0)
        Call Dac366Out(Retvalcol, "Current #1 Speed", "currentspeed", Last + 51, 8, 1)
        Call Dac366Out(Retvalcol, "Current #1 Direction", "currentdirection", Last + 59, 9, 0)
        Call Dac366Out(Retvalcol, "Current #1 Measuring Level", "currentlevel", Last + 68, 9, 0)
        Call Dac366Out(Retvalcol, "Current Reading #2 Distance", "currentlevel", Last + 77, 9, 0)
        Call Dac366Out(Retvalcol, "Current #2 Speed", "currentspeed", Last + 85, 8, 1)
        Call Dac366Out(Retvalcol, "Current #2 Direction", "currentdirection", Last + 94, 9, 0)
        Call Dac366Out(Retvalcol, "Current #2 Measuring Level", "currentlevel", Last + 103, 9, 0)
        Call DetailOut(Retvalcol, "Spare", "spare", Last + 112, 1)
    Case Is = 7      'sea state not verified
        Call Dac366Out(Retvalcol, "Swell height", "waveheight", Last + 28, 8, 1)
        Call DetailOut(Retvalcol, "Swell period", "second", Last + 36, 6)
        Call Dac366Out(Retvalcol, "Swell direction", "currentdirection", Last + 42, 9, 0)
        Call DetailOut(Retvalcol, "Sea State", "beaufortenvironmental", Last + 51, 4)
        Call Dac366Out(Retvalcol, "Sensor Data Description", "sensordata", Last + 55, 3, , "Swells")
        Call Dac366Out(Retvalcol, "Water Temperature", "temperature", Last + 58, 10, 1, "water")
        Call Dac366Out(Retvalcol, "Water Temperature Depth", "depth", Last + 68, 7, 1)
        Call Dac366Out(Retvalcol, "Sensor Data Description", "sensordata", Last + 75, 3, , "Temperature")
        Call Dac366Out(Retvalcol, "Significant wave height", "waveheight", Last + 78, 8, 1)
        Call DetailOut(Retvalcol, "Wave period", "second", Last + 86, 6)
        Call Dac366Out(Retvalcol, "Wave Direction", "currentdirection", Last + 92, 9, 0)
        Call Dac366Out(Retvalcol, "Sensor Data Description", "sensordata", Last + 101, 3, , "Waves")
        Call Dac366Out(Retvalcol, "Salinity", "salinity", Last + 104, 9, 1)

    Case Is = 8      'salinity not verified
        Call Dac366Out(Retvalcol, "Water Temperature", "temperature", Last + 28, 10, 1, "water")
        Call Dac366Out(Retvalcol, "Conductivity", "conductivity", Last + 38, 10, 2)
        Call Dac366Out(Retvalcol, "Water Pressure", "decibars", Last + 48, 16, 1)
        Call DetailOut(Retvalcol, "Salinity", "salinity", Last + 64, 9, 1)
        Call Dac366Out(Retvalcol, "Salinity Type", "salinitytype", Last + 73, 2)
        Call Dac366Out(Retvalcol, "Sensor Data Description", "sensordata", Last + 75, 3)
        Call DetailOut(Retvalcol, "Spare", "spare", Last + 78, 35)
    
    Case Is = 9 'weather      'not verified
        Call Dac366Out(Retvalcol, "Air temperature", "temperature", Last + 28, 11, 1, "air")
        Call Dac366Out(Retvalcol, "Sensor Data Description", "sensordata", Last + 39, 3)
        Call DetailOut(Retvalcol, "Precipitation type", "precipitation", Last + 42, 2)
'not MSB 366-33-9
        Call Dac366Out(Retvalcol, "Horizontal visibility", "visibility", Last + 44, 8, 1)
        Call Dac366Out(Retvalcol, "Dew Point", "temperature", Last + 52, 10, 1, "dew")
        Call Dac366Out(Retvalcol, "Sensor Data Description", "sensordata", Last + 62, 3)
        Call Dac366Out(Retvalcol, "Air Pressure", "pressure", Last + 65, 9)
        Call DetailOut(Retvalcol, "Air Pressure tendency", "tendency", Last + 74, 2)
        Call Dac366Out(Retvalcol, "Sensor Data Description", "sensordata", Last + 76, 3)
        Call Dac366Out(Retvalcol, "Salinity", "salinity", Last + 79, 9)
        Call DetailOut(Retvalcol, "Spare", "spare", Last + 88, 25)
    Case Is = 10      'air gap/air draught not verified
        Call DetailOut(Retvalcol, "Air Draught", "airdraught", Last + 28, 13, 1)
        Call DetailOut(Retvalcol, "Air Gap", "airdraught", Last + 41, 13, 1)
        Call DetailOut(Retvalcol, "Air Gap Trend", "tendency", Last + 54, 2)
        Call DetailOut(Retvalcol, "Forecast Air Gap", "airdraught", Last + 56, 13, 1)
        Call DetailOut(Retvalcol, "Valid UTC Day", "day", Last + 69, 5)
        Call DetailOut(Retvalcol, "Valid UTC Hour", "hour", Last + 74, 5)
        Call DetailOut(Retvalcol, "Valid UTC Minute", "minute", Last + 79, 6)
        Call DetailOut(Retvalcol, "Spare", "spare", Last + 85, 28)
    Case Else   '11-15
'        Call Rowout("", "", "Unknown Report Type " & ReportType)
'        Call DetailLineOut("", "", "Unknown Report Type " & ReportType)
        Call DetailOut(Retvalcol, "Unknown Report Type", "literal", , , CStr(ReportType))
    End Select
    Last = Last + 112 '112 is size of repeated section
Loop

End Function

'Replaces Environmental from 13Aug2012
Function Environmental367(Retvalcol As Long, From As Long, DataBits As Long)
Dim Last As Long
Dim i As Long
Dim ReportType As Integer
Dim Precision As String
Dim Version As Long

Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, "Environmental")
Last = From + 16 - 1
Do Until Last >= DataBits
    i = i + 1
'    Call Rowout("Sensor no " & i, "", "", True)
'    Call DetailLineOut("Sensor no " & i, "", "", True)
    Call DetailOut(Retvalcol, "Sensor Report No", "literal", , , CStr(i))
    ReportType = Dac366Out(Retvalcol, "Report Type", "reporttype", Last + 1, 4)
    Call DetailOut(Retvalcol, "UTC Day", "day", Last + 5, 5)
    Call DetailOut(Retvalcol, "UTC Hour", "hour", Last + 10, 5)
    Call DetailOut(Retvalcol, "UTC Minute", "minute", Last + 15, 6)
    Call DetailOut(Retvalcol, "Site ID", "number", Last + 21, 7)
    Select Case ReportType
    Case Is = 0 'Location
        Version = DetailOut(Retvalcol, "Version", "number", Last + 28, 6)
'The field is after the field where we wish to use it
'so we need it before, to keep the detail display order the same as the field order
        Precision = pLong(Last + 89, 3)
        Call DetailOut(Retvalcol, "Longitude", "lon", Last + 34, 28)
        Call DetailOut(Retvalcol, "Latitude", "lat", Last + 62, 27)
        Call DetailOut(Retvalcol, "Precision", "number", Last + 89, 3)
        Call Dac366Out(Retvalcol, "Altitude", "sialtitude", Last + 92, 12, 1)
        Call Dac366Out(Retvalcol, "Sensor Owner", "owner", Last + 104, 4)
        Call Dac366Out(Retvalcol, "Data Timeout", "timeout", Last + 108, 3)
        Call DetailOut(Retvalcol, "Spare", "spare", Last + 111, 2)
    Case Is = 1 'Station ID
        Call DetailOut(Retvalcol, "Name", "text", Last + 28, 84)
        Call DetailOut(Retvalcol, "Spare", "spare", Last + 112, 1)
    Case Is = 2 'wind
        Call Dac366Out(Retvalcol, "Average Wind Speed", "windspeed", Last + 28, 7, 0, "Average for last 10 minutes") '1Kt
        Call Dac366Out(Retvalcol, "Wind Gust", "windspeed", Last + 35, 7, 0, "Average for last 10 minutes") '1Kt
        Call Dac366Out(Retvalcol, "Wind Direction", "winddirection", Last + 42, 9, 0)
        Call Dac366Out(Retvalcol, "Wind Gust Direction", "winddirection", Last + 51, 9, 0)
        Call Dac366Out(Retvalcol, "Sensor Data Description", "sensordata", Last + 60, 3)
        Call Dac366Out(Retvalcol, "Forecast Wind Speed", "windspeed", Last + 63, 7, 0) '1Kt
        Call Dac366Out(Retvalcol, "Forecast Gust", "windspeed", Last + 70, 7, 0) '1Kt
        Call Dac366Out(Retvalcol, "Forecast Direction", "winddirection", Last + 77, 9, 0)
        Call DetailOut(Retvalcol, "Valid Day of Forecast", "day", Last + 86, 5)
        Call DetailOut(Retvalcol, "Valid Hour of Forecast", "hour", Last + 91, 5)
        Call DetailOut(Retvalcol, "Valid Minute of Forecast", "minute", Last + 96, 6)
        Call Dac366Out(Retvalcol, "Duration of Forecast", "duration", Last + 102, 8)
        Call DetailOut(Retvalcol, "Spare", "spare", Last + 110, 3)
    Case Is = 3     'water level
        Call Dac366Out(Retvalcol, "Water Level Type", "level", Last + 28, 1)
        Call DetailOut(Retvalcol, "Water Level", "simeters", Last + 29, 16, 2)
        Call DetailOut(Retvalcol, "Trend", "tendency", Last + 45, 2)
        Call Dac366Out(Retvalcol, "Reference Datum", "datum", Last + 47, 5)
        Call Dac366Out(Retvalcol, "Sensor Data Description", "sensordata", Last + 52, 3)
        Call Dac366Out(Retvalcol, "Forecast Water Level Type", "level", Last + 55, 1)
        Call DetailOut(Retvalcol, "Forecast Water Level", "simeters", Last + 56, 16, 2)
        Call DetailOut(Retvalcol, "Valid Day of Forecast", "day", Last + 72, 5)
        Call DetailOut(Retvalcol, "Valid Hour of Forecast", "hour", Last + 77, 5)
        Call DetailOut(Retvalcol, "Valid Minute of Forecast", "minute", Last + 82, 6)
        Call Dac366Out(Retvalcol, "Duration", "duration", Last + 88, 8)
        Call DetailOut(Retvalcol, "Spare", "spare", Last + 96, 17)
    Case Is = 4      'vertical current profile (2D) not verified
        Call Dac366Out(Retvalcol, "Current speed, #1", "currentspeed", Last + 28, 8, 1)
        Call Dac366Out(Retvalcol, "Current direction, #1", "currentdirection", Last + 36, 9, 0)
        Call Dac366Out(Retvalcol, "Current measuring level, #1", "currentlevel", Last + 45, 9, 0)
        Call Dac366Out(Retvalcol, "Current speed, #2", "currentspeed", Last + 54, 8, 1)
        Call Dac366Out(Retvalcol, "Current direction, #2", "currentdirection", Last + 62, 9, 0)
        Call Dac366Out(Retvalcol, "Current measuring level, #2", "currentlevel", Last + 71, 9, 0)
        Call Dac366Out(Retvalcol, "Current speed, #3", "currentspeed", Last + 80, 8, 1)
        Call Dac366Out(Retvalcol, "Current direction, #3", "currentdirection", Last + 88, 9, 0)
        Call Dac366Out(Retvalcol, "Current measuring level, #3", "currentlevel", Last + 97, 9, 0)
        Call Dac366Out(Retvalcol, "Sensor Data Description", "sensordata", Last + 106, 3)
        Call DetailOut(Retvalcol, "Spare", "spare", Last + 109, 4)
    Case Is = 5      'vertical current profile (3D) not verified
        Call Dac366Out(Retvalcol, "Current Vector North, #1", "currentvector", Last + 28, 9, 1)
        Call Dac366Out(Retvalcol, "Current Vector East, #1", "currentvector", Last + 37, 9, 1)
        Call Dac366Out(Retvalcol, "Current Vector Up, #1", "currentvector", Last + 46, 9, 1)
        Call Dac366Out(Retvalcol, "Current measuring level, #1", "currentlevel", Last + 55, 9, 0)
        Call Dac366Out(Retvalcol, "Current Vector North, #2", "currentvector", Last + 64, 9, 1)
        Call Dac366Out(Retvalcol, "Current Vector East, #2", "currentvector", Last + 73, 9, 1)
        Call Dac366Out(Retvalcol, "Current Vector Up, #2", "currentvector", Last + 82, 9, 1)
        Call Dac366Out(Retvalcol, "Current measuring level, #2", "currentlevel", Last + 91, 9, 0)
        Call Dac366Out(Retvalcol, "Sensor Data Description", "sensordata", Last + 100, 3)
        Call DetailOut(Retvalcol, "Spare", "spare", Last + 103, 10)
    Case Is = 6      'horizontal current profile not verified
        Call Dac366Out(Retvalcol, "Current Reading Bearing", "currentdirection", Last + 28, 9, 0)
'        Call Dac366Out(RetValCol, "Vertical Reference Datum", "datum", Last + 37, 5, 0)
        Call Dac366Out(Retvalcol, "Current Reading #1 Distance", "currentlevel", Last + 37, 9, 0)
        Call Dac366Out(Retvalcol, "Current #1 Speed", "currentspeed", Last + 46, 8, 1)
        Call Dac366Out(Retvalcol, "Current #1 Direction", "currentdirection", Last + 54, 9, 0)
        Call Dac366Out(Retvalcol, "Current #1 Measuring Level", "currentlevel", Last + 63, 9, 0)
        Call Dac366Out(Retvalcol, "Current Reading #2 Distance", "currentlevel", Last + 72, 9, 0)
        Call Dac366Out(Retvalcol, "Current #2 Speed", "currentspeed", Last + 81, 8, 1)
        Call Dac366Out(Retvalcol, "Current #2 Direction", "currentdirection", Last + 89, 9, 0)
        Call Dac366Out(Retvalcol, "Current #2 Measuring Level", "currentlevel", Last + 98, 9, 0)
        Call DetailOut(Retvalcol, "Spare", "spare", Last + 107, 6)
    Case Is = 7      'sea state not verified
        Call Dac366Out(Retvalcol, "Swell height", "waveheight", Last + 28, 8, 1)
        Call DetailOut(Retvalcol, "Swell period", "second", Last + 36, 6)
        Call Dac366Out(Retvalcol, "Swell direction", "currentdirection", Last + 42, 9, 0)
        Call DetailOut(Retvalcol, "Sea State", "beaufortenvironmental", Last + 51, 4)
        Call Dac366Out(Retvalcol, "Sensor Data Description", "sensordata", Last + 55, 3, , "Swells")
        Call Dac366Out(Retvalcol, "Water Temperature", "temperature", Last + 58, 10, 1, "water")
        Call Dac366Out(Retvalcol, "Water Temperature Depth", "depth", Last + 68, 7, 1)
        Call Dac366Out(Retvalcol, "Sensor Data Description", "sensordata", Last + 75, 3, , "Temperature")
        Call Dac366Out(Retvalcol, "Significant wave height", "waveheight", Last + 78, 8, 1)
        Call DetailOut(Retvalcol, "Wave period", "second", Last + 86, 6)
        Call Dac366Out(Retvalcol, "Wave Direction", "currentdirection", Last + 92, 9, 0)
        Call Dac366Out(Retvalcol, "Sensor Data Description", "sensordata", Last + 101, 3, , "Waves")
        Call Dac366Out(Retvalcol, "Salinity", "salinity", Last + 104, 9, 1)

    Case Is = 8      'salinity not verified
        Call Dac366Out(Retvalcol, "Water Temperature", "temperature", Last + 28, 10, 1, "water")
        Call Dac366Out(Retvalcol, "Conductivity", "conductivity", Last + 38, 10, 2)
        Call Dac366Out(Retvalcol, "Water Pressure", "decibars", Last + 48, 16, 1)
        Call DetailOut(Retvalcol, "Salinity", "salinity", Last + 64, 9, 1)
        Call Dac366Out(Retvalcol, "Salinity Type", "salinitytype", Last + 73, 2)
        Call Dac366Out(Retvalcol, "Sensor Data Description", "sensordata", Last + 75, 3)
        Call DetailOut(Retvalcol, "Spare", "spare", Last + 78, 35)
    
    Case Is = 9 'weather      'not verified
        Call Dac366Out(Retvalcol, "Air temperature", "temperature", Last + 28, 11, 1, "air")
        Call Dac366Out(Retvalcol, "Sensor Data Description", "sensordata", Last + 39, 3)
        Call DetailOut(Retvalcol, "Precipitation type", "precipitation", Last + 42, 2)
'not MSB 366-33-9
        Call Dac366Out(Retvalcol, "Horizontal visibility", "visibility", Last + 44, 8, 1)
        Call Dac366Out(Retvalcol, "Dew Point", "temperature", Last + 52, 10, 1, "dew")
        Call Dac366Out(Retvalcol, "Sensor Data Description", "sensordata", Last + 62, 3)
        Call Dac366Out(Retvalcol, "Air Pressure", "pressure", Last + 65, 9)
        Call DetailOut(Retvalcol, "Air Pressure tendency", "tendency", Last + 74, 2)
        Call Dac366Out(Retvalcol, "Sensor Data Description", "sensordata", Last + 76, 3)
        Call Dac366Out(Retvalcol, "Salinity", "salinity", Last + 79, 9)
        Call DetailOut(Retvalcol, "Spare", "spare", Last + 88, 25)
    Case Is = 10      'air gap/air draught not verified
        Call DetailOut(Retvalcol, "Air Draught", "airdraught", Last + 28, 13, 2)    'feb17 srt cms
        Call DetailOut(Retvalcol, "Air Gap", "airdraught", Last + 41, 13, 2)    'feb17 srt cms
        Call DetailOut(Retvalcol, "Air Gap Trend", "tendency", Last + 54, 2)
        Call DetailOut(Retvalcol, "Predicted Air Gap", "airdraught", Last + 56, 13, 2)    'feb17 srt cms
        Call DetailOut(Retvalcol, "Predicted UTC Day", "day", Last + 69, 5)
        Call DetailOut(Retvalcol, "Predicted UTC Hour", "hour", Last + 74, 5)
        Call DetailOut(Retvalcol, "Predicted UTC Minute", "minute", Last + 79, 6)
        Call Dac366Out(Retvalcol, "Sensor Data Description", "sensordata", Last + 85, 3)
        Call DetailOut(Retvalcol, "Spare", "spare", Last + 88, 25)
    Case Else   '11-15
'        Call Rowout("", "", "Unknown Report Type " & ReportType)
'        Call DetailLineOut("", "", "Unknown Report Type " & ReportType)
        Call DetailOut(Retvalcol, "Unknown Report Type", "literal", , , CStr(ReportType))
    End Select
    Last = Last + 112 '112 is size of repeated section
Loop

End Function

Function FiUsaCanada(Retvalcol As Long, From As Long, DataBits As Long)
'Dim Dac As String
'Dim Fi As String
'Dim UsaCanId As String
Dim Last  As Integer 'repeated sections
Dim i As Integer
Dim TargType As Integer

'Dac = pLong(from, 10)
'Dac = clsSentence.AisMsgDac
'Fi = clsSentence.AisMsgFi
Call DetailOut(Retvalcol, "FI", "usacanfi", From + 10, 6)
Call DetailOut(Retvalcol, "Spare", "spare", From + 16, 2)
'UsaCanId = DetailOut(retvalcol,"Message ID", "usacanid", from + 18, 6, Fi)

Select Case clsSentence.AisMsgFi
Case Is = "1"
    Select Case clsSentence.AisMsgFiId
    Case "1"
        Call DetailOut(Retvalcol, "Message ID", "fiid", From + 18, 6, "Weather Station Message")
        Last = From + 24 - 1
        Do Until Last >= DataBits
            i = i + 1
'            Call Rowout("Station " & i, "", "", True)
'            Call DetailLineOut("Station " & i, "", "", True)
            Call DetailOut(Retvalcol, "Station", "literal", , , CStr(i))
            Call DetailOut(Retvalcol, "Month", "month", Last + 1, 4)
            Call DetailOut(Retvalcol, "UTC Day", "day", Last + 5, 5)
            Call DetailOut(Retvalcol, "UTC Hour", "hour", Last + 10, 5)
            Call DetailOut(Retvalcol, "UTC Minute", "minute", Last + 15, 6)
            Call DetailOut(Retvalcol, "Sensor ID", "text", Last + 21, 42)
            Call DetailOut(Retvalcol, "Longitude", "lon", Last + 63, 25)
            Call DetailOut(Retvalcol, "Latitude", "lat", Last + 88, 24)
            Call DetailOut(Retvalcol, "Wind Speed", "speed", Last + 112, 10, 1)
            Call DetailOut(Retvalcol, "Wind Gust", "speed", Last + 122, 10, 1)
            Call DetailOut(Retvalcol, "Wind Direction", "heading", Last + 132, 9, 0)
            Call Dac366Out(Retvalcol, "Atmospheric Pressure", "millibars", Last + 141, 14)
'316/336-1-1 (St Lawrence)
            Call DetailOut(Retvalcol, "Air Temperature", "temperature", Last + 155, 10, 1)
            Call DetailOut(Retvalcol, "Dew Point", "temperature", Last + 165, 10, 1)
            Call DetailOut(Retvalcol, "Visibility", "meters", Last + 175, 8, 1, "Kilometers")
            Call DetailOut(Retvalcol, "Water Temperature", "temperature", Last + 183, 10, 1)
            Last = Last + 192 'size of repeated section
        Loop
    Case "2"
        Call DetailOut(Retvalcol, "Message ID", "fiid", From + 18, 6, "Wind Information Message")
        Last = From + 24 - 1
        Do Until Last >= DataBits
            i = i + 1
'            Call Rowout("Sensor " & i, "", "", True)
'            Call DetailLineOut("Sensor " & i, "", "", True)
            Call DetailOut(Retvalcol, "Sensor Report No", "literal", , , CStr(i))
            Call DetailOut(Retvalcol, "Month", "month", Last + 1, 4)
            Call DetailOut(Retvalcol, "UTC Day", "day", Last + 5, 5)
            Call DetailOut(Retvalcol, "UTC Hour", "hour", Last + 10, 5)
            Call DetailOut(Retvalcol, "UTC Minute", "minute", Last + 15, 6)
            Call DetailOut(Retvalcol, "Sensor ID", "text", Last + 21, 42)
            Call DetailOut(Retvalcol, "Longitude", "lon", Last + 63, 25)
            Call DetailOut(Retvalcol, "Latitude", "lat", Last + 88, 24)
            Call DetailOut(Retvalcol, "Wind Speed", "speed", Last + 112, 10, 1)
            Call DetailOut(Retvalcol, "Wind Gust", "speed", Last + 122, 10, 1)
            Call Dac366Out(Retvalcol, "Wind Direction", "compasspoint", Last + 132, 9)
            Call DetailOut(Retvalcol, "Spare", "spare", Last + 141, 4)
            Last = Last + 144 'size of repeated section
        Loop
    Case "3"
        Call DetailOut(Retvalcol, "Message ID", "fiid", From + 18, 6, "Water Level Message")
        Last = From + 24 - 1
        Do Until Last >= DataBits
            i = i + 1
            Call DetailOut(Retvalcol, "Sensor Report No", "literal", , , CStr(i))
            Call DetailOut(Retvalcol, "Month", "month", Last + 1, 4)
            Call DetailOut(Retvalcol, "UTC Day", "day", Last + 5, 5)
            Call DetailOut(Retvalcol, "UTC Hour", "hour", Last + 10, 5)
            Call DetailOut(Retvalcol, "UTC Minute", "minute", Last + 15, 6)
            Call DetailOut(Retvalcol, "Station ID", "text", Last + 21, 42)
            Call DetailOut(Retvalcol, "Longitude", "lon", Last + 63, 25)
            Call DetailOut(Retvalcol, "Latitude", "lat", Last + 88, 24)
            Call Dac366Out(Retvalcol, "Water Level Type", "level", Last + 112, 1)
            Call DetailOut(Retvalcol, "Water Level", "simeters", Last + 113, 16, 2)
            Call Dac366Out(Retvalcol, "Reference Datum", "datum", Last + 129, 2)
            Call DetailOut(Retvalcol, "Spare", "spare", Last + 131, 14)
            Last = Last + 144 'size of repeated section
        Loop
    Case "4"
        Call DetailOut(Retvalcol, "Message ID", "fiid", From + 18, 6, "Hydro/Current Message", "Awaiting Verification")
        Last = From + 24 - 1
        Do Until Last >= DataBits
            i = i + 1
'            Call Rowout("Sensor " & i, "", "", True)
'            Call DetailLineOut("Sensor " & i, "", "", True)
            Call DetailOut(Retvalcol, "Sensor Report No", "literal", , , CStr(i))
            Call DetailOut(Retvalcol, "Month", "month", Last + 1, 4)
            Call DetailOut(Retvalcol, "UTC Day", "day", Last + 5, 5)
            Call DetailOut(Retvalcol, "UTC Hour", "hour", Last + 10, 5)
            Call DetailOut(Retvalcol, "UTC Minute", "minute", Last + 15, 6)
            Call DetailOut(Retvalcol, "Station ID", "text", Last + 21, 42)
            Call DetailOut(Retvalcol, "Longitude", "lon", Last + 63, 25)
            Call DetailOut(Retvalcol, "Latitude", "lat", Last + 88, 24)
            Call DetailOut(Retvalcol, "Current Speed", "speed", Last + 112, 8, 1)
            Call DetailOut(Retvalcol, "Current Direction Toward", "heading", Last + 120, 9, 0)
            Call DetailOut(Retvalcol, "Spare", "spare", Last + 129, 16)
            Last = Last + 144 'size of repeated section
        Loop
    Case "5"
        Call DetailOut(Retvalcol, "Message ID", "fiid", From + 18, 6, "Hydro/Salinity Temp Message", "Awaiting Verification")
        Last = From + 24 - 1
        Do Until Last >= DataBits
            i = i + 1
'            Call Rowout("Sensor " & i, "", "", True)
'            Call DetailLineOut("Sensor " & i, "", "", True)
            Call DetailOut(Retvalcol, "Sensor Report No", "literal", , , CStr(i))
            Call DetailOut(Retvalcol, "Month", "month", Last + 1, 4)
            Call DetailOut(Retvalcol, "UTC Day", "day", Last + 5, 5)
            Call DetailOut(Retvalcol, "UTC Hour", "hour", Last + 10, 5)
            Call DetailOut(Retvalcol, "UTC Minute", "minute", Last + 15, 6)
            Call DetailOut(Retvalcol, "Station ID", "text", Last + 21, 42)
            Call DetailOut(Retvalcol, "Longitude", "lon", Last + 63, 25)
            Call DetailOut(Retvalcol, "Latitude", "lat", Last + 88, 24)
            Call DetailOut(Retvalcol, "Salinity", "number", Last + 112, 10, 1, " PSU")
'316/366-1-5 (PAWSS)
            Call DetailOut(Retvalcol, "Water Temperature", "temperature", Last + 122, 10, 1)
            Call DetailOut(Retvalcol, "Spare", "spare", Last + 132, 13)
            Last = Last + 144 'size of repeated section
        Loop
    Case "6"
        Call DetailOut(Retvalcol, "Message ID", "fiid", From + 18, 6, "Water Flow Message")
        Last = From + 24 - 1
        Do Until Last >= DataBits
            i = i + 1
'            Call Rowout("Station " & i, "", "", True)
'            Call DetailLineOut("Sensor " & i, "", "", True)
            Call DetailOut(Retvalcol, "Sensor Report No", "literal", , , CStr(i))
            Call DetailOut(Retvalcol, "UTC Month", "month", Last + 1, 4)
            Call DetailOut(Retvalcol, "UTC Day", "day", Last + 5, 5)
            Call DetailOut(Retvalcol, "UTC Hour", "hour", Last + 10, 5)
            Call DetailOut(Retvalcol, "UTC Minute", "minute", Last + 15, 6)
            Call DetailOut(Retvalcol, "Station ID", "text", Last + 21, 42)
            Call DetailOut(Retvalcol, "Longitude", "lon", Last + 63, 25)
            Call DetailOut(Retvalcol, "Latitude", "lat", Last + 88, 24)
            Call DetailOut(Retvalcol, "Water Flow", "number", Last + 112, 14, 0, "cubic meters/sec")
            Call DetailOut(Retvalcol, "Spare", "spare", Last + 126, 19)
            Last = Last + 144 'size of repeated section
        Loop
    Case Else
        Call DetailOut(Retvalcol, "Message ID", "fiid", From + 18, 6, "Information required")
    End Select
Case Is = "2"
    Select Case clsSentence.AisMsgFiId
    Case "1"
        Call DetailOut(Retvalcol, "Message ID", "fiid", From + 18, 6, "Lockage Order Message")
        Call DetailOut(Retvalcol, "Month", "month", From + 24, 4)
        Call DetailOut(Retvalcol, "Day", "day", From + 28, 5)
        Call DetailOut(Retvalcol, "Hour", "hour", From + 33, 5)
        Call DetailOut(Retvalcol, "Minute", "minute", From + 38, 6)
        Call Dac366Out(Retvalcol, "Lock ID", "lockid", From + 44, 42)
        Call DetailOut(Retvalcol, "Longitude", "lon", From + 86, 25)
        Call DetailOut(Retvalcol, "Latitude", "lat", From + 111, 24)
        Call DetailOut(Retvalcol, "Spare", "spare", From + 135, 9)
        Last = From + 144 - 1
        Do Until Last >= DataBits
            i = i + 1
            Call DetailOut(Retvalcol, "Vessel Name " & i, "vesselname", Last + 1, 90)
            Call Dac366Out(Retvalcol, "Direction", "updown", Last + 91, 1)
            Call DetailOut(Retvalcol, "ETA Month", "month", Last + 92, 4)
            Call DetailOut(Retvalcol, "ETA UTC Day", "day", Last + 96, 5)
            Call DetailOut(Retvalcol, "ETA UTC Hour", "hour", Last + 101, 5)
            Call DetailOut(Retvalcol, "ETA UTC Minute", "minute", Last + 106, 6)
            Call DetailOut(Retvalcol, "Spare", "spare", Last + 112, 9)
            
            Last = Last + 120 '120 is size of repeated section
        Loop
    Case "2"
        Call DetailOut(Retvalcol, "Message ID", "fiid", From + 18, 6, "Estimated Lock Times Message")
        Call DetailOut(Retvalcol, "Month", "month", From + 24, 4)
        Call DetailOut(Retvalcol, "Day", "day", From + 28, 5)
        Call DetailOut(Retvalcol, "Hour", "hour", From + 33, 5)
        Call DetailOut(Retvalcol, "Minute", "minute", From + 38, 6)
        Call DetailOut(Retvalcol, "Vessel Name", "vesselname", From + 44, 90)
        Call DetailOut(Retvalcol, "Last Location", "text", From + 134, 42)
        Call DetailOut(Retvalcol, "Last ATA Month", "month", From + 176, 4)
        Call DetailOut(Retvalcol, "Last ATA Day", "day", From + 180, 5)
        Call DetailOut(Retvalcol, "Last ATA Hour", "hour", From + 185, 5)
        Call DetailOut(Retvalcol, "Last ATA Minute", "minute", From + 190, 6)
        Call Dac366Out(Retvalcol, "First Lock", "lockid", From + 196, 42)
        Call DetailOut(Retvalcol, "First Lock ETA Month", "month", From + 238, 4)
        Call DetailOut(Retvalcol, "First Lock ETA Day", "day", From + 242, 5)
        Call DetailOut(Retvalcol, "First Lock ETA Hour", "hour", From + 247, 5)
        Call DetailOut(Retvalcol, "First Lock ETA Minute", "minute", From + 252, 6)
        Call Dac366Out(Retvalcol, "Second Lock", "lockid", From + 258, 42)
        Call DetailOut(Retvalcol, "Second Lock ETA Month", "month", From + 300, 4)
        Call DetailOut(Retvalcol, "Second Lock ETA Day", "day", From + 304, 5)
        Call DetailOut(Retvalcol, "Second Lock ETA Hour", "hour", From + 309, 5)
        Call DetailOut(Retvalcol, "Second Lock ETA Minute", "minute", From + 314, 6)
        Call DetailOut(Retvalcol, "Delay", "text", From + 320, 42)
        Call DetailOut(Retvalcol, "Spare", "spare", From + 362, 4)
    Case "3"
        Call DetailOut(Retvalcol, "Message ID", "fiid", From + 18, 6, "Vessel Procession Order Message", "Awaiting Verification")
        Call DetailOut(Retvalcol, "Month", "month", From + 24, 4)
        Call DetailOut(Retvalcol, "Day", "day", From + 28, 5)
        Call DetailOut(Retvalcol, "Hour", "hour", From + 33, 5)
        Call DetailOut(Retvalcol, "Minute", "minute", From + 38, 6)
        Call DetailOut(Retvalcol, "Direction ID", "text", From + 44, 96)
        Call DetailOut(Retvalcol, "Longitude", "lon", From + 140, 25)
        Call DetailOut(Retvalcol, "Latitude", "lat", From + 165, 24)
        Call DetailOut(Retvalcol, "Spare", "spare", From + 189, 3)
        Last = From + 192 - 1
        Do Until Last >= DataBits
            i = i + 1
'            Call Rowout("Vessel Order Procession Report " & i, "", "", True)
'            Call DetailLineOut("Vessel Order Procession Report " & i, "", "", True)
            Call DetailOut(Retvalcol, "Vessel Order Procession Report", "literal", , , CStr(i))
            Call DetailOut(Retvalcol, "Order", "number", Last + 1, 5, 0, "1 is first vessel")
            Call DetailOut(Retvalcol, "Vessel Name " & i, "vesselname", Last + 6, 90)
            Call DetailOut(Retvalcol, "Position Name", "text", Last + 96, 72)
            Call DetailOut(Retvalcol, "Call-in UTC Hour", "hour", From + 168, 5)
            Call DetailOut(Retvalcol, "Call-in UTC Minute", "minute", From + 173, 6)
            Call DetailOut(Retvalcol, "Spare", "spare", Last + 179, 6)
            Last = Last + 184 'size of repeated section
        Loop
    Case Else
        Call DetailOut(Retvalcol, "Message ID", "fiid", From + 18, 6, "Information required")
    End Select
Case Is = "32"
    Select Case clsSentence.AisMsgFiId
    Case "1"
        Call DetailOut(Retvalcol, "Message ID", "fiid", From + 18, 6, "Version Message")
        Call DetailOut(Retvalcol, "Major Version", "number", From + 24, 8)
        Call DetailOut(Retvalcol, "Minor Version", "number", From + 32, 8)
        Call DetailOut(Retvalcol, "Spare", "spare", From + 40, 8)
    Case Else
        Call DetailOut(Retvalcol, "Message ID", "fiid", From + 18, 6, "Information required")
    End Select
Case Else
    Call DetailOut(Retvalcol, "FI", "fi", From + 10, 6, , "Information required")
End Select

End Function
'detail out outputs all "members" if used in more than one message-ifm type
'if not the code can be in AivdmDetail()
'the function return is used if the calling routine needs to decide which branch to take next
'Arg is normally the range (eq .1 or .01)
'Arg1 is normally valdes
Function DetailOut(Retvalcol As Long, Des As String, _
Member As String, _
Optional From As Long, _
Optional reqbits As Long, _
Optional Arg As String, _
Optional Arg1 As String) As String   'returned value is (1,des,2=val,3=valdes)
Static length As Long
Static Width As Long
Static AisMsgType As Long    'so we can see later which msg no

Dim Arg1s() As String    'Splits arg1 into an array of arguments delimited with |

Dim wlong As Long   'working long us if called more than once
Dim wSi As Single
Dim Val As String
Dim Valdes As String
Dim Minus As Boolean
Dim MSB As Long     'MSB striped from wlong
Dim LSBits  As Long 'wlong after MSB is removed
Dim kb As String
Dim RadioMode As Integer     '0=SOTDMA, 1=ITDMA or N/A (blank)
'Dim CommState As Long       '19 bits
'Dim RadioData As Boolean    'true= we have data
'Dim SubMessageNo As Integer
'Dim SubMessageVal As Integer
Dim Fi As Integer
Dim Dac As Integer
'Dim Payloadbits As Long
Dim i As Integer
Dim j As Integer
Dim Bold As Boolean
Dim Bits As Long
Dim remainder As Long   'no of spare bits when outpt is 6 bit ascii
Dim fmtMins As String   'Format for Mins if Degrees + Mins
Dim fmtDegs As String   'Format for Degrees and decimal Degrees
Dim Multiplier As Long 'Used with lat/lon to convert minutes into degrees
Dim Precision As Long   'No of decimal places lat.lon is accurate to
Dim posMins As Long
Dim posDegs As Long     'in the units sent

Dim strMins As String
Dim SpareBitsToBoundary As Long

'    clsField.CallingRoutine = "DetailOut"
'    clsField.Des = Des
'    clsField.Member = Member
'    clsField.From = From
'    clsField.reqbits = reqbits
'    clsField.Arg = Arg
'    clsField.Arg1 = Arg1
    
'where the from can have an offset, when DetailOut is called
'stand alone to output a column ie where the record length
'before the required field can vary in length, We must
'calculate the offset here. This will only occur if RetValCol
'is not 0, and from = 0
        
    Select Case Member
    Case Is = "aismsgtype"  'this is NOT necessarily the received message type
        AisMsgType = pLong(From, reqbits)
        Val = AisMsgType
        If AisMsgType <= 0 Or AisMsgType > 27 Then
            Valdes = "{Invalid message type}"
        Else
            Valdes = AisMsgTypeName(AisMsgType)
        End If
    
    Case Is = "mmsi"    'this is NOT necessarily the FromMMSI
        wlong = pLong(From, reqbits)  'set separate variable (using more than once)
        Val = MmsiFmt(wlong)
        If Arg1 = "" Then
            Valdes = MmsiFmt(wlong, "D")
        Else
            Valdes = GetVessel(Val)
            If Valdes = "" Then
                Valdes = MmsiFmt(wlong, "D")    'try for a generic description
            End If
        End If
    Case Is = "mid"    'this is NOT necessarily the FromMMSI
        wlong = pLong(From, reqbits)  'set separate variable (using more than once)
        Val = MmsiFmt(wlong, "M")
        If Val <> 0 Then
            If wlong = 45133333 Then
                Valdes = "VTS Drechtsteden"
            Else
                Valdes = DacName(Val)
            End If
        Else
            Valdes = "Not defined"
        End If
    Case Is = "repeat"
            Val = pLong(From, reqbits)
            Valdes = RepeatName(Val)
    Case Is = "status"
        Val = pLong(From, reqbits)
        Valdes = StatusName(Val)
    Case Is = "turn"
        wlong = pSi(From, reqbits)
        If wlong < 0 Then
            Minus = True
        Else
            Minus = False
        End If
        Val = wlong
        wlong = (wlong / 4.733) ^ 2
        If Minus Then wlong = wlong * -1
        Select Case CInt(Val)
        Case Is = 127
            Valdes = "Turning to Starboard at more than 5" & Chr$(176) & "/30sec (No TI available)"
        Case Is = -127
            Valdes = "Turning to Port at more than 5" & Chr$(176) & "/30sec (No TI available)"
        Case Is = -128
            Valdes = "No turn information available (default)"
        Case Else
            Valdes = Abs(Int(wlong)) & Chr$(176) & "/min"
            If CInt(Val) > 0 Then Valdes = Valdes & " to Starboard"
            If CInt(Val) < 0 Then Valdes = Valdes & " to Port"
        End Select
    Case Is = "speed" 'should handle boat & water speeds
        wlong = pLong(From, reqbits)
'If wlong = 128 Then MsgBox Arg1 & ":" & ItoBits(pLong(From, reqbits))
        Val = Format$(wlong / (10 ^ Arg), ArgFmt(Arg))
        Valdes = "Knots"
        Select Case Arg1
        Case Is = "251"     'max permissible 251
            Select Case wlong
            Case Is = CInt(Arg1)    '251
                Valdes = "> " & Format$(wlong / (10 ^ Arg), ArgFmt(Arg)) & " Knots"
            Case Is = (2 ^ reqbits - 1)    '255
                Valdes = "Not available"
            Case Is > CInt(Arg1)    '252 to 254
                Valdes = "Reserved for future use"
            End Select
        Case Else
            If wlong = (2 ^ reqbits - 2) Then Valdes = "> " & Format$(wlong / (10 ^ Arg), ArgFmt(Arg)) & " Knots"
            If wlong = (2 ^ reqbits - 1) Then Valdes = "Not available"
        End Select
    Case Is = "accuracy"
        Val = pLong(From, reqbits)
        Valdes = AccuracyName(Val)
    Case Is = "lon"
        Select Case reqbits
        Case Is = 28    '0.0001 min = .000002 degree
            Multiplier = 4
            fmtDegs = "0.000000"
            fmtMins = "0.0000"
        Case Is = 25    '0.001 min = .00002 degree
            Multiplier = 3
            fmtDegs = "0.00000"
            fmtMins = "0.000"
        Case Is = 18    '0.1 min = .002 degree
            Multiplier = 1
            fmtDegs = "0.000"
            fmtMins = "0.0"
        Case Else
'MsgBox "Invalid Longitude"
'Stop
        End Select
'arg is the precision (0-4) for 8-1-22
        If Arg <> "" Then
            Precision = CLng(Arg)
            Select Case Precision
            Case Is = 4
                fmtDegs = "0.0000"
                fmtMins = "0.000"
            Case Is = 3
                fmtDegs = "0.000"
                fmtMins = "0.00"
            Case Is = 2
                fmtDegs = "0.00"
                fmtMins = "0.0"
            Case Is = 1
                fmtDegs = "0.0"
                fmtMins = "0"
            Case Is = 0
                fmtDegs = "0"
                fmtMins = "0"
                'use same format as for val
            End Select
        End If
        wSi = pSi(From, reqbits)                            '4233984
        posDegs = Int(Abs(wSi / (60 * 10 ^ Multiplier)))    '70
        posMins = Abs(wSi) - posDegs * 60 * 10 ^ Multiplier '33984
        wSi = wSi / (60 * 10 ^ Multiplier)
'Always display all bits passed
'force dot if french
        Val = Replace(Format$(wSi, fmtDegs), ",", ".")
        If wSi = 181 Then
            Valdes = "Not available (default)"
        Else
            Valdes = Abs(posDegs) & Chr$(176) & " " _
             & Format$(CSng(posMins / 10 ^ Multiplier), fmtMins) & "'"
            If wSi >= 0 Then
                Valdes = Valdes & " E"
            Else
                Valdes = Valdes & " W"
            End If
        End If
    Case Is = "lat"
        Select Case reqbits
        Case Is = 27
            Multiplier = 4
            fmtDegs = "0.000000"
            fmtMins = "0.0000"
        Case Is = 24
            Multiplier = 3
            fmtDegs = "0.00000"
            fmtMins = "0.000"
        Case Is = 17
            Multiplier = 1
            fmtDegs = "0.000"
            fmtMins = "0.0"
        Case Else
'MsgBox "Invalid Latitude"
'Stop
        End Select
'arg is the precision (0-4) for 8-1-22
        If Arg <> "" Then
            Precision = CLng(Arg)
            Select Case Precision
            Case Is = 4
                fmtDegs = "0.0000"
                fmtMins = "0.000"
            Case Is = 3
                fmtDegs = "0.000"
                fmtMins = "0.00"
            Case Is = 2
                fmtDegs = "0.00"
                fmtMins = "0.0"
            Case Is = 1
                fmtDegs = "0.0"
                fmtMins = "0"
            Case Is = 0
                fmtDegs = "0"
                fmtMins = "0"
                'use same format as for val
            End Select
        End If
        wSi = pSi(From, reqbits)                            '4233984
        posDegs = Int(Abs(wSi / (60 * 10 ^ Multiplier)))    '70
        posMins = Abs(wSi) - posDegs * 60 * 10 ^ Multiplier '33984
        wSi = wSi / (60 * 10 ^ Multiplier)
'Always display all bits passed
'force dot if french
        Val = Replace(Format$(wSi, fmtDegs), ",", ".")
        If wSi = 91 Then
            Valdes = "Not available (default)"
        Else
            Valdes = Abs(posDegs) & Chr$(176) & " " _
             & Format$(CSng(posMins / 10 ^ Multiplier), fmtMins) & "'"
            If wSi >= 0 Then
                Valdes = Valdes & " N"
            Else
                Valdes = Valdes & " S"
            End If
        End If
    Case Is = "course", "heading"
        wlong = pLong(From, reqbits)
                'if bits are 10 max val is 740 limit is 1023
                'hence must be in 1/2 degree increments
        If reqbits = 10 Then
            wlong = wlong * 5
            Arg = 1             'display XXX.X
        End If
        Val = Format$(wlong / (10 ^ Arg), ArgFmt(Arg))
        Valdes = Chr$(176) & " (degrees)"
        Select Case CLng(Val)
        Case Is < 360
        Case Is = 360, 511, 720, 2 ^ reqbits - 1
            Valdes = "Not available (default)"
        Case Else
            Valdes = "Invalid (should not be used)"
        End Select
    Case Is = "manoeuvre"
        Val = pLong(From, reqbits)
            If CInt(Val) = 0 Then Valdes = "not available (default)"
            If CInt(Val) = 1 Then Valdes = "No Special Manoeuvre"
            If CInt(Val) = 2 Then Valdes = "Special Manoeuvre (such as regional passing arrangement)"
    Case Is = "partno"
        Val = pLong(From, reqbits)
        If CInt(Val) <= 1 Then
            Valdes = PartnoName(Val)
        Else
            Valdes = "Invalid"
        End If
'        DetailOut = Val
    Case Is = "radiomode"
        Val = pLong(From, reqbits)
        Valdes = RadioModeName(Val)
'        DetailOut = Val
    Case Is = "mothership_mmsi"
        wlong = pLong(From, reqbits)
        Val = MmsiFmt(wlong)    ', "V")
        Valdes = MmsiFmt(wlong, "D")
    Case Is = "raim"
        Val = pLong(From, reqbits)
        Valdes = RaimName(Val)
    Case Is = "year"
        Val = Format$(pLong(From, reqbits), "0000")
        If reqbits = 8 Then Val = Val + 2000
        If Val = 0 Then Valdes = "not available (default)"
    Case Is = "month"
        Val = Format$(pLong(From, reqbits), "00")
        If Val = 0 Then Valdes = "not available (default)"
    Case Is = "day"
        Val = Format$(pLong(From, reqbits), "00")
        If Val = 0 Then Valdes = "not available (default)"
    Case Is = "hour"
        Val = Format$(pLong(From, reqbits), "00")
        If Val = 24 Then Valdes = "not available (default)"
    Case Is = "minute"
        Val = Format$(pLong(From, reqbits), "00")
        If Val = 60 Then Valdes = "not available (default)"
    Case Is = "second"
        Val = Format$(pLong(From, reqbits), "00")
        If Val > 60 Then Valdes = "not available (default)"
        Valdes = SecondDes(Val)
    Case Is = "epfd"
        Val = pLong(From, reqbits)
        Valdes = EpfdName(Val)
    Case Is = "txcontrol"
        Val = pLong(From, reqbits)
        Valdes = TxControlName(Val)
    Case Is = "ais_version"
        Val = pLong(From, reqbits)
        Valdes = Ais_VersionName(Val)
    Case Is = "imo"
        Val = pLong(From, reqbits)
    Case Is = "ship_type"
        Val = pLong(From, reqbits)
        Select Case CInt(Val)
        Case Is = 0
            Valdes = "not available or no ship (default)"
        Case 50 To 59
            Valdes = Ship_Type5NName(CInt(Val) - 50)
        Case 30 To 39
            Valdes = "Vessel" & Ship_Type3NName(CInt(Val) - 30)
        Case 10 To 29, 40 To 49, 60 To 99
            Valdes = Ship_TypeNxName(Int(CInt(Val) / 10)) & _
            Ship_TypexNName(CInt(Val) Mod 10)
        Case 100 To 199
            Valdes = "Reserved for regional use"
        Case Else
            Valdes = "Reserved for future use"
        End Select
    Case Is = "to_bow"
        Val = pLong(From, reqbits)
        Valdes = "meters"
'       Length = Val
        If Val = 0 Then Valdes = "not available (default)"
        If Val = 511 Then Valdes = "511 meters or greater"
    Case Is = "to_stern"
        Val = pLong(From, reqbits)
        Valdes = "meters"
'        Length = Length + Val
        If Val = 0 Then Valdes = "not available (default)"
        If Val = 511 Then Valdes = "511 meters or greater"
    Case Is = "clength"
        Val = pLong(From, reqbits) + pLong(From + reqbits, reqbits)
        Valdes = "meters, {calculated}"
        If Val = 0 Then Valdes = "not available (default)"
        If Val = 1022 Then Valdes = "511 meters or greater"
    Case Is = "to_port"
        Val = pLong(From, reqbits)
        Valdes = "meters"
'        Width = Val
        If Val = 0 Then Valdes = "not available (default)"
        If Val = 63 Then Valdes = "63 meters or greater"
    Case Is = "to_starboard"
        Val = pLong(From, reqbits)
        Valdes = "meters"
'        Width = Width + Val
        If Val = 0 Then Valdes = "not available (default)"
        If Val = 63 Then Valdes = "63 meters or greater"
    Case Is = "cbeam"
        Val = pLong(From, reqbits) + pLong(From + reqbits, reqbits)
        Valdes = "meters, {calculated}"
        If Val = 0 Then Valdes = "not available (default)"
        If Val = 126 Then Valdes = "63 meters or greater"
     Case Is = "draught"
        Val = CSng(pLong(From, reqbits)) / 10 ^ Arg
        Valdes = "meters"
        If Val = 0 Then Valdes = "not available (default)"
    Case Is = "dte"
        Val = pLong(From, reqbits)
        Valdes = DteName(Val)
    Case Is = "alt" 'msg 9 SAR
        Val = pLong(From, reqbits)
        If Val = 4094 Then Valdes = " 4094 meters or over"
        If Val = 4095 Then Valdes = " not available"
        Val = Val & " meters"
    Case Is = "cs"
        Val = pLong(From, reqbits)
        Valdes = CsName(Val)
    Case Is = "display"
        Val = pLong(From, reqbits)
        Valdes = DisplayName(Val)
    Case Is = "dsc"
        Val = pLong(From, reqbits)
        Valdes = DscName(Val)
    Case Is = "band"
        Val = pLong(From, reqbits)
        Valdes = BandName(Val)
    Case Is = "msg22"
        Val = pLong(From, reqbits)
        Valdes = Msg22Name(Val)
    Case Is = "aid_type"
        Val = pLong(From, reqbits)
        Valdes = AtoNName(Val)
    Case Is = "off_position"
        Val = pLong(From, reqbits)
        Valdes = Off_PositionName(Val)
    Case Is = "virtual_aid"
        Val = pLong(From, reqbits)
        Valdes = Virtual_AidName(Val)
    Case Is = "gnssposition"
        Val = pLong(From, reqbits)
        Valdes = GnssPositionName(CInt(Val))
'Dont think data is used
'data justs outputs Binary Head with Hex data - data is used to output unstructured data
    Case Is = "data", "aidata"
'dont output if csv
        If Retvalcol = 4 Then GoTo Check_Additional
        Val = pHex(From, reqbits)
        Valdes = "Hex Binary Data (" & reqbits & " bits)"
        Bold = True
        If Member = "aidata" Then
'            Payloadbits = (UBound(PayloadBytes) + 1) * 8
'            If Payloadbits - From + 1 < 16 Then
            If reqbits < 16 Then
                Valdes = "Application Identifier invalid (" _
                & 16 - reqbits & " bits too short)"
            End If
        End If
    Case Is = "dgnssdata"
        Val = pHex(From, reqbits)
        Valdes = Valdes + "Hex Binary Data (" & reqbits & " bits)"
    Case Is = "addressed"
        Val = pLong(From, reqbits)
        Valdes = AddressedName(Val)
'        DetailOut = Val
    Case Is = "structured"
        Val = pLong(From, reqbits)
        Valdes = StructuredName(Val)
'        DetailOut = Val
    Case Is = "band_width"
        Val = pLong(From, reqbits)
        Valdes = Band_WidthName(Val)
    Case Is = "zonesize"
        Val = pLong(From, reqbits)
        Valdes = (Val + 1) & " Nautical miles transitional zone"
    Case Is = "stationtype"
        Val = pLong(From, reqbits)
        Valdes = StationTypeName(Val)
    Case Is = "interval"
        Val = pLong(From, reqbits)
        Valdes = IntervalName(Val)
    Case Is = "altitudesensor"
        Val = pLong(From, reqbits)
        Valdes = AltitudeName(Val)
    Case Is = "assigned"
        Val = pLong(From, reqbits)
        Valdes = AssignedName(Val)
    Case Is = "quiet"
        Val = pLong(From, reqbits)
        If Val = 0 Then Valdes = "no quiet time commanded"
        Val = Val & " minutes"
    Case Is = "channel"
        Val = pLong(From, reqbits)
        Select Case Val
        Case Is = "2087"
            Valdes = "161.975 MHz"
        Case Is = "2088"
            Valdes = "162.025 MHz"
        Case Else
            Valdes = "See ITU-R M1084"
        End Select
    Case Is = "txrx"
        Val = pLong(From, reqbits)
        Valdes = TxrxName(Val)
    Case Is = "power"
        Val = pLong(From, reqbits)
        Valdes = PowerName(Val)
    Case Is = "temperature"
        If Arg = "" Then Arg = 1
        wSi = pSi(From, reqbits)
        Val = Format$(wSi / (10 ^ Arg), ArgFmt(Arg))
        Select Case Arg1
        Case Is = "water"
'SI -10.0 to 50.0 n/a=50.1  (10 bits)
            Select Case wSi
            Case Is = 501
                Valdes = "Not available"
            Case Is >= 502, Is <= -101
                Valdes = "Reserved for future use"
            Case Else
                Valdes = Val & " " & Chr$(176) & "C"
            End Select
        Case Is = "airdraught"
        wlong = pLong(From, reqbits)
'If wlong = 128 Then MsgBox Arg1 & ":" & ItoBits(pLong(From, reqbits))
        Val = Format$(wlong / (10 ^ Arg), ArgFmt(Arg))
        Valdes = "Meters"
        Select Case wlong
        Case Is = 8191
            Valdes = ">= " & Format$(wlong / (10 ^ Arg), ArgFmt(Arg)) & " " & Valdes
        Case Is = 0
            Valdes = "Not available (default)"
        End Select
        Case Is = "air"
'SI -60.0 to 60.0 n/a=-1024 (11 bits)
            Select Case wSi
            Case Is = -1024
                Valdes = "Not available"
            Case Is >= 601, Is <= -601
                Valdes = "Reserved for future use"
            Case Else
                Valdes = Val & " " & Chr$(176) & "C"
            End Select
        Case Is = "dew"
'SI -20 to +50 n/a=501  (10 bits)
            Select Case wSi
            Case Is = 501
                Valdes = "Not available"
            Case Is >= 502, Is <= -201
                Valdes = "Reserved for future use"
            Case Else
                Valdes = Val & " " & Chr$(176) & "C"
            End Select
'1-11 (I think this is correct but not documented)
        Case Is = "-600", "-200", "-100"
            wlong = pLong(From, reqbits)
            Val = wlong
            If wlong = (2 ^ (reqbits)) - 1 Then
                Valdes = "Not available"
            Else
                Valdes = Format$((wlong + CLng(Arg1)) / (10 ^ Arg), ArgFmt(Arg)) & " " & Chr$(176) & "C"
            End If
        Case Else
'si -51.2=n/a -51.1= or less,+511= or greater (10 bits)
            If wSi = -(2 ^ (reqbits - 1)) Then
                Valdes = "Not available"
            Else
                Valdes = Val & " " & Chr$(176) & "C"
            End If
        End Select
    Case Is = "airdraught"    '1-24,
        wlong = pLong(From, reqbits)
'If wlong = 128 Then MsgBox Arg1 & ":" & ItoBits(pLong(From, reqbits))
        Val = Format$(wlong / (10 ^ Arg), ArgFmt(Arg))
        Valdes = "Meters"
        Select Case wlong
        Case Is = 8191
            Valdes = ">= " & Format$(wlong / (10 ^ Arg), ArgFmt(Arg)) & " " & Valdes
        Case Is = 0
            Valdes = "Not available (default)"
        End Select
'1-11, 1-31
    Case Is = "chartdatum"
'Dac 1-11   'not used after 1/1/13
'si -51.2=n/a -51.1= or less,+511= or greater (9 (.1) or 12 (.01) bits)
        
        Select Case reqbits
        Case Is = 12    '1-31
            wlong = pLong(From, reqbits)
            Val = wlong
            Select Case Val
            Case Is <= 4000
                Valdes = Format$(wlong / (10 ^ Arg) - 10, ArgFmt(Arg)) & " meters"
            Case Is = 4001
                Valdes = "Not available (default)"
            Case Else
                    Valdes = "Reserved for future use"
            End Select
        Case Else   '1-11
            wSi = pSi(From, reqbits)
            Val = Format$(wSi / (10 ^ Arg), ArgFmt(Arg))
                Select Case wSi
                Case Is = 501
                    Valdes = "Not available"
                Case Is >= 502, Is <= -201
                    Valdes = "Reserved for future use"
                Case Else
                Valdes = Val & " meters"
                End Select
        End Select
    Case Is = "miles"
'1-11 & 316/366-1-6 & 366-33-9 Visibility
        wlong = pLong(From, reqbits)
        If Arg = "" Then
            If reqbits = 8 Then Arg = 1
        End If
        Val = Format$(wlong / (10 ^ Arg), ArgFmt(Arg))
        If wlong = (2 ^ reqbits - 1) Then
            Valdes = "Not available"
        Else
            Valdes = "Nautical Miles"
        End If
    Case Is = "meters"
        wlong = pLong(From, reqbits)
'If wlong = 128 Then MsgBox Arg1 & ":" & ItoBits(pLong(From, reqbits))
        Val = Format$(wlong / (10 ^ Arg), ArgFmt(Arg))
        Valdes = "Meters"
        Select Case Arg1
        Case Is = "Kilometers"  '316/366-1-1 visibility
            Valdes = Arg1
            If wlong = (2 ^ reqbits - 2) Then Valdes = "> " & Format$(wlong / (10 ^ Arg), ArgFmt(Arg)) & " kilometers"
            If wlong = (2 ^ reqbits - 1) Then Valdes = "Not available"
        Case Is = "251"     'max permissible 251
            Select Case wlong
            Case Is = CInt(Arg1)    '251
                Valdes = "> " & Format$(wlong / (10 ^ Arg), ArgFmt(Arg)) & " meters"
            Case Is = (2 ^ reqbits - 1)    '255
                Valdes = "Not available"
            Case Is > CInt(Arg1)    '252 to 254
                Valdes = "Reserved for future use"
                End Select
        Case Is = "2001"     'max permissible 2001
            Select Case wlong
            Case Is = CInt(Arg1)    '2002-2046
                Valdes = "> " & Format$(wlong / (10 ^ Arg), ArgFmt(Arg)) & " meters"
            Case Is = (2 ^ reqbits - 1)    '2047
                Valdes = "Not available"
            Case Is > CInt(Arg1)    '(not used here) 2001 to 2047
                Valdes = "Reserved for future use"
            End Select
        Case Else
            If wlong = (2 ^ reqbits - 2) Then Valdes = "> " & Format$(wlong / (10 ^ Arg), ArgFmt(Arg)) & " meters"
            If wlong = (2 ^ reqbits - 1) Then Valdes = "Not available"
        End Select
    Case Is = "simeters"
        wSi = pSi(From, reqbits)
        Val = Format$(wSi / (10 ^ Arg), ArgFmt(Arg))
        Valdes = "Meters"
        If wSi = (2 ^ (reqbits - 1)) - 1 Then Valdes = "=> " & Format$((2 ^ (reqbits - 2)) - 1, ArgFmt(Arg)) & " meters" '2*15=32768
        If wSi = -(2 ^ (reqbits - 1)) Then Valdes = "Not available" 'all bits set
    Case Is = "humidity"
       wlong = pLong(From, reqbits)
        Val = Format$(wlong / (10 ^ Arg), ArgFmt(Arg))
'100% is max humidity
        If wlong = (2 ^ reqbits - 1) Or wlong / (10 ^ Arg) > 100 Then
            Valdes = "Not available"
        Else
            Valdes = "% (percent)"
        End If
    Case Is = "pressure"
       wlong = pLong(From, reqbits)
       Val = wlong
        Select Case wlong
        Case Is = 0
            Valdes = "799 hPa or less!"
        Case Is = 402
            Valdes = "1201 hPa or greater"
        Case Is = 403
            Valdes = "Not available (default)"
        Case Is >= 404
            Valdes = "reserved for future use"
        Case Else
            Valdes = (wlong + 799) & " hPa"
        End Select
    Case Is = "tendency"
        Val = pLong(From, reqbits)
        Valdes = TendencyName(Val)
    Case Is = "ituwmotendency"
        Val = pLong(From, reqbits)
        Valdes = ItuWmoTendencyName(Val)
    Case Is = "ituimoweather"
        Val = pLong(From, reqbits)
        Valdes = ItuWmoWeatherName(Val)
    Case Is = "wmosecond"
        Val = Format$(pLong(From, reqbits), "00")
        Select Case Val
        Case Is = 63
            Valdes = "not available (default)"
        Case Is > 61, 62
            Valdes = "reserved for future use"
        Case Else
            Valdes = Val & " seconds"
        End Select
    Case Is = "wmometers"
        wlong = pLong(From, reqbits)
        If Arg1 <> "" Then
            Valdes = Arg1
        Else
            Valdes = "Meters"
        End If
        Val = Format$(wlong / (10 ^ Arg), ArgFmt(Arg))
        If wlong = (2 ^ reqbits - 1) Then
            Valdes = "Not available"
        Else
            If wlong = wlong / (10 ^ Arg) + 1 Then Valdes = "> "    '251
            If wlong > wlong / (10 ^ Arg) + 1 Then
                Valdes = "reserved for future use"    '252-254
            Else
                Valdes = Valdes & "Meters" '<=250
            End If
        End If
    Case Is = "visibility"
'Dac-FI 1-21-0 and 1-31 format
        wlong = pLong(From, reqbits)
        If Arg = "" Then
            If reqbits = 8 Then Arg = 1
        End If
'If 8 bits Mask Bit 8 then Shift right 7 bits
        MSB = (wlong And 2 ^ (reqbits - 1)) / 2 ^ (reqbits - 1)
'Mask lowest 7 bits
        LSBits = wlong And (2 ^ (reqbits - 1) - 1)
'value ignores MSB - if MSB is set then max range is exceeded
        Val = Format$(LSBits / 10 ^ Arg, ArgFmt(Arg))
'if all excepting MSB is set then reading is not available
        If LSBits = (2 ^ (reqbits - 1) - 1) Then
            Valdes = "Not available"    'if reqbits=8 0-126 is vis,
        Else
'if MSB is set the max range exceeded
            If MSB = 1 Then
                Valdes = "> "
            End If
            Valdes = Valdes + Val & " Nautical Miles"
        End If
    Case Is = "beaufort"
        Val = pLong(From, reqbits)
        Valdes = BeaufortName(Val)
    Case Is = "beaufortenvironmental"
        Val = pLong(From, reqbits)
        Valdes = BeaufortEnvironmentalName(Val)
    Case Is = "precipitation"
        Val = pLong(From, reqbits)
        Valdes = PrecipitationName(Val)
    Case Is = "yesno"
        Val = pLong(From, reqbits)
        Valdes = YesNoName(Val)
    Case Is = "yesno2"
        Val = pLong(From, reqbits)
        Valdes = YesNo2Name(Val)
    Case Is = "salinity"
        wlong = pLong(From, reqbits)
'If wlong = 128 Then MsgBox Arg1 & ":" & ItoBits(pLong(From, reqbits))
        If Arg = "" Then Arg = 1 'default .1
        Val = Format$(wlong / (10 ^ Arg), ArgFmt(Arg))
        Valdes = "%% (ppt)"
        Select Case Arg1
        Case Is = "501"     'max permissible 501
            Select Case wlong
            Case Is = CInt(Arg1)    '501
                Valdes = "> " & Format$(wlong / (10 ^ Arg), ArgFmt(Arg)) & " %% (ppt)"
            Case Is = (2 ^ reqbits - 2)    '510
                Valdes = "Not available (default)"
            Case Is = (2 ^ reqbits - 1)    '511
                Valdes = "Sensor not available"
            Case Is > CInt(Arg1)    '502 to 509
                Valdes = "Reserved for future use"
            End Select
        Case Else
            If wlong = (2 ^ reqbits - 2) Then Valdes = "Sensor not available"
            If wlong = (2 ^ reqbits - 1) Then Valdes = "Not available (default)"
        End Select
    Case Is = "targtype"
        Val = pLong(From, reqbits)
        Valdes = TargTypeName(Val)
'        DetailOut = Val
    Case Is = "eni"
        Val = p6bit(From, reqbits)
        If Len(Val) = 7 Then Val = "0" & Val
        Val = Format$("0" & Val, "000 00000")
        Valdes = EniName(Val)
    Case Is = "eri"
        Val = pLong(From, reqbits)
        Valdes = EriName(CInt(Val))
    Case Is = "cone"
        Val = pLong(From, reqbits)
        Valdes = ConeName(Val)
    Case Is = "loaded"
        Val = pLong(From, reqbits)
        Valdes = LoadedName(Val)
    Case Is = "qualspeed"
        Val = pLong(From, reqbits)
        Valdes = QualSpeedName(Val)
    Case Is = "qualhead"
        Val = pLong(From, reqbits)
        Valdes = QualHeadName(Val)
    Case Is = "onoff"
        Val = pLong(From, reqbits)
        Valdes = OnOffName(Val)
    Case Is = "bccode"
        Val = pLong(From, reqbits)
        Valdes = BcCodeName(Val)
    Case Is = "dgunits"
        Val = pLong(From, reqbits)
        Valdes = DgUnitsName(Val)
    Case Is = "marpol1"
        Val = pLong(From, reqbits)
        Valdes = Marpol1Name(Val)
    Case Is = "marpol2"
        Val = pLong(From, reqbits)
        Valdes = Marpol2Name(Val)
    Case Is = "dgcode"
        Val = pLong(From, reqbits)
        Valdes = DgCodeName(Val)
'        DetailOut = Val
    Case Is = "imdgcode"
        wlong = pLong(From, reqbits)
        Val = wlong
        Select Case wlong
            Case Is = 0
                Valdes = "Not available (default)"
            Case 10 To 99
                Valdes = "Main class=" & Int(Val / 10) & ", Sub class or division=" & Val Mod 10
            Case Else
                Valdes = "invalid"
        End Select
    Case Is = "igccode"
        wlong = pLong(From, reqbits)
        Val = wlong
        Select Case wlong
            Case Is = 0
                Valdes = "Not available (default)"
            Case 1 To 3363
                Valdes = "UN number"
            Case Else
                Valdes = "reserved for future use"
        End Select
    Case Is = "persons"
        wlong = pLong(From, reqbits)
        Val = wlong
        Valdes = wlong & " persons"
        If wlong = 0 Then Valdes = "not available (default)"
        If wlong = (2 ^ reqbits - 1) Then Valdes = " more than " & Val
    Case Is = "solasequipment"
        wlong = pLong(From, reqbits)
        Val = wlong
        Valdes = OperationalName(Val)
    Case Is = "iceclass"
        wlong = pLong(From, reqbits)
        Val = wlong
        Valdes = IceClassName(Val)
    Case Is = "bits"    'debugging
        Val = pbits(From, reqbits)
        Valdes = ", " & Len(Val) & " bits"
        Bits = From
        If Arg <> "" Then Valdes = Valdes & " " & Arg
        Do Until Len(Val) <= 16   'split into lines of 16 bits
             Valdes = Bits & "-" & Bits + 15 & Valdes
'            Call Rowout(Des, Left$(Val, 8) & " " & Mid$(Val, 9, 8), ValDes)
            Call DetailLineOut(Des, Left$(Val, 8) & " " & Mid$(Val, 9, 8), Valdes)
            Val = Right$(Val, Len(Val) - 16)
            Bits = Bits + 16    'next bit to output
            Des = ""
            Valdes = ""
        Loop
            'left & mid can return nul strings
        If Len(Val) > 8 Then
            Val = Left$(Val, 8) & " " & Mid$(Val, 9, 8)
            Valdes = Bits & "-" & Bits + Len(Val) - 2   'includes a space)
        Else
            Val = Left$(Val, 8)
            Valdes = Bits & "-" & Bits + Len(Val) - 1
        End If
    Case Is = "fiavailability"
        kb = pbits(From, reqbits)
        For i = 1 To Len(kb) Step 2
            If Mid$(kb, i, 1) = "1" Then
                Des = "FI Capability"
                Val = (i - 1) / 2
                Valdes = "bit (" & i & ") set"
'                Call Rowout(Des, Val, Valdes)
'                Call DetailLineOut(Des, Val, Valdes)
                Call DetailOut(Retvalcol, "Function Capability", "literal", , , Val, Valdes)
            End If
        Next i
'dont output the last FI again
        GoTo Check_Additional
    Case Is = "usacanfi"
        wlong = pLong(From, reqbits)
        Val = wlong
        Valdes = UsaCanFiName(Val)
    Case Is = "usacanid"
        Val = pLong(From, reqbits)
        Select Case Arg
            Case Is = "1"
                Valdes = UsaCan1IdName(Val)
            Case Is = "2"
                Valdes = UsaCan2IdName(Val)
            Case Is = "32"
                Valdes = UsaCan32IdName(Val)
            End Select
'        DetailOut = Val
    Case Is = "app_id"
        Val = pLong(From, reqbits)
        Fi = Val And 63     'remove 6 LSB
        Dac = (Val - Fi) / 2 ^ 6   'top 10 bits
        Valdes = "DAC is " & Dac & ", FI is " & Fi
    Case Is = "dac"
        Val = pLong(From, reqbits)
        Valdes = DacName(Val)
        If TestDac = True Then
            If DacMap(1) <> Val Then
                Valdes = Val & " " & Valdes & " mapped to " & DacMap(1) & " " & DacName(DacMap(1))
            End If
            Val = DacMap(1)
        End If
        If Arg1 <> "" Then
            Valdes = Arg1
        End If
    Case Is = "ifi"
        Val = pLong(From, reqbits)
        If Val <= 63 Then   'international fi'f
            Valdes = IfiName(Val)
        Else        'regional fi's
            Valdes = "Regional Specific"
        End If
        If Arg <> "" Then Valdes = Arg    'replace default name
        If Arg1 <> "" Then Valdes = Valdes & " {" & Arg1 & "}"
'        DetailOut = Val
    Case Is = "fiid"
        Val = pLong(From, reqbits)
'        DetailOut = Val
        If Arg <> "" Then Valdes = Arg    'replace default name
        If Arg1 <> "" Then Valdes = Valdes & " {" & Arg1 & "}"
    Case Is = "fi"
        Val = pLong(From, reqbits)
'        DetailOut = Val
        If Arg <> "" Then Valdes = Arg    'replace default name
        If Arg1 <> "" Then      'just tidies the formatting
            If Valdes <> "" Then
                Valdes = Valdes & " {" & Arg1 & "}"
            Else
                Valdes = Arg1
            End If
        End If
    Case Is = "scale"
        Val = pLong(From, reqbits)
        Valdes = 10 ^ (Val)
'        DetailOut = Val
    Case Is = "literal"     'used for counter eg sensor no
        If Arg = "" Then Arg = 0
        Val = Arg
'        DetailOut = Val
        Valdes = Arg1
'        Bold = True
    Case Is = "number"
        If Arg = "" Then Arg = 0
        wlong = pLong(From, reqbits)
        Val = Format$(wlong / (10 ^ Arg), ArgFmt(Arg))
        If Arg1 <> "" Then Valdes = Arg1
        If wlong = (2 ^ reqbits - 1) Then Valdes = "Not available (default)"
    Case Is = "vhfchannel"
        If Arg = "" Then Arg = 0
        wlong = pLong(From, reqbits)
        Val = Format$(wlong / (10 ^ Arg), ArgFmt(Arg))
        If Arg1 <> "" Then Valdes = Arg1
        If wlong = 0 Then Valdes = "Not available (default)"
    Case Is = "precision"
        If Arg = "" Then Arg = 0
        wlong = pLong(From, reqbits)
        Val = wlong
        Valdes = Val & " Decimal Places"
        If Val = 4 Then
            Valdes = Valdes & " (default)"
        End If
    Case Is = "zero"
        If Arg = "" Then Arg = 0
        wlong = pLong(From, reqbits)
        Val = wlong
        If Val <> 0 Then
            Valdes = "invalid"
        End If
    Case Is = "scalenumber"
        If Arg = "" Then Arg = 1    'Scale Factor 1,10,100,1000
        wlong = pLong(From, reqbits)
        Val = wlong
        Valdes = Val * 10 ^ Arg
        If Arg1 <> "" Then Valdes = Valdes & " " & Arg1
        If wlong = 0 Then Valdes = "Point (default) " & Valdes
    Case Is = "callsign", "vesselname", "destination", "name", "vendorid"
'Debug.Print Member
'v136        Val = RTrim$(Replace(p6bit(From, reqbits), "@", ""))
        Val = p6bit(From, reqbits)  'v136
'V 144 Jan 2017 unless $AITAG(Jason time stamp)has been found make CSV output compatible with V129
'Change Vessel@Name@@@@@@@@@ to Vessel Name - LEAVE leading spaces (use Rtrim not Trim)
        If Left$(GroupSentence.NmeaSentence, 6) = "$AITAG" Then
            Val = RTrim$(Replace(Val, "@", " "))
        End If
        Valdes = Int(reqbits / 6) & " character"
        If Int(reqbits / 6) <> 1 Then Valdes = Valdes & "s"
        remainder = reqbits - Int(reqbits / 6) * 6
        If remainder > 0 Then
            Valdes = Valdes & " + " & remainder & " fill bit"
            If remainder <> 1 Then Valdes = Valdes & "s"
        End If
    
'jnadebug
'Call DetailOut(RetValCol, Member, "bits", From, reqbits)
        
    Case Is = "text"    'Output 6-bit text
'Debug.Print "text(" & From & ":" & reqbits & ")"

'variable length text Msg 12,14,21.
'reqbits is set to the maximum possible length of the data to the end of the record
'This is the size that is saved
'If the req bits are within the length of the payload output all the bits
'v136    If reqbits <= clsSentence.AisPayloadBits - (From - 1) Then
'v136        Val = RTrim$(Replace(p6bit(From, reqbits), "@", ""))
'v136    Else
'The maximum size of the Text is larger than the size of the payload
'so only output up to the length of the payload.
'Spare bits may be required to fill to an 8-bit boundary (can be 0,2,4,6) see msg 21 definition
'v136        SpareBitsToBoundary = (clsSentence.AisPayloadBits - (From - 1)) - Int(((clsSentence.AisPayloadBits - (From - 1)) / 6)) * 6
'v136        Val = Replace(p6bit(From, Int(((clsSentence.AisPayloadBits - (From - 1)) / 6)) * 6), "@", "")
'v136    End If
    
'test If reqbits - Int(reqbits / 6) * 6 <> SpareBitsToBoundary Then Stop
    SpareBitsToBoundary = reqbits - Int(reqbits / 6) * 6
    Val = p6bit(From, reqbits)
        
        
        If Arg1 <> "" Then
            If Len(Val) = 0 Then
                Valdes = "[blank] " & Arg1
            Else
                Valdes = Arg1
            End If
        Else
            If Len(Val) = 0 Then
                Valdes = "[blank]"
            Else
                Valdes = Len(Val) & " characters"
            End If
        End If
'line is only split on Detail Form
        kb = Val 'keep original val
'MUST only call detailLineOut at end of this routine
'otherwise second call (at the end) does not output source
'        Call DetailLineOut(Des, Val, ValDes)
'        Val = kb 'recover val
'        Do Until Len(Val) <= 20   'split into lines of 20 characters
'            Call Rowout(Des, Left$(Val, 20), Valdes)
'            kb = Val 'keep original val
'            Call DetailLineOut(Des, Left$(Val, 20), ValDes)
'            Val = kb    'recover val
'            Val = Right$(Val, Len(Val) - 20)
'            Des = ""
'            ValDes = ""
'        Loop
    Case Is = "testfortext"
        For i = 0 To 5
        Call DetailOut(Retvalcol, "Test of text, offset (" & i & ")", "text", From + i, reqbits)
        Next i
    Case Is = "seqno", "increment", "retransmit", "offset"
        Val = pLong(From, reqbits)
    Case Is = "ack"
        Val = pLong(From, reqbits)
        If Val = 0 Then
            Valdes = "Negative"
        Else
            Valdes = "Affirmative"
        End If
    Case Is = "spare", "reserved", "numeric"
        If reqbits <= 30 Then
            Val = pLong(From, reqbits)
        Else
            Val = pHex(From, reqbits)
        End If
        Valdes = reqbits & " bit"
        If reqbits <> 1 Then Valdes = Valdes & "s"
        Valdes = Valdes & Arg
   Case Is = "awaitingdecoding"
        Val = pHex(From, reqbits)
        Valdes = "Data Awaiting Decoding (" & reqbits & " bits)"
    Case Is = "message"
        If Arg <> "" Then Valdes = Arg
    Case Is = "excessbits"
        Val = reqbits   'no of bits rather than bit value required
        If Val > 0 Then
            Valdes = pHex(From, reqbits)
            Valdes = "Message " & reqbits & " bits too long (" & Valdes & ") hex"
        End If
        If Val < 0 Then
            Valdes = "Message " & -reqbits & " bits too short"
        End If
'Any members not defined, try setting up a different member
'and recursively call DetailOut. This is required if the to and from
'vary even for the same message, when DetailOut is called to output a field.

    Case Else
'        Select Case Member
'        Case Is = "fillbits"
'            Val = reqbits
'    If Val >= 8 Then ValDes = "{Station inserted more fill bits than required}"
'    If Val <> 0 Then
'        Call Rowout(Des, Val, Valdes)
'            Call DetailLineOut(Des, Val, ValDes)
'        End If
'        Case Else
        Val = Member
        Valdes = "(" & From & ":"
        If IsMissing(reqbits) Then
            Valdes = Valdes & "none)"
        Else
            Valdes = Valdes & reqbits & ")"
        End If
        If IsMissing(Arg) Then
            Valdes = Valdes & " no Arg"
        Else
            Valdes = Valdes & " arg= " & Arg
        End If
        If IsMissing(Arg1) Then
            Valdes = Valdes & " no Arg1"
        Else
            Valdes = Valdes & " arg1= " & Arg1
        End If
        If IsMissing(Retvalcol) Then
            Valdes = Valdes & " no RetValCol"
        Else
            Valdes = Valdes & " RetValCol= " & Retvalcol
        End If
        Exit Function   'Member not found
    End Select
#If jnasetup = True Then
Des = Des & "(" & From & "-" & From + reqbits - 1 & ")"
#End If
    clsField.CallingRoutine = "DetailOut"
    clsField.Des = Des
    clsField.Member = Member
    clsField.From = From
    clsField.reqbits = reqbits
    clsField.Arg = Arg
    clsField.Arg1 = Arg1
'   If RetValCol <> 0 Then Call DetailLineOut(Des, Val, ValDes)
'this will be 0 even if arg is not passed, note cant then return col 0
    Select Case Retvalcol
        Case Is = 0
            If NmeaRcv.Option1(4).Value = False Then
                Call DetailLineOut(Des, Val, Valdes, Bold) 'main output
            End If
            DetailOut = Val
        Case Is = 1
            DetailOut = Des
        Case Is = 2
            DetailOut = Val
        Case Is = 3
            DetailOut = Valdes
        Case Is = 4
            If AllFieldsNo > UBound(AllFields) Then ReDim Preserve AllFields(AllFieldsNo)
            AllFields(AllFieldsNo) = Val
            AllFieldsNo = AllFieldsNo + 1
            DetailOut = Val
    End Select

Check_Additional:
'Have we to output additional info from the passed field
    Select Case Member
'output the MID
    Case Is = "mmsi"    'this is NOT necessarily the FromMMSI
        Call DetailOut(Retvalcol, "MID", "mid", From, reqbits)
    
    Case Is = "aidata"
'this call is re-entrant into detailout
'the data part is after the 16 bits of the application identifier (dac+fi)
        If reqbits >= 16 Then
            Call Application(Retvalcol, From, reqbits)
        End If
'#If False Then
    Case Is = "text"
    If Len(Val) > 18 Then   'Only in display split into lines of 18 characters
                            'val must remain the complete string
        With DetailDisplay
            .TextMatrix(.Rows - 1, 2) = Left$(Val, 18)
            i = 18 + 1  'Next line start
            Do While i <= Len(Val)
                If i + 18 <= Len(Val) Then
                    j = i + 18
                Else
                    j = Len(Val)
                End If
'Debug.Print Mid$(Val, i, j - i + 1)
                Call DetailLineOut("", Mid$(Val, i, j - i + 1), "")
                i = j + 1
            Loop
        End With
    End If
    If SpareBitsToBoundary > 0 Then
        Call DetailOut(Retvalcol, "Byte boundary filler", "spare", clsSentence.AisPayloadBits - SpareBitsToBoundary + 1, SpareBitsToBoundary)
    End If
'#End If
    End Select  'additional output
'return value depens on col required (if any). Normally used for Field output
'also used for re-entrant calls
End Function

Function Dac366Out(Retvalcol As Long, _
Des As String, _
Member As String, From As Long, _
Optional reqbits As Long, _
Optional Arg As String, _
Optional Arg1 As String) As String

Dim wlong As Long   'working long us if called more than once
Dim wSi As Single
Dim Val As String
Dim Valdes As String
Dim Minus As Boolean
Dim kb As String
Dim Bold As Boolean
    Select Case Member
    Case Is = "reporttype"
        Val = pLong(From, reqbits)
        Valdes = ReportType366Name(Val)
    Case Is = "owner"
        Val = pLong(From, reqbits)
        Valdes = Owner366Name(Val)
    Case Is = "windspeed"
        wlong = pLong(From, reqbits)
        If Arg = "" Then Arg = 1 'default .1
        Val = Format$(wlong / (10 ^ Arg), ArgFmt(Arg))
        Select Case wlong
        Case Is <= 120
            Valdes = "Knots"
        Case Is = 121
            Valdes = "> " & Format$(wlong / (10 ^ Arg), ArgFmt(Arg)) & " Knots"
        Case Is = 122
            Valdes = "Not available (default)"
        Case Is = 127
            Valdes = "(don't use)"
        Case Else
            Valdes = "Reserved for future use"
        End Select
    Case Is = "winddirection"
        wlong = pLong(From, reqbits)
        Val = Format$(wlong / (10 ^ Arg), ArgFmt(Arg))
        Select Case wlong
        Case Is <= 359
            Valdes = "Degrees"
'        Case Is = 360  'cant be over 360 for direction
'            ValDes = "> " & Format$(wlong / (10 ^ Arg), ArgFmt(Arg)) & " Knots"
        Case Is = 360
            Valdes = "Not available (default)"
        Case Is = 511
            Valdes = "(don't use)"
        Case Else
            Valdes = "Reserved for future use"
        End Select
    Case Is = "visibility"  '366-33-9 visibility
        wlong = pLong(From, reqbits)
'If wlong = 128 Then MsgBox Arg1 & ":" & ItoBits(pLong(From, reqbits))
        If Arg = "" Then Arg = 1 'default .1
        Val = Format$(wlong / (10 ^ Arg), ArgFmt(Arg))
        Select Case wlong
        Case Is <= 240
            Valdes = "Nautical Miles"
        Case Is = 241
            Valdes = "> " & Format$(wlong / (10 ^ Arg), ArgFmt(Arg)) & " Nautical Miles"
        Case Is = 242
             Valdes = "Not available"
        Case Is = 243
             Valdes = "sensor not available (default)"
        Case Is <= 254
                Valdes = "Reserved for future use"
        Case Is = 255
                Valdes = "(don't use)"
        End Select
    Case Is = "temperature"
        If Arg = "" Then Arg = 1 'default .1
'signed integer
        wSi = pSi(From, reqbits)
        Val = Format$(wSi / (10 ^ Arg), ArgFmt(Arg))
        Select Case Arg1
            Case Is = "air" '366-33-9
'SI -60.0 to 60.0 n/a=-1024 (11 bits)
            Select Case wSi
            Case Is = -1024
                Valdes = "Not available"
            Case Is >= 601, Is <= -601
                Valdes = "Reserved for future use"
            Case Else
                Valdes = Val & " " & Chr$(176) & "C"
            End Select
        Case Is = "dew" '366-33-9   (20 degree offset)
            wlong = pLong(From, reqbits)
            Val = Format$(wlong / (10 ^ Arg), ArgFmt(Arg))
'wlong = 0 to 1023
            Val = Format$(wlong / (10 ^ Arg) - 20, ArgFmt(Arg))
'val = -20.0 to 50.0
            Select Case wlong
            Case Is <= 700
                Valdes = Val & " " & Chr$(176) & "C"
            Case Is = 701
                Valdes = "data unavailable"
            Case Is = 1023
                Valdes = "(don't use)"
            Case Else
                Valdes = "Reserved for future use"
            End Select
        Case Is = "water" '366-33-9   (10 degree offset)
            wlong = pLong(From, reqbits)
            Val = Format$(wlong / (10 ^ Arg), ArgFmt(Arg))
'wlong = 0 to 1023
            Val = Format$(wlong / (10 ^ Arg) - 10, ArgFmt(Arg))
'val = -10.0 to 60.0
            Select Case wlong
            Case Is <= 600
                Valdes = Val & " " & Chr$(176) & "C"
            Case Is = 601
                Valdes = "data unavailable (default)"
'            Case Is = 1023
'                ValDes = "(don't use)"
            Case Else
                Valdes = "Reserved for future use"
            End Select
        End Select
    Case Is = "pressure"    '366-33-9
       wlong = pLong(From, reqbits)
       Val = wlong + 799
        Select Case wlong
        Case Is = 0
            Valdes = "799 hPa or less!"
        Case Is <= 401
            Valdes = Val & " hPa" '800-1200
        Case Is = 402
            Valdes = "1201 hPa or greater"
        Case Is = 403
            Valdes = "Not available (default)"
        Case Is = 511
            Valdes = "don't use"
        Case Else
            Valdes = "reserved for future use"
        End Select
    Case Is = "salinity"    '366-33-(7,8,9)
        wlong = pLong(From, reqbits)
'If wlong = 128 Then MsgBox Arg1 & ":" & ItoBits(pLong(From, reqbits))
        If Arg = "" Then Arg = 1 'default .1
        Val = Format$(wlong / (10 ^ Arg), ArgFmt(Arg))
        Valdes = "%% (ppt)"
        Select Case wlong
        Case Is <= 500
            Valdes = Val & " " & Valdes
        Case Is = 501
            Valdes = "> " & Val & " %% (ppt)"
        Case Is = 502
            Valdes = "data not available"
        Case Is = 503
            Valdes = "sensor not available (default)"
        Case Is = 511
            Valdes = "don't use"
        Case Else
            Valdes = "Reserved for future use"
        End Select
    Case Is = "currentspeed" '366-33-4
        wlong = pLong(From, reqbits)
'If wlong = 128 Then MsgBox Arg1 & ":" & ItoBits(pLong(From, reqbits))
        Val = Format$(wlong / (10 ^ Arg), ArgFmt(Arg))
        Valdes = "Knots"
        Select Case wlong
        Case Is = 246
                Valdes = "> " & Format$(wlong / (10 ^ Arg), ArgFmt(Arg)) & " " & Valdes
        Case Is = 247
                Valdes = "Not available (default)"
        Case Is = 255
            Valdes = "don't use"
        Case Is >= 248   '248 to 254
                Valdes = "Reserved for future use"
        End Select
    Case Is = "currentdirection"    '366-33-4
        wlong = pLong(From, reqbits)
                'if bits are 10 max val is 740 limit is 1023
                'hence must be in 1/2 degree increments
        Val = Format$(wlong / (10 ^ Arg), ArgFmt(Arg))
        Valdes = Chr$(176) & " (degrees)"
        Select Case wlong
        Case Is = 360
            Valdes = "Not available (default)"
        Case Is = 511
            Valdes = "don't use"
        Case Is >= 361   '361-510
                Valdes = "Reserved for future use"
        End Select
    Case Is = "currentlevel"    '366-33-4
        wlong = pLong(From, reqbits)
'If wlong = 128 Then MsgBox Arg1 & ":" & ItoBits(pLong(From, reqbits))
        Val = Format$(wlong / (10 ^ Arg), ArgFmt(Arg))
        Valdes = "Meters"
        Select Case wlong
        Case Is = 361
                Valdes = "> " & Format$(wlong / (10 ^ Arg), ArgFmt(Arg)) & " " & Valdes
        Case Is = 362
            Valdes = "Not available (default)"
        Case Is = 511
            Valdes = "don't use"
        Case Is >= 361   '361-510
                Valdes = "Reserved for future use"
        End Select
    Case Is = "currentvector"   '366-33-5
        wSi = pSi(From, reqbits)
        Val = Format$(wSi / (10 ^ Arg), ArgFmt(Arg))
        Valdes = "Knots"
        Select Case wSi
        Case Is = -251
            Valdes = "< " & Format$(wSi + 1, ArgFmt(Arg)) & " " & Valdes
        Case Is = 251
            Valdes = "> " & Format$(wSi - 1, ArgFmt(Arg)) & " " & Valdes
        Case Is = -256
            Valdes = "Data unavailable (default)"
        Case Is > 251, Is < -251 '252 to 256 or -255 to -252
            Valdes = "undefined in specification"
        End Select
    Case Is = "waveheight"    '366-33-7
        wlong = pLong(From, reqbits)
'If wlong = 128 Then MsgBox Arg1 & ":" & ItoBits(pLong(From, reqbits))
        Val = Format$(wlong / (10 ^ Arg), ArgFmt(Arg))
        Valdes = "Meters"
        Select Case wlong
        Case Is = 246
                Valdes = ">= " & Format$(wlong / (10 ^ Arg), ArgFmt(Arg)) & " " & Valdes
        Case Is = 247
            Valdes = "Not available (default)"
        Case Is = 255
            Valdes = "don't use"
        Case Is >= 248   '248-254
                Valdes = "Reserved for future use"
        End Select
    Case Is = "depth"    '366-33-7
        wlong = pLong(From, reqbits)
'If wlong = 128 Then MsgBox Arg1 & ":" & ItoBits(pLong(From, reqbits))
        Val = Format$(wlong / (10 ^ Arg), ArgFmt(Arg))
        Valdes = "Meters"
        Select Case wlong
        Case Is = 121
                Valdes = ">= " & Format$(wlong / (10 ^ Arg), ArgFmt(Arg)) & " " & Valdes
        Case Is = 122
            Valdes = "Not available (default)"
        Case Is = 127
            Valdes = "don't use"
        Case Is >= 123   '248-254
            Valdes = "Reserved for future use"
        End Select
    Case Is = "air"    '366-33-7
        wlong = pLong(From, reqbits)
'If wlong = 128 Then MsgBox Arg1 & ":" & ItoBits(pLong(From, reqbits))
        Val = Format$(wlong / (10 ^ Arg), ArgFmt(Arg))
        Valdes = "Meters"
        Select Case wlong
        Case Is = 8191
            Valdes = ">= " & Format$(wlong / (10 ^ Arg), ArgFmt(Arg)) & " " & Valdes
        Case Is = 0
            Valdes = "Not available (default)"
        End Select
    Case Is = "sialtitude"   '367-33-0
        wSi = pSi(From, reqbits)
        Val = Format$(wSi / (10 ^ Arg), ArgFmt(Arg))
        Valdes = "Meters"
        Select Case wSi
        Case Is < -2000
            Valdes = "Undefined in specification"
        Case Is = 2001
            Valdes = "> " & Format$(wSi - 1, ArgFmt(Arg)) & " " & Valdes
        Case Is = 2002
            Valdes = "Data unavailable (default)"
        Case Is = 2047
            Valdes = "don't use"
        Case Is >= 2003
            Valdes = "Reserved for future use"
        End Select
'1-11 (I think this is correct but not documented)
    Case Is = "timeout"
        Val = pLong(From, reqbits)
        Valdes = Timeout366Name(Val)
    Case Is = "sensordata"
        Val = pLong(From, reqbits)
        Valdes = SensorData366Name(Val)
    Case Is = "duration"
        Val = pLong(From, reqbits)
        Valdes = "Minutes"
        If Val = 0 Then
            Valdes = "Cancel Forecast (default)"
        Else
            Val = Val & " minutes"
        End If
    Case Is = "updown"
        Val = pLong(From, reqbits)
        Valdes = UpDown366Name(Val)
    Case Is = "lockid"
'v136        Val = RTrim$(Replace(p6bit(From, reqbits), "@", ""))
        Val = p6bit(From, reqbits)  'v136
        Select Case Val
            Case Is = "SLS_L01"
                Valdes = "Welland Canal Lock 1"
            Case Is = "SLS_L02"
                Valdes = "Welland Canal Lock 2"
            Case Is = "SLS_L03"
                Valdes = "Welland Canal Lock 3"
            Case Is = "SLS_L4E"
                Valdes = "Welland Canal Lock 4 East"
            Case Is = "SLS_L4W"
                Valdes = "Welland Canal Lock 4 West"
            Case Is = "SLS_L6E"
                Valdes = "Welland Canal Lock 6 East"
            Case Is = "SLS_L6W"
                Valdes = "Welland Canal Lock 6 West"
            Case Is = "SLS_L07"
                Valdes = "Welland Canal Lock 7"
            Case Is = "SLS_L08"
                Valdes = "Welland Canal Lock 8"
            Case Is = "SLS_IRP"
                Valdes = "Iroquois Lock"
            Case Is = "SLS_IKE"
                Valdes = "Eisenhower Lock"
            Case Is = "SLS_SNL"
                Valdes = "Snell Lock"
            Case Is = "SLS_BO3"
                Valdes = "Beauharnois Lock 3"
            Case Is = "SLS_BO4"
                Valdes = "Beauharnois Lock 4"
            Case Is = "SLS_CSC"
                Valdes = "Cote Saint Catherine Lock"
            Case Is = "SLS_SLB"
                Valdes = "Saint Lambert Lock"
            Case Else
                Valdes = "Unknown Lock"
           End Select
    Case Is = "compasspoint"
        Val = pLong(From, reqbits)
        Select Case Val
            Case CInt(0)
                Valdes = "N"
            Case 23
                Valdes = "NNE"
            Case 45
                Valdes = "NE"
            Case 68
                Valdes = "ENE"
            Case 90
                Valdes = "E"
            Case 113
                Valdes = "ESE"
            Case 135
                Valdes = "SE"
            Case 158
                Valdes = "SSE"
            Case 180
                Valdes = "S"
            Case 203
                Valdes = "SSW"
            Case 225
                Valdes = "SW"
            Case 248
                Valdes = "WSW"
            Case 270
                Valdes = "W"
            Case 293
                Valdes = "WNW"
            Case 315
                Valdes = "NW"
            Case 338
                Valdes = "NNW"
            Case Else
                Valdes = "Invalid"
        End Select
    Case Is = "level"
        Val = pLong(From, reqbits)
        Valdes = Level366Name(Val)
    Case Is = "datum"
        Val = pLong(From, reqbits)
        If reqbits = 2 Then
            Valdes = DatumA366Name(Val)
        Else            '5 BITS
            Valdes = DatumB366Name(Val)
        End If
    Case Is = "conductivity"    '266-33-8
        wlong = pLong(From, reqbits)
'If wlong = 128 Then MsgBox Arg1 & ":" & ItoBits(pLong(From, reqbits))
        Val = Format$(wlong / (10 ^ Arg), ArgFmt(Arg))
        Valdes = "Siemens/meter"
        Select Case wlong
        Case Is = 701
                Valdes = "> " & Format$(wlong / (10 ^ Arg) - 1, ArgFmt(Arg)) & " " & Valdes
        Case Is = 702
            Valdes = "Data unavailable"
        Case Is = 703
            Valdes = "Sensor not available (default)"
        Case Is = 1023
            Valdes = "don't use"
        Case Is >= 704   '704-1022
                Valdes = "Reserved for future use"
        End Select
    Case Is = "decibars"
        wlong = pLong(From, reqbits)
        Val = Format$(wlong / (10 ^ Arg), ArgFmt(Arg))
        Valdes = "decibars"
        Select Case wlong
        Case Is = 60001
                Valdes = "> " & Format$(wlong / (10 ^ Arg) - 1, ArgFmt(Arg)) & " " & Valdes
        Case Is = 60002
            Valdes = "Data unavailable"
        Case Is = 60003
            Valdes = "Sensor not available (default)"
        Case Is = 65535
            Valdes = "don't use"
        Case Is >= 60004   '60004-65534
                Valdes = "Reserved for future use"
        End Select
    Case Is = "salinitytype"
        Val = pLong(From, reqbits)
        Valdes = SalinityType366Name(Val)
    Case Else
        Val = Member
        Valdes = "(" & From & ":"
        If IsMissing(reqbits) Then
            Valdes = Valdes & "none)"
        Else
            Valdes = Valdes & reqbits & ")"
        End If
        If IsMissing(Arg) Then
            Valdes = Valdes & " no Arg"
        Else
            Valdes = Valdes & " arg= " & Arg
        End If
        Exit Function
    End Select
    clsField.CallingRoutine = "Dac366Out"
    clsField.Des = Des
    clsField.Member = Member
    clsField.From = From
    clsField.reqbits = reqbits
    clsField.Arg = Arg
    clsField.Arg1 = Arg1
    Select Case Retvalcol
        Case Is = 0
            Call DetailLineOut(Des, Val, Valdes, Bold) 'main output
            Dac366Out = Val
        Case Is = 1
            Dac366Out = Des
        Case Is = 2
            Dac366Out = Val
        Case Is = 3
            Dac366Out = Valdes
        Case Is = 4
            If AllFieldsNo > UBound(AllFields) Then ReDim Preserve AllFields(AllFieldsNo)
            AllFields(AllFieldsNo) = Val
            AllFieldsNo = AllFieldsNo + 1
            Dac366Out = Val
    End Select
End Function

Function Dac1Out(Retvalcol As Long, _
Des As String, _
Member As String, From As Long, _
Optional reqbits As Long, _
Optional Arg As String, _
Optional Arg1 As String) As String

Dim wlong As Long   'working long us if called more than once
Dim wSi As Single
Dim Val As String
Dim Valdes As String
Dim Minus As Boolean
Dim kb As String
Dim Bold As Boolean
    Select Case Member
    Case Is = "aiavailable"
        Val = pLong(From, reqbits)
        Valdes = AiAvailable1Name(Val)
    Case Is = "airesponse"
        Val = pLong(From, reqbits)
        Valdes = AiResponse1Name(Val)
    Case Is = "signalstatus"
        Val = pLong(From, reqbits)
        Valdes = SignalStatus1Name(Val)
    Case Is = "signalservice"
        Val = pLong(From, reqbits)
        Valdes = "Coding is unclarified"
    Case Is = "shape"
        Val = pLong(From, reqbits)
        Valdes = Shape1Name(Val)
    Case Is = "areatype"
        Val = pLong(From, reqbits)
        Valdes = AreaType1Name(Val)
    Case Is = "sender"
        Val = pLong(From, reqbits)
        Valdes = Sender1Name(Val)
    Case Is = "route"
        Val = pLong(From, reqbits)
        Valdes = Route1Name(Val)
    Case Else
        Val = Member
        Valdes = "(" & From & ":"
        If IsMissing(reqbits) Then
            Valdes = Valdes & "none)"
        Else
            Valdes = Valdes & reqbits & ")"
        End If
        If IsMissing(Arg) Then
            Valdes = Valdes & " no Arg"
        Else
            Valdes = Valdes & " arg= " & Arg
        End If
        Exit Function
    End Select
    clsField.CallingRoutine = "Dac1Out"
    clsField.Des = Des
    clsField.Member = Member
    clsField.From = From
    clsField.reqbits = reqbits
    clsField.Arg = Arg
    clsField.Arg1 = Arg1
    Select Case Retvalcol
        Case Is = 0
            Call DetailLineOut(Des, Val, Valdes, Bold) 'main output
            Dac1Out = Val
        Case Is = 1
            Dac1Out = Des
        Case Is = 2
            Dac1Out = Val
        Case Is = 3
            Dac1Out = Valdes
        Case Is = 4
            If AllFieldsNo > UBound(AllFields) Then ReDim Preserve AllFields(AllFieldsNo)
            AllFields(AllFieldsNo) = Val
            AllFieldsNo = AllFieldsNo + 1
            Dac1Out = Val
    End Select
End Function

Function Dac200Out(Retvalcol As Long, _
Des As String, _
Member As String, From As Long, _
Optional reqbits As Long, _
Optional Arg As String, _
Optional Arg1 As String) As String
Dim wlong As Long   'working long us if called more than once
Dim wSi As Single
Dim Val As String
Dim Valdes As String
Dim Minus As Integer    '1 in LSB = + ,0 = -
Dim kb As String
Dim Bold As Boolean
    Select Case Member
    Case Is = "tug"
        Val = pLong(From, reqbits)
        Valdes = Tug200Name(Val)
'        Dac200Out = Val
    Case Is = "airdraught"
        wlong = pLong(From, reqbits)
        If Arg1 <> "" Then
            Valdes = Arg1
        Else
            Valdes = "Meters"
        End If
        If wlong = 0 Then Valdes = "Not used (default)"
        If wlong > 4000 Then Valdes = "Not used"
        Val = Format$(wlong / (10 ^ Arg), ArgFmt(Arg))
    Case Is = "number"
        wlong = pLong(From, reqbits)
        Valdes = Arg1
        If wlong = (2 ^ reqbits - 1) Then Valdes = "Unknown (default)"
        Val = wlong
    Case Is = "negative"
        wlong = pLong(From, reqbits - 1)
        Minus = pLong(From + reqbits - 1, 1)
        If Minus = 0 Then wlong = wlong * -1
        Valdes = Arg1
        If wlong = 0 Then Valdes = "Unknown (default)"
        Val = Format$(wlong / (10 ^ Arg), ArgFmt(Arg))
    Case Is = "signalform"
        Val = pLong(From, reqbits)
        Valdes = SignalForm200Name(Val)
    Case Is = "signalimpact"
        Val = pLong(From, reqbits)
        Valdes = SignalImpact200Name(Val)
    Case Is = "signalstatus"
        Val = pLong(From, reqbits)
        Valdes = SignalStatus200Name(Val)
    Case Else
        Val = Member
        Valdes = "(" & From & ":"
        If IsMissing(reqbits) Then
            Valdes = Valdes & "none)"
        Else
            Valdes = Valdes & reqbits & ")"
        End If
        If IsMissing(Arg) Then
            Valdes = Valdes & " no Arg"
        Else
            Valdes = Valdes & " arg= " & Arg
        End If
        Exit Function
    End Select
    clsField.CallingRoutine = "Dac200Out"
    clsField.Des = Des
    clsField.Member = Member
    clsField.From = From
    clsField.reqbits = reqbits
    clsField.Arg = Arg
    clsField.Arg1 = Arg1
    Select Case Retvalcol
        Case Is = 0
            Call DetailLineOut(Des, Val, Valdes, Bold) 'main output
            Dac200Out = Val
        Case Is = 1
            Dac200Out = Des
        Case Is = 2
            Dac200Out = Val
        Case Is = 3
            Dac200Out = Valdes
        Case Is = 4
            If AllFieldsNo > UBound(AllFields) Then ReDim Preserve AllFields(AllFieldsNo)
            AllFields(AllFieldsNo) = Val
            AllFieldsNo = AllFieldsNo + 1
            Dac200Out = Val
    End Select
End Function

Function Dac235Out(Retvalcol As Long, _
Des As String, _
Member As String, From As Long, _
Optional reqbits As Long, _
Optional Arg As String, _
Optional Arg1 As String) As String

Dim wlong As Long   'working long us if called more than once
Dim wSi As Single
Dim Val As String
Dim Valdes As String
Dim Minus As Boolean
Dim kb As String
Dim Bold As Boolean
    Select Case Member
    Case Is = "volts"
        wlong = pLong(From, reqbits)
        Val = wlong
        If Val = 0 Then
            Valdes = "Not Used"
        Else
            Valdes = Format$(wlong * 5 / (10 ^ Arg), ArgFmt(Arg)) & " Volts"
        End If
    Case Is = "racon"
        Val = pLong(From, reqbits)
        Valdes = Racon235Name(Val)
    Case Is = "light"
        Val = pLong(From, reqbits)
        Valdes = Light235Name(Val)
    Case Is = "alarm"
        Val = pLong(From, reqbits)
        Valdes = Alarm235Name(Val)
    Case Is = "msgid"
        Val = pLong(From, reqbits)
        Valdes = Arg
    Case Is = "unitid"
        Val = pLong(From, reqbits)
        If Val = 0 Then Valdes = "peer-to-peer connection (default)"
        If Val = 63 Then Valdes = "reserved"
     Case Is = "biit"
        Val = pLong(From, reqbits)
        Valdes = Biit235Name(Val)
     Case Is = "extfld"
        Val = pLong(From, reqbits)
        Valdes = ExtFld235Name(Val)
     Case Is = "highlow"
        Val = pLong(From, reqbits)
        Valdes = HighLow235Name(Val)
     Case Is = "contentcontrol"
        Val = pLong(From, reqbits)
        Valdes = reqbits & " bit"
        If reqbits <> 1 Then Valdes = Valdes & "s"
        Valdes = Valdes & Arg
  Case Else
        Val = Member
        Valdes = "(" & From & ":"
        If IsMissing(reqbits) Then
            Valdes = Valdes & "none)"
        Else
            Valdes = Valdes & reqbits & ")"
        End If
        If IsMissing(Arg) Then
            Valdes = Valdes & " no Arg"
        Else
            Valdes = Valdes & " arg= " & Arg
        End If
        Exit Function
    End Select
    clsField.CallingRoutine = "Dac235Out"
    clsField.Des = Des
    clsField.Member = Member
    clsField.From = From
    clsField.reqbits = reqbits
    clsField.Arg = Arg
    clsField.Arg1 = Arg1
    Select Case Retvalcol
        Case Is = 0
            Call DetailLineOut(Des, Val, Valdes, Bold) 'main output
            Dac235Out = Val
        Case Is = 1
            Dac235Out = Des
        Case Is = 2
            Dac235Out = Val
        Case Is = 3
            Dac235Out = Valdes
        Case Is = 4
            If AllFieldsNo > UBound(AllFields) Then ReDim Preserve AllFields(AllFieldsNo)
            AllFields(AllFieldsNo) = Val
            AllFieldsNo = AllFieldsNo + 1
            Dac235Out = Val
    End Select
End Function

'example of a dac specific call
#If False Then
Function Dac999Out(Retvalcol As Long, _
Des As String, _
Member As String, From As Long, _
Optional reqbits As Long, _
Optional Arg As String, _
Optional Arg1 As String) As String

Dim wlong As Long   'working long us if called more than once
Dim wSi As Single
Dim Val As String
Dim Valdes As String
Dim Minus As Boolean
Dim kb As String
Dim Bold As Boolean
    Select Case Member
    Case Is = "example"
        Val = pLong(From, reqbits)
        Valdes = Example999Name(Val)
    Case Else
        Val = Member
        Valdes = "(" & From & ":"
        If IsMissing(reqbits) Then
            Valdes = Valdes & "none)"
        Else
            Valdes = Valdes & reqbits & ")"
        End If
        If IsMissing(Arg) Then
            Valdes = Valdes & " no Arg"
        Else
            Valdes = Valdes & " arg= " & Arg
        End If
        Exit Function
    End Select
    clsField.CallingRoutine = "Dac999Out"
    clsField.Des = Des
    clsField.Member = Member
    clsField.From = From
    clsField.reqbits = reqbits
    clsField.Arg = Arg
    clsField.Arg1 = Arg1
    Select Case Retvalcol
        Case Is = 0
            Call DetailLineOut(Des, Val, Valdes, Bold) 'main output
            Dac999Out = Val
        Case Is = 1
            Dac999Out = Des
        Case Is = 2
            Dac999Out = Val
        Case Is = 3
            Dac999Out = Valdes
        Case Is = 4
            If AllFieldsNo > UBound(AllFields) Then ReDim Preserve AllFields(AllFieldsNo)
            AllFields(AllFieldsNo) = Val
            AllFieldsNo = AllFieldsNo + 1
            Dac999Out = Val
    End Select
End Function
#End If

'uses PayloadBytes

'set up here if called more than once from detail and not from any other form
'If called from more than one form set up in DecodeDefs as Public Function
'des is used because the value must be formatted as well as the decription
Function TurnFmt(ByVal Turn As Long, ByVal Des As String) As String
Dim OutStr As String
If Des = "V" Then
Else
End If
TurnFmt = OutStr
End Function

Function SecondDes(Val As String)
SecondDes = "Second of UTC timestamp"
If CInt(Val) = 60 Then SecondDes = "not available (default)"
If CInt(Val) = 61 Then SecondDes = "Positioning system is in manual input mode"
If CInt(Val) = 62 Then SecondDes = "Electronic Positioning Fixing System operates in estimated (dead reckoning) mode"
If CInt(Val) = 63 Then SecondDes = "System is inoperative"
End Function

Function DetailLineOut(Des As String, Val As String, Valdes As String, Optional Bold As Boolean, Optional Tag As String, Optional RcvTime As String)
Dim arry(10) As String

#If jnasetup = True Then
'If Des = "Fill bits" Then Stop
#End If

If NmeaRcv.Option1(4).Value = False Then  'some detail is to be OutputFieldList
    With DetailDisplay
        If Des <> "" Or Val <> "" Or Valdes <> "" Then
            .AddItem (vbTab & Des & vbTab & Val & vbTab & Valdes)
            If .TextMatrix(1, 1) = "" Then .RemoveItem 1
            If Bold = True Then
                .Row = .Rows - 1
                .col = 1
                .CellFontBold = True
               .FocusRect = flexFocusLight
                Bold = False
            End If
        
'keep the source in col 0, no dynamic information pertaining to this sentence
'only where to get the information from, as this sentence will not be the same
'as the sentence
'used to get the information from the actual sentence were decoding
'next line temp to see all set up detail must be on
'if no calling routine we cannot output this field
            If clsField.CallingRoutine <> "" Then
                arry(0) = FieldKeyFromSentence(clsField.CallingRoutine) 'gets key from current clsSentence
                arry(1) = clsField.Des      'Required when calling routine is called
                                        'and to display on Field List
                arry(2) = clsField.CallingRoutine   'PayloadDetail,NmeaOut,DetailOut
                arry(3) = clsField.Member           'Arguments for calling routine
                arry(4) = clsField.From
                arry(5) = clsField.reqbits
                arry(6) = clsField.Arg
                arry(7) = clsField.Arg1
                .TextMatrix(.Rows - 1, 0) = Join(arry, vbTab)
                Set clsField = Nothing
            Else
DetailDisplay.Redraw = True

'Stop    'debug to see what we cant output
            End If
        End If      'some values
    End With
End If      'detail out turned on
'Des = ""
'Val = ""
'ValDes = ""
End Function

Public Sub Clear_DetailDisplay()
Dim i As Integer

With DetailDisplay
    Do While .Rows > 2
        .RemoveItem 2
    Loop
    For i = 1 To 3  'clear first row
        .TextMatrix(1, i) = ""
    Next i
End With
End Sub

Public Sub Fill_DetailDisplay()
With DetailDisplay
    Do While .Rows < 40
        .AddItem ""
    Loop
End With

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'set display to none
'V142 Call DisplayQueryUnload(Me.Name, Cancel, UnloadMode)

    If UnloadMode = vbFormControlMenu Then  'V3.4.143 User clicked (X)
        If FormLoaded("NmeaRcv") Then   'to prevent NmeaRcv being reloaded
            If NmeaRcv.Visible = True Then
                NmeaRcv.Option1(4).Value = True
                If Detail.Visible = True Then Detail.Hide
                Cancel = True   'just hide
            End If
        End If
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbNormal Then
        With DetailDisplay
            .Move 0, 0, ScaleWidth, ScaleHeight
            .ColWidth(.Cols - 1) = ScaleWidth
        End With
    End If
End Sub

Private Sub mnuCopyToClipBoard_Click() 'when Copy to Clipboard is clicked
Dim i As Long
Dim j As Long
Dim kb As String

With DetailDisplay
    For i = 1 To .Rows - 1
        If Not (.TextMatrix(i, 0) = "" And .TextMatrix(i, 2) = "") Then
            For j = 1 To .Cols - 1
            kb = kb + QuotedString(.TextMatrix(i, j), ",") & ","
            Next j
            kb = kb + vbCrLf
        End If
    Next i
End With
Clipboard.Clear
Clipboard.SetText kb
End Sub

Private Sub mnuCopy_Click()
    With DetailDisplay
        Clipboard.Clear
        Clipboard.SetText .TextMatrix(.RowSel, .ColSel)
    End With
End Sub

Private Sub DetailDisplay_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
    If Button = vbRightButton Then
        With DetailDisplay
'.BackColor = vbRed
'Select the Cell when the Cell is clicked
'not when the Popup menu is clicked
            .Row = .MouseRow
            .col = .MouseCol
            .RowSel = .Row    'force selection to one cell
            .ColSel = .col
            .CellBackColor = vbYellow
'Columns not displayed are all to the laft of the AisMsgType column
            PopupMenu mnuPopup
            .CellBackColor = vbWhite
        End With
    End If
End Sub

Private Sub mnuCreateTag_Click()
Dim arry() As String
Dim FieldKey As String   'arry(0)
'Des  & Name                       'arry(1)
Dim Source As String        'arry(2)
'Member= arry(3). From=Arry(4), ReqBits=Arry(5), Arg-Arry(6) Arg1=Arry(7)
Dim AisMsgType As String    'From FieldKey
Dim Dac As String           'diito
Dim Fi As String           'diito
Dim Fiid As String           'diito
Dim Retvalcol As String
Dim DefaultTag As String
Dim UserMsgName As String
Dim kb As String
Dim FieldKeyItem As String

'suspend here rather than on loading fieldInput form
'because problem with modal form loosing focus
'    If DecodingState <> 0 Then Exit Sub
'    DecodingState = DecodingState Or 6
    Processing.Suspended = True
    Processing.InputOptions = True
    Call TreeFilter.HideTreeFilter  'to stop clicking through form without focus, also Hides ListFilters
    With DetailDisplay
'Columns not displayed are all to the laft of the AisMsgType column
        arry = Split(.TextMatrix(.Row, 0), vbTab)
 'debug .ColWidth(0) = 5000
        If UBound(arry) = -1 Then ReDim arry(1) 'create a blank first element
        If arry(0) = "" Then   'blank column"
            MsgBox "There is no source Field for this selection" & vbCrLf & "Tag will not be created", vbInformation, "Create Field Tag"
            .CellBackColor = vbRed
            Exit Sub
        End If
        .CellBackColor = vbGreen
        Source = arry(2)
        kb = "Source=" & Source & vbCrLf
        FieldKey = arry(0)  'V130 was wrong (was getting from FieldKeyFromSentence(Source))
        kb = kb & "FieldKey=" & FieldKey & vbCrLf
        FieldInput.Label2 = .TextMatrix(.Row, 1) & " - " & .TextMatrix(0, .col) 'loads form input
        kb = kb & "FieldInput=" & FieldInput.Label2 & vbCrLf
        If Left$(arry(0), 2) <> "  " Then
'Ais Message
            DefaultTag = arry(3) & "_" & Trim$(Left$(arry(0), 2)) & "_" & .col
        Else
'Nmea Message
            DefaultTag = arry(3) & "_" & arry(4) & "_" & .col
        End If
        kb = kb & "DefaultTag=" & DefaultTag & vbCrLf
        
        FieldInput.Text1 = DefaultTag  'member (this is the default)
        FieldInput.Show vbModal, Detail
'        FieldInput.Text1.SetFocus
'MsgBox "return from field input " & FieldInput.Cancel & " " & FieldInput.Command2.Value
        If FieldInput.Cancel = False Then
'        If InputBox("Add field to Output List ? ", vbOKCancel, "Add") = vbOK Then
'0=msgkey,1=source,2=member,3=from,4=reqbits,5=arg,6=arg1,7=column
'8=MsgType,9=dac,10=fi,11=fiid      'displayed fields
'12=TAG,13=valdes - name,14=value probably not required

'If the first 2 characters of the field key not blank
'These must only be used for the Display of the Fields
'replace this with SplitFieldKey
            Call SplitFieldKey(FieldKey, AisMsgType, Dac, Fi, Fiid)
            Retvalcol = .col
            kb = kb & "FieldKey=" & FieldKey & vbCrLf
            kb = kb & "Dac=" & Dac & vbCrLf
            kb = kb & "Fi=" & Fi & vbCrLf
            kb = kb & "Fiid=" & Fiid & vbCrLf
            kb = kb & "Column=" & Retvalcol & vbCrLf
'Note Field 2 (Des on detail display) is not kept on FieldList
'This means that the mapping from ini file to FieldList differ
            FieldKeyItem = FieldKey & vbTab & Source & vbTab _
            & arry(3) & vbTab & arry(4) & vbTab & arry(5) & vbTab _
            & arry(6) & vbTab & arry(7) & vbTab & Retvalcol & vbTab _
            & AisMsgType _
            & vbTab & Dac & vbTab & Fi & vbTab & Fiid & vbTab _
            & UserFieldTagName & vbTab & arry(1)
            
            TreeFilter.FieldList.AddItem FieldKeyItem
            kb = kb & "FieldKeyItem=" & Replace(FieldKeyItem, vbTab, "<tab>") & vbCrLf
'V130 FieldInput.Text1 changed to UserFieldTagName because FieldInput modal form now unloaded
'in v128 so UserFieldTagName must be kept on form Detail

'remove initial blank row
'TreeFilter.Show 'debug to check added correctly (dont leave set else scheduled secs is wrong)
            If TreeFilter.FieldList.TextMatrix(1, 0) = "" Then TreeFilter.FieldList.RemoveItem 1
            TreeFilter.Show     'display edit form
        End If
        Unload FieldInput
#If jnasetup = True Then
MsgBox kb
#End If
        Call ResetTags("FieldList")     'Set TagList to FieldList
        Call TreeFilter.CheckOutputOptions
    End With
       
'    DecodingState = DecodingState Xor 6
    Call ResumeProcessing("InputOptions")
End Sub

