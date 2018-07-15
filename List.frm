VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form List 
   Caption         =   "Summary"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12645
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "List.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   12645
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid NmeaDecodeList 
      Height          =   5025
      Left            =   360
      TabIndex        =   0
      Top             =   0
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   8864
      _Version        =   393216
      Rows            =   20
      FixedCols       =   0
      BackColorBkg    =   -2147483636
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "Popup Menu"
      Enabled         =   0   'False
      Visible         =   0   'False
      Begin VB.Menu mnuCut 
         Caption         =   "Cut"
         Enabled         =   0   'False
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy to Clop Board as CSV"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
   End
End
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Adds clsSentence To NmeaDecodeList
'Called by ProcessSentence, SchedulerOutput
Public Sub AddToNmeaSummaryList() 'faster than above
'if there are any items in the nmea buffer not decoded, they will ber decoded
'and put in the decoded list. Items decoded are placed into decodelistrow
'by the nmeadecode routine. DecodeListRow.UnpackedNmea must contain the original nmea sentence.
Dim Comments As String
Dim i As Long
Dim iRow As Long 'row were inserting into
Dim LowCount As Long    'Low Buffer Level

'Show form if required when data received (UDP & Serial)
'Hide ONLY when option clicked or in QueryUnload
    If cmdNoWindow = False And List.Visible = False And NmeaRcv.Option1(0).Value = False Then
        List.Show
    End If
    
    With List.NmeaDecodeList
        .ScrollTrack = False
        
        iRow = .Rows  '1 fixed row at the top  (base is 0)
'If row before we the one we are going top use is blank, use the one before
        If .TextMatrix(iRow - 1, 0) = "" Then
'there is at least one blank row before the end of the rows
'find first blank line (intial fill is top to bottom)
            iRow = 1    '1 changed v125 1st entry was not displayed
            Do Until .TextMatrix(iRow, 0) = ""
                iRow = iRow + 1
            Loop
            .TopRow = 1    'Top Visible Row (excepting fixed rows)
        Else
            .AddItem clsSentence.FullSentence   'was nmea
            .TopRow = .Rows - .FixedRows
        End If
'note list displays non AIS sentences
        .TextMatrix(iRow, 0) = clsSentence.FullSentence 'was nmea
        If clsSentence.NmeaSentenceType <> "" Then
            .TextMatrix(iRow, 1) = clsSentence.NmeaSentenceType
        Else
            .TextMatrix(iRow, 1) = "None"
'MsgBox CtrlToString(clsSentence.FullSentence)
            If clsCb.Block <> "" Then
                Comments = "Comments Only " & clsCb.Block
            End If
        End If
        .TextMatrix(iRow, 2) = clsSentence.AisMsgFromMmsi
        .TextMatrix(iRow, 3) = clsSentence.AisMsgType
        .TextMatrix(iRow, 4) = clsSentence.AisMsgDac
        .TextMatrix(iRow, 5) = clsSentence.AisMsgFi
        .TextMatrix(iRow, 6) = clsSentence.AisMsgFiId
        .TextMatrix(iRow, 7) = clsSentence.VesselName
        With clsSentence
            If .CRCerrmsg = "" Then
                If .IsAisSentence Then
                    Comments = AisMsgTypeName(NullToZero(.AisMsgType))
                Else
                    Comments = IecFormatDes(.IecFormat)
                End If
#If jnasetup = True Then    'Display Payload bits on summary
    Comments = Comments & "(" & clsSentence.AisPayloadBits & ")"
#End If
                If .IecEncapsulationComments <> "" Then
                    Comments = Comments & " " & .IecEncapsulationComments
                End If
                If .PayloadReassemblerComments <> "" Then
                    Comments = Comments & " " & .PayloadReassemblerComments
                End If
            Else
                Comments = .CRCerrmsg
            End If
        End With
'Note if string starts with ( then rest of text is not displayed
'.TextMatrix(iRow, 8) = "(Test"
        .TextMatrix(iRow, 8) = Comments
 'Debug.Print clssentence.nmeasentence
'not sure what this is for
'        If NmeaRcv.cbDetail.Caption = "Detail Off" Then
'            .Row = iRow
'        End If

        If .Rows > MaxNmeaDecodeListCount Then
            LowCount = MaxNmeaDecodeListCount  '4/5 of static variable causes overflow
            LowCount = LowCount * 4 / 5
            Do Until .Rows < LowCount
                .RemoveItem 1
            Loop
        End If
End With    'NmeaRcv.DecodeList

End Sub

Private Sub Form_Load()
Dim i As Long
Dim tot As Long

'position form at below nmearcv
Top = NmeaRcv.Top + NmeaRcv.Height '- 200    '-200 is temp
Left = NmeaRcv.Left

With NmeaDecodeList
    .Cols = 9
    .FormatString = "<Nmea|^Sentence|^MMSI|^Message Type|^DAC|^FI|^ID|Vessel Name|Comments"
    .ColWidth(0) = 0
    .ColWidth(2) = 1000
    .ColWidth(4) = 500
    .ColWidth(5) = 500
    .ColWidth(6) = 500
    .ColWidth(7) = 2500
    .ColWidth(8) = 5000
   tot = 0
    For i = 0 To .Cols - 1
        tot = tot + .ColWidth(i)
    Next i
    .Width = tot
'    Tot = 350
'    For i = 0 To .Rows - 1
'        Tot = Tot + .RowHeight(i)
'    Next i
'    .Height = Tot
'code from ReceivedData
'the vertical scrollbar (when it appears) will be positioned against the RHS
'of the .width of the flexgrid
'If the rightmost column is against the rhs of the flexgrid, the scroll bar will
'effctively reduce the size of the rightmost column.
    .Width = Width - 450    'alter to position vertical scroll bar
                            'to RHS of form frame
    .ColWidth(.Cols - 1) = .Width   'set last col width leaves
                            'just enough room to fit vertical scroll bar
    .Rows = 20              'last row must just fit in frame
    Height = .Height + 450 - 200 '75 'alter to position bottom of vertical
                            'scroll bar just within form frame
'default when displaying as received
    MaxNmeaDecodeListCount = MAXNMEADECODELISTCOUNT_Rcv
    .FocusRect = flexFocusLight
    .Row = 0
    .Col = 0

End With
List.Hide   'Always hide on inital load
End Sub


'when user click X hide form
'show form if user request it
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'set list display option button to none
'V142 Call DisplayQueryUnload(Me.Name, Cancel, UnloadMode)   'V142

    If UnloadMode = vbFormControlMenu Then  'V3.4.143 User clicked (X)
        If FormLoaded("NmeaRcv") Then   'stop NmeaRcv being reloaded on exit
            If NmeaRcv.Visible = True Then
                NmeaRcv.Option1(0).Value = True
                List.Hide
                Cancel = True
            End If
        End If
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbNormal Then
        With NmeaDecodeList
            .Move 0, 0, ScaleWidth, ScaleHeight
            .ColWidth(.Cols - 1) = ScaleWidth
        End With
    End If
End Sub

Private Sub mnuCopy_Click()
Dim i As Long
Dim j As Long
Dim kb As String
With NmeaDecodeList
    For i = 1 To .Rows - 1
        If .TextMatrix(i, 0) <> "" Then
            For j = 1 To .Cols - 1
            kb = kb + .TextMatrix(i, j) & ","
            Next j
            kb = kb + vbCrLf
        End If
    Next i
End With
Clipboard.Clear
Clipboard.SetText kb

End Sub

Private Sub NmeaDecodeList_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'suspend processing to prevent list display being updated
'SuspendProcessList = True
'ProcessSuspended = True
Processing.Suspended = True
Processing.InputOptions = True
End Sub

Private Sub NmeaDecodeList_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim InputSentence As String
Dim i As Integer
Dim j As Integer
Dim kb As String

'If DecodingState <> 0 Then Exit Sub
'DecodingState = DecodingState Or 3
If Button = vbLeftButton Then
    With List.NmeaDecodeList
        If Len(.TextMatrix(.Row, 0)) <> 0 Then
'create clsSentence, PayloadBytes() and Ship(names)
'Note if clicked when the last sentence was encapsulated, the encapsulation will remain set
'This is probably the correct way to handle it as we have no way of unsetting it reliably
'It is specifically unset at the EOF if reading a file, and by design does not get set
'if just clicking an encapsulated sentence in the list
            Call DecodeSentence(.TextMatrix(.Row, 0))
'if incomplete decode previous 9 rows (if enough)
'this is simpler than trying to work out how far back to go
'to see if the missing part is in the PayloadReassembler buffer
            If clsSentence.AisMsgPartsComplete = False Then
                If .Row - 9 < 1 Then
                    i = 1
                Else
                    i = .Row - 9
                End If
                For j = i To .Row
Debug.Print "Row " & j & " " & clsSentence.AisMsgPartsComplete
                    Call DecodeSentence(.TextMatrix(j, 0))
                Next j
            End If
Debug.Print "Last 9 Loaded, bits=" & clsSentence.AisPayloadBits
'turn on detail display (Select) to enable output
            NmeaRcv.Option1(13).Value = True
            Detail.Show     'must also show the form
            Call Detail.SentenceAndPayloadDetail(0)       'clsSentence, PayloadBytes required
            Detail.Show
            If Detail.DetailDisplay.TextMatrix(1, 1) = "" Then
                kb = "Stop Encountered in NmeaDecodeList > MouseUp" & vbCrLf
    Call WriteErrorLog(kb & vbCrLf & clsSentence.NmeaSentence)
'                Stop 'debug
            End If
        Else
            MsgBox "No Data for this line"
        End If
    End With
    Call NmeaOutBufClear
End If

If Button = vbRightButton Then
'This code is actioned when Clciking he list with the left button !!
'    For i = 1 To List.NmeaDecodeList.Rows - 1
'        If List.NmeaDecodeList.TextMatrix(i, 0) <> "" Then
'            For j = 0 To List.NmeaDecodeList.Cols - 1
'            kb = kb + QuotedString(List.NmeaDecodeList.TextMatrix(i, j), ",") & ","
'            kb = kb + List.NmeaDecodeList.TextMatrix(i, j) & ","
'            Next j
'            kb = kb + vbCrLf
'        End If
'    Next i
'    Clipboard.Clear
'    Clipboard.SetText kb
End If

'restart processing to enable list display updating
'SuspendProcessList = False
'Call ResumeProcess
'no point in clicking list if detail is hidden
'DecodingState = DecodingState Xor 4
Call ResumeProcessing("InputOptions")
End Sub
