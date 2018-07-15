VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form ReceivedData 
   Caption         =   "Nmea Input"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6885
   Icon            =   "ReceivedData.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   6885
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid NmeaRcvList 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   3836
      _Version        =   393216
      Rows            =   9
      Cols            =   1
      FixedCols       =   0
      BorderStyle     =   0
   End
End
Attribute VB_Name = "ReceivedData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim tot As Integer
Dim i As Integer
With NmeaRcvList
'the vertical scrollbar (when it appears) will be positioned against the RHS
'of the .width of the flexgrid
'If the rightmost column is against the rhs of the flexgrid, the scroll bar will
'effctively reduce the size of the rightmost column.
    .Width = Width - 100    'forces scroll bar to RHS of form
    .FormatString = "<Nmea Sentences Received"
    .ColWidth(.Cols - 1) = .Width
    
    .Rows = 9
    Height = .Height + 500  'force form height to enclose vertical scroll bar
    .GridLines = flexGridNone
End With
Left = NmeaRcv.Width    'place at right of nmearcv form
Top = NmeaRcv.Top       'in line with top of nmearcv
ReceivedData.Hide   'Always hide on initial load
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'V142 Call DisplayQueryUnload(Me.Name, Cancel, UnloadMode)
    If UnloadMode = vbFormControlMenu Then  'V3.4.143 User clicked (X)
        If ReceivedData.Visible = True Then
'removed v79 modal form error on exit cause by this ?     NmeaRcv.Option1(15).Value = True
            ReceivedData.Hide
            Cancel = True   'just hide
            NmeaRcv.Option1(15).Value = True    'Display NMEA Input = None
        End If
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbNormal Then
        With NmeaRcvList
            .Move 0, 0, ScaleWidth, ScaleHeight
            .ColWidth(.Cols - 1) = ScaleWidth
        End With
    End If
End Sub

Private Sub NmeaRcvList_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'suspend processing to prevent list display being updated
'SuspendProcessList = True
'ProcessSuspended = True
Processing.Suspended = True
Processing.InputOptions = True
End Sub

'adds the received sentence to the first available slot in the NmeaRcvList
'V129 moved from .Main
Public Function AddToNmeaRcvList(ByVal SentenceRcv As String)
Dim iRow As Long
Dim i As Long

'Show form only when when data received
'Hide ONLY when option clicked or in QueryUnload

'v129 If cmdNoWindow = False And ReceivedData.Visible = False And NmeaRcv.Option1(15).Value = False Then
    If cmdNoWindow = False And Visible = False Then   'v129
        Show
    End If

    Const MinRows = 9     'min permitted rows

    If NmeaRcv.Option1(16).Value = True Then  'nmea to Output
'if buffer is full reject any more
'v129
'v129        If Waiting < MAXRCVLISTCOUNT Then

        With ReceivedData.NmeaRcvList
'remove in chunks of 20%
            If .Rows >= MAXRCVLISTCOUNT Then
                .Redraw = False
                For i = 1 To MAXRCVLISTCOUNT / 5
'dont remove a row that is waiting to be decoded
'must be at least 1 waiting and 2 rows (1 is fixed)
'v129                        If .Rows > Waiting And .Rows > MinRows Then
                    If .Rows > MinRows Then
                        .RemoveItem 1
                    End If
                Next i
                .Redraw = True
            End If
    
            iRow = .Rows  '1 fixed row at the top  (base is 0)
            If .TextMatrix(iRow - 1, 0) = "" Then
'there is at least one blank row before the end os the rows
'find first blank line (intial fill is top to bottom)
                iRow = 1
                Do Until .TextMatrix(iRow, 0) = ""
                    iRow = iRow + 1
                Loop
                .TextMatrix(iRow, 0) = SentenceRcv
                .TopRow = 1
            Else
'Debug.Print Asc(Mid$(SentenceRcv, 4, 1))
                .AddItem SentenceRcv
                .TopRow = .Rows - .FixedRows - MinRows + 2
            End If
'v129            Waiting = Waiting + 1 'waiting
        End With    'ReceivedData.NmeaRcvList
'v129        End If  'rcv buffer is not full
    End If

End Function

Private Sub NmeaRcvList_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim InputSentence As String
Dim i As Integer
Dim j As Integer

'    If DecodingState <> 0 Then Exit Sub
'    DecodingState = DecodingState Or 6
    With ReceivedData.NmeaRcvList
        If Button = vbRightButton Then
'    Filter.Show
        Else
                        
            If Len(.TextMatrix(.Row, 0)) <> 0 Then
'create clsSentence, PayloadBytes() and Ship(names)
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
                        Call DecodeSentence(.TextMatrix(j, 0))
                    Next j
                End If

'turn on detail display (Select) to enable output
                NmeaRcv.Option1(13).Value = True
                Detail.Show     'must also show the form
                Call Detail.SentenceAndPayloadDetail(0)       'clsSentence, PayloadBytes required
            Else
                MsgBox "No Data for this line"
            End If
        End If
    End With

'Must clear the NmeaOutBuf when SentanceAndPayloadDetail has finished using it
    Call NmeaOutBufClear
'restart processing to enable list display updating
'    SuspendProcessList = False
'    Call ResumeProcess
    Call ResumeProcessing("InputOptions")
'no point in clicking list if detail is hidden
    Detail.Show
    Call ResumeProcessing("InputOptions")
'    DecodingState = DecodingState Xor 8
'If Detail.DetailDisplay.TextMatrix(1, 1) = "" Then Stop 'debug

End Sub
