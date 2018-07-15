VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Begin VB.Form frmFTP 
   Caption         =   "FTP"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7155
   BeginProperty Font 
      Name            =   "@Arial Unicode MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFTP.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer StartTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   6720
      Top             =   600
   End
   Begin VB.Timer PollTimer 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   6120
      Top             =   600
   End
   Begin VB.TextBox txtSizeLocal 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6120
      TabIndex        =   16
      Top             =   1800
      Width           =   855
   End
   Begin VB.TextBox txtSizeRemote 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6120
      TabIndex        =   15
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox txtRemote 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   14
      Text            =   "output.csv"
      Top             =   1320
      Width           =   4935
   End
   Begin VB.TextBox txtLocal 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "frmFTP.frx":058A
      Top             =   1800
      Width           =   4935
   End
   Begin InetCtlsObjects.Inet inetFTP 
      Left            =   6000
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
      RequestTimeout  =   0
   End
   Begin VB.TextBox txtLog 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2400
      Width           =   6975
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5400
      TabIndex        =   8
      Top             =   480
      Width           =   555
   End
   Begin VB.TextBox txtDir 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1080
      TabIndex        =   7
      Text            =   "/var/www/html/web/docs/ais/test"
      Top             =   930
      Width           =   4935
   End
   Begin VB.TextBox txtPW 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   3720
      PasswordChar    =   "*"
      TabIndex        =   5
      Text            =   "hx45red"
      Top             =   480
      Width           =   1575
   End
   Begin VB.TextBox txtUser 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1080
      TabIndex        =   3
      Text            =   "webftp"
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox txtHost 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1080
      TabIndex        =   1
      Text            =   "mailgate.arundale.co.uk"
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label Label7 
      Caption         =   "Size"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   17
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label6 
      Caption         =   "Remote File"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Local File"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   855
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   0
      TabIndex        =   10
      Top             =   6120
      Width           =   9975
   End
   Begin VB.Label Label4 
      Caption         =   "Directory:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "User ID:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "FTP Host:"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmFTP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'I see people are still downloading this from time to time so I thought
'I should clean it up a little.  While the example is meant to show
'DIR retrieval it could be modified to handle other sequences of
'commands including uploading and downloading.


Private Const SECONDS_PER_CMD = 20
Private Const SECONDS_PER_CD_CMD = 30   'Include login
Private Const SECONDS_PER_DIR_CMD = 20
Private Const SECONDS_PER_TRANSFER_CMD = 300  '300 = 5 minutes.  Can be longer of course.
Private StartExecuteWait As Single
Private StartCmd As Single
Private StartFTP As Single
Private CmdTimeout As Single
Private AbortFlag As Boolean

Private CmdQ As New Collection

'This is used to set the command timeout interval, signify "idle"
'status (command Q empty), and can be used in the icResponseCompleted
'case to decide what to do with the response text returned by various
'comands.
Private CurrCmdVerb As String
Private strData As String
Private Connected As Boolean
'Private Flushed As Boolean
Private ExitMsg As String
Private ExecuteTimer As clsWaitableTimer
Dim TempRemoteName As String

'Passive Form Level declarations
Dim hOpen As Long, hConnection As Long, hFile As Long
Dim dwType As Long
Dim dwSeman As Long

'Use the Form to do the upload (for testing)
Private Sub cmdGo_Click()
    Call Initialise
    Call Next_Command
    Call Terminate
End Sub
    
'Use the Call from AisDecoder to do the upload
Public Sub FtpUpload()
If FtpUploadExecuting = True Then
MsgBox "FtpUpload still executing"
    Exit Sub
End If
    Set ExecuteTimer = New clsWaitableTimer
    FtpUploadExecuting = True
    txtUser.Text = FtpUserName
    txtPW.Text = FtpPassword
    txtHost.Text = HostFromPath(PathFromFullName(FtpRemoteFileName, "/"))
    txtDir = FolderFromPath(PathFromFullName(FtpRemoteFileName, "/"))
    txtRemote.Text = NameFromFullPath(FtpRemoteFileName, "/")
'If were going to FTP the Output file, it is copied to
'OutputFile + .ftp when it is closed. This allows the file to be read
'whilst a new output file is being created.
    txtLocal.Text = FtpLocalFileName & ".ftp"   'includes path
    TempRemoteName = NameFromFullPath(txtLocal.Text, "\")
    SetUI Active:=False
    If True Then    'Allways using passive now
        Call PassiveSend
    Else
        Call Initialise
        Call Next_Command
        Call Terminate
    End If
    NmeaRcv.StatusBar.Panels(1) = ExitMsg
    NmeaRcv.ClearStatusBarTimer.Enabled = True
    FtpUploadExecuting = False
    If FileExists(FtpLocalFileName & ".ftp") Then Kill FtpLocalFileName & ".ftp"
    Set ExecuteTimer = Nothing
End Sub

'Connect & set up the remote directory
Private Sub Initialise(Optional Caller As String)   'Only if not using Passive
    Connected = False
    AbortFlag = False
    txtSizeLocal.Text = ""
    txtSizeRemote.Text = ""
    LogClear

'clear command buffer
    Do While CmdQ.Count > 0
        CmdQ.Remove 1
    Loop
'    TempRemoteName = NameFromFullPath(txtLocal.Text, "\")
 'add commands to be used when we have verified a connection
    CmdQ.Add "PUT """ & txtLocal.Text & """ """ & TempRemoteName & """"
'if remote differs from local the rename (probably a .tmp file)
    If txtRemote.Text <> TempRemoteName Then
        CmdQ.Add "SIZE """ & TempRemoteName & """"
        CmdQ.Add "RENAME """ & TempRemoteName & """ """ & txtRemote.Text & """"
    Else
        CmdQ.Add "SIZE """ & txtRemote.Text & """"
    End If
    
    StartFTP = Timer
    lblStatus.Caption = "Working..."
    LogLine Now() & " Local"
'Prepare FTP logon.
    With inetFTP
    ' You must set the URL before the user name and
    ' password. Otherwise the control cannot verify
    ' the user name and password and you get the error:
    '
    '       Unable to connect to remote host
         .Cancel  'clear any error mmessage
        .RemoteHost = txtHost.Text
'LogLine "URL: " & .Url
        LogLine "Remote Host: " & .RemoteHost & ":" & .RemotePort
        .UserName = txtUser.Text
        .Password = txtPW.Text
        .Protocol = icFTP
        .RequestTimeout = 60
        LogLine "User Name: " & .UserName
        LogLine "Password: " & .Password
'LogLine "URL: " & .Url
    End With
    
'CD makes the connection. If root directory outputing CD "" generates error
'so just output CD
    If txtDir.Text <> "" Then
        Execute_Command ("CD """ & txtDir.Text & """")
    Else
        Execute_Command ("CD")
    End If
End Sub

'Execute a Command
Private Sub Execute_Command(cmd As String)   'Only if not using Passive
Dim CmdErrorMsg As String
Dim ret As Long

    CurrCmdVerb = UCase$(Split(cmd, " ")(0))

'set the timeouts
    Select Case CurrCmdVerb
    Case "CD"
        CmdTimeout = SECONDS_PER_CD_CMD
    Case "DIR", "LS"
        CmdTimeout = SECONDS_PER_DIR_CMD
    Case "GET", "PUT"
        CmdTimeout = SECONDS_PER_TRANSFER_CMD
    Case Else
        CmdTimeout = SECONDS_PER_CMD
    End Select
    
    LogLine "> " & cmd
'Begin polling for completion (or timeout).
    StartCmd = Timer
    StartExecuteWait = Timer
    
'log the size of the file were uploading
    Select Case CurrCmdVerb
    Case "PUT"
        On Error GoTo Cmd_err
        txtSizeLocal.Text = FileLen(txtLocal.Text)  '.kmz.ftp
        On Error GoTo 0
    Case "RENAME"
'don't rename if PUT may have partially uploaded a file
        If txtSizeRemote.Text <> txtSizeLocal.Text Then
            CmdErrorMsg = "Remote File Size differs from Local File Size"
            If inetFTP.StillExecuting Then LogLine "Waiting for response"
            Call Execute_Wait("Command " & CurrCmdVerb)
            LogLine CmdErrorMsg
            AbortFlag = True
            GoTo SkipCmd
        End If
    End Select


'trap Cant connect error
 'Debug.Print "wait=" & inetFTP.StillExecuting & Cmd
    On Error GoTo Cmd_err
    inetFTP.Execute , cmd
    On Error GoTo 0
'return here if an error
SkipCmd:
 'Debug.Print "wait=" & inetFTP.StillExecuting & Cmd
'add condition Only wait if still executing Dino
mySleep 100
    If inetFTP.StillExecuting Then
 'Debug.Print "wait"
        LogLine "Waiting for response"
'wait until last operation is finished
        Call Execute_Wait("Command " & CurrCmdVerb)
    End If
'Display the Output (if any) from the last command inet must have stopped executing
'for DIR it will be the directory listing, for Size it's the uploaded file
'strdata contains the data returned by Get_Chunk
    If strData <> "" Then
        Select Case CurrCmdVerb
        Case "SIZE"
            txtSizeRemote.Text = strData
        Case Else
            LogLine strData
        End Select
        strData = ""
    End If
Exit Sub
'keep for debugging command speed
'    LogLine "LastFTPOperation (" & CStr(CmdQ.Count) & ") " & CurrCmdVerb _
'    & " " & inetFTP.StillExecuting & ", Time Taken:" & CStr(Elapsed(StartExecuteWait))
Cmd_err:
'we must keep the message as execute_wait will destroy it
'and we want this message to be display after the StateChanged message
    CmdErrorMsg = "FTP Command Error " & Str(err.Number) & " " & err.Description
    If inetFTP.StillExecuting Then LogLine "Waiting for response"
    Call Execute_Wait("Command " & CurrCmdVerb)
    LogLine CmdErrorMsg
    AbortFlag = True
End Sub

'Execute the Commands in the queue (except CD and CLOSE)
Private Sub Next_Command()   'Only if not using Passive
Dim CmdNo As Long
Dim cmd As String

    Do While CmdNo < CmdQ.Count And AbortFlag = False
        CmdNo = CmdNo + 1
        cmd = CmdQ.Item(CmdNo)
        Call Execute_Command(cmd)
    Loop

End Sub

'Close down the connection
Private Sub Terminate()   'Only if not using Passive
'If not connected, close generates another not connected error
    If Connected Then
        Call Execute_Command("CLOSE")
    End If
    lblStatus.Caption = "Completed"
'SetUI Active:=True
    If AbortFlag Then
        ExitMsg = "FTP operations failed after " & Format$(Elapsed(StartFTP), "###0.00") & " seconds"
        LogLine ExitMsg
        WriteErrorLog ("FTP Error" & vbCrLf & txtLog.Text)
        inetFTP.Cancel 'added 19/3 not sure if we need this
    Else
        ExitMsg = "All FTP operations completed successfully in " & Format$(Elapsed(StartFTP), "###0.00") & " seconds"
        LogLine ExitMsg
    End If
End Sub

Private Sub Log(ByVal Text As String)
On Error GoTo Log_error
    With txtLog
        .SelStart = Len(.Text)
        .SelText = Text
    End With
    Exit Sub
Log_error:      'buffer is full
    Call LogClear
End Sub

Private Sub LogClear()
    txtLog.Text = ""
    txtLog.SelStart = 0
    Refresh 'ensure display is cleared
End Sub

Private Sub LogLine(Optional ByVal Text As String)
    
    Log Text
    Log vbNewLine
End Sub

Private Sub SetUI(ByVal Active As Boolean)
    txtHost.Enabled = Active
    txtUser.Enabled = Active
    txtPW.Enabled = Active
    txtDir.Enabled = Active
    txtLocal.Enabled = Active
    txtRemote.Enabled = Active
    txtSizeRemote.Enabled = Active
    txtSizeLocal.Enabled = Active
    cmdGo.Enabled = Active
    cmdGo.Visible = Active
End Sub

Private Sub Form_Load()
    
#If jnasetup Then
    SetUI Active:=True
#Else
    SetUI Active:=False
#End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'V142 Call DisplayQueryUnload(Me.Name, Cancel, UnloadMode)
    If UnloadMode = vbFormControlMenu Then  'V3.4.143 User clicked (X)
        If FormLoaded("NmeaRcv") Then   'stop NmeaRcv being reloaded on exit
            If NmeaRcv.Visible = True Then
                NmeaRcv.Check1(4).Value = 0 'FTP Output
                If frmFTP.Visible = True Then frmFTP.Hide
                Cancel = True   'just hide
            End If
        End If
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbNormal Then
        txtLog.Width = ScaleWidth
    End If
End Sub

Private Sub inetFTP_StateChanged(ByVal State As Integer)   'Only if not using Passive

'This is useful to debug response
'LogLine "State:" & CStr(State) & " " & inetFTP.StillExecuting & ", Elapsed:" & CStr(Elapsed(StartExecuteWait)) _
'& ", Timeout:" & inetFTP.RequestTimeout
    
   On Error GoTo myerror
    Select Case State
        Case icError
LogLine "FTP State Error " & CStr(inetFTP.ResponseCode) & " " & inetFTP.ResponseInfo & " Aborting"
'reports
'icError 12007 Name not resolved
'icError 12029 Cannot connect, if incorrect IP
'icError 12014 Incorrect password  if username or password incorrect
'icError 12003 550 Failed to change directory.
'icError 2 The system cannot find the file specified.
'icError Aborting 12030 Connection aborted
'flushQ adds another close
        
 '       If CurrCmdVerb <> "CLOSE" Then
 '           FlushQ
 '       End If
'            CancelFTP
            AbortFlag = True
        Case icResolvingHost
        Case icHostResolved
        Case icReceivingResponse
'            If AbortFlag = False Then GetData  'inserted dino
'            If strData <> "" Then LogLine strData    'display received data
'        Stop
        Case icResponseReceived
            If AbortFlag = False Then GetData  'inserted dino
            If strData <> "" Then LogLine strData    'display received data
        Case icConnected
            Connected = True
        Case icDisconnected
            Connected = False
        Case icResponseCompleted
            'We have a completed response!  We just log any responses here, but
            'we could use the value of CurrCmdVerb to send different results
            'different places or perform different kinds of processing.
            If AbortFlag = False Then Call GetData
            If strData <> "" Then LogLine strData    'display received data
'            FtpCmdTicksLeft = 0
    End Select
'Only display last error, some errors are duplicated eg failed to change directory
'where the message text is generated by (5) icRequesting but the response code
'stays at zero. The ResponseInfo is repeated until (11).
            
'keep for debugging
'   If inetFTP.ResponseCode <> 0 Then   'don't display earlier duplicated errors
'       LogLine "StateChanged Info " & CStr(inetFTP.ResponseCode) & " (" & CStr(State) & ") " & inetFTP.ResponseInfo
'   End If

'    If inetFTP.ResponseInfo <> "" Then
'        LogLine inetFTP.ResponseInfo
'    End If
'reports
'StateChanged Info 0 (8) The operation completed successfully.
'StateChanged Info 12007 (11) Name not resolved
'StateChanged Info 12029 (11) Cannot connect, if incorrect IP
'StateChanged Info 12014 (11) Incorrect password
'StateChanged Info 0 (5) 550 Failed to change directory.
'StateChanged Info 0 (8) 550 Failed to change directory.
'StateChanged Info 12003 (11) 550 Failed to change directory.
'StateChanged Info 2 (11) The system cannot find the file specified.

    Exit Sub
myerror:
    MsgBox "State Changed Error: " & err.Description & " " & err.Number
End Sub

'this code is taken from the Microsoft Example
Private Sub GetData()   'Only if not using Passive
        Dim vtData As Variant ' Data variable.
        strData = ""
        Dim bDone As Boolean: bDone = False
        
        On Error GoTo myerror
'removed 19/3        Call Execute_Wait("Receiving Data ")
       ' Get first chunk.
        vtData = inetFTP.GetChunk(1024, icString)
        DoEvents
        If Len(vtData) = 0 Then bDone = True
        Do While Not bDone
'LogLine "Get Chunk, AbortFlag=" & AbortFlag
           strData = strData & vtData
           DoEvents
           ' Get next chunk.
           vtData = inetFTP.GetChunk(1024, icString)
           If Len(vtData) = 0 Then
              bDone = True
           End If
        Loop
'        Call Execute_Wait("Receiving Data")
        Exit Sub
myerror:
'added 19/3
    Select Case err.Number
    Case Is = 35764 'Still Executing last request
        Resume Next
    End Select
    MsgBox "GetData Error: " & err.Description & " " & err.Number
End Sub


Private Sub StartTimer_Timer()
    StartTimer.Enabled = False
    Call FtpUpload
End Sub

Private Sub txtPW_Validate(Cancel As Boolean)
    If Len(txtPW.Text) = 0 Then
        Beep
        lblStatus.Caption = "Password is required"
        Cancel = True
    Else
        lblStatus.Caption = ""
    End If
End Sub

Private Sub Execute_Wait(Text As String)   'Only if not using Passive
Dim Step As Long
Dim Total As Single
Dim Aborting As Boolean 'if true don't redisplay Timout message

'leave for debugging
 '   If Left$(Text, 8) = "Command " Then
 '       LogLine Text & " - Waiting for response"
 '   Else
'        LogLine Text   'waiting for Chunk
 '   End If
    Step = 100
    StartExecuteWait = Timer
    Do Until inetFTP.StillExecuting = False
        DoEvents
        If Elapsed(StartCmd) > CmdTimeout Then
            If Aborting = False Then LogLine "Command " & CurrCmdVerb _
            & " Timeout (" & CStr(CmdTimeout) & " seconds) timed out after " _
            & Format$(Elapsed(StartCmd), "###0.00") & " seconds"
            Aborting = True 'dont repeat abort message
'AbortFlag must be set immediately otherwise getchunk does not see it
'as the wait continues until GetChunk has finished
            AbortFlag = True
'If AbortFlag Then Exit Do
 'Debug.Print "Abortflag"
            Step = 1000
'MUST NOT inetFTP.Cancel whilst still executing (gets stuck in loop)
        End If
'If AbortFlag = False Then
'Sleep 1000
        mySleep Step
'End If
    Loop

'get all state changes before returning to main routine
    DoEvents
    If Aborting Then
        LogLine Text & " - Aborted after " & Format$(Elapsed(StartExecuteWait), "###0.00") & " seconds"
    Else
'leave in for debugging
'        If Len(strData) = 0 Then
'            LogLine Text & " - No data"
'        Else
'            LogLine Text & " - Resuming after " & CStr(Elapsed(StartExecuteWait)) & " secs"
'        End If
    End If
'LogLine "Execute_Wait " & CStr(inetFTP.ResponseCode) & " " & inetFTP.ResponseInfo
End Sub

Private Function Elapsed(StartTime As Single) As Single
Dim FinishTime As Single
Dim ElapsedTime As Single
FinishTime = Timer
ElapsedTime = FinishTime - StartTime
'60# is to prevent overflow (otherwise I think it assumes an integer)
If ElapsedTime < 0 Then ElapsedTime = ElapsedTime + CSng(60# * 60# * 24#)
Elapsed = ElapsedTime
End Function

'Required to continue with interupts whilst sleeping
Private Sub mySleep(dwMilliseconds As Long)   'Only if not using Passive
Dim initTickCount As Double
    
    initTickCount = LongToUnsigned(GetTickCount)
    Do Until GetTickCount - initTickCount >= dwMilliseconds
        DoEvents
    Loop
End Sub

Private Function PassiveSend()
Dim Data(99) As Byte ' array of 100 elements 0 to 99
Dim Written As Long
Dim SIZE As Long
Dim sum As Long
Dim j As Long
Dim FreeCh As Long
Dim hFind As Long
Dim nLastError As Long
Dim pData As WIN32_FIND_DATA

    LogClear
'Set tracking flags
    Connected = False
    AbortFlag = False
    txtSizeLocal.Text = ""
    txtSizeRemote.Text = ""
    StartFTP = Timer
    LogLine Now() & " Local"

'--Open session
    LogLine "Opening Internet Connection"
    hOpen = InternetOpen(App.ProductName, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    If hOpen = 0 Then
        ErrorOut err.LastDllError, "InternetOpen"
        GoTo Send_Error
    End If
    dwType = FTP_TRANSFER_TYPE_BINARY
    dwSeman = INTERNET_FLAG_PASSIVE
    
'--make connection with the host
    txtUser.Text = FtpUserName
    txtPW.Text = FtpPassword
    LogLine "Connecting to " & txtHost.Text
    Select Case dwSeman
    Case Is = INTERNET_FLAG_PASSIVE
        LogLine "Using Passive"
    Case Else
        LogLine "Using Active"
    End Select
    hConnection = InternetConnect(hOpen, _
    txtHost.Text, INTERNET_INVALID_PORT_NUMBER, _
    txtUser.Text, txtPW.Text, INTERNET_SERVICE_FTP, _
    dwSeman, 0)
    If hConnection = 0 Then
        ErrorOut err.LastDllError, "InternetConnect"
        GoTo Send_Error
    Else
        LogLine "Connected to server " & txtHost.Text
    End If
    TempRemoteName = NameFromFullPath(txtLocal.Text, "\")
'--change the directory
'    txtHost.Text = HostFromPath(PathFromFullName(FtpRemoteFileName, "/"))
'    txtDir.Text = FolderFromPath(PathFromFullName(FtpRemoteFileName, "/"))
'    txtRemote.Text = NameFromFullPath(FtpRemoteFileName, "/")
'    txtLocal.Text = FtpLocalFileName & ".ftp"   'includes path
    If (FtpSetCurrentDirectory(hConnection, txtDir.Text) = False) Then
        ErrorOut err.LastDllError, "FtpSetCurrentDirectory"
        GoTo Send_Error
    Else
        LogLine "Remote directory changed to " & txtDir.Text
    End If

    Select Case dwType
    Case Is = FTP_TRANSFER_TYPE_BINARY
        LogLine "Opening remote file " & TempRemoteName & " for Binary transfer"
    Case Is = FTP_TRANSFER_TYPE_ASCII
        LogLine "Opening remote file " & TempRemoteName & " for ASCII transfer"
    Case Is = FTP_TRANSFER_TYPE_UNKNOWN
        LogLine "Opening remote file " & TempRemoteName & " for Unknown transfer"
    End Select
'--copy the file in host
    hFile = FtpOpenFile(hConnection, TempRemoteName, GENERIC_WRITE, dwType, 0)
    If hFile = 0 Then
        ErrorOut err.LastDllError, "FtpOpenFile"
        GoTo Send_Error
    End If
    LogLine "Remote file " & TempRemoteName & " opened"
    LogLine "Opening local file " & txtLocal.Text
'MsgBox txtLocal.Text
    On Error GoTo OpenLocalFile_Error
'    txtSizeLocal.Text = FileLen(txtLocal.Text)  '.kmz.ftp
    txtSizeLocal.Text = FileLen(FtpLocalFileName & ".ftp")  '.kmz.ftp
    FreeCh = FreeFile
    Open txtLocal.Text For Binary Access Read As #FreeCh
    On Error GoTo 0
    LogLine "Local file " & txtLocal.Text & " opened"
    SIZE = LOF(FreeCh)
    For j = 1 To SIZE \ 100
        Get #FreeCh, , Data
        If (InternetWriteFile(hFile, Data(0), 100, Written) = 0) Then
            ErrorOut err.LastDllError, "InternetWriteFile"
            GoTo Send_Error
        End If
        DoEvents
        sum = sum + 100
    If sum Mod 100000 = 0 Then lblStatus.Caption = Str(sum)
    Next j
    Get #FreeCh, , Data
    If (InternetWriteFile(hFile, Data(0), SIZE Mod 100, Written) = 0) Then
        ErrorOut err.LastDllError, "InternetWriteFile"
        GoTo Send_Error
    End If
    sum = sum + (SIZE Mod 100)
    lblStatus.Caption = Str(sum)   'byte count
    LogLine sum & " bytes transferred"
    Close #FreeCh
    FreeCh = 0
    If hFile Then InternetCloseHandle hFile

'--rename remote file
    If txtRemote.Text <> TempRemoteName Then
        If (FtpRenameFile(hConnection, TempRemoteName, txtRemote.Text) = False) Then
            ErrorOut err.LastDllError, "FtpRenameFile"
            GoTo Send_Error
        Else
            LogLine "Remote file " & TempRemoteName & " renamed " & txtRemote.Text
        End If
    End If


'--get size of remote file
    hFind = FtpFindFirstFile(hConnection, txtRemote.Text, pData, 0, 0)
    nLastError = err.LastDllError
    If hFind = 0 Then
        If (nLastError = ERROR_NO_MORE_FILES) Then
            LogLine "Remote file " & txtRemote.Text & "not found"
        Else
            ErrorOut err.LastDllError, "FtpFindFirstFile"
        End If
        GoTo Send_Error
    End If
    If hFind Then InternetCloseHandle (hFind)
    txtSizeRemote.Text = pData.nFileSizeLow

    If txtSizeRemote.Text <> txtSizeLocal.Text Then
        LogLine "Local and Remote file size differs"
    End If


'--close connection and session
    
Close_Session:
    If FreeCh Then Close #FreeCh
    If hFind Then InternetCloseHandle hFind
    If hFile Then InternetCloseHandle hFile
    If hConnection Then InternetCloseHandle hConnection
    If hOpen Then InternetCloseHandle hOpen
    LogLine "Connection closed"
    
'Clean up
    lblStatus.Caption = "Completed"
'SetUI Active:=True
    If AbortFlag Then
        ExitMsg = "FTP operations failed after " & Format$(Elapsed(StartFTP), "###0.00") & " seconds"
        LogLine ExitMsg
        WriteErrorLog ("FTP Error" & vbCrLf & txtLog.Text)
'        inetFTP.Cancel 'added 19/3 not sure if we need this
    Else
        ExitMsg = "All FTP operations completed successfully in " & Format$(Elapsed(StartFTP), "###0.00") & " seconds"
        LogLine ExitMsg
    End If

    
    Exit Function
OpenLocalFile_Error:
        LogLine txtLocal.Text & " " & err.Description
        AbortFlag = True
    GoTo Close_Session
Send_Error:
    AbortFlag = True
    GoTo Close_Session
End Function

Private Sub ErrorOut(ByVal dwError As Long, ByRef szFunc As String)
Dim dwRet As Long
Dim dwTemp As Long
Dim szString As String * 2048
Dim szErrorMessage As String

dwRet = FormatMessage(FORMAT_MESSAGE_FROM_HMODULE, _
                  GetModuleHandle("wininet.dll"), dwError, 0, _
                  szString, 256, 0)
szErrorMessage = szFunc & " error code: " & dwError & vbCrLf & "Message: " & szString
 'Debug.Print szErrorMessage
LogLine szErrorMessage
If (dwError = 12003) Then
    ' Extended error information was returned
    dwRet = InternetGetLastResponseInfo(dwTemp, szString, 2048)
 'Debug.Print szString
    LogLine szString
End If
End Sub



