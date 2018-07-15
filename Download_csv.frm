VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form Download_csv 
   Caption         =   "Internet Timeout"
   ClientHeight    =   2085
   ClientLeft      =   3885
   ClientTop       =   4665
   ClientWidth     =   5835
   Icon            =   "Download_csv.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   2085
   ScaleWidth      =   5835
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   1080
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   1200
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   720
      Width           =   4695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Enabled         =   0   'False
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "Progress"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Downloading"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "Download_csv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DownloadOK As Boolean
Dim Outfil As String

Private Sub Command1_Click()
    Timer1.Enabled = True
    Unload Me
End Sub

Private Sub Command2_Click()
    Inet1.Cancel
    Command1.Enabled = True
    Command2.Enabled = False
    Text1.Text = "Download Cancelled"
    Unload Me
'need to take action if user cancels during download
End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
Dim i As Long
Dim iFile As Integer

'check all the states an list them in the listbox
    Select Case State
    Case 1
        Text1.Text = "Resolving Host"
    Case 2
        Text1.Text = "Host Resolved"
    Case 3
        Text1.Text = "Connecting"
    Case 4
        Text1.Text = "Connected"
    Case 5
        Text1.Text = "Requesting"
    Case 6
        Text1.Text = "Request Sent"
        Text1.Text = "Preparing to download"
    Case 7
'        Text1.Text = "Receiving Response"
    Case 8
'        Text1.Text = "Response Received"
    Case 9
        Text1.Text = "Disconnecting"
    Case 10
        Text1.Text = "Disconnected"
        Text1.Text = "Downloaded"
    Case 11
'        If cmdStart = False Then
'            MsgBox "Error connecting to internet " & Str(Inet1.ResponseCode) & ": " & Inet1.ResponseInfo, vbOKOnly, "MSInet error"
'        End If
        Text1.Text = "Error connecting to internet"
    Case 12  'request complete get the data
        Text1.Text = "Response Completed"
        Dim sHeader As String
        ' look in the headers for a 401 or 407 error
        'If we get them we will then need to try the request with a username and password
        sHeader = Inet1.GetHeader()
'        MsgBox sHeader, vbOKOnly, "Header info that was returned"
        If InStr(1, sHeader, "407") Or InStr(1, sHeader, "401") Then 'we check for both proxy and IIS Access denied
            MsgBox "Access is denied. Try adding a Username and Password"
        End If
        
        Dim vtData As Variant ' Data variable.
        Dim strData As String: strData = ""
        Dim bDone As Boolean: bDone = False
        Dim b() As Byte
        Dim bData() As Byte
        Outfil = Environ("TEMP") & "\" & UrlToName(Inet1.Url)
'MsgBox Outfil
        Text1.Text = "Creating " & Outfil
        If FileExt(Inet1.Url) = "exe" Or FileExt(Inet1.Url) = "csv" Then
            iFile = FreeFile
           Open Outfil For Binary Access _
Write As #iFile
            b() = Inet1.GetChunk(1024, icByteArray)
            Do While Not bDone
                i = i + 1
                Put #iFile, , b()
                b() = Inet1.GetChunk(1024, icByteArray)
            If UBound(b) = -1 Then
                  bDone = True
            End If
            If i * 1024 <= ProgressBar1.Max Then ProgressBar1.Value = i * 1024
            Loop
            Close #iFile
            DownloadOK = True
        Else
        vtData = Inet1.GetChunk(1024, icString)
        Do While Not bDone
           strData = strData & vtData
           ' Get next chunk.
           vtData = Inet1.GetChunk(1024, icString)
           If Len(vtData) = 0 Then
              bDone = True
           End If
        Loop
        iFile = FreeFile
        Open Outfil For Output As #iFile
        Write #iFile, vtData
        Close #iFile
        End If
    End Select
'        Command1.Enabled = True
'        Command2.Enabled = False

End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False 'turn the timer off
    If Inet1.StillExecuting Then 'are we still working on the request
        Text1.Text = "Time-out"      ' vbModal
    End If
End Sub

Public Function FileExt(FileName As String)
Dim i As Integer
Dim j As Integer
j = 0
Do
i = j + 1
j = InStr(i, FileName, ".")
Loop Until j = 0
FileExt = Mid$(FileName, i, Len(FileName) - i + 1)

End Function

'returns 3.1.0.67, blank if no internet access
Public Function GetWebCurrentVersion() As String
Dim fileBytes() As Byte
Dim arry() As String
Dim ch As Integer
Dim kb As String
'MsgBox "In GetWebCurrentVersion"
    On Error Resume Next     'file may not exist
    Kill Environ("TEMP") & "\CurrentVersion.csv"
    On Error GoTo DownloadError
    DoEvents
    
'MsgBox "Using inet1"
    fileBytes() = Inet1.OpenURL("http://" & DownloadURL & "CurrentVersion.csv", icByteArray)
'MsgBox "Exit inet1"
'    On Error GoTo 0
    ch = FreeFile
    Open Environ("TEMP") & "\CurrentVersion.csv" For Binary Access Write As #ch
    Put #ch, , fileBytes()
    Close #ch
    ch = FreeFile
    Open Environ("TEMP") & "\CurrentVersion.csv" For Input As #ch
    Do Until EOF(ch)
        Line Input #ch, kb
        arry() = Split(kb, ",")
        If arry(0) = App.EXEName Then
            GetWebCurrentVersion = arry(1)
        End If
    Loop
    Close #ch
    Unload Me
'MsgBox "exiting GetWebCurrentVersion"
    Exit Function
DownloadError:
    Exit Function       'don't display any errors, uncomment to debug
'we want to check for MSInet control errors
    If err.Number > 35749 And err.Number < 35805 Then 'MSInet error
        MsgBox "error number: " & Str(err.Number) & vbCrLf & _
            "error discription: " & err.Description & vbCrLf & _
            "Headers: " & Inet1.GetHeader() & vbCrLf & _
            "Response code: " & Str(Inet1.ResponseCode) & vbCrLf & _
            "Response Info: " & Inet1.ResponseInfo, vbOKOnly, "MSInet error"
    Else 'some other error
        MsgBox "error number: " & err.Number & vbCrLf & _
            "error discription: " & err.Description, vbOKOnly, "Error"
    End If
End Function
