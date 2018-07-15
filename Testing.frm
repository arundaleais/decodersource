VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form Testing 
   Caption         =   "Testing"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4125
   Icon            =   "Testing.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   4125
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MsFlexGrid1 
      Height          =   6255
      Left            =   0
      TabIndex        =   1
      Top             =   960
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   11033
      _Version        =   393216
      Cols            =   4
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Testing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Long
Dim kb As String

Private Sub Command1_Click()
Call Variables
End Sub

Public Sub Variables()
With MSFlexGrid1
    .Clear
    For i = 2 To .Rows - 1
        .RemoveItem 1
    Next i
    kb = "TaggedOutputOn" & vbTab
    For i = 1 To 2
        kb = kb & vbTab & TaggedOutputOn(i)
    Next i
    .AddItem kb
    .RemoveItem 1    'first blank line
    kb = "ChannelMethod" & vbTab
    For i = 1 To 2
        kb = kb & vbTab & ChannelMethod(i)
    Next i
    .AddItem kb
    kb = "ChannelFormat" & vbTab
    For i = 1 To 2
        kb = kb & vbTab & ChannelFormat(i)
    Next i
    .AddItem kb
    kb = "ChannelEncoding" & vbTab
    For i = 1 To 2
        kb = kb & vbTab & ChannelEncoding(i)
    Next i
    .AddItem kb
    kb = "OutputFile Names" & vbTab & _
        NameFromFullPath(OutputFileNameNmea) & vbTab & _
        NameFromFullPath(OutputFileNameCsv) & vbTab & _
        NameFromFullPath(OutputFileNameTagged)
    .AddItem kb
    kb = "OutputFile Rollover" & vbTab
    For i = 1 To 2
        kb = kb & vbTab & OutputFileRollover(i)
    Next i
    .AddItem kb
    kb = "NmeaOutput"
    For i = 0 To 2
        kb = kb & vbTab & NmeaOutput(i)
    Next i
    .AddItem kb
    kb = "CsvOutput"
    For i = 0 To 2
        kb = kb & vbTab & CsvOutput(i)
    Next i
    .AddItem kb
    kb = "CsvAll"
    For i = 0 To 2
        kb = kb & vbTab & CsvAll(i)
    Next i
    .AddItem kb
    kb = "TaggedOutput"
    For i = 0 To 2
        kb = kb & vbTab & TaggedOutput(i)
    Next i
    .AddItem kb
    kb = "ShellOn"
    For i = 0 To 2
        kb = kb & vbTab & ShellOn(i)
    Next i
    .AddItem kb
    kb = "TagsReq"
    For i = 0 To 2
        kb = kb & vbTab & TagsReq(i)
    Next i
    .AddItem kb
    kb = "CsvDelim"
    For i = 0 To 2
        kb = kb & vbTab & CsvDelim(i)
    Next i
    .AddItem kb
    kb = "Tag Character"
    For i = 0 To 1
        kb = kb & vbTab & TagChr(i)
    Next i
    .AddItem kb
    kb = "ChannelOutput"
    For i = 0 To 2
        kb = kb & vbTab & ChannelOutput(i)
    Next i
    .AddItem kb
'    kb = "ChannelOutputOk"
'    For i = 0 To 2
'        kb = kb & vbTab & ChannelOutputOk(i)
'    Next i
'    .AddItem kb
    kb = "RangeReq"
    For i = 0 To 2
        kb = kb & vbTab & RangeReq(i)
    Next i
    .AddItem kb
    kb = "ScheduledReq"
    For i = 0 To 2
        kb = kb & vbTab & ScheduledReq(i)
    Next i
    .AddItem kb
    kb = "FenReq"
    For i = 0 To 2
        kb = kb & vbTab & FenReq(i)
    Next i
    .AddItem kb
    kb = "GisReq"
    For i = 0 To 2
        kb = kb & vbTab & GisReq(i)
    Next i
    .AddItem kb
'    kb = "GisTag-Lat-Long OK" & vbTab & GisTagOk & vbTab & LatOk & vbTab & LonOK
'    kb = "-Lat-Long OK" & vbTab & vbTab & LatOK & vbTab & LonOK
'    .AddItem kb
    kb = "TimeStampReq"
    For i = 0 To 2
        kb = kb & vbTab & TimeStampReq(i)
    Next i
    .AddItem kb
    kb = "DisplayOutput" & vbTab
    For i = 1 To 2
        kb = kb & vbTab & DisplayOutput(i)
    Next i
    .AddItem kb
    kb = "MethodOutput"
    For i = 0 To 2
        kb = kb & vbTab & MethodOutput(i)
    Next i
    .AddItem kb
    kb = "FTPOutput"
    For i = 0 To 2
        kb = kb & vbTab & FtpOutput(i)
    Next i
    .AddItem kb
    
End With
End Sub


Private Sub Form_Load()
With MSFlexGrid1
    .FormatString = "<Variable       |Element 0|Element 1|Element 2"
    .ColWidth(0) = 1500
End With
Testing.Show
End Sub

