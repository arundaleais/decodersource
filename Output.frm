VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form Output 
   Caption         =   "Output Display"
   ClientHeight    =   2472
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   7680
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Output.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2475
   ScaleMode       =   0  'User
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2535
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   13575
      _ExtentX        =   23940
      _ExtentY        =   4466
      _Version        =   393216
      Rows            =   11
      Cols            =   1
      FixedCols       =   0
      ScrollBars      =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "Output.frx":058A
      Top             =   120
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2292
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   7215
   End
End
Attribute VB_Name = "Output"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim tot As Integer
Dim i As Integer

MaxOutputDisplayCount = 20000      '32k limit before overflow

#If True Then
With MSFlexGrid1
'the vertical scrollbar (when it appears) will be positioned against the RHS
'of the .width of the flexgrid
'If the rightmost column is against the rhs of the flexgrid, the scroll bar will
'effctively reduce the size of the rightmost column.
    .Width = Width - 200    'forces scroll bar to RHS of form
    .ColWidth(.Cols - 1) = .Width
    
'    .Rows = 9
''    Height = .Height + 540  'force form height to enclose vertical scroll bar
'set the msflexgrid to the inner size of the form (ScaleHeight)
    .Height = ScaleHeight
'we need a top row to hold the format (Shift Left)
'otherwise foramstring sdoesnt work
    .GridLines = flexGridNone
    .FormatString = "< |"
    .RowHeight(0) = 0       'hide top row
End With
#End If
#If False Then
With Text1 'set the properties for the displaying textbox
'.BackColor = vbCyan
    .Locked = True
    .Text = ""
End With 'Text1
#End If
'Left = NmeaRcv.Width    'place at right of nmearcv form
'Top = NmeaRcv.Top       'in line with top of nmearcv
Hide   'Always hide on initial load

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Channel As Long
Dim oForm As Variant
'Dim kb As String
'V142 Call DisplayQueryUnload(Me.Name, Cancel, UnloadMode)

    If UnloadMode = vbFormControlMenu Then  'V3.4.143 User clicked (X)
        Hide    'hide this form first
'as I cant find out how you find out the index of the form were closing
'scan through each form in the collection and set the checkbox
'in nmearcv
        If FormLoaded("NmeaRcv") Then   'stop NmeaRcv being reloaded on exit
            If NmeaRcv.Visible = True Then
                For Each oForm In Outputs
                    Channel = Channel + 1
                    DisplayOutput(Channel) = oForm.Visible
                    Select Case Channel
                    Case Is = 1
                        NmeaRcv.Check1(2).Value = 0 'File Output (channel 1)
                    Case Is = 2
                        NmeaRcv.Check1(3).Value = 0 'UDP Output (channel 2)
                    End Select
'           kb = oForm.Visible
                Next
            End If
        End If
    End If
End Sub


Private Sub Form_Resize()
    If Me.WindowState = vbNormal Then
        With MSFlexGrid1
            .ColWidth(.Cols - .FixedCols - 1) = ScaleWidth
            .Move 0, 0, ScaleWidth, ScaleHeight
            .ColWidth(0) = .Width
        End With
    End If

End Sub

Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
MsgBox List1.Text, , "Complete Outputted String"
End Sub

Public Sub OutputDisplay(Channel As Long, kb As String)
Const MaxLin = 13
Const MinRows = 10     'min permitted rows
Dim i As Long
Dim iRow As Long
Static lin As Long
Dim s As Long
Dim t As Long
Dim Line As String

'MsgBox "in " & Channel & ":" & Outputs(1).MSFlexGrid1.TopRow
'Debug.Print "OutputDisplay=" & Screen.ActiveForm.Name
With Outputs(Channel)
    
        s = 1
        Do
            t = InStr(s, kb, vbCrLf)
'May not be a crlf at the end of the last line
'Check we do not have a null length sentence or not termination crlf
            If t > i Then
                Line = Mid$(kb, s, t + i)   'includes crlf
            Else
                Line = Mid$(kb, s) & vbCrLf  'add crlf if not one at end of kb
            End If
    
            With .MSFlexGrid1
'remove in chunks of 20%
                If .Rows >= MaxOutputDisplayCount Then
                    For i = 1 To MaxOutputDisplayCount / 5
                        .RemoveItem 1
                    Next i
                End If
'cycling through this chunk
                If .TextMatrix(.Rows - 1, 0) = "" Then  'at least 1 blank row
                    iRow = .FixedRows
                    Do Until .TextMatrix(iRow, 0) = ""
                        iRow = iRow + 1
                    Loop
                    .TextMatrix(iRow, 0) = Line
                    .TopRow = .FixedRows
                Else
                    .AddItem Line
'.GridLines = flexGridFlat
'.ColWidth(0) = .Width
'MsgBox .TextMatrix(.Rows - 1, 0)
'MsgBox "add out " & Channel & ":" & Outputs(1).MSFlexGrid1.TopRow
                End If
'position scroll bar at the bottom
                .TopRow = .Rows - 1 ' - MinRows
            End With
            
            s = t + 2       'Skip crlf
        Loop Until t = 0  'end of all sentences

'there is some output - note if you want to supress update
'for speed, only redraw=false should be set
'Outputs(Channel).Show
'Show form if required when data received
'Hide ONLY when option clicked or in QueryUnload
    If cmdNoWindow = False And .Visible = False And DisplayOutput(Channel) = True Then
        .Show
    End If
End With
'MsgBox "out " & Channel & ":" & Outputs(1).MSFlexGrid1.TopRow
'Debug.Print "OutputDisplayEnd=" & Screen.ActiveForm.Name
End Sub

