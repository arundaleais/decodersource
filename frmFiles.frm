VERSION 5.00
Begin VB.Form frmFiles 
   Caption         =   "Sample Initialisation Files"
   ClientHeight    =   1815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4170
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFiles.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1815
   ScaleWidth      =   4170
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
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
      Left            =   3360
      TabIndex        =   4
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
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
      Left            =   2400
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
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
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label Label3 
      Caption         =   "User Files differ from downloaded files"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "\Setings\default.ini is always replaced"
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
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Untick to retain, Tick to replace "
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "frmFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AllUsersPath As String
Dim CurrentUserPath As String
Dim i As Integer
Dim ChangedFiles As Long
Dim kb As String
Dim kbout As String
Public Cancel As Boolean

Private Sub Command1_Click()    'cancel
Cancel = True
Call WriteStartUpLog("User Cancelled copying files from All User to Current User")
Hide
End Sub

Private Sub Command2_Click()    'ok
Cancel = False
Call CopyAllToCurrentFiles
Hide
End Sub

Public Sub UpdateUserFiles()
        CurrentUserPath = QueryValue(HKEY_CURRENT_USER, "Software\Arundale\" & App.EXEName & "\Settings", "CurrentUserPath")
        If FolderExists(CurrentUserPath) = False Then
            Call WriteStartUpLog(CurrentUserPath & " - Path does not exist, exiting UpdateUserFiles)")
            Exit Sub
        End If
        AllUsersPath = GetSpecialFolderA(CSIDL_COMMON_APPDATA) & "Arundale\Ais Decoder"
        If FolderExists(AllUsersPath) = False Then
            Call WriteStartUpLog(AllUsersPath & " - Path does not exist, exiting UpdateUserFiles)")
            Exit Sub
        End If
        i = 0
        Do
            kb = QueryValue(HKEY_LOCAL_MACHINE, "Software\Arundale\" & App.EXEName & "\Settings\ReservedFiles", CStr(i))
            If kb = "" Then Exit Do
'        call writestartuplog( Now() & vbTab &
Call WriteStartUpLog(AllUsersPath & kb & vbTab & FileExists(AllUsersPath & kb))
Call WriteStartUpLog(CurrentUserPath & kb & vbTab & FileExists(CurrentUserPath & kb))
'check file specified in Registry actually exists
'I may have made a mistake setting up in INNO
            If (FileCompare(AllUsersPath & kb, CurrentUserPath & kb) = False _
            Or FileExists(CurrentUserPath & kb) = False) _
            And FileExists(AllUsersPath & kb) = True Then
'MsgBox FileExists(AllUsersPath & kb)
Call WriteStartUpLog(kb & " differ")
                ChangedFiles = ChangedFiles + 1
                If ChangedFiles > 1 Then
                    Load Check1(ChangedFiles - 1)
                    Check1(ChangedFiles - 1).Top = Check1(ChangedFiles - 2).Top + 300
                    Command1.Top = Command1.Top + 300
                    Command2.Top = Command2.Top + 300
                    Height = Height + 300
                End If
                Check1(ChangedFiles - 1).Caption = kb
                Check1(ChangedFiles - 1).Visible = True
                Check1(ChangedFiles - 1).Value = 1
                Check1(ChangedFiles - 1).Enabled = True
                If kb = "\Settings\default.ini" Then
                    Check1(ChangedFiles - 1).Enabled = False
                    Label2.Visible = True
                End If
            End If
            i = i + 1
        Loop Until kb = ""
        If ChangedFiles <> 0 Then
            If EditOptions = True Then
                Show vbModal    'files are copied if user clicks ok
            Else
'if not Admin the copy files anyway
                Call CopyAllToCurrentFiles
            End If
        Else
Call WriteStartUpLog("User Files same as All User Files")
        
        End If
    Unload frmFiles

End Sub

Public Sub CopyAllToCurrentFiles()
For i = Check1.LBound To Check1.UBound
    If Check1(i).Value <> 0 Then
Call WriteStartUpLog(Now() & vbTab _
& "Copying " _
& "from " & AllUsersPath & Check1(i).Caption _
& "to   " & CurrentUserPath & Check1(i).Caption)
    FileCopy AllUsersPath & Check1(i).Caption, CurrentUserPath & Check1(i).Caption
    End If
Next i
End Sub

Private Sub Form_Load()
Hide
End Sub

