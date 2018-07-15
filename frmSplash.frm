VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   ClientHeight    =   5265
   ClientLeft      =   3090
   ClientTop       =   1320
   ClientWidth     =   6570
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrSplash 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   5700
      Top             =   4620
   End
   Begin VB.Label lblAction 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2775
      TabIndex        =   1
      Top             =   3960
      Width           =   1185
   End
   Begin VB.Label lblAisDecoder 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AisDecoder"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2520
      TabIndex        =   0
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   5280
      Left            =   0
      Picture         =   "frmSplash.frx":058A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6600
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub tmrSplash_Timer()

    tmrSplash.Enabled = False
    Unload Me

End Sub
