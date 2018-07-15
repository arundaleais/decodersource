VERSION 5.00
Begin VB.Form FieldInput 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input Field"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5100
   ControlBox      =   0   'False
   Icon            =   "FieldInput.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   5100
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   495
      Left            =   3120
      TabIndex        =   3
      Top             =   1320
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3960
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4815
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2280
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Change to your name (if you wish)"
      Height          =   195
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   2385
   End
End
Attribute VB_Name = "FieldInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Cancel As Boolean

Private Sub Command1_Click()    'cancel
Cancel = True
'v128 Hide
Unload Me 'v128
End Sub

Private Sub Command2_Click()    'ok
Cancel = False
'v128 Hide
'Required to create user name for tag
Detail.UserFieldTagName = Text1.Text  'v131
Unload Me 'v128
End Sub

'v128 Private Sub Form_Activate()
'v128 SuspendProcessOptionsInput = True
'v128 ProcessSuspended = True
'v128 End Sub

'v128 Private Sub Form_Deactivate()
'v128 SuspendProcessOptionsInput = False
'v128 Call ResumeProcess
'v128 End Sub

Private Sub Form_Load() 'v128
'SuspendProcessOptionsInput = True 'v128
'ProcessSuspended = True 'v128

End Sub

Private Sub Form_Unload(Cancel As Integer)  'v128
'SuspendProcessOptionsInput = False  'v128
'Call ResumeProcess  'v128
End Sub 'v128
