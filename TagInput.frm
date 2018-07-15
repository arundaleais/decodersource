VERSION 5.00
Begin VB.Form TagInput 
   Caption         =   "Edit Tag Output Filter"
   ClientHeight    =   1830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3585
   Icon            =   "TagInput.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1830
   ScaleWidth      =   3585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   495
      Left            =   1440
      TabIndex        =   6
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   495
      Left            =   600
      TabIndex        =   5
      Top             =   1200
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2280
      TabIndex        =   4
      Top             =   1200
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label1 
         Caption         =   "Max Value"
         Height          =   255
         Left            =   2040
         TabIndex        =   3
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   480
         Width           =   3015
      End
   End
End
Attribute VB_Name = "TagInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Cancel As Boolean

Private Sub Command1_Click()    'cancel
Cancel = True
Hide
End Sub

Private Sub Command2_Click()    'ok
            
Cancel = False
Hide
End Sub

Private Sub Command3_Click()    'clear to ("") = dont check
TagInput.Text1.Text = ""

Cancel = False
Hide
End Sub

Private Sub Form_Activate()
'SuspendProcessOptionsInput = True
'ProcessSuspended = True
End Sub

Private Sub Form_Deactivate()
'SuspendProcessOptionsInput = False
'Call ResumeProcess
End Sub


