VERSION 5.00
Begin VB.Form TestFileSelect 
   Caption         =   "TestFileSelect"
   ClientHeight    =   1320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8220
   Icon            =   "TestFileSelect.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1320
   ScaleWidth      =   8220
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Must Exist"
      Height          =   255
      Left            =   5280
      TabIndex        =   5
      Top             =   360
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   360
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SetFileName"
      Height          =   615
      Index           =   1
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   7935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "AskFileName"
      Height          =   615
      Index           =   0
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "FileName"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "TestFileSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
Select Case Index
Case Is = 0
    If Check1.Value = 0 Then
        Text1.Text = FileSelect.AskFileName(Text2.Text)
    Else
        Text1.Text = FileSelect.AskFileName(Text2.Text, True)
    End If
Case Is = 1
    Text1.Text = FileSelect.SetFileName(Text2.Text)
End Select
'Call Files.RefreshData
End Sub

