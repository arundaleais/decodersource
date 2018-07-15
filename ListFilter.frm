VERSION 5.00
Begin VB.Form ListFilter 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mmsi Filter"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2955
   Icon            =   "ListFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   2955
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.CommandButton Close 
      Caption         =   "Close"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   2640
      Width           =   975
   End
   Begin VB.CommandButton Clear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton Remove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Add 
      Caption         =   "Add"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   1620
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Enter MMSI "
      Height          =   735
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   2535
   End
End
Attribute VB_Name = "ListFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Add_Click()
List1.AddItem Text1.Text
Text1.Text = ""
'Call frmDebug.Clear
Call TreeFilter.CheckAll(TreeFilter.TreeView1.Nodes("InputFilter"), "ListFilter.Add")
Text1.SetFocus
End Sub

Private Sub Clear_Click()
List1.Clear
Remove.Enabled = False
'v118 if required untick on Options and copy to NmeaRcv
Call TreeFilter.CheckAll(TreeFilter.TreeView1.Nodes("InputFilter"), True)
End Sub

Private Sub Close_Click()
Me.Hide
End Sub

Private Sub Form_Load()
'Me.Show
'Text1.Text = ""
End Sub

Private Sub List1_Click()
Remove.Enabled = List1.ListIndex <> -1
End Sub

Private Sub Remove_Click()
Dim i As Long
i = List1.ListIndex
If i >= 0 Then List1.RemoveItem i
Remove.Enabled = (List1.ListIndex <> -1)    'disable button
'Sync Displayed filter
Call TreeFilter.CheckAll(TreeFilter.TreeView1.Nodes("InputFilter"), "ListFilter.Remove")
End Sub

Private Sub Text1_Change()
Add.Enabled = (Len(Text1.Text) > 0) 'enable if > 1 char
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'V142 Call DisplayQueryUnload(Me.Name, Cancel, UnloadMode)

    If UnloadMode = vbFormControlMenu Then  'V3.4.143 User clicked (X)
        If Me.Visible = True Then
            Me.Hide
            Cancel = True   'dont close the form
        End If
    End If
End Sub


Sub Initialise()
End Sub

