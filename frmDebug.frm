VERSION 5.00
Begin VB.Form frmDebug 
   Caption         =   "Debug Window"
   ClientHeight    =   6915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7425
   Icon            =   "frmDebug.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6915
   ScaleWidth      =   7425
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txt 
      Height          =   6735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   7215
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function Display(Data As String)
    
         If Len(txt.Text) > 60000 Then
            txt.Text = Right$(txt.Text, 50000)
        End If
        txt.SelStart = Len(txt.Text)
        txt.SelText = Data & vbCrLf
End Function


Public Sub Clear()
    txt.Text = ""
End Sub

Private Sub Form_Load()
    Me.Visible = True
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbNormal Then
        With txt
            .Move 0, 0, ScaleWidth, ScaleHeight
            .Width = ScaleWidth
        End With
    End If
End Sub

