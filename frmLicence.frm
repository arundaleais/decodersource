VERSION 5.00
Begin VB.Form frmLicence 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Licence"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5055
   Icon            =   "frmLicence.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.Label lblValue 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   8
         Left            =   1920
         TabIndex        =   18
         Top             =   2520
         Width           =   2775
      End
      Begin VB.Label lblValue 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   7
         Left            =   1920
         TabIndex        =   17
         Top             =   2280
         Width           =   2775
      End
      Begin VB.Label lblValue 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   6
         Left            =   1920
         TabIndex        =   16
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label lblValue 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   5
         Left            =   1920
         TabIndex        =   15
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lblValue 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   14
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblValue 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   3
         Left            =   1920
         TabIndex        =   13
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblValue 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   2
         Left            =   1920
         TabIndex        =   12
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblValue 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   11
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblName 
         BackColor       =   &H80000009&
         Caption         =   "User Name"
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   10
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label lblName 
         BackColor       =   &H80000009&
         Caption         =   "Computer Name"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   9
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label lblName 
         BackColor       =   &H80000009&
         Caption         =   "Issued to"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label lblName 
         BackColor       =   &H80000009&
         Caption         =   "Date Issued"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label lblName 
         BackColor       =   &H80000009&
         Caption         =   "Licence Expires"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lblName 
         BackColor       =   &H80000009&
         Caption         =   "Max Input File Size"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblName 
         BackColor       =   &H80000009&
         Caption         =   "Max Input File Speed"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblName 
         BackColor       =   &H80000009&
         Caption         =   "Max Receive Speed"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblValue 
         BackColor       =   &H80000009&
         Height          =   255
         Index           =   0
         Left            =   1920
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label lblName 
         BackColor       =   &H80000009&
         Caption         =   "Install Valid to"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   1
         Top             =   1560
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmLicence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim i As Long

    With UserLicence
'Change lbl to transparent otherwise Win7 onwards displays box as black
        For i = 0 To 8
            lblName(i).BackStyle = vbTransparent
            lblValue(i).BackStyle = vbTransparent
        Next i
        
        If .MaxRcvPerMin = 0 Then
            lblValue(0).Caption = "Unlimited"
        Else
            lblValue(0).Caption = .MaxRcvPerMin & " sentences/min"
        End If
        If .MaxFilePerMin = 0 Then
            lblValue(1).Caption = "Unlimited"
        Else
            lblValue(1).Caption = .MaxFilePerMin & " sentences/min"
        End If
        If NullToZero(.sMaxInputFileSize) = "0" Then
            lblValue(2).Caption = "Unlimited"
        Else
            lblValue(2).Caption = aByte(.sMaxInputFileSize)
        End If
        If .DateIssued = "00:00:00" Then
            lblValue(3).Caption = "None"
        Else
            lblValue(3).Caption = .DateIssued
        End If
        If .UpdatingValidTo = "00:00:00" Then
            lblValue(4).Caption = "None"
        Else
            lblValue(4).Caption = .UpdatingValidTo
            If Date <= .UpdatingValidTo Then lblValue(4).BackColor = vbGreen
        End If
        If .ExpiryDate = "00:00:00" Then
            lblValue(5).Caption = "None"
        Else
            lblValue(5).Caption = .ExpiryDate
            If Date > .ExpiryDate Then lblValue(5).BackColor = vbRed
        End If
        lblValue(6).Caption = .IssuedTo
        lblValue(7).Caption = .ComputerName
        If .ComputerName <> "" Then 'Licence.ini may not exist
            If Environ$("ComputerName") <> .ComputerName Then lblValue(7).BackColor = vbRed
        End If
        lblValue(8).Caption = .UserName
    End With
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'V142 Call DisplayQueryUnload(Me.Name, Cancel, UnloadMode)
'V3.4.143 Always unload - Never Cancel
    TreeFilter.Check1(23).Value = vbUnchecked
End Sub
