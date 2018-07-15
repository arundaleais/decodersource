VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSComm32.ocx"
Begin VB.Form NmeaRcv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ais Decoder"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7860
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   Icon            =   "NmeaRcv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   7860
   Begin VB.Timer PollTimer 
      Enabled         =   0   'False
      Interval        =   55
      Left            =   8280
      Top             =   840
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Spare Hidden"
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
      Index           =   17
      Left            =   7560
      TabIndex        =   55
      Top             =   3120
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.OptionButton Option1 
      Caption         =   "External"
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
      Index           =   8
      Left            =   7680
      TabIndex        =   50
      Top             =   4320
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.OptionButton Option1 
      Caption         =   "NMEA File"
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
      Index           =   9
      Left            =   7680
      TabIndex        =   49
      Top             =   4080
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Message File"
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
      Index           =   12
      Left            =   7680
      TabIndex        =   48
      Top             =   3840
      Visible         =   0   'False
      Width           =   1245
   End
   Begin VB.Timer FileNextBlockTimer 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   7920
      Top             =   840
   End
   Begin VB.Timer ScheduledTimer 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   7560
      Top             =   840
   End
   Begin VB.Timer TcpTimer 
      Enabled         =   0   'False
      Interval        =   4000
      Left            =   8280
      Top             =   120
   End
   Begin VB.Timer ClearStatusBarTimer 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   8280
      Top             =   480
   End
   Begin VB.Timer RcvSpeedTimer 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   7920
      Top             =   480
   End
   Begin MSWinsockLib.Winsock ClientTCP 
      Left            =   7920
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame10 
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   38
      Top             =   0
      Width           =   5055
      Begin VB.OptionButton Option1 
         Caption         =   "UTC"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   1440
         TabIndex        =   41
         Top             =   120
         Value           =   -1  'True
         Width           =   765
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Local"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   1440
         TabIndex        =   40
         Top             =   360
         Width           =   3435
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Timer TimeZoneTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7560
      Top             =   120
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   37
      Top             =   4920
      Width           =   7860
      _ExtentX        =   13864
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   6879
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   6879
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer StatsTimer 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7920
      Top             =   120
   End
   Begin VB.CommandButton cbUpdates 
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
      Left            =   720
      TabIndex        =   33
      Top             =   4560
      Width           =   975
   End
   Begin VB.Timer OpenSerialTimer 
      Enabled         =   0   'False
      Left            =   7560
      Top             =   480
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   7560
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Frame Frame6 
      Caption         =   "Input"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   30
      Top             =   720
      Width           =   1620
      Begin VB.CheckBox Check1 
         Caption         =   "File"
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
         Index           =   6
         Left            =   120
         TabIndex        =   47
         Top             =   960
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "TCP"
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
         Index           =   5
         Left            =   120
         TabIndex        =   43
         Top             =   480
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Serial"
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
         Index           =   1
         Left            =   120
         TabIndex        =   32
         Top             =   720
         Width           =   735
      End
      Begin VB.CheckBox Check1 
         Caption         =   "UDP"
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
         TabIndex        =   31
         Top             =   240
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.Label lblBuffer 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   840
         TabIndex        =   54
         Top             =   960
         Width           =   45
      End
      Begin VB.Label lblBuffer 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   840
         TabIndex        =   53
         Top             =   720
         Width           =   45
      End
      Begin VB.Label lblBuffer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   840
         TabIndex        =   52
         Top             =   480
         Width           =   45
      End
      Begin VB.Label lblBuffer 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   840
         TabIndex        =   51
         Top             =   240
         Width           =   45
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Spare Hidden"
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
      Index           =   19
      Left            =   7560
      TabIndex        =   26
      Top             =   3600
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Spare Hidden"
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
      Index           =   18
      Left            =   7560
      TabIndex        =   25
      Top             =   3360
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Frame Frame7 
      Caption         =   "Input Filter"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   1800
      TabIndex        =   19
      Top             =   600
      Width           =   3375
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   1815
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   3201
         _Version        =   393217
         Style           =   7
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Display"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   1800
      TabIndex        =   7
      Top             =   2880
      Width           =   6015
      Begin VB.Frame Frame11 
         Caption         =   "MyShip"
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
         Left            =   240
         TabIndex        =   45
         Top             =   1320
         Width           =   1455
         Begin VB.Label Label4 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Index           =   0
            Left            =   120
            TabIndex        =   46
            Top             =   240
            Width           =   1275
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Output"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   5040
         TabIndex        =   34
         Top             =   240
         Width           =   975
         Begin VB.CheckBox Check1 
            Caption         =   "FTP"
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
            Index           =   4
            Left            =   120
            TabIndex        =   42
            Top             =   720
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "UDP"
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
            Index           =   3
            Left            =   120
            TabIndex        =   36
            Top             =   480
            Width           =   735
         End
         Begin VB.CheckBox Check1 
            Caption         =   "File"
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
            Index           =   2
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   735
         End
      End
      Begin VB.CommandButton cbSpawnGis 
         Caption         =   "GIS"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   27
         Top             =   1440
         Width           =   735
      End
      Begin VB.Frame Frame9 
         Caption         =   "Nmea Input"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   21
         Top             =   240
         Width           =   1455
         Begin VB.OptionButton Option1 
            Caption         =   "Received"
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
            Index           =   16
            Left            =   120
            TabIndex        =   23
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "None"
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
            Index           =   15
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   975
         End
         Begin VB.Label Label3 
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
            Left            =   120
            TabIndex        =   44
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Detail"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   3480
         TabIndex        =   11
         Top             =   240
         Width           =   1455
         Begin VB.OptionButton Option1 
            Caption         =   "Select"
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
            Index           =   13
            Left            =   120
            TabIndex        =   18
            Top             =   1200
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Scheduled"
            Enabled         =   0   'False
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
            Index           =   7
            Left            =   120
            TabIndex        =   17
            Top             =   960
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Filtered"
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
            Index           =   6
            Left            =   120
            TabIndex        =   16
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Unfiltered"
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
            Index           =   5
            Left            =   120
            TabIndex        =   15
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "None"
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
            Index           =   4
            Left            =   120
            TabIndex        =   14
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Summary"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   1800
         TabIndex        =   8
         Top             =   240
         Width           =   1575
         Begin VB.OptionButton Option1 
            Caption         =   "Range Filtered"
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
            Index           =   10
            Left            =   120
            TabIndex        =   28
            Top             =   1200
            Visible         =   0   'False
            Width           =   1485
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Scheduled"
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
            Index           =   3
            Left            =   120
            TabIndex        =   13
            Top             =   960
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Input Filtered"
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
            Index           =   2
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Unfiltered"
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
            Index           =   1
            Left            =   120
            TabIndex        =   10
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "None"
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
            TabIndex        =   9
            Top             =   240
            Width           =   1215
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Control"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   2
      Top             =   2040
      Width           =   1620
      Begin VB.CommandButton cbDownload 
         Caption         =   "Update"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   1680
         Width           =   855
      End
      Begin VB.CommandButton Help 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   2040
         Width           =   855
      End
      Begin VB.CommandButton cbOptions 
         Caption         =   "Options"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton cbStop 
         Caption         =   "Stop"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   855
      End
      Begin VB.CommandButton cbPause 
         Appearance      =   0  'Flat
         Caption         =   "Pause"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   600
         Width           =   855
      End
      Begin VB.CommandButton cbStart 
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   8160
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid Stats 
      Height          =   2835
      Left            =   5280
      TabIndex        =   0
      Top             =   0
      Width           =   2520
      _ExtentX        =   4445
      _ExtentY        =   5001
      _Version        =   393216
      Rows            =   11
      FixedRows       =   0
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSWinsockLib.Winsock ClientUDP 
      Left            =   8280
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      LocalPort       =   29421
   End
   Begin MSWinsockLib.Winsock ServerUDP 
      Left            =   7560
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
      LocalPort       =   29421
   End
   Begin VB.Label Label1 
      Caption         =   "Version"
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
      Left            =   120
      TabIndex        =   1
      Top             =   4560
      Width           =   615
   End
End
Attribute VB_Name = "NmeaRcv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#Const UseComm = -1     '-1 = use clsComm for serial

'Private Cancel As Boolean   'Must be defined in this form otherwise
                            'application will not close on shutdown

Private mblnCancel As Boolean
Private WithEvents InputFile As HugeBinaryFile
Attribute InputFile.VB_VarHelpID = -1
Private InputFileCancel As Boolean
Private CycleClock As clsTimer  'Used to reject Rcv Sentences and Block File input
Private RunClock As clsTimer    'Used to display the Received/min (Rcv and File)
Private RcvClock As clsTimer    'Used to get TotalTime any non-file inputs are open
Private RcvWaitableTimer As clsWaitableTimer
Private FileWaitableTimer As clsWaitableTimer
'Private StopWaitableTimer As clsWaitableTimer   'Delays stop until NmeaBufCleared

Private Const DEFAULT_BLOCK_SIZE = 2 ^ 16 ' 131,072=2^17   1,048,576=2^20
Private Const UDP_RCV_MAX = 1000000     'Size input is closed (red, Yellow at 50%)
Private Const UDP_RCV_INUSE = 2000     'Size input is flagged as in use (Green)
Private Const TCP_RCV_MAX = 255000     'Size input is closed (red, Yellow at 50%)
Private Const TCP_RCV_INUSE = 10000     'Size input is flagged as in use (Green)
Private Const SERIAL_RCV_MAX = 10000     'Size input is closed (red, Yellow at 50%)
Private Const SERIAL_RCV_INUSE = 1000     'Size input is flagged as in use (Green)
Private Const FILE_RCV_MAX = 2 ^ 16 '1000000     'Size input is closed (red, Yellow at 50%)
Private Const FILE_RCV_INUSE = 2000     'Size input is flagged as in use (Green)


Private MaxRcvPerTimerCycle As Long
Private MaxFilePerTimerCycle As Long
Private RcvThisTimerCycle As Long       '
Private FileThisTimerCycle As Long
Private FileCycleCount As Long       '
'Private RcvCycleCount As Long
Private RcvSpeedClock As New clsTimer          'To get Msgs received ino Rcv Buffers/min
'Private RcvAcceptedThisTimerCycle As Long
'Private RcvRejectedThisTimerCycle As Long

Private TcpXoff As Boolean 'False = leave data in windows TCP buffer
Private UdpXoff As Boolean
Private FileXoff As Boolean
'Private SerialXoff As Boolean
Private DataArrivalCalls As Long    'debug
Private InputBuffersToNmeaBufCalls As Long
Private MyInputThrottleTimer(0 To 6) As clsWaitableTimer    '0=UDP,5=TCP,6=File

Private InputLogFile As HugeBinaryFile

'  Private hbfFile As HugeBinaryFile
'15Sep  Private WithEvents InputFile As HugeBinaryFile
'  Private hbfOutFile As HugeBinaryFile
'  Private blnWriting As Boolean
'  Private lngBlocks As Long
'  Private bStopRun As Boolean
'  Private BlocksRead As Long
'  Private bytBuf() As Byte
'  Private Const BIG_BLOCK = 1000000
  
'  Private Const MAX_BLOCKS As Long = 5000
'  Private Const SMALL_BLOCK = 123456

'  Dim iFullReads As Long
'  Dim cLeftBytes As Currency
'  Dim iLeftBytes As Long

'-------- My decarations
Dim LastRemoteHost As String
Dim LastRemotePort As String

'Receive buffers
'May be a partial sentence if no delimiter
'Any data received after terminator is left in these buffers until the next call
Dim SerialRcvBuffer As String
Dim TCPRcvBuffer As String
Dim UDPRcvBuffer As String
Dim FileRcvBuffer As String

Dim ClientTCPError As Integer 'Generated by ClientTCP_Error event
Dim ClientTCPErrorDescription As String

Dim CrLfDetected As Boolean 'To Check baud rate
Dim FramingCount As Long    'No of framing errors since last CrLf

Dim ScheduledTimerElapsedMins As Long    'No of minutes since timer started
Dim ScheduledTimerElapsedMilliSecs As Long

'Together these equal UserScheduledSecs
Dim UserScheduledMins As Long 'Minutes=60000 milliseconds
Dim UserScheduledMilliSecs As Long 'Milliseconds

Private Sub cbDownload_Click()
Dim retval As Long
Dim BrowserPage As String

    If InputState <> 0 Then cbStop.Value = True
'button only enabled if a new version available
    BrowserPage = "http://" & DownloadURL & "ais_decoder_v3_downloads.html"
'this is microsofts way of starting default browser
'last arg 5 activates the window
        If Environ("windir") <> "" Then
            retval = ShellExecute(0, "open", BrowserPage, 0, 0, 5)
        Else
'try for linux compatibility
            Command = "winebrowser " & BrowserPage & " ""%1"""
            Shell (Command)
        End If
'V142        QueryQuit = False
        Unload NmeaRcv  'Terminates and saves vessels.dat
        Exit Sub
End Sub

Private Sub cbOptions_Click()
'WriteLog "cbOption_Click", LogForm
    Call TreeFilter.ScanCommPorts
    TreeFilter.Show
    TreeFilter.WindowState = vbNormal
'Don't show list Filters unless parent Node is ticked
'even if there are entries in the list
'To show these the user has to untick then tick the parent node
'If there are List entries but no parent node these will not be shown
End Sub

'Test cbPause.Caption = "Continue"  'input IS Paused
'Test cbPause.Caption <> "Continue" 'input NOT Paused
'Continue reading data into Input Buffers even when paused, Only Call StopInputOutput stops
'input
'(FileRcvBufffer, UDPRcvBuffer, TCPRcvBuffer, SerialRcvBuffer)
'and Call InputBuffersToNmeaBuf, but don't call ProcessNmeaBuf
'If NmeaBuffer is Full (Xoff=true), then continue reading into the Input Buffers
'but DoEvents after
Private Sub cbPause_Click() 'toggles each time pressed

'Change InputState first
    Select Case InputState
    Case Is = 1
        InputState = 2
        cbPause.Caption = "Continue"
    Case Is = 2
        InputState = 1
        cbPause.Caption = "Pause"
    Case Else
        MsgBox "Invalid InputState " & aInputState, , "cbPause_Click"
    End Select
 
 Debug.Print "-" & aInputState & "-"
    
    Select Case InputState
    Case Is = 1 'Started
        Call UpdateStats
        ScheduledTimer.Enabled = False
        List.NmeaDecodeList.ScrollTrack = False
'Unproceeed sentences are left in NmeaBuf when paused
'Clear NmeaBufFirst as it may be full as it fills first when Paused
        
 'Empty NmeaBuf first
        Do Until NmeaBufUsed = 0
            Call ProcessNmeaBuf
            Call UpdateStats
            Do Until RcvBufferedMsgs = 0 Or InputState <> 1
                Call InputBuffersToNmeaBuf  'Calls ProcessNmeaBuf - but not for file input
                Call UpdateStats
                Call ProcessNmeaBuf
                Call UpdateStats
                DoEvents    'Force stats to be updated
            Loop
            Call UpdateStats
        Loop
'Scheduledtimer not implemented
        Call ResumeProcessing("Paused")
        Call UpdateStats
'If any buffered messages then process these, even when Inputs have been closed
'because the Input buffers are full
'We have to exit the loop if pause has been pressed again while the input buffer
'is being cleared (It can be beacause DoEvents may have been activated)
'UDP
        If Check1(0).Value = vbGrayed Then
            Check1(0).Value = vbChecked 'Causes the Input to be re-opened
        End If
'TCP
        If Check1(5).Value = vbGrayed Then
            Check1(5).Value = vbChecked
        End If
'Serial
        If Check1(1).Value = vbGrayed Then
            Check1(1).Value = vbChecked
        End If
'File
        If Check1(6).Value = vbGrayed Then
'Occurs when the FileInputBuffer is emptied following a InputFile.Pause
            Check1(6).Value = vbChecked
'The File is not reopened by Check1(6) because it is already open
            InputFile.Pause = False         'InputFile must be open
        End If
        If ScheduledReq(0) = True Then ScheduledTimer.Enabled = True
'Will not re-start after pause
'        If Check1(5).Value = vbChecked Then
'May have been closed if Max_Rcv_buffer exceeeded
            If TcpXoff = True Then
                TcpXoff = False
            End If
'            If ClientTCP.State = sckClosed Then
'                Call OpenClientTCP
'            End If
'        End If
'Will not re-start after pause
'        If Check1(0).Value = vbChecked Then
'May have been closed if Max_Rcv_buffer exceeeded
'            If ClientUDP.State = sckClosed Then
'                Call OpenClientUdp
'            End If
            If UdpXoff = True Then
                UdpXoff = False
            End If
        RunClock.StartTimer
'        End If
    Case Is = 2 'Paused
        List.NmeaDecodeList.ScrollTrack = True
        ScheduledTimer.Enabled = False
        Processing.Suspended = True
        Processing.Paused = True
        RunClock.StopTimer
    End Select
End Sub

Private Sub cbSpawnGis_Click()
'this is microsofts way of starting the default browser
Dim r As Long
Dim Command As String

If Environ("windir") <> "" Then
    r = ShellExecute(0, "open", OutputFileName, 0, 0, 1)
Else
'try for linux compatibility
    Command = "winebrowser " & OutputFileName & " ""%1"""

    Shell (Command)
End If
End Sub

Private Sub cbStart_Click()

    If InputsCount = 0 Then
        MsgBox "You must select at least one Input Source" & vbCrLf _
        & "for AisDecoder to start processing", vbInformation, "Start"
        Exit Sub
    End If
'Change InputState first
    Select Case InputState
    Case Is = 0
        InputState = 1
        cbStart.Enabled = False
        cbStop.Enabled = True
'if cbStart is clicked in the .main because "Start" is set as command line option
'the form will not be visible and it must be otherwise set focus causes an error
        If NmeaRcv.Visible = True Then
            cbStop.SetFocus
        End If
        cbPause.Enabled = True
    Case Else
        MsgBox "Invalid InputState " & aInputState, , "cbStart_Click"
    End Select
 Debug.Print "-" & aInputState & "-"
     
    Select Case InputState
    Case Is = 1
        Call StartInputOutput
'Debug.Print "#cbStartExit " & InputState
    End Select
End Sub

Private Sub cbStop_Click()
'Change InputState first
Dim bytesTotal As Long

'Debug.Print "#cbStop was " & aInputState

'The TCP device Buffer will not be empty if data just stops being received
    bytesTotal = ClientTCP.BytesReceived
    If bytesTotal > 0 Then
        Call ClientTCP_DataArrival(bytesTotal)
    End If
    
'UDP cannot return if there is data in buffer, so just call it to empty buffer
    If ClientUDP.State <> sckClosed Then
        Call ClientUDP_DataArrival(bytesTotal)
    End If
    
    Select Case InputState
    Case Is = 1, 2
        InputState = 0
'If cmdNoWindows = true, this call is invalid
        cbStop.Enabled = False
        cbPause.Caption = "Pause"   'Continue to allow routines paused to exit
        cbPause.Enabled = False
    Case Else
        MsgBox "Invalid InputState " & aInputState, , "cbStop_Click"
    End Select
    Select Case InputState
    Case Is = 0
'dont re-enable until FileRcvBuffer and NmeaBuf is emptied
'        Set StopWaitableTimer = New clsWaitableTimer
'        Do Until NmeaBufUsed = 0
'            Call ProcessNmeaBuf
                Do Until Len(FileRcvBuffer) = 0
                    Call FlushInputBuffer(FileRcvBuffer)
                Loop
'            StopWaitableTimer.Wait 50
'        Loop
'        Set StopWaitableTimer = Nothing
        cbStart.Enabled = True
        If NmeaRcv.Visible = True Then
            cbStart.SetFocus
        End If
        Call UpdateStats
        StatusBar.Panels(1).Text = StatusBar.Panels(1).Text & "Terminating Input"
        Call StopInputOutput
    End Select
End Sub

'Displays the Latest Version Updates - does not terminate AisDecoder
Private Sub cbUpdates_Click()
'    Call cbDownload_Click
    Call HttpSpawn("http://" & DownloadURL & "ais_decoder_v3_updates.html")
    cbUpdates.Enabled = False   'remove dotted line round button
    cbUpdates.Enabled = True
End Sub

Private Sub Check1_Click(Index As Integer)
Dim Channel As Long

'Forms are "Shown" when data is received or output
'We remove the window irrespective of if we are started or stopped
'Here we only Hide the form if the check box is not ticked
    Select Case Index
    Case Is = 0, 1, 5, 6
'Close Inputs if not stopped
'vbGrayed is clicked when RcvBuffer exceeds maximum bytes
        If InputState <> 0 Then
            Select Case Index
'UDP
            Case Is = 0 'UDP Input
                Select Case Check1(Index).Value
                Case Is = vbChecked
                    Call OpenClientUdp
                Case Is = vbUnchecked, vbGrayed
                    Call CloseUDPInput
               End Select
'TCP
            Case Is = 5 'TCP Input
                Select Case Check1(Index).Value
                Case Is = vbChecked
                    Call OpenClientTCP
                Case Is = vbUnchecked, vbGrayed
                    Call CloseTCPInput
                End Select
'Serial
            Case Is = 1 'Serial Input
                Select Case Check1(Index).Value
                Case Is = vbChecked
                    If SerialPortCount > 0 Then
                        If MSComm1.PortOpen = False Then
                            Call OpenSerialInput
                        End If
                    End If
                Case Is = vbUnchecked, vbGrayed
                    Call CloseSerialInput
                End Select
 'File
            Case Is = 6 'File Input
                Select Case Check1(Index).Value
                Case Is = vbChecked
'if InputFile.Pause is set - then dont try and re-open file
'Because it will already be open
                    If InputFile Is Nothing Then
                        Call OpenFileInput
                    End If
                Case Is = vbUnchecked
'Happens when FileInputBuffer is full (because of Pause) and stop clicked
'as well as when FileInput is stopped by unticking
'Must CloseFileInput before cbStop because cbStop will only stop Ticked inputs
'and File is unticked when Check1 was called
                                        
                    Call CloseFileInput 'StopInput only
                        'InputState=1
                        'FileNextBlockTimer now disabled
                    If InputsCount = 0 Then 'FileInput is the only input
'Although the Inputs have been closed, whith FileInputOnly we also Close the Output files
                        cbStop.Value = True 'Close OUTPUTS (no inputs open)
                    End If
                Case Is = vbGrayed
'Occurs when the FileInputBuffer is filled
                    InputFile.Pause = True         'InputFile must be open
                End Select  'Check1
           End Select  'Input Source
    
            If NmeaBufXoff = True Then
                If Processing.Suspended = False Then Call ProcessNmeaBuf
            End If
        End If  'Inputs
    
    Case Is = 2    'Output File Display
        Channel = 1
        If Check1(2).Value = 0 Then
            DisplayOutput(Channel) = False
            Outputs(Channel).Hide
        Else
            DisplayOutput(Channel) = True
            Call TreeFilter.SetOutputOptions
        End If
    Case Is = 3     'Output UDP Display
        Channel = 2
        If Check1(3).Value = 0 Then
            DisplayOutput(Channel) = False
            Outputs(Channel).Hide
        Else
            DisplayOutput(Channel) = True
            Call TreeFilter.SetOutputOptions
        End If
    Case Is = 4     'Output FTP Display
        If Check1(4).Value = 0 Then
            frmFTP.Hide
        Else
            If cmdNoWindow = False Then frmFTP.Show
        End If
    End Select

End Sub

Public Sub ClearStatusBarTimer_Timer()
    StatusBar.Panels(1).Text = ""
    ClearStatusBarTimer.Enabled = False
End Sub

Private Sub ClientTCP_Connect()
'Send any login commands
Dim kb As String
Dim ch As Long

    If FileExists(TcpLoginFileName) Then
        ch = FreeFile
        Open TcpLoginFileName For Input As #ch
        Do Until EOF(ch)
            Line Input #ch, kb
            ClientTCP.SendData (kb & vbCrLf)
            DoEvents
        Loop
        Close #ch
    End If
End Sub

'Private Sub FileData_FileDataArrival()
'Stop
'End Sub

Private Sub Form_Load()         ' The control's name is NmeaRcv
Dim i As Long
Dim tot As Long
#If UseComm Then
    Dim Idx As Long
#End If

'MsgBox "in nmearcv"

ReDim NmeaBuf(NMEABUF_MAX)

With NmeaRcv
    .Caption = App.EXEName & " - Control/Stats [" & NameFromFullPath(IniFileName) & "]"
    .StatusBar.Panels(1).AutoSize = sbrContents
'    If ScaleWidth <> 0 Then StatusBar.Panels(1).Width = ScaleWidth
    .StatusBar.Panels(1) = "Directory is " & Environ("APPDATA") & "\Arundale\Ais Decoder\"
    .ClearStatusBarTimer.Enabled = True
End With
With Stats
    .TextMatrix(0, 0) = "Total Bytes Rx"
    .TextMatrix(1, 0) = "Buffered Bytes"
    .TextMatrix(2, 0) = "Rejected"
    .TextMatrix(3, 0) = "Received"
    .TextMatrix(4, 0) = "Waiting"
    .TextMatrix(5, 0) = "Processed"
    .TextMatrix(6, 0) = "Filtered"
    .TextMatrix(7, 0) = "Outputted"
    .TextMatrix(8, 0) = "Scheduled"
    .TextMatrix(9, 0) = "Named Vessels"
    .TextMatrix(10, 0) = "Last Output"
   tot = 50
    .ColWidth(0) = 1200
    .ColWidth(1) = 1300
    For i = 0 To .Cols - 1
        tot = tot + .ColWidth(i)
    Next i
    .Width = tot
    tot = 70
    For i = 0 To .Rows - 1
        tot = tot + .RowHeight(i) + 1
     .TextMatrix(i, 1) = 0  'set long
    Next i
    .Height = tot
End With

StatusBar.Panels(1).MinWidth = 1
StatusBar.Panels(1) = IniFileName
ClearStatusBarTimer.Enabled = True
cbUpdates.Caption = App.Major & "." & App.Minor & "." & App.Revision
If EditOptions = False Then
    cbOptions.Enabled = False
End If
Option1(14).Value = True
    
#If UseComm Then
    Set LogForm = Me
'Test WriteLog and display on Status Bar
'    Call WriteLog("Log Message", LogForm)
    ReDim sockets(1 To 1)    'Defined in modCommHandler
    ReDim Comms(1 To 1)      'Defined in modConnHandler
    Idx = 1
#End If

End Sub

'Forwards all complete sentences in one go to InputBuffersToNmeaBuf
'Leaves any remaining bytes in RcvBuffer

Private Sub ClientUDP_DataArrival(ByVal bytesTotal As Long)
Dim DataRcv As String
Dim cpos As Long
Dim Buf As String

'Call UpdateStats
DataArrivalCalls = DataArrivalCalls + 1
'    On Error GoTo DataArrival_err

'With UDP we must put the received data into the RCV buffer because when the data is recieved
'we cannot leave it in the Windows buffer. Receiving it automatically removes it
    
'This is unusual - we keep getting the data until we have emptied the Windows
'UDP buffer. This is to avoid at fast data rates, generating a Arrival event every time a couple of UDP sentences
'are received. This can cause a stack overflow, because while processing

    Do
        DataRcv = ""
        ClientUDP.GetData DataRcv
'Debug.Print "#ClientUDP_DataArrival(" & DataArrivalCalls & ")" & LFCount(DataRcv) & "+" & Len(DataRcv) ' - (InStrRev(DataRcv, vbCrLf) + 1)
        UDPRcvBuffer = UDPRcvBuffer & DataRcv
        cBytesRx = cBytesRx + Len(DataRcv) * 0.0001@
    Loop Until DataRcv = "" Or Len(UDPRcvBuffer) > UDP_RCV_MAX
'If buffer is full exit the loop to prevent 100% cpu usage
        
'If buffer full Display to user (UDP input will also be closed - not sure why)
    If Len(UDPRcvBuffer) > UDP_RCV_MAX Then Call UpdateStats
    If UdpXoff = True Then
'Must not doevents because UDP data arrival can be called again before the sub is exited
'causing out of stack space
DataArrivalCalls = DataArrivalCalls - 1
'        DoEvents
'        If Len(UDPRcvBuffer) > 1 Then Stop
        Exit Sub
    Else
        UdpXoff = True
'Debug.Print "Xoff (UDP)"    'stop data
    End If

'If there is any space in NmeaBuf, remove as many complete sentences as we can from
'Rcv buffer, otherwise we continue adding then into the Rcv buffer
    If NmeaBufXoff = False Then
        Call InputBuffersToNmeaBuf
    Else
Debug.Print "UDPRcvBuffer " & LFCount(UDPRcvBuffer)
        DoEvents    'Otherwise Click Pause is not actioned
'        MyInputThrottleTimer(0).Wait 50
    End If
    
    If UdpXoff = True Then  'Start accepting UDP data from the Windows UDP buffer
        UdpXoff = False
'Debug.Print "Xon (UDP)"
    End If
 DataArrivalCalls = DataArrivalCalls - 1

 Exit Sub

DataArrival_err:
    On Error GoTo 0
    Select Case err.Number
    Case Is = sckBadState
        StatusBar.Panels(1).Text = err.Description
        ClearStatusBarTimer.Enabled = True
        'Believe this occurs when the buffer is full
        'Wrong protocol or connection state for the requeste
        'transaction or request
    Case Is = sckMsgTooBig
        StatusBar.Panels(1).Text = "ClientUDP_DataArrival " & err.Description
        ClearStatusBarTimer.Enabled = True
    Case Is = sckConnectionReset
        StatusBar.Panels(1).Text = "ClientUDP_DataArrival " & err.Description
        ClearStatusBarTimer.Enabled = True
        'Reset by remote side is OK
    Case Else
'Display the sentence we're processing that caused the error (if not trapped in the routine
'in error
        StatusBar.Panels(1).Text = "ClientUDP_DataArrival, Error " & err.Number & " - " & err.Description
        ClearStatusBarTimer.Enabled = True
        'Call WriteErrorLog(StatusBar.Panels(1).Text & vbCrLf & clsSentence.NmeaSentence)
'Try & stop Invald Proceedure call or argument (Ulf)
'        MsgBox "Client_UDP DataArrival Error " & Str(err.Number) & " " & err.Description
    End Select
    err.Clear
    Resume Next
End Sub

Private Sub ClientTCP_ConnectionRequest(ByVal requestID As Long)
    If ClientTCP.State <> sckClosed Then ClientTCP.Close
    ClientTCP.Accept requestID
End Sub

'Forwards all complete sentences in one go to InputBuffersToNmeaBuf
'Leaves any remaining bytes in RcvBuffer
'If Paused Data continues being put into Rcv buffer
'If NmeaBuf is full with TCP you keep getting data into the Windows TCP buffer
'When Windows TCP buffer is full, Windows stops accepting TCP data

Private Sub ClientTCP_DataArrival(ByVal bytesTotal As Long)
Dim DataRcv As String
Dim cpos As Long
Dim Buf As String
Static CallCount As Long    'To stop recursive calls

If CallCount > 0 Then
    Track.TCPCallCountExit = Track.TCPCallCountExit + 1
'    Exit Sub
End If
CallCount = CallCount + 1

Track.TCPArrival = Track.TCPArrival + 1

'Do not remove from the Windows buffer, if more data arrives while we are processing
'the current call (potential stack overflow)

    On Error GoTo DataArrival_err
    If TcpXoff = True Then
'If you don't then Tick File is unresponsive
        DoEvents
        CallCount = CallCount - 1
        Exit Sub
    Else
        TcpXoff = True
'Debug.Print "Xoff (TCP)"    'stop data
    End If

'Must loop round GetData otherwise data in DataRcv is lost
    Track.TCPDataRcvLoops = Track.TCPDataRcvLoops - 1
    Do
        ClientTCP.GetData DataRcv
        Track.TCPDataRcvLFCount = Track.TCPDataRcvLFCount + LFCount(DataRcv)
'Debug.Print "#ClientTCP_DataArrival " & LFCount(DataRcv) & "+" & Len(DataRcv) - (InStrRev(DataRcv, vbCrLf) + 1)
        cBytesRx = cBytesRx + Len(DataRcv) * 0.0001@
        TCPRcvBuffer = TCPRcvBuffer & DataRcv
'If there is any space in NmeaBuf, remove as many complete sentences as we can from
'Rcv buffer, otherwise we continue adding then into the Rcv buffer
        Track.TCPDataRcvLoops = Track.TCPDataRcvLoops + 1
    
    Loop Until DataRcv = ""
    If NmeaBufXoff = False Then
        Call InputBuffersToNmeaBuf
    Else
Debug.Print "TcpRcvBuffer " & Len(TCPRcvBuffer)
        If Len(TCPRcvBuffer) > TCP_RCV_MAX Then
            If Check1(5).Value = vbChecked Then
                Check1(5).Value = vbGrayed
                Call CloseTCPInput
            End If
        End If
    End If
    
    If TcpXoff = True Then  'Start accepting TCP data from the Windows TCP buffer
        TcpXoff = False
'Debug.Print "Xon (TCP)"
    End If
    
'   DoEvents    'MUST NOT in data arrival - causes loss of data into DataRcv
    
    CallCount = CallCount - 1
Exit Sub

DataArrival_err:
'    On Error GoTo 0
'Stop    'v131debug
    Select Case err.Number
    Case Is = sckBadState
        StatusBar.Panels(1).Text = err.Description
        ClearStatusBarTimer.Enabled = True
        'Believe this occurs when the buffer is full
        'Wrong protocol or connection state for the requeste
        'transaction or request
'Stop
    Case Is = sckMsgTooBig
        StatusBar.Panels(1).Text = err.Description
        ClearStatusBarTimer.Enabled = True
    Case Is = sckConnectionReset
        StatusBar.Panels(1).Text = err.Description
        ClearStatusBarTimer.Enabled = True
        'Reset by remote side is OK
    Case Else
'        Call frmDpyBox.DpyBox("ClientTCP_DataArrival Error " & Str(err.Number) & " " & err.Description, 5, "ClientTCP_DataArrival")
'        ClearStatusBarTimer.Enabled = True
'Display the sentence we're processing that caused the error (if not trapped in the routine
'in error
        StatusBar.Panels(1).Text = "ClientTCP_DataArrival, Error " & err.Number & " - " & clsSentence.NmeaSentence ' err.Description
        ClearStatusBarTimer.Enabled = True
'        Call WriteErrorLog(StatusBar.Panels(1).Text & vbCrLf & clsSentence.NmeaSentence)
    End Select
    err.Clear
    Resume Next    'Quit routine
End Sub

'FileNextBlockTimer keeps trying to get another block until File is closed
'More data is only got when file rcvbuffer is empty
Private Sub FileNextBlock()
Dim DataRcv As String
Dim CloseInput As Boolean
Dim cpos As Long
    
Debug.Print "'";
    
    On Error GoTo FileNextBlock_err
Debug.Print vbCrLf & "#FileNextBlock " & Len(FileRcvBuffer) & "+" & LFCount(DataRcv) & "+" & Len(DataRcv) - (InStrRev(DataRcv, vbLf))
'Will not be open if paused, Can be closed before NextBlockTimer is disabled
    If Not (InputFile Is Nothing) Then
'Only get more data if FileRcvBuffer is empty except partial sentences
'If some data has been read
        If Len(FileRcvBuffer) < 128000 Then
'MsgBox FileRcvBuffer & vbCrLf & Len(FileRcvBuffer), , "FileNextBlock" 'FileRcvBuffer has any partial sentences still in it
            InputFile.NextFileBlock DataRcv   'Includes incomplete sentences, clears ascBuf
'Debug.Print Len(DataRcv)
                                                'DataRcv is ascBuf (byRef)
            cBytesRx = cBytesRx + Len(DataRcv) * 0.0001@
'Debug.Print "#FileNextBlock " & Len(FileRcvBuffer) & "+" & LFCount(DataRcv) & "+" & Len(DataRcv) - (InStrRev(DataRcv, vbCrLf) + 1)
Else
            If InputsCount > 1 Then
'Allow other inputs and user Clicks to interrupt
                MyInputThrottleTimer(6).Wait 50
            End If
        End If
'Must call FileDataArrival even if no new data, to re-try InputBuffersToNmeaBuf
        Call FileDataArrival(DataRcv)
'The Input file can have been closed by cbStop interrupt, while in FileDataArrival (called above)
        If Not (InputFile Is Nothing) Then
            If InputFile.EOF Then
                If InputsCount = 1 Then 'FileInput is the only input
                    cbStop.Value = True 'Stop Input and Output
                Else
                    Call CloseFileInput 'StopInput only
                End If
            End If
        Else
'Stop   'Input File Closed by cbStop interrupt, while in FileDataArrival (called above)
        End If
    Else    'InputFile Is Nothing
'Stop
    End If
'Debug.Print "#DataNextBlock " & Len(FileRcvBuffer) & "+" & LFCount(DataRcv) & "+" & Len(DataRcv) - (InStrRev(DataRcv, vbLf))
'The FileBlockTimer will not fire while we are still in FileNextBlock
Exit Sub

FileNextBlock_err:
    Select Case err.Number
    Case Else
        StatusBar.Panels(1).Text = "FileNeatBlock, Error " & err.Number & " - " & err.Description
        ClearStatusBarTimer.Enabled = True
'        Call WriteErrorLog(StatusBar.Panels(1).Text & vbCrLf & clsSentence.NmeaSentence)
    End Select
    err.Clear
    Resume Next    'Quit routine
End Sub

'_1Oct14
Private Sub FileDataArrival(ByRef DataRcv As String)
Dim cpos As Long
'Dim DataRcv As String
Dim FileRcvBufferCount As Long
Dim DataRcvCount As Long
Dim EOF As Boolean
'Do not remove from the Windows buffer, if more data arrives while we are processing
'the current call (potential stack overflow)
'MsgBox DataRcv

    On Error GoTo DataArrival_err
    If FileXoff = True Then
'If you don't then Tick File is unresponsive
        DoEvents
        Exit Sub
    Else
        FileXoff = True
'Debug.Print "Xoff (File)"    'stop data
    End If

'Input file could have been closed since last called
    If InputFile Is Nothing Then
        EOF = True
    End If
        
        Do
'we exit this loop when the file is EOF or has been closed
'        Do Until CRLFCount(FileRcvBuffer) >= MaxFilePerTimerCycle
'We exit this loop when we have enough sentences in the FileRcvBuffer to output
'All sentences in the FileRcvBuffer are output
 'only get more data into datarcv when all data had been tranferred to FileRcvBuffer
'31/3/15            If Len(DataRcv) = 0 Then
            If Len(DataRcv) = 0 And InputState <> 2 Then
                InputFile.NextFileBlock DataRcv   'Includes incomplete sentences, clears ascBuf
                                                'DataRcv is ascBuf (byRef)
                cBytesRx = cBytesRx + Len(DataRcv) * 0.0001@
Debug.Print "FileNextBlock " & LFCount(DataRcv) & "+" & Len(DataRcv) - (InStrRev(DataRcv, vbCrLf) + 1)
            Else
Debug.Print "FileBlock " & LFCount(DataRcv) & "+" & Len(DataRcv) - (InStrRev(DataRcv, vbCrLf) + 1)
            End If
            DataRcvCount = LFCount(DataRcv)
            If MaxFilePerTimerCycle = 0 Then
                FileRcvBuffer = FileRcvBuffer & DataRcv
FileRcvBufferCount = LFCount(FileRcvBuffer)
                DataRcv = ""
            Else
'This is the maximum no of sentences we can leave in the FileRcvBuffer when
'all sentences in DataRcv are transferred to the FileRcvBuffer
                FileRcvBufferCount = LFCount(FileRcvBuffer)
'If we have at least enough transfer the minimum no of sentences to Output
                If DataRcvCount + FileRcvBufferCount > MaxFilePerTimerCycle Then
'MaxFilePerTimerCycle = 60
                    cpos = CRLFPos(DataRcv, MaxFilePerTimerCycle - FileRcvBufferCount)
'cpos = CRLFPos(DataRcv, 650)
                    FileRcvBuffer = FileRcvBuffer + Left$(DataRcv, cpos + 1)
FileRcvBufferCount = LFCount(FileRcvBuffer)
                    DataRcv = Mid$(DataRcv, cpos + 2)
DataRcvCount = LFCount(DataRcv)
                Else
'If not enough data, transfer what weve got
                    FileRcvBuffer = FileRcvBuffer & DataRcv
FileRcvBufferCount = LFCount(FileRcvBuffer)
                    DataRcv = ""
                End If
            End If
'We now have all the sentences we can transfer in oneTimerCycle in FileRcvBuffer
'There may or may not be some left in DataRcv
            If NmeaBufXoff = False Then
                Call InputBuffersToNmeaBuf
Else
'Stop
            End If
'The Stop event may have closed the Input file

'Allow events - Could have closed input file
            If InputFile Is Nothing Then
                EOF = True
                Exit Do
            Else
                If InputFile.EOF = True Then
                    Exit Do
                End If
            End If
'We must wait in this loop with EventsEnabled in order to move the remaining data
'into the FileRcvBuffer
'31/3/15            If Len(DataRcv) > 0 Then
            If Len(DataRcv) > 0 Or InputState = 2 Then
                MyInputThrottleTimer(6).Wait SPEED_INTERVAL
            End If
'Events during the wait - Could have closed input file
            If InputFile Is Nothing Then
                EOF = True
                Exit Do
            End If
'Call UpdateStats
        Loop
'Exits this loop when we have an EOF
'But the input file may not actually have been closed
        If Not (InputFile Is Nothing) Then
            If InputFile.EOF Then
            'Here if EOF InputState=1 File needs closing
            'CloseInput = True
            'Timer enabled
            'InputFile not nothing
                If InputsCount = 1 Then 'FileInput is the only input
                    cbStop.Value = True 'Stop Input and Output
                Else
                    Call CloseFileInput 'StopInput only
                End If
            End If
        End If
    
'Call even if no data in DataRcv, may still be data in other Input Buffers
    If NmeaBufXoff = False Then
        Call InputBuffersToNmeaBuf
    End If
    
    If FileXoff = True Then  'Start accepting File data from the File buffer
        FileXoff = False
'Debug.Print "Xon (File)"
    End If
 
Exit Sub

DataArrival_err:
    Select Case err.Number
    Case Else
        StatusBar.Panels(1).Text = "FileDataArrival, Error " & err.Number & " - " & err.Description
        ClearStatusBarTimer.Enabled = True
'        Call WriteErrorLog(StatusBar.Panels(1).Text & vbCrLf & clsSentence.NmeaSentence)
    End Select
    err.Clear
    Resume Next    'Quit routine
End Sub

#If False Then 'V142
Private Sub Form_Unload(Cancel As Integer)
Dim i As Long

    For i = 0 To UBound(MyInputThrottleTimer)
        Set MyInputThrottleTimer(i) = Nothing
    Next i
    Set RcvSpeedClock = Nothing
'Must not force END otherwise NmeaRcv is not unloaded properly
'    End 'v140 terminate program

End Sub
#End If

'Adds valid data into SerialRcvBuffer
'Does not require a Xoff because data is not received fast enough
Private Sub MSComm1_OnComm()
Dim b() As Byte
Dim Num_Bytes As Long
Dim InBuf As String
Dim noEOL As Boolean
Dim i As Long
Dim cpos As Long                            '
Dim ByteNo As Long
Dim HexString As String
Dim AsciiString As String
Dim BufferString As String
Dim DevSettings() As String
Dim FramingErrorCount As Long
Dim DataRcv As String

With MSComm1
'test for incoming event
    If .CommEvent <> comEvReceive Then
 Debug.Print "CE=" & .CommEvent
    End If
    
    Select Case .CommEvent
   ' Errors
      Case comEventBreak   ' A Break was received.
        err.Clear
'        MsgBox "Event Break"   'happens when input baud rate is changed
      Case comEventFrame   ' Framing Error
        FramingCount = FramingCount + 1

'If you anticipate that the data you are receiving will encounter (usually
'many are required) parity errors you can do one of two things.  When you
'detect the error, you can close and reopen the port (via a timer, not
'directly), or... You can add a Timer control that polls the Input property
'instead of using OnComm.
      Case comEventOverrun   ' Data Lost.
 'see http://www.xtremevbtalk.com/showthread.php?t=254354
 'You MUST react to an overrun by minimizing the overhead of your program immediately.
 'No updating graphics or text boxes. This can cause more overruns
 '       MsgBox "Event Overrun" 'v3.2.135 Tim Last Model Form Error
      Case comEventRxOver   ' Receive buffer overflow.
'        MsgBox "Event Rx Over"
            err.Clear
      Case comEventRxParity   ' Parity Error.
'        MsgBox "Event Rx Parity"
            err.Clear
      Case comEventTxFull   ' Transmit buffer full.
        MsgBox "Event Tx Full"
      Case comEventDCB   ' Unexpected error retrieving DCB]
        MsgBox "Event DCB "

   ' Events
      Case comEvCD   ' Change in the CD line.
      Case comEvCTS   ' Change in the CTS line.
      Case comEvDSR   ' Change in the DSR line.
      Case comEvRing   ' Change in the Ring Indicator.
      Case comEvReceive   ' Received RThreshold # of
'ensure we clear the buffer
'time to wait for rest of sentence
        OpenSerialTimer.Enabled = False   'must disable the restart timer
'Now detecting <LF>
'        OpenSerialTimer.Interval = 100   'reset time to wait for data
'        OpenSerialTimer.Enabled = True
        Num_Bytes = .InBufferCount
                    '// Number of Bytes received
        cBytesRx = cBytesRx + .InBufferCount * 0.0001@
        b() = .Input
 'Debug.Print UBound(b)
 'Debug.Print Num_Bytes
                    '// Byte Array containing Bytes received
'Can happen if Oncomm Event called before previous one has terminated
        If UBound(b) <> -1 Then
            DataRcv = DataRcv & StrConv(b, vbUnicode)
        End If
'        For i = 0 To UBound(b)
'            DataRcv = DataRcv & Chr(b(i))
'        Next i
'Remove all in the SerialRcvBuffer before the first CRLF is detected since the CommPort
'was opened - CrLfDetected is defined in NmeaRcv
        If CrLfDetected = False Then
 '           .RThreshold = 200
            cpos = InStr(1, DataRcv, vbCrLf)
'Got the first CrLf
            If cpos > 0 Then
'Clear upto first cflf
                DataRcv = (Mid$(DataRcv, cpos + 2))
                CrLfDetected = True
                DevSettings = Split(.Settings, ",")
                TreeFilter.Combo1(1).Text = DevSettings(0)
'                .RThreshold = 1
            End If
        End If
        
'Check if DataRcv is valid
        If CrLfDetected = True Then
            FramingCount = 0
'Debug.Print "#SerialDataArrival " & LFCount(DataRcv) & "+" & Len(DataRcv) - (InStrRev(DataRcv, vbCrLf) + 1)
            SerialRcvBuffer = SerialRcvBuffer & DataRcv
            If NmeaBufXoff = False Then
                Call InputBuffersToNmeaBuf
            End If
        Else    'No CRLF detected - reject until we get the first CRLF
            If Len(SerialRcvBuffer) > 200 Or FramingCount > 3 Then
 '          If FramingCount > 3 Then
 'Debug.Print "Buflen=" & Len(SerialRcvBuffer)
 'Debug.Print "FE=" & FramingCount
                .PortOpen = False
                DataRcv = ""    'Clear junk from buffer
                CrLfDetected = False
                FramingCount = 0
                DevSettings = Split(.Settings, ",")
                Select Case DevSettings(0)  'cycle baud rate
                Case Is = 1200
                    TreeFilter.Combo1(1).Text = 2400
                Case Is = 2400
                    TreeFilter.Combo1(1).Text = 4800
                Case Is = 4800
                    TreeFilter.Combo1(1).Text = 9600
                Case Is = 9600
                    TreeFilter.Combo1(1).Text = 19200
                Case Is = 19200
                    TreeFilter.Combo1(1).Text = 38400
                Case Is = 38400
                    TreeFilter.Combo1(1).Text = 2400
                Case Else
                    TreeFilter.Combo1(1).Text = 4800
                End Select
                DevSettings(0) = TreeFilter.Combo1(1).Text
                .Settings = Join(DevSettings, ",")
Debug.Print .Settings
                .PortOpen = True
            End If   'cycling baud rate
        End If       'got at least one CRLF
    Case comEvSend   ' There are SThreshold number of
                     ' characters in the transmit
                     ' buffer.
    Case comEvEOF   ' An EOF charater was found in
                     ' the input stream
    End Select
End With 'MSComm1
    
End Sub

'Called from NmeaRcv.CommRcv
'CommRcv is Called from CommBuffer using the CallbackObject which is set as
'NmeaRcv in AisDecoder. CommBuffer is common to NmeaRouter.
'The DataArrival Code is separated from CommRcv because the structure is the same
'for TCP,UDP,Serial,File etc
Private Sub CommDataArrival(DataRcv As String, ByVal bytesTotal As Long)
'insert code from oncomm that is not specific to MSComm
'Dim b() As Byte
'Dim Num_Bytes As Long
'Dim InBuf As String
'Dim noEOL As Boolean
'Dim i As Long
Dim cpos As Long                            '
'Dim ByteNo As Long
'Dim HexString As String
'Dim AsciiString As String
'Dim BufferString As String
'Dim DevSettings() As String
'Dim FramingErrorCount As Long

'ensure we clear the buffer
'time to wait for rest of sentence
        OpenSerialTimer.Enabled = False   'must disable the restart timer
'Now detecting <LF>
'        OpenSerialTimer.Interval = 100   'reset time to wait for data
'        OpenSerialTimer.Enabled = True
                    '// Number of Bytes received
        cBytesRx = cBytesRx + bytesTotal * 0.0001@
'Remove all in the SerialRcvBuffer before the first CRLF is detected since the CommPort
'was opened - CrLfDetected is defined in NmeaRcv
        If CrLfDetected = False Then
            cpos = InStr(1, DataRcv, vbCrLf)
'Got the first CrLf
            If cpos > 0 Then
'Clear upto first cflf
                DataRcv = (Mid$(DataRcv, cpos + 2))
                CrLfDetected = True
'Display BaudRate (Autodetect may have change it)
'                DevSettings = Split(.Settings, ",")
'                TreeFilter.Combo1(1).Text = DevSettings(0)
            End If
        End If
        
'Check if DataRcv is valid
        If CrLfDetected = True Then
'Debug.Print "#SerialDataArrival " & LFCount(DataRcv) & "+" & Len(DataRcv) - (InStrRev(DataRcv, vbCrLf) + 1)
            SerialRcvBuffer = SerialRcvBuffer & DataRcv
            If NmeaBufXoff = False Then
                Call InputBuffersToNmeaBuf
            End If
        Else    'No CRLF detected - reject until we get the first CRLF
            If Len(SerialRcvBuffer) > 200 Then
 'Debug.Print "Buflen=" & Len(SerialRcvBuffer)
                WriteLog "No <cr><lf> detected", LogForm
                DataRcv = ""    'Clear junk from buffer
                CrLfDetected = False
'Autodetect baud rate
'               .PortOpen = False
'                DevSettings = Split(.Settings, ",")
'                Select Case DevSettings(0)  'cycle baud rate
'                Case Is = 1200
'                    TreeFilter.Combo1(1).Text = 2400
'                Case Is = 2400
'                    TreeFilter.Combo1(1).Text = 4800
'                Case Is = 4800
'                    TreeFilter.Combo1(1).Text = 9600
'                Case Is = 9600
'                    TreeFilter.Combo1(1).Text = 19200
'                Case Is = 19200
'                    TreeFilter.Combo1(1).Text = 38400
'                Case Is = 38400
'                    TreeFilter.Combo1(1).Text = 2400
'                Case Else
'                    TreeFilter.Combo1(1).Text = 4800
'                End Select
'                DevSettings(0) = TreeFilter.Combo1(1).Text
'                .Settings = Join(DevSettings, ",")
'Debug.Print .Settings
'                .PortOpen = True
            End If   'cycling baud rate
        End If       'got at least one CRLF
End Sub


'If more than one InputBuffer has data, remove the least processor intensive first
'We dont't need to pass the buffer address, testv131
'Can contain incomplete sentence
Private Sub InputBuffersToNmeaBuf()

    If NmeaBufXoff = True Then  'NmeaBuf is full
        If Processing.Suspended = False Then Call ProcessNmeaBuf
        Exit Sub
    End If
    
'    If InStrRev(SerialRcvBuffer, vbCrLf) > 0 Then
'        Call RejectSentences(SerialRcvBuffer)
'        Call InputToNmeaBuf(SerialRcvBuffer, "Serial")
'    ElseIf InStrRev(UDPRcvBuffer, vbCrLf) > 0 Then
'        Call InputToNmeaBuf(UDPRcvBuffer, "UDP")
'    ElseIf InStrRev(TCPRcvBuffer, vbCrLf) > 0 Then
'        Call RejectSentences(TCPRcvBuffer)
'        Call InputToNmeaBuf(TCPRcvBuffer, "TCP")
'    ElseIf InStrRev(FileRcvBuffer, vbCrLf) > 0 Then
'FileRcv does not reject any sentences, speed is determined in FileDataArrival
'because the file speed is set higher and they must pass through because otherwise
'we would reject sentences from the file input
'        Call InputToNmeaBuf(FileRcvBuffer, "File")
    If InStrRev(SerialRcvBuffer, vbCrLf) > 0 Then
        Call InputToNmeaBuf(SerialRcvBuffer, "Serial")
    End If
    If InStrRev(UDPRcvBuffer, vbCrLf) > 0 Then
        Call InputToNmeaBuf(UDPRcvBuffer, "UDP")
    End If
    If InStrRev(TCPRcvBuffer, vbCrLf) > 0 Then
        Call InputToNmeaBuf(TCPRcvBuffer, "TCP")
    End If
   If InStrRev(FileRcvBuffer, vbCrLf) > 0 Then
'FileRcv does not reject any sentences, speed is determined in FileDataArrival
'because the file speed is set higher and they must pass through because otherwise
'we would reject sentences from the file input
        Call InputToNmeaBuf(FileRcvBuffer, "File")
    End If
    DoEvents
'Stop
End Sub


'Extracts all complete sentences from the Input Buffers
'(FileRcvBufffer, UDPRcvBuffer, TCPRcvBuffer, SerialRcvBuffer)
'These are string buffers and can not overflow
'Puts the complete sentences into NmeaRcvBuf
'If NmaRcvBuf exceeds 80% full Exits leaving any remaining sentences in the Input Buffer
'and sets Xoff to stop InputBuffersToNmeaBuf being called until Xoff=False
'Leaves any remaining bytes in the Input Buffers including incomplete sentences

'If Paused or Scheduler is currently running all incoming sentences are held in
'NmeaBuf (ProcessNmeaBuf is not called). In addition DoEvents is called each time
'a Sentence is inserted into NmeaBuf to ensure the program remains responsive to
'User interaction and Stats are updated by the StatsTimer.

'MaxBytes are the maximum, to take out of the buffer
Private Sub InputToNmeaBuf(ByRef InputBuffer As String, Optional InputName As String)  'Must be byref
Dim i As Long
Dim j As Long
Dim SentenceRcv As String  'current sentence were decoding
Dim NmeaBufOverflow As Boolean
Dim InCount As Long     'Sentences were putting into NmeaBuf for debugging
Dim OutCount As Long
Dim arry() As String
Dim cpos As Long
Dim Added As Long
Dim Reject As Long
Dim MaxThisCycle As Long

'    If InputBuffersToNmeaBufCalls > 0 Then Stop  'v131Debug
'    InputBuffersToNmeaBufCalls = InputBuffersToNmeaBufCalls + 1
'If InputBuffersToNmeaBufCalls > 1 Then
'Debug.Print "calls = " & InputBuffersToNmeaBufCalls
'End If
    
    If InputName <> "File" Then
        MaxThisCycle = MaxRcvPerTimerCycle - RcvThisTimerCycle
    Else
        MaxThisCycle = MaxFilePerTimerCycle - FileThisTimerCycle
    End If
'Ensure we have enough room maxnmeabuf - NmeaBufUsedfor the sentences we want to add this cycle
    If NMEABUF_MAX - NmeaBufUsed Or MaxThisCycle = 0 < MaxThisCycle Then
        MaxThisCycle = NMEABUF_MAX - NmeaBufUsed
    End If
'If LFCount = 0 and there is data in the buffer
'then all data is left in buffer because 0 > 0 is false
    InCount = LFCount(InputBuffer)
    If InCount > MaxThisCycle Then   'returns 0 if
'        If InputName <> "File" Then
'            cpos = CRLFPos(InputBuffer, MaxThisCycle)
'Position of start of first sentence after the Max no we can process this cycle
'            Reject = LFCount(Mid$(InputBuffer, cpos + 2))    'Latest
'Debug.Print "Reject " & Reject
'Debug.Print Len(InputBuffer)
'            Rejected = Rejected + Reject    'for stats
'            InputBuffer = Left$(InputBuffer, cpos + 1)    'Earliest
'        End If
    End If
'when j=cpos we must stop outputting as this is when the maxThiscycle has been reached
    If InputBuffer = "" Then
        cpos = 0
    Else
        cpos = CRLFPos(InputBuffer, MaxThisCycle) + 2 'Start of next crlf
    End If                           'crlfpos  returns -1 if none found

'If we cant get any more data in NmeaBuf exit (will remain in FileInput)
'Dont move before above line, could be repeated call when NmeaBufXoff is true
    If NmeaBufXoff = True Then
'        Call ProcessNmeaBuf
'        InputBuffersToNmeaBufCalls = InputBuffersToNmeaBufCalls - 1
        Exit Sub
    End If
    
    Added = 0
    i = 1
    j = InStr(i, InputBuffer, vbCrLf)
'Debug.Print "InputBuffer " & LFCount(InputBuffer)
    Do
        If j Then
            If j > i Then   'If blank j=0 therefore it is ignored
'must be at least 1 spare slot in the RcvList
'When Buffer is over 80% full process buffer until 50% full
'                If NmeaBufUsed * 1.2 > NMEABUF_MAX Then
                If NmeaBufUsed = NMEABUF_MAX Then
                    XoffNmeaBuf
                    Exit Do
                Else
                    If TreeFilter.Option1(1).Value <> 0 Then
'                        Call WriteNmeaLog(Mid$(InputBuffer, i, j - i) & "," & NowUtc())
                        Call WriteInputLogFile(Mid$(InputBuffer, i, j - i) & "," & NowUtc())
                    End If
'Sentence over 200 bytes (with CRLF)
                    If j - i > 200 Then
                        StatusBar.Panels(1).Text = "InputToNmeaBuf - Sentence rejected over 200 bytes " & i - j & " bytes received"
                        ClearStatusBarTimer.Enabled = True
                    Else
                        Received = Received + MESSAGE_COUNT_STEP
'use to debug overflow                        incr Received
                        
'You must only incr if we are using the timer, otherwise you'll get an overflow
'because RcvThisTimerCycle will never get reset
                        If InputName <> "File" Then
                            If MaxRcvPerTimerCycle > 0 Then
                                RcvThisTimerCycle = RcvThisTimerCycle + MESSAGE_COUNT_STEP
                                 incr RcvThisTimerCycle
                            End If
                        Else
                            If MaxFilePerTimerCycle > 0 Then
                                FileThisTimerCycle = FileThisTimerCycle + MESSAGE_COUNT_STEP
                                incr FileThisTimerCycle
                            End If
                        End If
'CrLf is not put into NmeaBuf
                        SentenceRcv = Mid$(InputBuffer, i, j - i) & "," & NowUtc()
'v3.2.135 Replace any nuls as it will terminate the string
                        SentenceRcv = Replace(SentenceRcv, Chr(0), "<nul>")
                        NmeaBuf(NmeaBufNxtIn) = SentenceRcv
                        NmeaBufNxtIn = NmeaBufNxtIn + 1
                        NmeaBufUsed = NmeaBufUsed + 1
                        If NmeaBufNxtIn > NMEABUF_MAX Then NmeaBufNxtIn = 0
                        Added = Added + MESSAGE_COUNT_STEP
'use to debug overflow                        incr Added
                        If Option1(16).Value = True Then   'Show NmeaRcv 'v129
                            Call ReceivedData.AddToNmeaRcvList(SentenceRcv)
                        End If  'Sentence OK
                    End If  'check Len of sentence
                End If  'Check Xoff
            End If  'Check blank
        i = j + 2
        End If  'Find next Crlf
    j = InStr(i, InputBuffer, vbCrLf)
    If j = cpos Then    'Last CRLF in buffer
        Exit Do
    End If
    Loop Until j = 0

    If Added > 0 Then   'at least one sentence moved to NmeaRcvBuffer
'Remove Sentence added to NmeaTcvBuffer from Input Buffer
'MsgBox InputBuffer
        InputBuffer = Mid$(InputBuffer, i)
    Else
'Check No CRLF received at all & input buffer longer than 200 bytes
'        If LFCount(InputBuffer) > 200 Then 'v144 this was wrong
        If Len(InputBuffer) > 200 Then
            StatusBar.Panels(1).Text = "InputToNmeaBuf - No <CRLF> " & Len(InputBuffer) & " bytes rejected"
            ClearStatusBarTimer.Enabled = True
'Clear the buffer
            InputBuffer = ""
        End If
    End If
    
'Debug.Print InputName & " add/rej/left " & Added & "/" & Reject & "/" & LFCount(InputBuffer); "+" & Len(InputBuffer) - InStrRev(InputBuffer, vbLf)
    
'Create some space in NmeaBuf - while we've got the processor
'NmeaBufXoff = True 'testing if NmeaBuf full
    If NmeaBufXoff = False Then
        If Processing.Suspended = False Then Call ProcessNmeaBuf
Else
'Stop
    End If

'InputBuffersToNmeaBufCalls = InputBuffersToNmeaBufCalls - 1
    
End Sub


Private Sub FlushInputBuffer(ByRef InputBuffer As String)
    
'If Paused, we cannot Flush the buffer because Pause stops InputBuffersToNmeaBuf
'transferring sentences to NmeaBuf so we have to clear the input buffer
'Anf because it is Paused, we cannot process and sentences in NmeaBuf

    If InputState <> 2 Then
        Do Until LFCount(InputBuffer) = 0
'Debug.Print "#FlushInputBuffer (" & LFCount(InputBuffer) & ")"
            Call InputBuffersToNmeaBuf
            If Processing.Suspended = False Then ProcessNmeaBuf
'keep user informed
            Call UpdateStats
            Me.Refresh
        Loop
    Else
Debug.Print "Clear InputBuffer (" & LFCount(InputBuffer) & ")"
    End If
    InputBuffer = ""    'Clear any data after CRLF
End Sub

Private Sub ClientUDP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "Socket Error " & Number & ": " & Description       ' show some "debug" info
    ClientUDP.Close ' close the erraneous connection
End Sub

Private Sub ClientTCP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ClientTCPError = Number
    ClientTCPErrorDescription = Description
    Select Case Number
    Case Is = sckConnectionRefused
'        Call frmDpyBox.DpyBox("TCP Client connection refused" & vbCrLf & "Reconfiguring as TCP Server" & vbCrLf, 10, "TCP Client Error")
'        ClientTCP.Close
'        ClientTCP.LocalPort = ClientTCP.RemotePort
'        ClientTCP.Listen
'        Check1(5).Caption = "TCP Server"
    Case Else
        frmDpyBox.DpyBox "Socket Error " & Number & ": " & Description       ' show some "debug" info
        ClientTCP.Close ' close the erraneous connection
        Check1(5).Caption = "TCP"
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Call WriteStartUpLog(Me.Name & ".Form_QueryUnload")
    If QueryQuit(UnloadMode) = vbOK Then  'User click OK
        If Not (InputLogFile Is Nothing) Then InputLogFile.CloseFile
        If InputState <> 0 Then
            cbStop.Value = True
        End If
        Call CheckExpiry
'prevent trying to save the position of the dpybox
'as it would prompt to save settings
        If frmDpyBox.Visible Then
            Unload frmDpyBox
        End If
        NmeaRcv.Hide
        frmSplash.lblAction = ""
'        If EditOptions = True Then
'Only save and ask changes if running in interactive mode
        If EditOptions = True And UnloadMode = vbFormControlMenu Then 'v140
            Call WriteStartUpLog("Existing " & IniFileName)
            Call FileToLog(IniFileName)
            Call TreeFilter.ReplaceIniFile
            Call WriteStartUpLog("New " & IniFileName)
            Call FileToLog(IniFileName)
'            Call EncryptFiles(Environ("ProgramFiles") & "\Arundale\Ais Decoder\", ".cfg_txt", ".cfg")
        End If
        If CacheVessels Then
            frmSplash.lblAction = "Saving Vessels"
            frmSplash.AutoRedraw = True
            frmSplash.Show
            frmSplash.Refresh
            If Vessels.Count Then
'Don't let timer from previous message clear this message, leave on the screen
                ClearStatusBarTimer.Enabled = False
                StatusBar.Panels(1).Text = "Writing " & Vessels.Count & " vessels to disk"
            End If
            Call SaveVessels
        End If
        frmSplash.lblAction = "Terminating"
        frmSplash.Show
        frmSplash.Refresh
        If CacheTrappedMsgs Then Call SaveTrappedMsgs
        Call Terminate
'V3.4.143 when here all forms except NmeaRcv should be closed so when end Sub is reached NmeaRcv will
'terminate and no task image will be left running.
'V142 MUST NOT call END to exit program otherwise
'V142 Exit in Batch or Task manager cannot force program exit
        
'moved here from end of terminate because gcoDebug .Count had not  ? had time ? to be set to 0
        Call LogDebugCls
    Else        'V142
        Cancel = True   'V142
    End If

End Sub


'Not resizable
Private Sub Form_Resize()
Dim sngWidth As Single, sngHeight As Single
Dim sngDisplayHeight As Single
Dim sngTxtWidth As Single
Dim sngCmdWidth As Single, sngCmdHeight As Single
'calculate the inner size of the form
sngWidth = ScaleWidth
sngHeight = ScaleHeight - StatusBar.Height
'If sngWidth <> 0 Then StatusBar.Panels(1).Width = sngWidth
End Sub

'Private Sub Form_Terminate()
'Call Terminate
'End Sub


'this is microsofts way of starting the default browser
Private Sub Help_Click()
Call HttpSpawn(App.Path & "\Help\AisDecoder.chm")
'Dim r As Long
'Dim Command As String

'If Environ("windir") <> "" Then
'    r = ShellExecute(0, "open", "http://web.arundale.co.uk/docs/ais/ais_decoder_v3.html", 0, 0, 1)
'Else
'try for linux compatibility
'    Command = "winebrowser http://web.arundale.co.uk/docs/ais/ais_decoder_v3.html ""%1"""

'    Shell (Command)
'End If

End Sub

#If False Then  'Leave till Scheduler is sorted
Public Sub RcvStart()

    Call TreeFilter.EnableDisableOptions
        
    StatsTimer.Enabled = True

'ScheduledTimer not implemented
    If ScheduledReq(0) = True Then
        ScheduledTimer.Enabled = True
    End If
    
    Call StartInputOutput
    
End Sub
#End If

'Opens required inputs, if not open, and starts getting data

Private Sub StartInputOutput()
Dim Elapsed As Double
Dim WaitTime As Long    'millisecs
Dim Channel As Long

'Debug.Print String(255, vbNewLine)
'Debug.Print "#StartInputOutput"
    strDebug = "#StartInputOutput"
    Set RunClock = New clsTimer
    Set RcvClock = New clsTimer
    Set CycleClock = New clsTimer
    Set RcvWaitableTimer = New clsWaitableTimer
    Set FileWaitableTimer = New clsWaitableTimer
    Call ClearStats
    Call ClearMyShip(MyShip)
    
    mblnCancel = False
    InputFileCancel = False
'Set up Timer interval (cant do until form loaded)
    If UserLicence.MaxRcvPerMin <= 0 Then
        MaxRcvPerTimerCycle = 0 '100000    '100k
    Else
'Watch overflow divide before multiply (1m/60k*6k)=100k
        MaxRcvPerTimerCycle = (UserLicence.MaxRcvPerMin / MILLISECS_PER_MIN) * SPEED_INTERVAL
    End If
    
'    If UserLicence.MaxFilePerMin > 0 Then
    If UserLicence.MaxFilePerMin <= 0 Then
        MaxFilePerTimerCycle = 0    '100000    '100k
    Else
'Watch overflow divide before multiply (1m/60k*6k)=100k
        MaxFilePerTimerCycle = (UserLicence.MaxFilePerMin / MILLISECS_PER_MIN) * SPEED_INTERVAL
    End If


'Initially we can process this number until the time fires
    RcvThisTimerCycle = 0
    FileThisTimerCycle = 0
    Call TreeFilter.EnableDisableOptions
                
'Re-read, it may have been changed
    If TaggedOutput(0) = True Then
        Call ReadTagTemplate(TagTemplateReadFile)
    End If
'UDP
    If Check1(0).Value = vbChecked Then  'open UDP input if closed
        If ClientUDP.State = sckClosed Then
            Call OpenClientUdp
        End If
    End If

'TCP
    If Check1(5).Value = vbChecked Then  'open TCP input if closed
        Call OpenClientTCP
    End If

'Serial
    If Check1(1).Value = vbChecked Then 'open serial input if closed
        If SerialPortCount > 0 Then
#If UseComm Then
            Call OpenSerialInput
#Else
Debug.Print "SerialPortOpen = " & MSComm1.PortOpen
            If MSComm1.PortOpen = False Then
                Call OpenSerialInput
Debug.Print "SerialPortOpen = " & MSComm1.PortOpen
            End If
#End If
        Else
Debug.Print "No Serial Ports Found"
        End If
    End If
            
'Start user stats display
    StatsTimer.Enabled = True

'ScheduledTimer not implemented
    If ScheduledReq(0) = True Then
        ScheduledTimer.Enabled = True
    End If
    
'File
    If Check1(6).Value = vbChecked Then
        Call OpenFileInput  'Used a Timer to start it so that rest of cbStart code is executed
    End If      'File Input
'InputLogFile
    If TreeFilter.Option1(0) = False Then   'An Input log is required
        Call OpenInputLogFile
    End If

    RunClock.StartTimer
    If InputsCount("Rcv") > 0 Then RcvClock.StartTimer

Exit Sub

    CycleClock.StartTimer

'This now loops until the Stop is pressed
    Do Until InputState = 0
'WaitTime=0 before first cycle
        If WaitTime = 0 Or CycleClock.Duration * 1000 >= CDbl(SPEED_INTERVAL) Then
            CycleClock.StopTimer
    'ProcessCycle
            WaitTime = SPEED_INTERVAL - CycleClock.Duration * 1000
'Duration will be longer than Interval when File is max speed
            If WaitTime <= 0 Then WaitTime = 50 'if 0 waits indefinately
            CycleClock.ResetTimer
            CycleClock.StartTimer
            Call ResetCycle
            If Not (InputFile Is Nothing) Then
'called once per timer cycle - returns after MaxPerTimerCycle sentences are sent
                Call InputFile.LongTask(MaxFilePerTimerCycle, SPEED_INTERVAL)
            End If
'Here when MaxPerTimerCycle sentences are sent (by PercentDone Event)
        End If
'Debug.Print "Wait " & WaitTime
        If Not (RcvWaitableTimer Is Nothing) Then   'Possibly closed by StopInputOutput
            RcvWaitableTimer.Wait WaitTime
        End If
        If InputsCount = 0 Then Exit Do
    Loop
    
End Sub
 

Private Function OpenFileInput() As String
Dim Cancel As Boolean   'Local to OpenFileInput
Dim i As Long
Dim cFileLen As Currency
Dim kb As String

'When first called by Click event starts the timer to allow click event procedure to complete
'    If OpenFileInputTimer.Enabled = False Then
'        OpenFileInputTimer.Enabled = True
'        Exit Sub
'    End If
'Debug.Print "#OpenFileInput"
'Clear last message (if still there)
    If Left$(StatusBar.Panels(1).Text, 22) = "File Input Completed. " Then
        StatusBar.Panels(1).Text = ""
    End If
    Set MyInputThrottleTimer(6) = New clsWaitableTimer

    If cmdStart = False Or NmeaReadFile = "" Then
'If nmeareadfile does not exists AskFileName fails with err=5 invalid call
        If FileExists(NmeaReadFile) = False Then
            NmeaReadFile = ""
        End If
        NmeaReadFile = FileSelect.AskFileName("NmeaReadFile", False, Cancel)
    End If
'Ensure Command Button changes are displayed before any processing commences
'And the File Dialog Box is not partially displayed within NmeaRcv window
'Due to reading the Input File grabbing all the processor time
    Me.Refresh
    
'MsgBox NmeaReadFile
'Need to trap if weve tried to open the currently open nmealogfile for input
    If Cancel = False Then
        If Not (InputLogFile Is Nothing) Then
'MsgBox "Input=" & NmeaReadFile & vbCrLf & "Log=" & NmeaLogFile, , "OpenFileInput"
            If NmeaReadFile = InputLogFile.FileName Then
                Cancel = True
                MsgBox NmeaReadFile & vbCrLf & vbCrLf & "Permission denied - The Input File cannot be the same as the Input Log File", vbExclamation, "Open Input File"
            End If
        End If
    End If
    
    If Cancel = False Then
        If FileExists(NmeaReadFile) = False Then
            Cancel = True
            MsgBox NmeaReadFile & vbCrLf & vbCrLf & "File not found", vbExclamation, "Open Input File"
        End If
    End If
    
    If Cancel = False Then
        Set InputFile = New HugeBinaryFile
'Open the file first
        If UserLicence.FileBlockSize = 0 Then
            UserLicence.FileBlockSize = DEFAULT_BLOCK_SIZE
        End If
    
        InputFile.OpenFile NmeaReadFile, , UserLicence.FileBlockSize
        cFileLen = CDec(InputFile.FileLen) / 10000@
'        cFileLen = StringToCurrency(InputFile.FileLen)
        If CDec(UserLicence.sMaxInputFileSize) <> CDec(0) And cFileLen > CDec(UserLicence.sMaxInputFileSize) / 10000@ Then
            Cancel = True
            InputFile.CloseFile
            Set InputFile = Nothing
            kb = "Input File Size is " & aByte(cFileLen * 10000@) & vbCrLf
            kb = kb & "To read a file over " & aByte(UserLicence.sMaxInputFileSize) & " a program upgrade is required" & vbCrLf
            kb = kb & "Please email myself neal@arundale.com for futher information" & vbCrLf
            MsgBox NmeaReadFile & vbCrLf & vbCrLf & kb, vbExclamation, "Open Input File"
'        Call CloseFileInput
'            InputFile.EOF = True
        End If
    End If
        
    If Cancel = False Then
 'We must kick off with a timer event otherwise it remains in the Click event
 'User needs to cbStop & cbStart
'9/10/14 now using RcvWaitableTimer
        FileNextBlockTimer.Enabled = True
    End If
'Note Start is not re-enabled if cancelled or file too big
'File Name mey be changed on open
If Files.Visible Then Call Files.RefreshData
End Function

'Called by StartInputOutput
Private Function OpenInputLogFile() As String
Dim cFileLen As Currency
Dim kb As String

'Stop
'Debug.Print "#OpenInputLogFile"
            
    On Error GoTo error_open
    NmeaLogFile = FileSelect.SetFileName("NmeaLogFile")
'MsgBox "Log=" & NmeaLogFile & vbCrLf & "Read=" & NmeaReadFile

'if were reading from a log file then add .log to the end of the nmea file name
    If Not (InputFile Is Nothing) Then
        NmeaLogFile = PathFromFullName(NmeaLogFile) & "\" & NameFromFullPath(InputFile.FileName)
'MsgBox "Log=" & NmeaLogFile & vbCrLf & "Read=" & NmeaReadFile
        If InputFile.FileName = NmeaLogFile Then
            NmeaLogFile = InputFile.FileName & ".log"
        End If
    End If
'MsgBox "Log=" & NmeaLogFile & vbCrLf & "Read=" & NmeaReadFile
        
    Set InputLogFile = New HugeBinaryFile
    Call InputLogFile.OpenFileAppend(NmeaLogFile)
'too slow    InputLogFile.AutoFlush = True   'File must be open first
    cFileLen = CDec(InputLogFile.FileLen) / 10000@
'Debug.Print "OpenInputLogFile " & NameFromFullPath(NmeaLogFile)
'File Name mey be changed on open
    If Files.Visible Then Call Files.RefreshData
Exit Function

error_open:
    MsgBox err.Description, vbExclamation, "Open Input LogFile"
End Function

Public Sub WriteInputLogFile(kb As String)
            
    On Error GoTo Error_WriteInputLogFile
'Need to check file is open, because it will have been closed
'if its too big
    If Not (InputLogFile Is Nothing) Then
        If TreeFilter.Option1(11).Value = True Then
            InputLogFile.WriteString kb & vbCrLf
        Else
            If IsDate(Mid$(kb, InStrRev(kb, ",") + 1)) Then
                InputLogFile.WriteString Left$(kb, InStrRev(kb, ",") - 1) & vbCrLf
            Else
                InputLogFile.WriteString kb & vbCrLf
            End If
        End If
    End If
Exit Sub

Error_WriteInputLogFile:
'Now reports when we try to open it
'    NmeaRcv.StatusBar.Panels(1).Text = "WriteInputLogFile, Error " & err.Number & " - " & err.Description
'    NmeaRcv.ClearStatusBarTimer.Enabled = True
End Sub

Public Sub OpenClientUdp()
Dim kb As String
'Debug.Print "#OpenClientUDP"
    kb = "OpenClientUdp Port " & ClientUDP.LocalPort
    WriteLog kb, LogForm
'MyInputThrottleTimer(0) is not actually used Apr17
'    Set MyInputThrottleTimer(0) = New clsWaitableTimer
    On Error GoTo ClientUDP_Bind_err
        ClientUDP.Bind CLng(TreeFilter.Text1(3).Text)
    On Error GoTo 0
'Feb 2017 v145 It seems we may have to call DataArrival manually once to trigger receiving data immediately
'    Call ClientUDP_DataArrival(0)
    kb = kb & ", State=" & ClientUDP.State & " (" & aState(ClientUDP.State) & ")"
    WriteLog kb, LogForm
Exit Sub

ClientUDP_Bind_err:
    MsgBox "Can't open port " & TreeFilter.Text1(3).Text        ' & vbCrLf & "Working off line"
    err.Clear
End Sub

Public Sub OpenClientTCP()
'If in Server mode (Listening) then
'Close and re-open as in Client Mode
    Set MyInputThrottleTimer(5) = New clsWaitableTimer
    If ClientTCP.State <> sckClosed Then
        ClientTCP.Close
    End If
            
    If ClientTCP.State = sckClosed Then
'Debug.Print "#OpenClientTCP"
        On Error GoTo ClientTCP_Connect_err
        ClientTCP.RemoteHost = TreeFilter.Text1(12).Text
        If IsNumeric(TreeFilter.Text1(13).Text) Then
            ClientTCP.RemotePort = CLng(TreeFilter.Text1(13).Text)
        End If

'Must set to 0 to ensure it is retried as a Client
        ClientTCP.LocalPort = 0
        ClientTCP.Connect
        TcpTimer.Interval = 5000    'Give it time to connect before trying as server
'Restart timer if not running
        If TcpTimer.Enabled = False Then
            TcpTimer.Enabled = True
        End If
        Check1(5).Caption = "Client"
    End If
    Exit Sub
ClientTCP_Connect_err:
    Select Case err.Number
    Case Else
        Call frmDpyBox.DpyBox(err.Number & " " & err.Description & TreeFilter.Text1(12).Text _
        & ":" & TreeFilter.Text1(13).Text & vbCrLf & "Working off line", 5, "TCP Open Client")
    End Select
End Sub

Private Sub OpenServerTCP()
'MsgBox TcpTimer.Enabled
'Debug.Print "#OpenServerTCP"
    Set MyInputThrottleTimer(5) = New clsWaitableTimer
    If ClientTCP.State <> sckClosed Then
        ClientTCP.Close
    End If
    
    On Error GoTo ServerTCP_Connect_err
    If IsNumeric(TreeFilter.Text1(13).Text) Then
        ClientTCP.LocalPort = CLng(TreeFilter.Text1(13).Text)
    End If
    ClientTCP.Listen
    Check1(5).Caption = "Servr"
        TcpTimer.Interval = 1000
'Restart timer if not running
        If TcpTimer.Enabled = False Then
            TcpTimer.Enabled = True
        End If
    Exit Sub

ServerTCP_Connect_err:
    Select Case err.Number
'Client is not finished connecting
    Case Is = sckInvalidOp
    Case Else
        Call frmDpyBox.DpyBox(err.Number & " " & err.Description & TreeFilter.Text1(12).Text _
        & ":" & TreeFilter.Text1(13).Text & vbCrLf, 5, "TCP Server")
    End Select
End Sub

'Only Called by OutputUDP
Public Sub OpenOutputUdp(Channel As Long)
    
'Debug.Print "#OpenOutputUDP(" & Channel & ")"
    
    With ServerUDP
'        .RemoteHost = "81.137.214.195" 'jna home
'not opening port every other time bug
        If LastRemoteHost <> TreeFilter.Text1(5).Text Then
            LastRemoteHost = TreeFilter.Text1(5).Text
            .RemoteHost = LastRemoteHost
        End If
        If LastRemotePort <> TreeFilter.Text1(4).Text Then
            LastRemotePort = TreeFilter.Text1(4).Text
            .RemotePort = LastRemotePort
        End If
'MsgBox .RemoteHost & ":" & .RemotePort & "state is " & .State
        If ServerUDP.State <> sckClosed Then GoTo ServerUDP_Bind_err
        On Error GoTo ServerUDP_Bind_err
        .Bind 0          ' Bind to a random local port.
        On Error GoTo 0
        TreeFilter.Text1(4).Enabled = False
        TreeFilter.Label1(4).Enabled = False
        TreeFilter.Text1(5).Enabled = False
        TreeFilter.Label1(5).Enabled = False
    End With
    
    NoDataOutput(Channel) = True
    
    If DisplayOutput(Channel) Then
        Call Output.OutputDisplay(Channel, "[Open UDP]")
    End If

Exit Sub
ServerUDP_Bind_err:
    MsgBox "Can't open port " & TreeFilter.Text1(5).Text _
    & ":" & TreeFilter.Text1(4).Text    ' & vbCrLf & "Working off line"
'    cbStart.Caption = "Off Line"

End Sub

'This is only opened by OutputToDataFile
Sub OpenOutputDataFile(Channel As Long)
Dim kb As String
Dim UnZippedFilename As String

'Debug.Print "#OpenOutputDataFile(" & Channel & ")"
'Stop
'May have been closed and set to Nothing when rolled Over
    If Not OutputDataFile Is Nothing Then
        Set OutputDataFile = Nothing
    End If

'Dont write out No Data
'check output file name (if rollover name will have changed)
    OutputFileName = FileSelect.SetFileName("OutputFileName")
'MsgBox OutputFileName
    On Error GoTo FileLocked    'user is reading output file
    If DisplayOutput(Channel) Then    'channel 1 is file output
        Call Output.OutputDisplay(Channel, "[Open file:" & OutputFileName & "]")
    End If

    
    Set OutputDataFile = New HugeBinaryFile
'If we want a Zipped up Output file, we still create it a KML first
    Select Case LCase(ExtFromFullName(OutputFileName))
    Case Is = "kmz", "zip"
        ZipOutputFile = True
    Case Else
        ZipOutputFile = False
    End Select
'We must zip up if we include an overlay file
    If OverlayReq(Channel) = True Then ZipOutputFile = True
'Stop
    
    If ZipOutputFile = True Then
        Select Case LCase(ExtFromFullName(OutputFileName))
        Case Is = "kmz"
            UnZippedFilename = Replace(LCase(OutputFileName), ".kmz", ".kml")
        Case Is = "zip"
            UnZippedFilename = Replace(LCase(OutputFileName), ".zip", ".uzip")
        End Select

        If FileExists(UnZippedFilename) Then
            Kill UnZippedFilename
        End If
        OutputDataFile.OpenFileAppend (UnZippedFilename)
    Else
        If OutputFileRollover(Channel) = False Then
            If FileExists(OutputFileName) Then
                Kill OutputFileName
            End If
        End If
        OutputDataFile.OpenFileAppend (OutputFileName)
    End If
    NoDataOutput(Channel) = True
    On Error GoTo 0
            
    If OverlayReq(Channel) = True Then
'Now only dependant og Template File name
'    OverlayOutputFileName = FileSelect.SetFileName("OverlayOutputFileName")
        OverlayOutputFileCh = FreeFile
 'Debug.Print "Open " & OverlayOutputFileCh
        On Error GoTo OverlayFileLocked    'user is reading output file
        Open OverlayOutputFileName For Output As #OverlayOutputFileCh
'MsgBox "Open Overlay Output" & vbCrLf & OverlayOutputFileName

        Print #OverlayOutputFileCh, "<?xml version=""1.0"" encoding=""UTF-8""?>"
        Print #OverlayOutputFileCh, "<kml xmlns=""http://earth.google.com/kml/2.0"">"
        Print #OverlayOutputFileCh, "<Document>"
    End If
Exit Sub

FileLocked:
    OutputFileErr = True
    MsgBox OutputFileName & vbCrLf & "In use by another user" & vbCrLf, vbExclamation, "Open Output File"
    Set OutputDataFile = Nothing
Exit Sub

OverlayFileLocked:
    OverlayOutputFileCh = 0
Exit Sub

End Sub

'Close open inputs and clear buffers (input and Nmea)
Sub StopInputOutput()
Dim Channel As Long

'Debug.Print "#StopInputOutput"
    
#If UseComm Then
    PollTimer.Enabled = False   'If not closed Comms(idx) does not get unloaded
                                'causing Serial pollong no to restart
#End If

    mblnCancel = True
    InputFileCancel = True
'Check if it should be open - so we can debug easily the individual close functions
'UDP
    Select Case Check1(0).Value
    Case Is = vbChecked
        Call CloseUDPInput
    Case Is = vbGrayed
        Check1(0).Value = vbChecked
    End Select
'TCP
    Select Case Check1(5).Value
    Case Is = vbChecked
        Call CloseTCPInput
    Case Is = vbGrayed
        Check1(5).Value = vbChecked
    End Select
'Serial
    Select Case Check1(1).Value
    Case Is = vbChecked
        Call CloseSerialInput
    Case Is = vbGrayed
        Check1(1).Value = vbChecked
    End Select
'File
    Select Case Check1(6).Value
    Case Is <> vbUnchecked
'Close big input file if Stop Pressed (If EOF file will already be closed)
'InputFile.Pause may also be set true
        If Not (InputFile Is Nothing) Then
            Call CloseFileInput
        End If
   End Select
    
'Only Clear the buffers when cbStop otherwise cbContinue will not process the buffer
'if it was full
    UDPRcvBuffer = ""
    TCPRcvBuffer = ""
    SerialRcvBuffer = ""
    FileRcvBuffer = ""
    NmeaBufUsed = 0
    Call UpdateStats

'Close Outputs
'   CloseOutput(Channel)
'Channels are Opened as required
'When Scheduled Output Runs CLoseOutput for any Scheduled Channel
'They will get re-opened when any data wants writing out
'They should also be closed when file is rolled over
    For Channel = 1 To 2
        If ChannelOutput(Channel) = True Then   'some output is selected
'TaggedOutput(to write out the Tail) wants doing in Both CloseOutput subroutines
'            If TaggedOutput(Channel) = True Then CloseTaggedOutput (Channel) 'write out tagged tail
            If ChannelMethod(Channel) = "file" Then
                Call CloseOutputDataFile(Channel)    'write out No Data if not opened
'Dont write out the Tail if not opened either
            End If
            If ChannelMethod(Channel) = "udp" Then
                Call CloseOutputUdp(Channel)
            End If
        End If
    Next Channel
    

'Re-enable disabled options (Input state will be stopped)
    Call TreeFilter.EnableDisableOptions

    If Not (InputLogFile Is Nothing) Then
        On Error Resume Next
        Call InputLogFile.CloseFile     'File may have been locked
        On Error GoTo 0
        Set InputLogFile = Nothing
    End If
                           
'Clear Multipart buffer so that new payloads are not mixed with
'old payloads
    ReDim Multipart(9, 9)
'Clear tag values
    Call ClearTagValues
    
    StatsTimer.Enabled = False
    Set CycleClock = Nothing
    If Not (RcvClock Is Nothing) Then
'Should not happen
        TotalRcvTime = TotalRcvTime + RcvClock.Duration
    End If
    TotalRejected = TotalRejected + Rejected
    Set RcvClock = Nothing
    Set RunClock = Nothing
    Set RcvWaitableTimer = Nothing
    Set FileWaitableTimer = Nothing
    OutputFileErr = False

    Label3.Caption = "" 'RcvSpeed
'Scheduledtimer not implemented
    ScheduledTimer.Enabled = False
        
    NmeaBufNxtIn = 0
    NmeaBufNxtOut = 0
    NmeaBufUsed = 0
    NmeaBufXoff = False
    With Processing
        .Paused = False
        .NmeaBuf = False
        .Scheduler = False
        .InputOptions = False
        .Suspended = False
    End With
    Call UpdateStats    'Remove any Processing message in status bar
End Sub

Private Sub CloseTCPInput()
'ensure timer doesn't try & re-open, will if its failed to open before Closing Handler
'Debug.Print "#CloseTCPInput"
    Set MyInputThrottleTimer(5) = Nothing
    TcpTimer.Enabled = False
    If ClientTCP.State = sckConnected Then
        ClientTCP.Close  'udp input
    End If
    Check1(5).Caption = "TCP"   'Would be TCPClient or TCPServer
'Dont clear the buffer, will require processing after cbContinue (after buffer full)
'    TCPRcvBuffer = ""
    TcpXoff = False
End Sub

Private Sub CloseUDPInput()
'Debug.Print "#CloseUDPInput"
'MyInputThrottleTimer(0) is not actually used Apr17
'    Set MyInputThrottleTimer(0) = Nothing
    If ClientUDP.State = sckOpen Then
        ClientUDP.Close  'udp input
    End If
'Dont clear the buffer, will require processing after cbContinue (after buffer full)
'        UDPRcvBuffer = ""
End Sub

Private Sub CloseSerialInput()


'Debug.Print "#CloseSerialInput"
    Set MyInputThrottleTimer(1) = Nothing
    OpenSerialTimer.Enabled = False
    OpenSerialTimer.Interval = 0  'millisecs

#If UseComm Then
    Call CloseHandler(1)
    CurrentSocket = 0
#Else
    With MSComm1
        If .PortOpen = True Then
 'Debug.Print "Close " & MSComm1.Settings
            .PortOpen = False
        End If
    End With
#End If

'Dont clear the buffer, will require processing after cbContinue (after buffer full)
    SerialRcvBuffer = ""    'Removed v131 (I think we must clear the buffer
                            'if resetting the speed
    CrLfDetected = False
End Sub

Private Sub CloseFileInput()
'Debug.Print "#CloseFileInput"
        
    Set MyInputThrottleTimer(6) = Nothing
    FileNextBlockTimer.Enabled = False
    If Not (InputFile Is Nothing) Then
        With InputFile
            .LongTaskCancel = True
'Don't let timer from previous message clear this message, leave on the screen
'           .Xoff   'stop output
            ClearStatusBarTimer.Enabled = False
'NOTE cleared by OpenFileInput if text is "File Input Completed. "
            StatusBar.Panels(1).Text = "File Input Completed. " & .TotalLinesRead & " Sentences (" _
            & aByte(.TotalBytesRead) & ") in " _
            & Format$(.aRunTime, "HH:NN:SS") & " "
            If .RunSeconds > 0 Then
                StatusBar.Panels(1).Text = StatusBar.Panels(1).Text & Format$(.TotalLinesRead * 60 / .RunSeconds, "###,###\/min")
            End If
Debug.Print StatusBar.Panels(1).Text
            .CloseFile
        End With
        Set InputFile = Nothing
    Else    'Already closed
'Stop    'debugv131
    End If
    
'Only flush if its the last input, because NmeaRcvXoff could be True and it will never flush
'Until the other input(s) have released Xoff
    If InputsCount = 1 Then 'FileInput is the only input
        Call FlushInputBuffer(FileRcvBuffer)    'If you wish to clear the buffer before stopping
    End If
'Dont clear the buffer, will require processing after cbContinue (after buffer full)
        FileRcvBuffer = ""          'To Clear buffer first
    If Check1(6).Value = vbGrayed Then
        Check1(6).Value = vbChecked
    End If
End Sub

Public Sub CloseOutputUdp(Channel As Long)

'Debug.Print "#CloseOutputUDP(" & Channel & ")"
        
    
'    If NmeaRcv.ServerUDP.State <> sckOpen Then
    If NoDataOutput(Channel) = True Then
        Call OutputUdp(Channel, "No Data " & NowUtc & " UTC")
    Else
        If TaggedOutputOn(Channel) = True Then
            Call WriteTagTail(Channel)  'Closes OutputOverlayFile
        End If
    End If
    
    If DisplayOutput(Channel) Then
        Call Output.OutputDisplay(Channel, "[Close UDP]")
    End If
    
    NmeaRcv.ServerUDP.Close ' close the erraneous connection
        
'Debug.Print "#OutputUDPClosed(" & Channel & ")"
End Sub

Public Sub CloseOutputDataFile(Channel As Long)   'only channel 1 at the moment
Dim sError      As String
Dim cmdLine As String
Dim ch As Long

'Debug.Print "#CloseOutputDataFile(" & Channel & ")"
'This writes No Data 29/08/2013 09:44:47 UTC
'If there has been no data written out when we attempt to close the file
'The File will not have been opened
'    If OutputDataFile Is Nothing Then
    If NoDataOutput(Channel) = True Then
'This will open the file
        Call OutputToDataFile(Channel, "No Data " & NowUtc & " UTC")
    Else
        If TaggedOutputOn(Channel) = True Then
            Call WriteTagTail(Channel)  'Closes OverlayOutput file (if open)
        End If
    End If
    
    If Not (OutputDataFile Is Nothing) Then
        Call OutputDataFile.CloseFile
        Set OutputDataFile = Nothing
    
        If DisplayOutput(Channel) Then
            Call Output.OutputDisplay(Channel, "[Close file:" & OutputFileName & "]")
        End If
    End If
    
'if not stopped
    If InputState <> 0 Then
'convert to kmz if required
        If ZipOutputFile = True Then
'Create Zip (will include overlay file & pics if file name is not nul
            Call MyZip.ZipUp(OutputFileName)
       End If

'MovedBack to SchedulerOutput
'If were going to FTP the Output file, it is copied to
'OutputFile + .ftp when it is closed. This allows the file to be read
'whilst a new output file is being created.
'    If FtpUploadExecuting = False Then
'            If FtpOutput(Channel) = True Then
'Error if OutputFile has not been created
'                On Error Resume Next
'                FileCopy OutputFileName, OutputFileName & ".ftp"
'MsgBox OutputFileName
'MsgBox "FtpUploadExecuting=" & FtpUploadExecuting
'                Call frmFTP.FtpUpload   'move to here from above
 'Debug.Print Time$() & " End FTP"
'                On Error GoTo 0
'            End If
'    End If
'Command prompt created in Same directory as ShellFileName
'/C closes the Console when finished /K doesnt
'Shell is Synchronous so files must be closed & input suspended
        If ShellOn(Channel) = True Then
'Debug.Print "#Shell(" & NameFromFullPath(ShellFileName) & ")"
            If FileExists(ShellFileName) Then
                ch = FreeFile
                Open ShellFileName For Input As #ch
'File may be zero length
                Do Until EOF(ch)
                    Line Input #ch, cmdLine
                    If cmdLine <> "" Then ExecCmd (cmdLine)
                Loop
                Close ch
            End If
        End If
    End If
'Debug.Print "#OutputDataFileClosed(" & Channel & ")"

End Sub

Private Sub Option1_Click(Index As Integer)
'Forms are "Shown" when data is received or output
'Here we only Hide the form if the user clicks "None"
Select Case Index
Case Is = 0, 3  'list None or Scheduled
    If Option1(Index).Value = True Then
        List.Hide
    End If
Case Is = 4, 7 'detail none or Scheduled, Nmea
    If Option1(Index).Value = True Then
        Detail.Hide
    End If
Case Is = 11    'Scheduled Field Source
    If TreeFilter.Text1(0) <= 0 Or TreeFilter.Text1(1) <= 0 Then
        MsgBox "You must set a Scheduled Time and Time to Live"
    End If

Case Is = 15    'Nmea input
    If Option1(Index).Value = True Then
        ReceivedData.Hide
    End If

End Select
End Sub

Sub SetTreeViewBackColor(TV As TreeView, ByVal BackColor As Long)
    Dim lStyle As Long
    Dim TVNode As Node
         
    ' set the BackColor for every node
    For Each TVNode In TV.Nodes
        TVNode.BackColor = BackColor
    Next

    ' set the BackColor for the TreeView's window
    SendMessage TV.hwnd, TVM_SETBKCOLOR, 0, ByVal BackColor
    ' get the current style
    lStyle = GetWindowLong(TV.hwnd, GWL_STYLE)
    ' temporary hide lines
    SetWindowLong TV.hwnd, GWL_STYLE, lStyle And (Not TVS_HASLINES)
    ' redraw lines
    SetWindowLong TV.hwnd, GWL_STYLE, lStyle
End Sub

'Splits the TagTemplate in Head,Content,Tail & establishes the TagChe <> or []
Sub ReadTagTemplate(TemplateFileName As String)   'Template file must exist
Static LastTemplateFileName As String
Dim kb As String
Dim TagLineState As String  'null,HEAD,CONTENT,TAIL
Dim NoOutput As Boolean  'don't output this line
Dim LineNoHead As Long          'reqd to append crlf if multiline
Dim LineNoContent As Long
Dim LineNoTail As Long
Dim TagNo As Long
Dim Channel As Long
Dim i As Long
Dim j As Long
Dim Matches(0 To 1) As Long '0=<>, 1=[]
Dim NetworkLink As Boolean  'see if these kml tags are active
Dim Link As Boolean
Dim href As Boolean
Dim hrefFile As String
Dim FileTags0() As String   '<tags>
Dim FileTags1() As String   '[tags]

            If FileExists(TemplateFileName) = False Then
MsgBox "Template File" & vbCrLf & TemplateFileName & vbCrLf & " not found"
            End If
'If we dont currently have any tags set up, we cant establisg
'the TagChr <> or []. When we set up a tag a new reset will be done
'When the .ini file first loads this will always happen
'but ResetTags at the end of the load will force the template
'to be read

'Exit if we do not have any tags set up (On the Options Window)
            If (Not TagArray) = -1 Then
                Exit Sub
            End If
'We call this routine whenever the TagTemplateReadFile is Set
'but it may not actually have changed, in which case there is no point in re-reading it
'            If TemplateFileName = LastTemplateFileName Then
'                Exit Sub
'            End If
'MsgBox "TemplateFileName Reqd" & vbCrLf & TemplateFileName _
'& vbCrLf & "Last File was" & vbCrLf & LastTemplateFileName
            Call WriteStartUpLog("")
            If LastTemplateFileName <> TemplateFileName Then
                Call WriteStartUpLog("TemplateFileName changed to " & TemplateFileName)
                Call WriteStartUpLog("Last File was " & LastTemplateFileName)
            End If
            LastTemplateFileName = TemplateFileName
'Check if we have a TagTemplate (if reqd)
'must always check - may be reqd on loading ini file
'            TagTemplateHead = "2"  'think this is a mistake
            For Channel = 1 To 2
                TagTemplateHead(Channel) = ""
                TagTemplateContent(Channel) = ""
                TagTemplateTail(Channel) = ""
            Next Channel
'            TagTemplateReadFile = FileSelect.GetFileName("TagTemplateReadFile")
'the TagTemplateReadfile name will already have been set
            Call WriteStartUpLog("Parsing Template File " & TemplateFileName)
            TagTemplateReadFileCh = FreeFile
            On Error GoTo BadFile   'File may heve been deleted
            Open TemplateFileName For Input As #TagTemplateReadFileCh

            TagLineState = ""   'treated as content if no head
            Do Until EOF(TagTemplateReadFileCh)
                Line Input #TagTemplateReadFileCh, kb
'MsgBox kb
                Select Case kb
                Case Is = "<![CDATA[/HEAD]]>", "<![CDATA[/CONTENT]]>" _
                , "<![CDATA[/TAIL]]>"
                    TagLineState = ""
                    NoOutput = True
                Case Is = "<![CDATA[HEAD]]>"
                    TagLineState = "HEAD"
                    NoOutput = True
                Case Is = "<![CDATA[CONTENT]]>"
                    TagLineState = "CONTENT"
                    NoOutput = True
                Case Is = "<![CDATA[TAIL]]>"
                    TagLineState = "TAIL"
                    NoOutput = True
                
                Case Is = "<Placemark>"
'if we see the placemark and we have only content, the content must
'actally be the head - so move content to head
                    If TagLineState = "" Then
                        For Channel = 1 To 2
                            TagTemplateHead(Channel) = TagTemplateContent(Channel)
                            TagTemplateContent(Channel) = ""
                        Next Channel
                    LineNoHead = LineNoContent
                    LineNoContent = 0
                    TagLineState = "CONTENT"
                    End If
                End Select
                If NoOutput = False Then
                    Select Case TagLineState
                    Case Is = "HEAD"
                        For Channel = 1 To 2
                            If LineNoHead = 0 Then
                                TagTemplateHead(Channel) = kb
                            Else
                                TagTemplateHead(Channel) = TagTemplateHead(Channel) & vbCrLf & kb
                            End If
                        Next Channel
                        LineNoHead = LineNoHead + 1
                    Case Is = "CONTENT", ""
                        For Channel = 1 To 2
                            If LineNoContent = 0 Then
                                TagTemplateContent(Channel) = kb
                            Else
                                TagTemplateContent(Channel) = TagTemplateContent(Channel) & vbCrLf & kb
                            End If
                        Next Channel
                        LineNoContent = LineNoContent + 1
                    Case Is = "TAIL"
                        For Channel = 1 To 2
                            If LineNoTail = 0 Then
                                TagTemplateTail(Channel) = kb
                            Else
                                TagTemplateTail(Channel) = TagTemplateTail(Channel) & vbCrLf & kb
                            End If
                        Next Channel
                        LineNoTail = LineNoTail + 1
                    End Select
'check the line to see if weve got any of our tags in it
                    TagChr(0) = "<"
                    TagChr(1) = ">"
                    For TagNo = 0 To UBound(TagArray, 1)
                        i = 1
                        Do
                            j = InStr(i, kb, TagChr(0) & TagArray(TagNo, 0) & TagChr(1))
                            If j = 0 Then Exit Do
                            i = j + 1
                            Matches(0) = Matches(0) + 1
 'Debug.Print TagChr(0) & TagArray(Tagno, 0) & TagChr(1)
                        Loop
                    Next TagNo
                    TagChr(0) = "["
                    TagChr(1) = "]"
                    For TagNo = 0 To UBound(TagArray, 1)
                    i = 1
                        Do
                            j = InStr(i, kb, TagChr(0) & TagArray(TagNo, 0) & TagChr(1))
                            If j = 0 Then Exit Do
                            i = j + 1
                            Matches(1) = Matches(1) + 1
 'Debug.Print TagChr(0) & TagArray(Tagno, 0) & TagChr(1)
                        Loop
                    Next TagNo

#If False Then
'see if weve a network link in the file
                    i = InStr(1, kb, "<NetworkLink>")
                    If (i) Then NetworkLink = True
                    i = InStr(1, kb, "</NetworkLink>")
                    If (i) Then NetworkLink = False
                    i = InStr(1, kb, "<Link>")
                    If (i) Then Link = True
                    i = InStr(1, kb, "</Link>")
                    If (i) Then Link = False
                    i = InStr(1, kb, "<href>")
                    If (i) Then href = True
                    j = InStr(1, kb, "</href>")
                    If NetworkLink = True And Link = True And href = True And j <> 0 Then
                        hrefFile = PathFromFullName(TagTemplateReadFile) _
                        & "\" & Mid$(kb, i + 6, j - i - 6)
                        If FileExists(hrefFile) Then
'MsgBox TagTemplateReadFile
                            OverlayTemplateReadFile = hrefFile
                        Else
'This code is only needed if the overlay template file is to be read and parsed
'                            MsgBox hrefFile & vbCrLf & "not found", vbExclamation, "Template File not Found"
                            OverlayTemplateReadFile = hrefFile
                        End If
                    
                    End If
                    If (j) Then href = False
#End If
                
                Else
                    NoOutput = False    'skipped this line
                End If
'end of placemark - next is tail
                Select Case kb
                Case Is = "</Placemark>"
                    TagLineState = "TAIL"
                End Select
            Loop    'next line to input
            Close TagTemplateReadFileCh
            TagTemplateReadFileCh = 0
            If Matches(0) > Matches(1) Then
                TagChr(0) = "<"
                TagChr(1) = ">"
            Else
                TagChr(0) = "["
                TagChr(1) = "]"
                End If
 'Debug.Print "TagChr is " & TagChr(0) & TagChr(1)
            TagTemplateReadFileCh = 0
'replace any template tags, at the moment only tag ranges
            If (Not TagArray) <> -1 Then
'if (Not TagArray) = True Then
                For TagNo = 0 To UBound(TagArray, 1)
'min range
                    If TagArray(TagNo, 3) <> "" Then
                        For Channel = 1 To 2
                            TagTemplateHead(Channel) = Replace(TagTemplateHead(Channel), TagChr(0) & TagArray(TagNo, 0) & "_min" & TagChr(1), AsciiToXml(Channel, TagArray(TagNo, 3)))
                            TagTemplateContent(Channel) = Replace(TagTemplateContent(Channel), TagChr(0) & TagArray(TagNo, 0) & "_min" & TagChr(1), AsciiToXml(Channel, TagArray(TagNo, 3)))
                            TagTemplateTail(Channel) = Replace(TagTemplateTail(Channel), TagChr(0) & TagArray(TagNo, 0) & "_min" & TagChr(1), AsciiToXml(Channel, TagArray(TagNo, 3)))
                        Next Channel
                    End If
'max range
                    If TagArray(TagNo, 4) <> "" Then
                        For Channel = 1 To 2
                        TagTemplateHead(Channel) = Replace(TagTemplateHead(Channel), TagChr(0) & TagArray(TagNo, 0) & "_max" & TagChr(1), AsciiToXml(Channel, TagArray(TagNo, 4)))
                        TagTemplateContent(Channel) = Replace(TagTemplateContent(Channel), TagChr(0) & TagArray(TagNo, 0) & "_max" & TagChr(1), AsciiToXml(Channel, TagArray(TagNo, 4)))
                        TagTemplateTail(Channel) = Replace(TagTemplateTail(Channel), TagChr(0) & TagArray(TagNo, 0) & "_max" & TagChr(1), AsciiToXml(Channel, TagArray(TagNo, 4)))
                        Next Channel
                    End If
                Next TagNo
'            Call CheckTagExists
            End If
            Call CloseStartupLogFile
            Call SetOverlayOutputFile
            Call XcheckTags
Exit Sub
BadFile:
    On Error GoTo 0
'We must clear this as the file must be invalid
'This happens on the initial load if the TagArray contains no tags
'as
    LastTemplateFileName = ""
    Exit Sub
End Sub

Sub XcheckTags()
Dim TagLineState As String
Dim FileArry() As String
Dim i As Long
Dim j As Long
Dim k As Long
Dim Channel As Long
Dim File As String
Dim NxtFileCount As Long
Dim FileTag As String
Dim FileTagCount As Long
Dim kb As String
Dim XCheckErr As Boolean
Dim MissingFileTags As String
Dim MissingFieldTags As String
    
    On Error GoTo ProcErr
'On Error GoTo 0
'Create FileArry containing all tags used in the Template File
    ReDim FileArry(10)
    For Channel = 1 To 2
        File = TagTemplateHead(Channel) & TagTemplateContent(Channel) & TagTemplateTail(Channel)
    Next Channel
'Because these GE tags are written out as part of the content, we must remove them
    i = 0
    File = Replace(File, "<![CDATA[", "")
    File = Replace(File, "]]>", "")
    File = Replace(File, "[LINK]", "")
    File = Replace(File, "<marker", "") 'in templates/data.xml
    File = Replace(File, "/markers>", "") 'in templates/data.xml
    
'Do Until i * 1023 > Len(File)
'    MsgBox Mid$(File, i * 1023 + 1, (i + 1) * 1023)
'    i = i + 1
'Loop
    
    i = 1
    Do
        j = InStr(i, File, TagChr(0))
        If j = 0 Or j > Len(File) - 1 Then Exit Do  'At end "<>"
        i = j + 1
        j = InStr(i, File, TagChr(1))
        If j > i + 1 Then   'ignore blank tags
            FileTag = Mid$(File, i, j - i)
            Do
                If Len(FileArry(k)) = 0 Then
                    FileArry(k) = FileTag
                    FileTagCount = FileTagCount + 1
                End If
                If FileArry(k) = FileTag Then Exit Do
                k = k + 1
                If k > UBound(FileArry) Then ReDim Preserve FileArry(UBound(FileArry) + 10)
            Loop
        End If
        i = j + 1
    Loop
    
    kb = ""
    XCheckErr = False
    kb = "Your Template File is :-" & vbCrLf & TagTemplateReadFile & vbCrLf
    kb = kb & "The template file will be used because you have selected Output - Tagged" & vbCrLf
'Check if there are any Tags in the Template File (TagTemplateReadFile)
    If FileTagCount = 0 Then
        kb = kb & "There are no Tags in your Template File"
        kb = kb & " - Please examine your Template File" & vbCrLf
        XCheckErr = True
    Else
    ReDim Preserve FileArry(FileTagCount - 1)
    End If
'Check if some tags have been set up in the TagArray
    If Len(TagArray(0, 0)) = 0 Then
        kb = kb & "You have not set up any Tags"
        kb = kb & " (see ""Output Fields and Tags"" on the Options Window)" & vbCrLf
        kb = kb & "Please setup Output Fields with Tags matching your Template File (see Help)" & vbCrLf
        XCheckErr = True
    End If

    If XCheckErr = False Then
'Check if all tags in the TagArray are in the Template File (FileTags)
        MissingFileTags = ""
        For i = 0 To UBound(TagArray, 1)
            For j = 0 To UBound(FileArry)
                If TagArray(i, 0) = FileArry(j) Then Exit For
            Next j
            If j > UBound(FileArry) Then
                Select Case TagArray(i, 0)
                Case Is = "to_bow", "to_stern", "to_starboard", "to_port"
'These are required the OverlayFile if used
                Case Else
                    MissingFileTags = MissingFileTags & TagChr(0) & TagArray(i, 0) & TagChr(1)
                End Select
            End If
        Next i

'Check if all tags in the Template File are in the Tag Array (FieldTags)
        MissingFieldTags = ""
        For i = 0 To UBound(FileArry)
            For j = 0 To UBound(TagArray, 1)
                If FileArry(i) = TagArray(j, 0) Then Exit For
            Next j
            If j > UBound(TagArray, 1) Then
                Select Case FileArry(i)
                Case Is = "IconHeading", "IconScale", "IconColor", "Now", "CRC"
'These are SpecialTags and are always available
                Case Else
                    MissingFieldTags = MissingFieldTags & TagChr(0) & FileArry(i) & TagChr(1)
                End Select
            End If
        Next i
        
'If the file is an html file, the tags will be missing so we cant check it
        If Left$(MissingFieldTags, 14) = "<!DOCTYPE HTML" Then
            MissingFieldTags = ""
        End If
'Construct error message
        If Len(MissingFieldTags) <> 0 Then
            kb = kb & "Tags in Template file not allocated to Fields are :-" & vbCrLf
            kb = kb & MissingFieldTags & vbCrLf
            kb = kb & "Please create an ""Output Field and Tag"" matching the Template File (see Help)" & vbCrLf
        End If
        If Len(MissingFileTags) <> 0 Then
            kb = kb & "Fields and Tags defined not in your Template file are :-" & vbCrLf
            kb = kb & MissingFileTags & vbCrLf
            kb = kb & "Removing unused ""Output Tags and Range"" will improve performance (see Help)" & vbCrLf
        End If
    
    End If
    If Len(MissingFileTags) <> 0 Then XCheckErr = True
    If Len(MissingFieldTags) <> 0 Then XCheckErr = True
    If XCheckErr Then
        MsgBox kb, vbInformation, "XCheck Tags"
    End If
Exit Sub

ProcErr:
'Stop
    MsgBox "Bad Parse"
End Sub

'see if weve a Link (to an overlay file) in the content
'If found:-
'Sets OverlayReq
'The Overlay Output File name (which is the OutputFilename & _link)
Sub SetOverlayOutputFile()
Dim Channel As Long
Dim i As Long

            For Channel = 1 To 1
                Call WriteStartUpLog("")
                Call WriteStartUpLog("Checking Tag Template Tail for Link ")
                OverlayReq(Channel) = False
                OverlayTemplateReadFile = ""
                i = InStr(1, TagTemplateTail(Channel), "[LINK]")
                If i <> 0 Then
'only read the license once
'MsgBox LicenseStatus & ":" & LicenseLevel
                    Call WriteStartUpLog("Link found in Template tail")
                    Call WriteStartUpLog("LicenseStatus=" & LicenseStatus)
                    If LicenseStatus = "" Then Load frmLockMain
                    Call WriteStartUpLog("LicenseLevel=" & LicenseLevel)
                    If LicenseLevel = "Test Version" Then
                        If InStr(1, OutputFileName, ".kmz", vbTextCompare) <> 0 Then
                            OverlayReq(Channel) = True
'force the default overlayoutputfilename. If you don't the
'previous file name will be hung on to if the overlaytemplate file
'is changed. Note the user has no control over the OverlayOutputFileName
                            OverlayOutputFileName = ""
                            OverlayOutputFileName = FileSelect.SetFileName("OverlayOutputFileName")
                            OverlayTemplateReadFile = "In Memory\Defined"
                        Else
                            Call WriteStartUpLog("The Output File must be a kmz file to link to another file")
MsgBox "The Output File must be a kmz file to link to another file"
                        End If
                    Else
                        Call WriteStartUpLog("Not Licensed for using Overlays")
                        frmLockMain.Show
MsgBox "Not Licensed for using Overlays"
                    End If
                    Call WriteStartUpLog("Link finished")
                Else
                    Call WriteStartUpLog("Link not found")
                End If
            Next Channel
            Call CloseStartupLogFile

'                        If InStr(1, OutputFileName, ".kmz", vbTextCompare) <> 0 Then
'see if we are trying to used the CloseUp view of vessels
'OverlayTemplateReadFile = ExtendFullName(TagTemplateReadFile, "_link")
End Sub

'May be used later
#If False Then
Sub CheckTagExists()
Dim kb As String
Dim i As Long
Dim O1 As Long  'open (<) 1
Dim O2 As Long
Dim C1 As Long  'close (>) 1
Dim C2 As Long
Dim Token As String
Dim Remove As Long

'NEED toCHECK BOTH CHANNELS
    kb = TagTemplateHead + TagTemplateContent + TagTemplateTail
    i = 1
    Do Until i >= Len(kb)
MsgBox kb
        Remove = 0
        O1 = InStr(i, kb, "<")
        C1 = InStr(O1, kb, ">")
'ensure token is complete
        If C1 <> 0 Then
'ensure its not a closing token
            If Mid$(kb, O1, 2) <> "</" Then
                Token = Mid$(kb, O1 + 1, C1 - O1 - 1)
'remove this token
'check if NEXT token is </>
                O2 = InStr(C1, kb, "<")
                If Mid$(kb, O2, 3) = "</>" Then
'remove both tokens closing first
                                        
                    Mid$(kb, O2, 3) = ""
                    Mid$(kb, O1, O2 - O1 + 1) = ""
                    i = O1
                    Remove = O2 - O1 + 1 + 3
MsgBox Token & vbCrLf & "has been removed </>"
                Else
'check if we can find a full closing tag, not necesarily
'next
                    O2 = InStr(C1, kb, "</" & Token & ">")
                    If (O2) <> 0 Then
'remove both tokens closing first
kb = Left$(kb, O2 - 1) & Right$(kb, Len(kb) - O2 - Len(Token) - 2)
'MsgBox kb
'O2 = O2 - Len(Token) - 3
kb = Left$(kb, O1 - 1) & Right$(kb, Len(kb) - O1 - Len(Token) - 1)
'MsgBox kb
                        
                        Remove = Len(Token) * 2 + 5
MsgBox Token & vbCrLf & "has been removed"
                        i = O1
                    Else
'tag not removed start 1 right and go again
MsgBox Token & vbCrLf & "has been left"
                        i = C1
                    End If
                End If
            End If
        Else        'no last >
            Exit Do
        End If
    Loop
End Sub
#End If

'This is enabled whenever Scheduled output is on
'and data is expected.
'If the time stamp on an sentence being processed
'calls for the scheduler to be run, the timer is
're-enabled (restarting the timer).

Private Sub ScheduledTimer_Timer()
    If ScheduledTimer.Interval = 60000 Then
        ScheduledTimerElapsedMins = ScheduledTimerElapsedMins + 1
        If ScheduledTimerElapsedMins = UserScheduledMins Then
'Now use up the remaining part minutes
            ScheduledTimer.Interval = UserScheduledMilliSecs
        End If
    Else
        ScheduledTimerElapsedMilliSecs = ScheduledTimerElapsedMilliSecs + ScheduledTimer.Interval
    End If
Debug.Print "Elapsed=" & ScheduledTimerElapsedMins & "." & ScheduledTimerElapsedMilliSecs / 60000
    
    If ScheduledTimerElapsedMins = UserScheduledMins _
    And ScheduledTimerElapsedMilliSecs = UserScheduledMilliSecs Then
        If Processing.Suspended = False Then Call SchedulerOutput("Timer")
'reset ElapsedMins & Interval
        Call ResetScheduledTimer(UserScheduledSecs)
    Else
        NmeaRcv.StatusBar.Panels(1) = "Next Scheduled in " & UserScheduledMins - ScheduledTimerElapsedMins & "." & (UserScheduledMilliSecs - ScheduledTimerElapsedMilliSecs) / 60000 & " mins"
        NmeaRcv.ClearStatusBarTimer.Enabled = True
    End If
End Sub

Public Sub ResetScheduledTimer(Secs As Long)
'UserScheduledSecs is in seconds
    UserScheduledMins = Int(UserScheduledSecs / 60)
    UserScheduledMilliSecs = (UserScheduledSecs - UserScheduledMins * 60) * 1000
    If UserScheduledMins > 0 Then
        ScheduledTimer.Interval = 60000
    Else
        ScheduledTimer.Interval = UserScheduledMilliSecs
    End If
    ScheduledTimerElapsedMins = 0
    ScheduledTimerElapsedMilliSecs = 0
'Reset when Started and Stoped occurs after
'    If Not ScheduledTimer.Enabled = True Then
'        NmeaRcv.StatusBar.Panels(1) = "Reset Scheduled in " & UserScheduledMins - ScheduledTimerElapsedMins & "." & (UserScheduledMilliSecs - ScheduledTimerElapsedMilliSecs) / 60000 & " mins"
'        NmeaRcv.ClearStatusBarTimer.Enabled = True
'    End If
End Sub

Private Sub ServerUDP_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    MsgBox "Socket Error " & Number & ": " & Description       ' show some "debug" info
    ServerUDP.Close ' close the erraneous connection
End Sub

'The Open_Timer keeps retrying the open if unsuccessful
Function OpenSerialInput() As Boolean
Dim kb As String

#If UseComm Then
    WriteLog "Using CommHandler", LogForm
    Call OpenCommHandler
    CurrentSocket = 1
    Exit Function
#Else
    WriteLog "Using MSComm", LogForm
#End If

'Debug.Print "#OpenSerialInput"
Set MyInputThrottleTimer(1) = New clsWaitableTimer
On Error GoTo Load_err
With MSComm1
'make sure the serial port is not open (by this program)
'And set crlfdetected
    Call CloseSerialInput
'set the active serial port
    CrLfDetected = False
    SerialRcvBuffer = ""
kb = Mid$(TreeFilter.Combo1(0).Text, 4)
    .CommPort = Mid$(TreeFilter.Combo1(0).Text, 4)
'.CommPort = "\\.\" & TreeFilter.Combo1(0).Text
'.CommPort = "\\.\COM6"
'set the badurate,parity,databits,stopbits for the connection
    .Settings = TreeFilter.Combo1(1).Text & ",N,8,1"

 Debug.Print "Port " & .CommPort
 Debug.Print "Open " & MSComm1.Settings


'set the DRT and RTS flags
    .DTREnable = True
'.RTSEnable = True
'enable the oncomm event for 100 every received characters
'Set a highenough figure to stand a good chance of finding a CrLF
'if the baud rate is OK
    .RThreshold = 100
'disable the oncomm event for send characters
    .SThreshold = 0
    .InputMode = comInputModeBinary
    .InBufferSize = 4096
    .InBufferCount = 0  'read all the buffer
'open the serial port
'Possible err 15 "Couldnt set state (if close and open too quick)
    On Error Resume Next
    .PortOpen = True
    On Error GoTo 0
    If .PortOpen = False Then
        kb = "Couldn't Open Serial Port(" & .CommPort & ") retrying in 5 Seconds"
StatusBar.Panels(1) = kb
OpenSerialTimer.Interval = 5000   'time to wait for retry
        OpenSerialTimer.Enabled = True    'try re-opening
    Else
StatusBar.Panels(1) = ""
        OpenSerialTimer.Enabled = False
    End If
End With 'MSComm1
Debug.Print "OpenSerialTimerEnabled " & OpenSerialTimer.Enabled
Exit Function

Load_err:
    OpenSerialTimer.Enabled = True
    Select Case err.Number
    Case Is = 8002   'invalid port
        OpenSerialTimer.Interval = 500  'millisecs
    Case Is = 8005   'port already open
        OpenSerialTimer.Interval = 3000
    Case Is = 8015    'set com state failed
        OpenSerialTimer.Interval = 3000
    Case Else
        MsgBox "The following error has occurred:" & vbNewLine _
         & "Err # " & err.Number & " - " & err.Description, _
           vbCritical, _
           "Open Error"
        OpenSerialTimer.Interval = 3000
    End Select
    OpenSerialTimer.Enabled = True    'keep trying

End Function

'The Open_Timer keeps retrying the open if unsuccessful
Function OpenCommHandler() As Boolean
Dim kb As String
Dim Idx As Long
     
    Idx = 1
'    Debug.Print "#OpenCommHandler"
    Set MyInputThrottleTimer(1) = New clsWaitableTimer
    On Error GoTo Load_err
'make sure the serial port is not open (by this program)
'And set crlfdetected
    Call CloseHandler(1)
'set the active serial port
    CrLfDetected = False
    SerialRcvBuffer = ""
    sockets(Idx).Handler = 1 'Serial
kb = TreeFilter.Combo1(0).Text
    sockets(Idx).Comm.Name = TreeFilter.Combo1(0).Text
    sockets(Idx).DevName = FriendlyName(sockets(Idx).Comm.Name)
'set the baudrate,parity,databits,stopbits for the connection
    sockets(Idx).Comm.BaudRate = TreeFilter.Combo1(1).Text
    sockets(Idx).Comm.AutoBaud = True
'open the serial port
    sockets(Idx).Enabled = True 'enable the handler
    On Error Resume Next
 'Now create the Serial Socket
    Call OpenHandler(Idx)
    On Error GoTo 0
#If False Then
    .PortOpen = True
    On Error GoTo 0
    If .PortOpen = False Then
        kb = "Couldn't Open Serial Port(" & .CommPort & ") retrying in 5 Seconds"
StatusBar.Panels(1) = kb
OpenSerialTimer.Interval = 5000   'time to wait for retry
        OpenSerialTimer.Enabled = True    'try re-opening
    Else
StatusBar.Panels(1) = ""
        OpenSerialTimer.Enabled = False
    End If
End With 'MSComm1
Debug.Print "OpenSerialTimerEnabled " & OpenSerialTimer.Enabled
#End If
Exit Function

Load_err:
    OpenSerialTimer.Enabled = True
    Select Case err.Number
    Case Is = 8002   'invalid port
        OpenSerialTimer.Interval = 500  'millisecs
    Case Is = 8005   'port already open
        OpenSerialTimer.Interval = 3000
    Case Is = 8015    'set com state failed
        OpenSerialTimer.Interval = 3000
    Case Else
        MsgBox "The following error has occurred:" & vbNewLine _
         & "Err # " & err.Number & " - " & err.Description, _
           vbCritical, _
           "Open Error"
        OpenSerialTimer.Interval = 3000
    End Select
    OpenSerialTimer.Enabled = True    'keep trying

End Function

Private Sub OpenSerialTimer_Timer()
'    If MSComm1.PortOpen = True Then
'        Call CloseSerialInput
'        MSComm1.PortOpen = False
'    End If

 'Debug.Print "Timer " & OpenSerialTimer.Interval
    Call OpenSerialInput     'resetting port
End Sub

'this fires every 1/2 second
Private Sub StatsTimer_Timer()
Dim UtcTime As SYSTEMTIME

'MsgBox "StatsTimer " & Now()
    Call UpdateStats
'Get Current UTC time
    Call GetSystemTime(UtcTime) 'SYSTEMTIME
       
'If an input buffer is full (eg UDP) then We must try and clear input buffer as it will not be
'triggered by more UDP data being received
    If InputState = 1 Then 'running and not paused
        If NmeaBufXoff = False Then
            Call InputBuffersToNmeaBuf
        Else
'NmeaBuf is full & not paused
            Call ProcessNmeaBuf
            Call UpdateStats    'Buffer emptied
''        DoEvents    'Otherwise Click Pause is not actioned
''        MyInputThrottleTimer(0).Wait 50
'            Call UpdateStats    'if input buffer was full and now < 50% will re-open handler
        End If
'        Call InputBuffersToNmeaBuf
'Turn on receiving data if buffers have been previously turned off because they were full
'Also empties Input buffers
        If Check1(0).Value = vbGrayed Then
            Check1(0).Value = vbChecked     'UDP
            Call UpdateStats    'Buffer emptied
        End If
        If Check1(5).Value = vbGrayed Then
            Check1(5).Value = vbChecked     'TCP
        End If
        If Check1(1).Value = vbGrayed Then
            Check1(1).Value = vbChecked     'Serial
        End If
'File is left open even when paused
    End If
End Sub

Private Sub TcpTimer_Timer()
Static ListenCount As Long
'Peer Closing Connection
'When acting as TCP Server and Client closes connect
'then close then server & try re-opening
'If it cannot open as a Client
'OpenCientTCP will re-open as a Server
 'Debug
'Call frmDpyBox.DpyBox(ClientTCP.State & vbCrLf, 5, "TCP Timer")
    
    Select Case ClientTCP.State
'We can stop the timer if the socket is closed
'Start the time when anythin happens
    Case Is = sckClosed '0
        TcpTimer.Enabled = False
    Case Is = sckOpen '1
    Case Is = sckClosing    '8 Peer is closing connection
        Check1(5).Caption = "TCP"
        Call OpenClientTCP
    Case Is = sckListening  '2
            Check1(5).Caption = "TCP"
            Call OpenClientTCP
    Case Is = sckConnecting '6
'Occurs when ClientTCP cannot Connect to Server
'try Server
        Call OpenServerTCP
    Case Is = sckConnected  '7
'Occurs when ClientTCP has connected to server
'Data should be now coming
    Case Is = sckClosed     '0
    Case Is = sckError      '9
        Select Case ClientTCPError
'When ClientTCP tries to connect to a TcpServer
'and the Server or Firewall will not allow the connection
'My message is more meaningful to a user
        Case Is = sckConnectionRefused  'err 10061
            frmDpyBox.DpyBox ClientTCPError & " TCP Connection is Refused by Remote Firewall or Server" & vbCrLf, 5, "TCP Error"
            Call OpenServerTCP
        Case Else
            frmDpyBox.DpyBox ClientTCPError & " " & ClientTCPErrorDescription & vbCrLf, 5, "TCP Error"
        End Select
    Case Else   '3,4,5
        Call frmDpyBox.DpyBox(ClientTCPError & " " & ClientTCPErrorDescription & vbCrLf, 5, "TCP Error")
    End Select
'Stop
End Sub

Private Sub TimeZoneTimer_Timer()
Dim UtcTime As SYSTEMTIME
    
'Once a second update Time displayed
    Option1(11).Caption = LocalTimeZoneName()
    Option1(14).Caption = "UTC"
    If Option1(11).Value = True Then    'Local
        Label2(0) = Time$
        Frame10.Caption = Date  'Local Formate, Date$ format american format
    Else
        Option1(14) = True
        Label2(0) = TimeValue(NowUtc())
        On Error Resume Next    'DateValue does not work with Japanese
        Frame10.Caption = DateValue(NowUtc())
        On Error GoTo 0
    End If
        
'MsgBox "StatsTimer " & Now()
'    Call UpdateStats

'Get Current UTC time
    Call GetSystemTime(UtcTime) 'SYSTEMTIME
        
'Dont bother checking if not running, or is paused
    If InputState = 1 Then
'Update Vessels.dat every hour
        If CacheVessels = True Then
            If UtcTime.wHour <> VesselsFileSyncTime.wHour _
            Or UtcTime.wDay <> VesselsFileSyncTime.wDay Then
                Call SaveVessels
            End If
        End If
                
'Check for updates every hour
        If TreeFilter.Check1(2).Value <> 0 Then
            If UtcTime.wHour <> LastProgramUpdateCheckTime.wHour _
            Or UtcTime.wDay <> LastProgramUpdateCheckTime.wDay Then
                Call CheckUpdates   'Updates LastProgramUpdateCheckTime
            End If
        End If
    
'Check if InputLofFile needs rolling over every second
        If Not (InputLogFile Is Nothing) Then
'If rolling over
            If NmeaLogFileDate <> "" Then
'if date change, close and re-open with new name
                If NmeaLogFileDate <> Format$(Now(), "yyyy-mm-dd") Then
                    InputLogFile.CloseFile
'v145                    NmeaLogFileDate = Format$(Now(), "yyyy-mm-dd")
                    Call OpenInputLogFile   'v145 this will set the new date on the file name
'v145                    Call InputLogFile.OpenFileAppend(NmeaLogFile)
                End If
            End If
        End If

        If OutputFileRollover(1) = True Then
                If OutputFileDate <> Format$(Now(), "yyyy-mm-dd") Then
'Close existing file (if open)
                    If Not (OutputDataFile Is Nothing) Then 'v137
                      OutputDataFile.CloseFile
                       Set OutputDataFile = Nothing
                       OutputFileDate = Format$(Now(), "yyyy-mm-dd")
'File will be re-opened on next Call OutputToDataFile with new date
                    End If
                End If
        End If
    End If  'Running
End Sub


Private Sub ResetCycle()

'Debug.Print "ReceivedLastCycle " & RcvThisTimerCycle & "+" & FileThisTimerCycle
    RcvThisTimerCycle = 0
    FileThisTimerCycle = 0
    FileCycleCount = 0
    RcvSpeedClock.ResetTimer
    RcvSpeedClock.StartTimer
Call UpdateStats
End Sub

'Used to check if ClickStart is valid and if ClickStop reqd at File Input EOF
Private Function InputsCount(Optional InputType As String) As Long
Dim Count As Long

    If Check1(0).Value = vbChecked Then Count = Count + 1
    If Check1(1).Value = vbChecked Then Count = Count + 1
    If Check1(5).Value = vbChecked Then Count = Count + 1
    If InputType <> "Rcv" Then
        If Check1(6).Value = vbChecked Then Count = Count + 1
    End If
    InputsCount = Count
End Function


'Called by InputBuffersToNmeaBuf reaches upper limit of buffer capacity
Public Sub XoffNmeaBuf()  'Call to turn off getting data (set Xoff=true)
    If NmeaBufXoff = False Then
'Must action immediately to force exit from within ReadSequential DO loop
        NmeaBufXoff = True
Debug.Print "Xoff (NmeaBuf)"
    End If
End Sub

Public Function RcvBufferedMsgs() As Long
Dim SerialCount As Long
Dim UDPCount As Long
Dim TCPCount As Long
Dim FileCount As Long
Dim cTemp As Currency
    SerialCount = LFCount(SerialRcvBuffer)
    UDPCount = LFCount(UDPRcvBuffer)
    TCPCount = LFCount(TCPRcvBuffer)
    FileCount = LFCount(FileRcvBuffer)
    RcvBufferedMsgs = SerialCount + UDPCount + TCPCount + FileCount
'StatusBar.Panels(1).Text = "Buffered = " & Count
'StatusBar.Panels(1).Text = "Serial = " & SerialCount _
'    & ", UDP = " & UDPCount _
'    & ", TCP = " & TCPCount _
'    & ", File = " & FileCount
End Function

'Used to kick off reading the File
Private Sub FileNextBlockTimer_Timer()
'Debug.Print "#FileNextBlockTimer"
    
'Will remain here until the FileInput is terminated
'    Call InputFile.LongTask(MaxFilePerTimerCycle, 0.1)
    Call FileNextBlock
    FileNextBlockTimer.Enabled = False

End Sub

Private Function RcvBufferBytes() As String
    RcvBufferBytes = Len(SerialRcvBuffer) _
    + Len(UDPRcvBuffer) _
    + Len(TCPRcvBuffer) _
    + Len(FileRcvBuffer)
End Function


Public Sub UpdateStats()
Dim kb As String
    With Processing
        If .NmeaBuf Then kb = kb & "NmeaBuf "
        If .Scheduler Then kb = kb & "Scheduler "
        If .InputOptions Then kb = kb & "InputOptions "
        If .Suspended Then
            kb = kb & "wait "
        End If
        If .Paused Then kb = kb & "paused"
If .Suspended = True And .NmeaBuf = False And .Scheduler = False And .InputOptions = False And .Paused = False Then Stop
If .Suspended = False And (.NmeaBuf = True Or .Scheduler = True Or .InputOptions = True Or .Paused = True) Then Stop
    End With
   
StatusBar.Panels(2).Text = kb
'If DecodingState <> 0 Then Stop
    If Not (RunClock Is Nothing) Then
        If RunClock.Duration > 0 Then     'div/0
            LastSpeed = (Received - LastReceived) / RunClock.Duration * 60
'First start
            If LastReceived = 0 Then ForecastSpeed = LastSpeed
            ForecastSpeed = LastSpeed * 0.1 + ForecastSpeed * 0.9
            Label3.Caption = Format$(ForecastSpeed, "#####0") & "/min"
            LastReceived = Received
            RunClock.ResetTimer
        Else
            LastReceived = Received
        End If
    Else
'        Stop    'Debug should not be called until running
    End If
    
    Select Case Len(UDPRcvBuffer)
    Case Is > UDP_RCV_MAX
        lblBuffer(0).Caption = aByte(Len(UDPRcvBuffer))
        Check1(0).BackColor = vbRed
        Check1(0).Value = vbGrayed
    Case Is > UDP_RCV_MAX / 2
        lblBuffer(0).Caption = aByte(Len(UDPRcvBuffer))
        Check1(0).BackColor = vbYellow
    Case Is > UDP_RCV_INUSE
        lblBuffer(0).Caption = aByte(Len(UDPRcvBuffer))
        Check1(0).BackColor = vbGreen
    Case Else
        lblBuffer(0).Caption = aByte(Len(UDPRcvBuffer))
        Check1(0).BackColor = &H8000000F
    End Select

    Select Case Len(TCPRcvBuffer)
    Case Is > TCP_RCV_MAX
        lblBuffer(1).Caption = aByte(Len(TCPRcvBuffer))
        Check1(5).BackColor = vbRed
        Check1(5).Value = vbGrayed
    Case Is > TCP_RCV_MAX / 2
        lblBuffer(1).Caption = aByte(Len(TCPRcvBuffer))
        Check1(5).BackColor = vbYellow
    Case Is > TCP_RCV_INUSE     'gets bigger chunks than udp
        lblBuffer(1).Caption = aByte(Len(TCPRcvBuffer))
        Check1(5).BackColor = vbGreen
    Case Else
        lblBuffer(1).Caption = aByte(Len(TCPRcvBuffer))
        Check1(5).BackColor = &H8000000F
    End Select

    Select Case Len(SerialRcvBuffer)
    Case Is > SERIAL_RCV_MAX
        lblBuffer(2).Caption = aByte(Len(SerialRcvBuffer))
        Check1(1).BackColor = vbRed
        Check1(1).Value = vbGrayed
    Case Is > SERIAL_RCV_MAX / 2
        lblBuffer(2).Caption = aByte(Len(SerialRcvBuffer))
        Check1(1).BackColor = vbYellow
    Case Is > SERIAL_RCV_INUSE
        lblBuffer(2).Caption = aByte(Len(SerialRcvBuffer))
        Check1(1).BackColor = vbGreen
    Case Else
        lblBuffer(2).Caption = aByte(Len(SerialRcvBuffer))
        Check1(1).BackColor = &H8000000F
    End Select

    Select Case Len(FileRcvBuffer)
    Case Is > FILE_RCV_MAX
        lblBuffer(3).Caption = aByte(Len(FileRcvBuffer))
        Check1(6).BackColor = vbRed
    Case Is > FILE_RCV_MAX / 2
        lblBuffer(3).Caption = aByte(Len(FileRcvBuffer))
        Check1(6).BackColor = vbYellow
    Case Is > FILE_RCV_INUSE
        lblBuffer(3).Caption = aByte(Len(FileRcvBuffer))
        Check1(6).BackColor = vbGreen
    Case Else
        lblBuffer(3).Caption = aByte(Len(FileRcvBuffer))
        Check1(6).BackColor = &H8000000F
    End Select
    
With Stats
    .TextMatrix(0, 1) = aByte(cBytesRx * 10000)
    If CDec(RcvBufferBytes) > UDP_RCV_INUSE + TCP_RCV_INUSE Then    'stop stat flashing
        .TextMatrix(1, 1) = aByte(RcvBufferBytes)
    Else
        .TextMatrix(1, 1) = ""
    End If
    .TextMatrix(2, 1) = Rejected ' + Buffered
    .TextMatrix(3, 1) = Received ' + Buffered
    .TextMatrix(4, 1) = NmeaBufUsed ' + Buffered 'was Waiting
    .TextMatrix(5, 1) = Processed
    .TextMatrix(6, 1) = Filtered
    .TextMatrix(7, 1) = Outputted
    .TextMatrix(8, 1) = AisMsgs.Count
    .TextMatrix(9, 1) = NamedVessels    'Vessels.Count
    .TextMatrix(10, 1) = SchedOut

'put some intelligence into when stats are updated
'    If (NmeaBufUsed And NmeaBufUsed Mod 100 = 0 And NmeaBufUsed <> NMEABUF_MAX) _
'    Or (Received And Received Mod 1000 = 0) _
'    Or (Processed And Processed Mod 10 = 0) Then
'        If .Redraw = False Then 'update even if not updating
'            .Redraw = True
'            .Refresh
'            .Redraw = False
'        Else
'            .Refresh
'        End If
'    End If
End With
If MyShip.Mmsi = "" Then
    NmeaRcv.Label4(0) = ""
    NmeaRcv.Frame11.Visible = False
End If

'processing of Input buffer is triggered by data arriving.
'if input is not stopped and Nmeabuf is full and no more data data has arrive
'input buffers must be processed
'If cbStop.Enabled = True Then
'    If TCPRcvBuffer <> "" Or UDPRcvBuffer <> "" Or SerialRcvBuffer <> "" Or FileRcvBuffer <> "" Then
'        If NmeaBufXoff = True Then
'            If Processing.Suspended = False Then Call ProcessNmeaBuf
'        Else
'            Call InputBuffersToNmeaBuf
'        End If
'    End If
'End If
End Sub

Public Sub ClearStats()
With Stats
    cBytesRx = 0@
    Rejected = 0
    Received = 0
    Processed = 0
    Filtered = 0
    Outputted = 0
    SchedOut = 0
    LastReceived = 0
    LastSpeed = 0
    ForecastSpeed = 0
End With
StatusBar.Panels(1).Text = ""
With Track
    .TCPArrival = 0
    .TCPDataRcvLFCount = 0
    .TCPCallCountExit = 0
    .TCPDataRcvLoops = 0
End With
Call UpdateStats
End Sub

Private Sub InputFile_PercentDone(ByRef Data As String, ByVal Sentences As Long, Cancel As Boolean)
Dim kb As String
Dim WaitTime As Long    'millisecs

'Debug.Print "#PercentDone"

            
    FileRcvBuffer = FileRcvBuffer & Data
    cBytesRx = cBytesRx + Len(Data) * 0.0001@
    FileCycleCount = FileCycleCount + Sentences

'If the FileRcvBuffer is full we need to keep trying here, otherwise
'LongTask will set stuck in a loop - note we also must stall LongTask in PercentDone
    If NmeaBufXoff = False Then
        Call InputToNmeaBuf(FileRcvBuffer, "File")
    End If
'NmeaBufXoff wil get set above when yhe buffer is full
    If NmeaBufXoff = True Then
        Do Until NmeaBufXoff = False
Debug.Print "FileRcvBuffer " & LFCount(FileRcvBuffer)
            If Processing.Suspended = False Then Call ProcessNmeaBuf
'keep trying to empty FileRcvBuffer into NmeaBuf
            FileWaitableTimer.Wait 50
        Loop
    End If

'If this is not at the end of the timer cycle wait
'NOTE complete this cycle to empty the NmeaBuf, even if we have an EOF
    If FileCycleCount = MaxFilePerTimerCycle Then
        WaitTime = SPEED_INTERVAL - CycleClock.Duration * 1000
        If WaitTime > 0 Then
Debug.Print "PercentDone.Wait " & WaitTime
            FileWaitableTimer.Wait WaitTime
        End If
    End If

'This stops any more PercentDone events getting queued
    If InputFileCancel Then
'Here if Stop clicked manually - this will close the input file
        Cancel = True
'        Stop
    Else
        If Sentences = 0 Then
'Here if EOF - but InputFile is still Open
            If InputsCount = 1 Then 'FileInput is the only input
                cbStop.Value = True 'Stop Input and Output
            Else
                Call CloseFileInput 'StopInput only
           End If
        End If
    End If
End Sub


Private Sub PollTimer_Timer()
Dim Hidx As Integer
Dim kb As String

'If comm has not yet been dimensioned we skip polling
    On Error GoTo Timer_error
    For Hidx = 1 To UBound(Comms)
        If Not Comms(Hidx) Is Nothing Then
'If Idx = 2 Then Stop
'        kb = kb & Idx & ":" & Sockets(Idx).State & " "
'Hidx may not yet have been set
            If Comms(Hidx).sIndex > 0 Then
                If sockets(Comms(Hidx).sIndex).State > 0 Then
                    Comms(Hidx).Poll    'Gets the serial data
               End If
            End If
        End If
    Next Hidx
Exit Sub

Timer_error:
'v45 changed from Sockets(Comms(Hidx).sIndex) to Hidx as Sockets() may not exist
'v146 change to non fatal error (Sergey)
'    MsgBox "Poll Timer Error " & Str(err.Number) & " " & err.Description & vbCrLf _
'    & "On Comm (" & Hidx & ")", , "Poll Timer"
        WriteLog "Poll Timer Error " & Str(err.Number) & " " & err.Description _
    & "On Comm (" & Hidx & ")", LogForm

End Sub

Public Sub CommRcv(commdata As String, Source As Long)
Dim bytestoforward As Long
'Stop
'commdata$ has a NULL appended at end of last input
        If Right$(commdata$, 1) = Chr$(0) Then
            bytestoforward = Len(commdata$) - 1
        Else
            bytestoforward = Len(commdata$)
        End If
        If bytestoforward > 0 Then
            If Not Comms(Source) Is Nothing Then
'                Call ForwardData(Left$(commdata$, bytestoforward), Comms(Source).sIndex)
'Debug                Call TermOutput(Left$(commdata$, bytestoforward), Comms(Source).sIndex)
                Call CommDataArrival(commdata$, bytestoforward)
            End If
        End If
'    Call ForwardData(commdata, Source)
End Sub


Public Function TermOutput(Data As String, Optional Source As Long) As Long

    WriteLog Data, LogForm
'Display Received Sentence
'    If txtTerm.Enabled = True Then
'        txtTerm.SelStart = Len(txtTerm.Text)
'        If Source <> 0 Then
'            txtTerm.SelText = CStr(Source) & "<" & Data
'        Else
'            txtTerm.SelText = Data
'        End If
'         If Len(txtTerm.Text) > 4096 Then
'            txtTerm.Text = Right$(txtTerm.Text, 2048)
'        End If
'    End If
'    TermOutput = SendMessageAsLong(txtTerm.hWnd, EM_GETLINECOUNT, 0, 0)
End Function




