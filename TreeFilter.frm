VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form TreeFilter 
   Caption         =   "Options"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12720
   BeginProperty Font 
      Name            =   "Microsoft Sans Serif"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "TreeFilter.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   12720
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Frame Frame18 
      Caption         =   "FTP server"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   3840
      TabIndex        =   88
      Top             =   5160
      Width           =   1815
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   10
         Left            =   120
         TabIndex        =   92
         Text            =   "mydirectory"
         Top             =   960
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   9
         Left            =   120
         PasswordChar    =   "*"
         TabIndex        =   91
         Tag             =   "encrypt"
         Text            =   "mypassword"
         Top             =   720
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   8
         Left            =   120
         TabIndex        =   90
         Text            =   "username"
         Top             =   480
         Width           =   1185
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   7
         Left            =   120
         TabIndex        =   89
         Text            =   "my.server.com"
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label1 
         Caption         =   "Dir"
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
         Left            =   1320
         TabIndex        =   96
         Top             =   960
         Width           =   345
      End
      Begin VB.Label Label1 
         Caption         =   "Pass"
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
         Left            =   1320
         TabIndex        =   95
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "User"
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
         Left            =   1320
         TabIndex        =   94
         Top             =   480
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Host"
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
         Left            =   1320
         TabIndex        =   93
         Top             =   240
         Width           =   465
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Input"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   74
      Top             =   0
      Width           =   1815
      Begin VB.Frame Frame20 
         Caption         =   "TCP Host"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   83
         Top             =   840
         Width           =   1575
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            IMEMode         =   3  'DISABLE
            Index           =   12
            Left            =   120
            TabIndex        =   85
            Text            =   "remote.server.com"
            Top             =   240
            Width           =   1305
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            IMEMode         =   3  'DISABLE
            Index           =   13
            Left            =   120
            TabIndex        =   84
            Tag             =   "Numeric"
            Text            =   "12345"
            Top             =   600
            Width           =   825
         End
         Begin VB.Label Label1 
            Caption         =   "Port"
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
            Left            =   1080
            TabIndex        =   86
            Top             =   600
            Width           =   375
         End
      End
      Begin VB.Frame Frame19 
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
         Height          =   975
         Left            =   120
         TabIndex        =   78
         Top             =   1800
         Width           =   1575
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   80
            Top             =   600
            Width           =   975
         End
         Begin VB.ComboBox Combo1 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   79
            Top             =   240
            Width           =   975
         End
         Begin VB.Label lblCombo1 
            Caption         =   "Speed"
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
            Left            =   1080
            TabIndex        =   82
            Top             =   600
            Width           =   465
         End
         Begin VB.Label lblCombo1 
            Caption         =   "Port"
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
            Left            =   1080
            TabIndex        =   81
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.Frame Frame8 
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
         Height          =   615
         Left            =   120
         TabIndex        =   75
         Top             =   240
         Width           =   1575
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   3
            Left            =   120
            TabIndex        =   76
            Text            =   "Text1"
            Top             =   240
            Width           =   720
         End
         Begin VB.Label Label1 
            Caption         =   "Port"
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
            Left            =   960
            TabIndex        =   77
            Top             =   240
            Width           =   435
         End
      End
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Spare"
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
      Left            =   11280
      TabIndex        =   71
      Top             =   6240
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Spare"
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
      Left            =   11280
      TabIndex        =   70
      Top             =   5880
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   840
   End
   Begin VB.Frame Frame14 
      Caption         =   "Filtered Output"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   5760
      TabIndex        =   31
      Top             =   2760
      Width           =   5415
      Begin VB.CheckBox Check1 
         Caption         =   "Cache Vessel Names"
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
         Index           =   21
         Left            =   3480
         TabIndex        =   68
         Tag             =   "csvfile"
         Top             =   240
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   11
         Left            =   1080
         TabIndex        =   64
         Text            =   "dd/mm/yyyy hh:nn:ss"
         Top             =   240
         Width           =   2235
      End
      Begin VB.Frame Frame11 
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
         Height          =   1815
         Left            =   2760
         TabIndex        =   51
         Top             =   600
         Width           =   2535
         Begin VB.CheckBox Check1 
            Caption         =   "Head"
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
            Index           =   20
            Left            =   840
            TabIndex        =   67
            Tag             =   "csvfile"
            Top             =   720
            Width           =   735
         End
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
            Index           =   18
            Left            =   1440
            TabIndex        =   63
            Tag             =   "csvfile"
            Top             =   240
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.OptionButton Option1 
            Caption         =   "CSV"
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
            TabIndex        =   60
            Top             =   720
            Value           =   -1  'True
            Width           =   765
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Tagged"
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
            Left            =   120
            TabIndex        =   59
            Top             =   960
            Width           =   915
         End
         Begin VB.OptionButton Option1 
            Caption         =   "NMEA"
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
            Left            =   120
            TabIndex        =   58
            Top             =   480
            Width           =   885
         End
         Begin VB.CheckBox Check1 
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
            Index           =   6
            Left            =   120
            TabIndex        =   57
            Tag             =   "csvfile"
            Top             =   1200
            Width           =   1200
         End
         Begin VB.CheckBox Check1 
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
            Index           =   7
            Left            =   120
            TabIndex        =   56
            Tag             =   "csvfile"
            Top             =   1440
            Width           =   1425
         End
         Begin VB.CheckBox Check1 
            Caption         =   "+Time Stamp"
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
            Left            =   1080
            TabIndex        =   55
            Tag             =   "csvfile"
            Top             =   480
            Width           =   1275
         End
         Begin VB.CheckBox Check1 
            Caption         =   "GIS Filtered"
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
            Index           =   11
            Left            =   1080
            TabIndex        =   54
            Tag             =   "csvfile"
            Top             =   960
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            Caption         =   "UDP Output"
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
            TabIndex        =   53
            Tag             =   "csvfile"
            Top             =   240
            Value           =   1  'Checked
            Width           =   1335
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   6
            Left            =   1560
            TabIndex        =   52
            Tag             =   "Char"
            Text            =   ","
            Top             =   720
            Width           =   225
         End
         Begin VB.Label Label1 
            Caption         =   "Delimiter"
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
            Left            =   1800
            TabIndex        =   61
            Top             =   720
            Width           =   675
         End
      End
      Begin VB.Frame Frame10 
         Caption         =   "Output UDP"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2760
         TabIndex        =   46
         Top             =   2760
         Width           =   2535
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   4
            Left            =   120
            TabIndex        =   48
            Tag             =   "UdpOut"
            Text            =   "39421"
            Top             =   600
            Width           =   720
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   5
            Left            =   120
            TabIndex        =   47
            Tag             =   "UdpOut"
            Text            =   "127.0.0.1"
            Top             =   240
            Width           =   1755
         End
         Begin VB.Label Label1 
            Caption         =   "Port"
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
            Left            =   1080
            TabIndex        =   50
            Top             =   600
            Width           =   405
         End
         Begin VB.Label Label1 
            Caption         =   "Client"
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
            Left            =   1920
            TabIndex        =   49
            Top             =   240
            Width           =   495
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "Output File"
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
         TabIndex        =   43
         Top             =   2400
         Width           =   2535
         Begin VB.CommandButton cbNewFile 
            Caption         =   "Set Shell Cmd File"
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
            TabIndex        =   98
            ToolTipText     =   "Name of Shell Command File"
            Top             =   960
            Width           =   1390
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Shell"
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
            Left            =   1560
            TabIndex        =   97
            Tag             =   "csvfile"
            ToolTipText     =   "Tick to exexute command file each time output file is closed"
            Top             =   960
            Width           =   880
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Rollover"
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
            Index           =   22
            Left            =   1560
            TabIndex        =   69
            Tag             =   "csvfile"
            ToolTipText     =   "Roll Over File Name at midnight"
            Top             =   600
            Width           =   880
         End
         Begin VB.CommandButton cbNewFile 
            Caption         =   "Set Output File"
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
            TabIndex        =   44
            ToolTipText     =   "Name of Output File"
            Top             =   600
            Width           =   1390
         End
         Begin VB.Label lblOutputFile 
            Caption         =   "File Name"
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
            TabIndex        =   45
            Top             =   240
            Width           =   2295
         End
      End
      Begin VB.Frame Frame2 
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
         Height          =   1815
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   2535
         Begin VB.CheckBox Check1 
            Caption         =   "Head"
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
            Left            =   840
            TabIndex        =   66
            Tag             =   "csvfile"
            Top             =   720
            Width           =   735
         End
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
            Index           =   17
            Left            =   1320
            TabIndex        =   62
            Tag             =   "csvfile"
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton Option1 
            Caption         =   "CSV"
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
            TabIndex        =   41
            Top             =   720
            Value           =   -1  'True
            Width           =   720
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Tagged"
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
            TabIndex        =   40
            Top             =   960
            Width           =   945
         End
         Begin VB.OptionButton Option1 
            Caption         =   "NMEA"
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
            Left            =   120
            TabIndex        =   39
            Top             =   480
            Width           =   840
         End
         Begin VB.CheckBox Check1 
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
            Index           =   4
            Left            =   120
            TabIndex        =   38
            Tag             =   "csvfile"
            Top             =   1200
            Width           =   1200
         End
         Begin VB.CheckBox Check1 
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
            Index           =   5
            Left            =   120
            TabIndex        =   37
            Tag             =   "csvfile"
            Top             =   1440
            Width           =   1425
         End
         Begin VB.CheckBox Check1 
            Caption         =   "+Time Stamp"
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
            Left            =   1080
            TabIndex        =   36
            Tag             =   "csvfile"
            Top             =   480
            Width           =   1275
         End
         Begin VB.CheckBox Check1 
            Caption         =   "GIS Filtered"
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
            Left            =   1080
            TabIndex        =   35
            Tag             =   "csvfile"
            Top             =   960
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox Check1 
            Caption         =   "File Output"
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
            Index           =   14
            Left            =   120
            TabIndex        =   34
            Tag             =   "csvfile"
            Top             =   240
            Value           =   1  'Checked
            Width           =   1095
         End
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   2
            Left            =   1560
            TabIndex        =   33
            Tag             =   "Char"
            Text            =   ","
            Top             =   720
            Width           =   225
         End
         Begin VB.Label Label1 
            Caption         =   "Delimiter"
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
            Left            =   1800
            TabIndex        =   42
            Top             =   720
            Width           =   615
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Time Stamp Format"
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
         Index           =   11
         Left            =   120
         TabIndex        =   65
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame17 
      Caption         =   "All Settings"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   11280
      TabIndex        =   25
      Top             =   2760
      Width           =   1335
      Begin VB.CommandButton cbOpenNew 
         Caption         =   "Open New"
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
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton cbSave 
         Caption         =   "Save"
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
         Left            =   120
         TabIndex        =   27
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cbSaveAs 
         Caption         =   "Save As"
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
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   975
      End
   End
   Begin VB.Frame Frame16 
      Caption         =   "Tag Template File"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3840
      TabIndex        =   24
      Top             =   4320
      Width           =   1815
      Begin VB.CommandButton cbNewFile 
         Caption         =   "New File"
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
         Left            =   480
         TabIndex        =   29
         Top             =   480
         Width           =   855
      End
      Begin VB.Label lblTemplateFile 
         Caption         =   "File Name"
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
         TabIndex        =   30
         Top             =   240
         Width           =   1545
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "Input Log File"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   2040
      TabIndex        =   14
      Top             =   0
      Width           =   1695
      Begin VB.Frame Frame13 
         Caption         =   "Format"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   21
         Top             =   1560
         Width           =   1455
         Begin VB.OptionButton Option1 
            Caption         =   "NMEA+Time Stamp"
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
            Index           =   11
            Left            =   120
            TabIndex        =   23
            Top             =   480
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton Option1 
            Caption         =   "NMEA"
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
            Index           =   10
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Rollover"
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
         TabIndex        =   20
         Tag             =   "csvfile"
         Top             =   2520
         Width           =   960
      End
      Begin VB.Frame Frame5 
         Caption         =   "Source"
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
         TabIndex        =   15
         Top             =   240
         Width           =   1455
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
            TabIndex        =   19
            Top             =   240
            Width           =   1095
         End
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
            Index           =   1
            Left            =   120
            TabIndex        =   18
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Processed"
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
            TabIndex        =   17
            Top             =   720
            Value           =   -1  'True
            Width           =   1095
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
            Index           =   3
            Left            =   120
            TabIndex        =   16
            Top             =   960
            Width           =   1095
         End
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Output Tags and Range"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   8520
      TabIndex        =   11
      Top             =   0
      Width           =   4095
      Begin VB.CheckBox Check1 
         Caption         =   "Synchronise Input Filter to Tags"
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
         Left            =   240
         TabIndex        =   13
         Tag             =   "csvfile"
         Top             =   2400
         Width           =   2955
      End
      Begin MSFlexGridLib.MSFlexGrid TagList 
         Height          =   2175
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   3836
         _Version        =   393216
         Rows            =   8
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Output Fields and Tags"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   3840
      TabIndex        =   9
      Top             =   0
      Width           =   4575
      Begin MSFlexGridLib.MSFlexGrid FieldList 
         Height          =   2175
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   3836
         _Version        =   393216
         Rows            =   8
         FixedCols       =   0
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Other"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   11280
      TabIndex        =   7
      Top             =   4440
      Width           =   1335
      Begin VB.CheckBox Check1 
         Caption         =   "Licence"
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
         Index           =   23
         Left            =   120
         TabIndex        =   99
         Tag             =   "csvfile"
         ToolTipText     =   "Show Licence Details"
         Top             =   960
         Width           =   880
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Show Files"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   120
         TabIndex        =   87
         Tag             =   "csvfile"
         Top             =   600
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check for Updates"
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
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Tag             =   "csvfile"
         Top             =   240
         Width           =   1065
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Scheduler"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   3840
      TabIndex        =   2
      Top             =   2760
      Width           =   1815
      Begin VB.CheckBox Check1 
         Caption         =   "Output on MMSI change"
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
         Index           =   16
         Left            =   120
         TabIndex        =   73
         Tag             =   "csvfile"
         Top             =   840
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Text            =   "15"
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Text            =   "1"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Time to Live"
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
         Left            =   720
         TabIndex        =   6
         Top             =   600
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Minutes"
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
         Left            =   720
         TabIndex        =   4
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Input Filter"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   2880
      Width           =   3735
      Begin VB.CheckBox Check1 
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
         Height          =   255
         Index           =   24
         Left            =   2280
         TabIndex        =   100
         Tag             =   "csvfile"
         ToolTipText     =   "Show VDO as MyShip"
         Top             =   240
         Width           =   880
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Reject Shore Stations"
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
         Left            =   120
         TabIndex        =   72
         Tag             =   "csvfile"
         Top             =   240
         Width           =   2175
      End
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   3135
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   5530
         _Version        =   393217
         LabelEdit       =   1
         Style           =   2
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "TreeFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TmpFileCh As Long
Dim BadIniFile As Boolean
Dim BadModule As String
Dim ClickedNode As Node
Dim TemplateCh As Long
Dim NewTemplate As Boolean
Dim SecForm(2) As Form   '0=Blank,1=TreeFilter,2=NmeaRcv
Dim SecFile(2, 1) As Boolean    '(SecForm, 0=Required,1=Found)

Dim SecNodes(2, 1) As Boolean
Dim SecList(2, 1) As Boolean
Dim SecCheckBox(2, 1) As Boolean
Dim SecComboBox(2, 1) As Boolean
Dim SecOptionButton(2, 1) As Boolean
Dim SecTextLabel(2, 1) As Boolean
Dim SecFlexGrid(2, 1) As Boolean
Dim SecFlexgrid1(2, 1) As Boolean

Private Sub cbNewFile_Click(Index As Integer)
Dim Channel As Long

Select Case Index
Case Is = 0     'Output file Name (Same name for both UDP and File) only used for file
    For Channel = 1 To 2
        If TaggedOutputOn(Channel) = True Then
            If ChannelMethod(Channel) = "file" Then
                Call NmeaRcv.CloseOutputDataFile(Channel)   'write tag tail out to file or udp
            End If
            If ChannelMethod(Channel) = "udp" Then
                Call NmeaRcv.CloseOutputUdp(Channel)
            End If
        End If
    Next Channel
    OutputFileName = FileSelect.AskFileName("OutputFileName")
Case Is = 1     'tag template file - Used to output UDP as well as file
    For Channel = 1 To 2
        If TaggedOutputOn(Channel) = True Then
            If ChannelMethod(Channel) = "file" Then
                Call NmeaRcv.CloseOutputDataFile(Channel)   'write tag tail out to file or udp
            End If
            If ChannelMethod(Channel) = "udp" Then
                Call NmeaRcv.CloseOutputUdp(Channel)
            End If
        End If
    Next Channel
    TagTemplateReadFile = FileSelect.AskFileName("TagTemplateReadFile", True)
Case Is = 2     'Shell file - only channel 1 at the moment
    For Channel = 1 To 2
    Next Channel
    ShellFileName = FileSelect.AskFileName("ShellFileName", False)
End Select

Call SetOutputOptions   'moved from below 31/5/11

If Visible Then
    Call CheckOutputOptions
End If

End Sub


Private Sub Check1_Click(Index As Integer)
Select Case Index
Case Is = 1
    
    Call SynchroniseTags
'Case Is = 21    'cached vessels has been turned on or off
'    If Check1(Index).Value = True And CacheVessels = False Then
'        Call ReadVessels    'turned on get cached vessels
'    End If
End Select
If Visible Then
    Call SetOutputOptions
    Call CheckOutputOptions
End If
End Sub

'Does not cause any user interaction
Public Sub SetOutputOptions()
Dim Channel As Long
'NMEA, CSV or Tagged
With TreeFilter
'set up the Public arrays
'clear all first
    ChannelMethod(1) = "file"
    ChannelMethod(2) = "udp"
'check csv delimiter is not same as decimal separataor
'testing only MsgBox "Dot." & GetDecimalSep & "comma," & .Text1(2) & "|"
    If GetDecimalSep = .Text1(2) Then
        .Text1(2) = "|"
    End If
    If GetDecimalSep = .Text1(6) Then .Text1(6) = "|"
    CsvDelim(1) = .Text1(2)
    CsvDelim(2) = .Text1(6)
'speeds up output if both delimeters the same
    If CsvDelim(1) = CsvDelim(2) Then
        CsvDelim(0) = CsvDelim(1)
    Else
        CsvDelim(0) = ""
    End If
    For Channel = 0 To 2
'        ChannelOutputOk(Channel) = False
        ChannelOutput(Channel) = False
        NmeaOutput(Channel) = False
        CsvOutput(Channel) = False
        CsvAll(Channel) = False
        TagsReq(Channel) = False
        TaggedOutput(Channel) = False
        ShellOn(Channel) = False
        RangeReq(Channel) = False
        ScheduledReq(Channel) = False
        FenReq(Channel) = False
        GisReq(Channel) = False
        TimeStampReq(Channel) = False
'        UdpOutput(Channel) = False
'        FileOutput(Channel) = False
'moved to nmearcv
'        DisplayOutput(Channel) = False
        MethodOutput(Channel) = False
        FtpOutput(Channel) = False
    Next Channel
    
    For Channel = 1 To 2
        CsvHeadOn(Channel) = False
'clear csvhead (to force re-create) as tags may have been changed
        CsvHead(Channel) = ""
        ChannelFormat(Channel) = ""
'fix if French on initialisation
        Outputs(Channel).Tag = Channel
        ChannelEncoding(Channel) = "none"
        OutputFileRollover(Channel) = False
        NoDataOutput(Channel) = True
    Next Channel
            
    If .Check1(4).Value <> 0 Then ScheduledReq(1) = True
    If .Check1(6).Value <> 0 Then ScheduledReq(2) = True
    If .Check1(5).Value <> 0 Then RangeReq(1) = True
    If .Check1(7).Value <> 0 Then RangeReq(2) = True
    NmeaOutput(1) = Option1(8).Value
    NmeaOutput(2) = Option1(9).Value
    If .Check1(8).Value <> 0 Then TimeStampReq(1) = True
    If .Check1(9).Value <> 0 Then TimeStampReq(2) = True
    CsvOutput(1) = Option1(5).Value
    CsvOutput(2) = Option1(13).Value
    TaggedOutput(1) = Option1(6).Value
    TaggedOutput(2) = Option1(12).Value
    If .Check1(10).Value <> 0 Then GisReq(1) = True
    If .Check1(11).Value <> 0 Then GisReq(2) = True
    If .Check1(12).Value <> 0 Then ShellOn(1) = True
    If .Check1(14).Value <> 0 Then MethodOutput(1) = True
    If .Check1(15).Value <> 0 Then MethodOutput(2) = True
'    if check1(16).Value <> 0 Output Tags on MMSI change
'the option for ftp on channel 2 is disabled, but i've included
'it because if any channel could output to a file, it would be required
    If Check1(17).Value <> 0 Then FtpOutput(1) = True
    If Check1(18).Value <> 0 Then FtpOutput(2) = True
    If Check1(19).Value <> 0 Then CsvHeadOn(1) = True
    If Check1(20).Value <> 0 Then CsvHeadOn(2) = True
End With
   
'clear outputs if there is none
    For Channel = 1 To 2
'moved to nmearcv
'        If DisplayOutput(Channel) = False And MethodOutput(Channel) = False Then
        If MethodOutput(Channel) = False Then
            NmeaOutput(Channel) = False
            CsvOutput(Channel) = False
            TaggedOutput(Channel) = False
            ShellOn(Channel) = False
            ScheduledReq(Channel) = False
            RangeReq(Channel) = False
'Added for to enable gis output for Nmea
            GisReq(Channel) = False
        End If
        If CsvOutput(Channel) = False And TaggedOutput(Channel) = False Then
'Removed when we required Gis Output for Nmea as well
'            GisReq(Channel) = False
        End If
        If NmeaOutput(Channel) = False Then
            TimeStampReq(Channel) = False
        End If
'csv output required but no tags have been set
        If CsvOutput(Channel) = True _
        And TreeFilter.FieldList.TextMatrix(1, 0) = "" Then
            CsvAll(Channel) = True
            ReDim AllFields(200)    'reset
            CsvOutput(Channel) = False
        End If
'dont shell if no file being output
        If ChannelMethod(Channel) <> "file" Then
            ShellOn(Channel) = False
        End If
     Next Channel

'set the (0) subscript if either 1 or 2 is true
'kept separately to make the code clearer and quicker
    
    If NmeaOutput(1) = True Or NmeaOutput(2) = True Then NmeaOutput(0) = True
    If CsvOutput(1) = True Or CsvOutput(2) = True Then CsvOutput(0) = True
    If CsvAll(1) = True Or CsvAll(2) = True Then CsvAll(0) = True
    If TaggedOutput(1) = True Or TaggedOutput(2) = True Then TaggedOutput(0) = True
    If ShellOn(1) = True Or ShellOn(2) = True Then ShellOn(0) = True
    If RangeReq(1) = True Or RangeReq(2) = True Then RangeReq(0) = True
    If ScheduledReq(1) = True Or ScheduledReq(2) = True Then ScheduledReq(0) = True
    If FenReq(1) = True Or FenReq(2) = True Then FenReq(0) = True
    If GisReq(1) = True Or GisReq(2) = True Then GisReq(0) = True
    If TimeStampReq(1) = True Or TimeStampReq(2) = True Then TimeStampReq(0) = True
'    If UdpOutput(1) = True Or UdpOutput(2) = True Then UdpOutput(0) = True
'    If FileOutput(1) = True Or FileOutput(2) = True Then FileOutput(0) = True
'moved to nmearcv
'    If DisplayOutput(1) = True Or DisplayOutput(2) = True Then DisplayOutput(0) = True
    If MethodOutput(1) = True Or MethodOutput(2) = True Then MethodOutput(0) = True
    If FtpOutput(1) = True Or FtpOutput(2) = True Then FtpOutput(0) = True
    
'set any other 1 or 2
    For Channel = 1 To 2
        If NmeaOutput(Channel) = True Then ChannelFormat(Channel) = "nmea"
        If CsvOutput(Channel) = True Then ChannelFormat(Channel) = "csv"
        If CsvAll(Channel) = True Then ChannelFormat(Channel) = "csv"
        If TaggedOutput(Channel) = True Then
            ChannelFormat(Channel) = "tag"
'must be set before OutputFile
            TagTemplateReadFile = FileSelect.SetFileName("TagTemplateReadFile")
        End If
        Select Case ChannelMethod(Channel)
'set up default output filenames
        Case Is = "file"
'only rollover if file output and NOT scheduled
            If TreeFilter.Check1(22).Value <> 0 And ScheduledReq(Channel) = False Then
                OutputFileRollover(Channel) = True
            Else
                OutputFileRollover(Channel) = False
            End If
            OutputFileName = FileSelect.SetFileName("OutputFileName")
'This is done when the tag template file is actually read
'until then we don't even know if an overlay is required
'            If OverlayReq(Channel) Then
'                OverlayOutputFileName = FileSelect.SetFileName("OverlayOutputFileName")
'            End If
'set up FTP (Before File Names)
            FtpUserName = TreeFilter.Text1(8)
            FtpPassword = TreeFilter.Text1(9)
            FtpLocalFileName = FileSelect.SetFileName("FtpLocalFileName")

'Local must be set before remote to get default remote file name
'Clear remote name to ensurelocal file name (default) is used
            FtpRemoteFileName = ""
            FtpRemoteFileName = FileSelect.SetFileName("FtpRemoteFileName")
            Outputs(Channel).Caption = "File Output [" _
                & NameFromFullPath(OutputFileName)
            If MethodOutput(Channel) = True Then
                Outputs(Channel).Caption = _
                Outputs(Channel).Caption & "]"
            Else
                Outputs(Channel).Caption = _
                Outputs(Channel).Caption & "-display only]"
            End If
        Case Is = "udp"
            Outputs(Channel).Caption = "UDP Output [" _
                & TreeFilter.Text1(5) & ":" _
                & TreeFilter.Text1(4)
            If MethodOutput(Channel) = True Then
                Outputs(Channel).Caption = _
                Outputs(Channel).Caption & "]"
            Else
                Outputs(Channel).Caption = _
                Outputs(Channel).Caption & "-display only]"
            End If
        End Select
    Next Channel
'set any interdepandant variables
    For Channel = 0 To 2
'some output has been selected
'moved to nmearcv
'        If MethodOutput(Channel) = True Or DisplayOutput(Channel) = True Then
        If MethodOutput(Channel) = True Then
            ChannelOutput(Channel) = True
        End If
    Next Channel

'set the Program Flow Variables
    For Channel = 1 To 2
'if we need to set range tags
'if there are no tags actually setup its dumb to set up tagsreq
        If ChannelOutput(Channel) = True And _
            (RangeReq(Channel) = True Or _
            CsvOutput(Channel) = True Or _
            TaggedOutput(Channel) = True) Then
                TagsReq(Channel) = True
                TagsReq(0) = True
'the Ranger will not output the current message because it needs the
'next message to trigger the output on change of mmsi
            Else
'Direct Nmea output (with no range check) in ProcesssInput
            End If
 
 'Need Fen buffering unless nmea only output with no tags required
        If (TagsReq(Channel) = True _
        Or ScheduledReq(Channel) = True) And NmeaOutput(Channel) = True Then
            FenReq(Channel) = True
            FenReq(0) = True
        End If
If NmeaOutput(Channel) = True Then
    FenReq(Channel) = True
    FenReq(0) = True
End If
    Next Channel

If OutputFileName <> "" Then SpawnGisOk = True
'If Right$(OutputFileName, 4) = ".kml" Then
'    SpawnGisOk = True
'Else
'    SpawnGisOk = False
'End If
'wait until first output after an option change before allowing gis
NmeaRcv.cbSpawnGis.Enabled = False
DateTimeOutputFormat = TreeFilter.Text1(11).Text
'update files display
'Call Files.RefreshData

If cmdJna Then
    Call Testing.Variables
End If

End Sub

'I think this should be used when there may be user interaction
Sub CheckOutputOptions()
Dim Channel As Long
Dim FieldNo As Long
Dim TagNo As Long
Dim kb As String
Dim kb1 As String

'Lat and lon are now always got fom the sentence as it is received
'by DecodeSentence so temp setting GisTagOK
    
    DecimalSeparator = GetDecimalSeparator

#If False Then
'The slow way
    With TreeFilter.FieldList
        If .TextMatrix(1, 0) <> "" Then 'if only 2rows and first blank there are no entries
            For FieldNo = 1 To .Rows - 1
                With TreeFilter.TagList
                    For TagNo = 1 To .Rows - 1
                        If .TextMatrix(TagNo, 1) = TreeFilter.FieldList.TextMatrix(FieldNo, 12) Then
'got tagno for this Field
                            Select Case .TextMatrix(TagNo, 1)
                            Case Is = "lat"
'Tagno is 1 less than matrix line (see ResetTags)
                                GisLatTagNo = TagNo - 1
                            Case Is = "lon"
                                GisLonTagNo = TagNo - 1
                            End Select
                            Exit For
                        End If
                    Next TagNo
                End With
            Next FieldNo
        End If
    End With
#End If
    
'"lat" & "lon" tags are not named by user
'lat/lon is always dot, force lata/lon & range to locale
'because Csng is locale aware
'Allow them to be checked separately and allow =0 or =181 or 91 to be chosen by user
    ChkVesselLatMin = False
    ChkVesselLatMax = False
    ChkVesselLonMin = False
    ChkVesselLonMax = False
    For TagNo = 0 To UBound(TagArray)
        Select Case TagArray(TagNo, 0)
        Case Is = "lat"
            If IsNumeric(TagArray(TagNo, 3)) Then
                ChkVesselLatMin = True
                VesselLatMin = CSng(Replace(TagArray(TagNo, 3), ".", DecimalSeparator))
            End If
            If IsNumeric(TagArray(TagNo, 4)) Then
                ChkVesselLatMax = True
                VesselLatMax = CSng(Replace(TagArray(TagNo, 4), ".", DecimalSeparator))
            End If
        Case Is = "lon"
            If IsNumeric(TagArray(TagNo, 3)) Then
                ChkVesselLonMin = True
                VesselLonMin = CSng(Replace(TagArray(TagNo, 3), ".", DecimalSeparator))
            End If
            If IsNumeric(TagArray(TagNo, 4)) Then
                ChkVesselLonMax = True
                VesselLonMax = CSng(Replace(TagArray(TagNo, 4), ".", DecimalSeparator))
            End If
        End Select
    Next TagNo
        
'    GisTagOk = True

#If False Then
 'Debug.Print "CheckOutputOptions"
'check we have field set up if any gis filter
If GisReq(0) = True _
    And (TaggedOutput(0) = True Or CsvOutput(0) = True) Then
'Check if both lat and lon in field list
    With TreeFilter.FieldList
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 2) = "lat" Then LatFound = True
            If .TextMatrix(i, 2) = "lon" Then LonFound = True
        Next i
    End With
    If LatFound = True And LonFound = True Then
        GisTagOk = True
    Else
        GisTagOk = False
    End If
'output meaningful description
    If GisTagOk = False Then
        If cmdNoWindow = False Then TreeFilter.Show
        kb = "Geographic Information System (GIS)" & vbCrLf _
        & "Output Fields and Tags are required" & vbCrLf
        If LatFound = False Then
            kb = kb & vbTab & "Latitude" & vbCrLf
        End If
        If LonFound = False Then
            kb = kb & vbTab & "Longitude" & vbCrLf & "For" & vbCrLf
        End If
        For Channel = 1 To 2
            If GisReq(Channel) = True Then
                Select Case ChannelMethod(Channel)
                Case Is = "file"
                    kb1 = vbTab & "File"
                Case Is = "udp"
                    kb1 = vbTab & "UDP"
                End Select
                Select Case ChannelFormat(Channel)
                Case Is = "csv"
                kb = kb & kb1 & "-CSV Output" & vbCrLf
                Case Is = "tag"
                kb = kb & kb1 & "-Tagged Output" & vbCrLf
                End Select
            End If
        Next Channel
        MsgBox kb
    End If
End If
#End If

'check if output file requires closing only 1 channel at the moment
    If OutputFileRollover(1) = False And OutputFileDate <> "" _
Or OutputFileRollover(1) = True And OutputFileDate = "" Then 'dont rollover output file
        Call NmeaRcv.CloseOutputDataFile(1)    'write out No Data if not opened
    End If
    
    For Channel = 1 To 2

'Set up output encoding, do BEFORE ChannelMethod as we need the
'template file before output file
        Select Case ChannelFormat(Channel)
        Case Is = "tag"
'if file does not exist then ask
'Files.Show

            If FileExists(TagTemplateReadFile) = False Then
                TagTemplateReadFile = FileSelect.AskFileName("TagTemplateReadFile", True)
            End If
            Select Case LCase$(ExtFromFullName(TagTemplateReadFile))
            Case Is = "html", "htm", "xml", "aspx"
                ChannelEncoding(Channel) = "xml"
            Case Is = "kml", "kmz"
                ChannelEncoding(Channel) = "kml"
            Case Else
                ChannelEncoding(Channel) = "none"
            End Select
        Case Else
            ChannelEncoding(Channel) = "none"
        End Select
'Check required files are known
        
        Select Case ChannelMethod(Channel)
        Case Is = "file"
'if uncommented will ask for file name
''           outputfilename = FileSelect.AskFileName("OutputFileName")
'            Select Case ChannelFormat(Channel)
'           Case Is = "nmea"
'                OutputFileNameNmea = OutputFileName
'            Case Is = "csv"
'                OutputFileNameCsv = OutputFileName
'            Case Is = "tag"
'                OutputFileNameTagged = OutputFileName
'            End Select
'             = NameFromFullPath(OutputFileName)
            Outputs(Channel).Caption = "File Output [" _
                & NameFromFullPath(OutputFileName)
            If MethodOutput(Channel) = True Then
                Outputs(Channel).Caption = _
                Outputs(Channel).Caption & "]"
'If FolderExists(PathFromFullName(OutputFileName)) = False Then
'    MsgBox "Output File Directory does not exists " & vbCrLf _
'    & PathFromFullName(OutputFileName)
'End If
            Else
                Outputs(Channel).Caption = _
                Outputs(Channel).Caption & "-display only]"
            End If
        Case Is = "udp"
            Outputs(Channel).Caption = "UDP Output [" _
                & TreeFilter.Text1(5) & ":" _
                & TreeFilter.Text1(4)
            If MethodOutput(Channel) = True Then
                Outputs(Channel).Caption = _
                Outputs(Channel).Caption & "]"
            Else
                Outputs(Channel).Caption = _
                Outputs(Channel).Caption & "-display only]"
            End If
        End Select
    Next Channel

'set up tag template file display
If TagTemplateReadFile = "" Then
    TreeFilter.lblTemplateFile = ""
Else
    TreeFilter.lblTemplateFile = NameFromFullPath(TagTemplateReadFile)
'Call NmeaRcv.ReadTagTemplate    'zzz
End If

TreeFilter.Caption = "Options [" & NameFromFullPath(IniFileName) & "]"

FtpUserName = TreeFilter.Text1(8)
FtpPassword = TreeFilter.Text1(9)
'if been ticked (this time) re-read cache (only when changed)
If Check1(21).Value <> 0 Then
    If CacheVessels = False Then
Call WriteStartUpLog("ReadVessels")
        Call ReadVessels
Call WriteStartUpLog("Vessels Read")
        CacheVessels = True
    End If
Else
'dont write out if changed from cached to uncached
    CacheVessels = False
End If

'set initial scheduler frequency in seconds
'User sets minutes (including decimals)
On Error Resume Next
UserScheduledSecs = CInt(CSng(TreeFilter.Text1(0).Text) * 60)
On Error GoTo 0
If UserScheduledSecs = 0 Then
    ScheduledReq(0) = 0
    ScheduledReq(1) = 0
    ScheduledReq(2) = 0
End If
    Call NmeaRcv.ResetScheduledTimer(UserScheduledSecs)
'always reqset adjustment (don't know what else user has changed
'that affects upload speed)
ScheduledFreqAdj = 0

'set key len required for mmsi change key length to determine when to output scheduled message
If Check1(16).Value <> 0 Then
    SchedKeyLen = 9
Else
    SchedKeyLen = 20
End If
If Check1(13).Value <> 0 Then
    Files.Show
    Call Files.RefreshData
Else
    Files.Hide
End If

If Check1(23).Value <> 0 Then
    frmLicence.Show
Else
    Unload frmLicence
End If

If Check1(24).Value = vbUnchecked Then  'Dont show ownship
    Call ClearMyShip(MyShip)
End If

With frmFTP
    .txtUser.Text = FtpUserName
    .txtPW.Text = FtpPassword
    .txtHost.Text = HostFromPath(PathFromFullName(FtpRemoteFileName, "/"))
    .txtDir = FolderFromPath(PathFromFullName(FtpRemoteFileName, "/"))
    .txtRemote.Text = NameFromFullPath(FtpRemoteFileName, "/")
    .txtLocal.Text = FtpLocalFileName    'includes path
End With

'NmeaRcv.StatusBar.Panels(1) = IniFileName
'NmeaRcv.ClearStatusBarTimer.Enabled = True

If cmdJna Then
    Call Testing.Variables
End If
End Sub

Private Sub cbSaveAs_Click()    'save as
Dim NewFileName As String
'remove must exist flag
NewFileName = FileSelect.AskFileName("IniFileName", False)
Call SaveIniFile(NewFileName)
IniFileName = NewFileName
SetKeyValue HKEY_CURRENT_USER, "Software\Arundale\" & App.EXEName & "\Settings", "InitialisationFile", IniFileName, REG_SZ
TreeFilter.Caption = "Options [" & NameFromFullPath(NewFileName) & "]"
NmeaRcv.StatusBar.Panels(1) = IniFileName
NmeaRcv.ClearStatusBarTimer.Enabled = True
NmeaRcv.Caption = App.EXEName & " - Control/Stats [" & NameFromFullPath(IniFileName) & "]"
Unload Me
End Sub

Private Sub cbSave_Click()    'save
Call SaveIniFile(IniFileName)
End Sub

'only called when user asks for changes to be saved on exit
Public Sub ReplaceIniFile()
Dim reply As Integer

Call SaveIniFile(IniFileName & ".new")
If FileCompare(IniFileName & ".new", IniFileName) = False Then
    If ExitLogFileCh <> 0 Then Print #ExitLogFileCh, Now() & vbTab & "Files differ: " & IniFileName & " *.tmp"
    reply = MsgBox("Do you wish to save your changed settings ? ", vbYesNo, "Settings Changed")
    If reply = vbYes Then
        FileCopy IniFileName, IniFileName & ".old"
        If ExitLogFileCh <> 0 Then Print #ExitLogFileCh, Now() & vbTab & "File *.ini copied to .old"
        FileCopy IniFileName & ".new", IniFileName
        If ExitLogFileCh <> 0 Then Print #ExitLogFileCh, Now() & vbTab & "File *.new copied to .ini"
    End If
End If
'Kill IniFileName & ".new"
If ExitLogFileCh <> 0 Then Print #ExitLogFileCh, Now() & vbTab & "File: "; IniFileName & ".new - deleted"
End Sub

'NewFileName can be .tmp for file compare on exit
Private Sub SaveIniFile(NewFileName As String)
Dim i As Long
Dim arry() As String
Dim j As Integer
Dim Replace As Boolean  'true when replacing current settings
Dim kb As String
Dim CurrentForm As Form        'the current form that contains this control
Dim CurrentFormNo As Long
Dim Outbuf As String

'Dim ComboBoxSeen As Boolean 'found on ini file
'This is done so that the Add and Write routines can use the same code
'all you need to do is to set the form in the ini file
'prior to defining the controls

    Call CheckAll(TreeView1.Nodes("InputFilter"), "SaveIniFile")    'Ensure changes are correct
'if Copy not set then must walk second leg of tree explicitly
'Call CheckAll(TreeView1.Nodes("OutputFilter"))
    
'Ensure settings in Sockets(1).comms.baud are set as the list index
'otherwise any changes are not saved
    CurrentSocket = 1   'Must be set to serial socket
    Call SetCommComboBoxDefault(Combo1(0), Combo1(1))  'Form can differ eg Treefilter

'Ensure the encrypted FtpPassword has been created before the file is written out
    ValidateControls
    
    Call ResetSections  'reset the required sections of the .ini file
    
'open the current ini file which were using as the template
    IniFileCh = FreeFile
    Open IniFileName For Input As #IniFileCh
    TmpFileCh = FreeFile
    Open NewFileName & ".tmp" For Output As #TmpFileCh
        
    Do Until EOF(IniFileCh)
        Line Input #IniFileCh, kb
        Select Case kb
        Case Is = "[NODES]"
            Replace = True
            Call WriteNodes(TmpFileCh)
            SecNodes(CurrentFormNo, 1) = True
        Case Is = "[LIST]"
            Replace = True
            Call WriteList(TmpFileCh)
            SecList(CurrentFormNo, 1) = True
        Case Is = "[FILE]"
            Replace = True
            Call WriteFile(TmpFileCh)
            SecFile(CurrentFormNo, 1) = True
        Case Is = "[CHECKBOX]"
            Replace = True
            Call WriteCheckBox(CurrentForm, TmpFileCh)
            SecCheckBox(CurrentFormNo, 1) = True
        Case Is = "[COMBOBOX]"
            Replace = True
            Call WriteComboBox(CurrentForm, TmpFileCh)
            SecComboBox(CurrentFormNo, 1) = True
        Case Is = "[OPTIONBUTTON]"
            Replace = True
            Call WriteOptionButton(CurrentForm, TmpFileCh)
            SecOptionButton(CurrentFormNo, 1) = True
        Case Is = "[TEXTLABEL]"
            Replace = True
            Call WriteTextLabel(TmpFileCh)
            SecTextLabel(CurrentFormNo, 1) = True
        Case Is = "[FLEXGRID]"
            Replace = True
            Call WriteFlexGrid(CurrentForm, TmpFileCh)
            SecFlexGrid(CurrentFormNo, 1) = True
        Case Is = "[FLEXGRID1]"
            Replace = True
            Call WriteFlexGrid1(CurrentForm, TmpFileCh)
            SecFlexgrid1(CurrentFormNo, 1) = True
        Case Else
            If Left$(kb, 6) = "[FORM=" Then
                Call CheckForm(kb, CurrentForm, CurrentFormNo)
                Replace = False
            End If
        End Select
'write out any comments
        If Replace = False Then Print #TmpFileCh, kb
        Select Case kb
        Case Is = "[/NODES]"
            Replace = False
        Case Is = "[/LIST]"
            Replace = False
        Case Is = "[/FILE]"
            Replace = False
        Case Is = "[/CHECKBOX]"
            Replace = False
        Case Is = "[/COMBOBOX]"
            Replace = False
        Case Is = "[/OPTIONBUTTON]"
            Replace = False
        Case Is = "[/TEXTLABEL]"
            Replace = False
        Case Is = "[/FLEXGRID]"
            Replace = False
        Case Is = "[/FLEXGRID1]"
            Replace = False
        End Select
    Loop

'check to see if any sections are missing on the .ini file
'this is required as I may have added new sections to existing
'ini files which the user may have saved.
kb = ""
For i = 0 To 2
    If SecFile(i, 0) <> SecFile(i, 1) Then
        kb = kb & SecForm(i).Name & " File required=" & SecFile(i, 0) & vbCrLf
'        Call WriteForm(SecForm(i), TmpFileCh)  'no form
        Call WriteFile(TmpFileCh)
    End If
    If SecNodes(i, 0) <> SecNodes(i, 1) Then
        kb = kb & SecForm(i).Name & " Nodes required=" & SecNodes(i, 0) & vbCrLf
'        Call WriteForm(SecForm(i), TmpFileCh)
        Call WriteNodes(TmpFileCh)  'no form
    End If
    If SecList(i, 0) <> SecList(i, 1) Then
        kb = kb & SecForm(i).Name & " List required=" & SecList(i, 0) & vbCrLf
'        Call WriteForm(SecForm(i), TmpFileCh)
        Call WriteList(TmpFileCh)   'no form
    End If
    If SecCheckBox(i, 0) <> SecCheckBox(i, 1) Then
        kb = kb & SecForm(i).Name & " CheckBox required=" & SecCheckBox(i, 0) & vbCrLf
        Call WriteForm(SecForm(i), TmpFileCh)
        Call WriteCheckBox(SecForm(i), TmpFileCh)
    End If
    If SecComboBox(i, 0) <> SecComboBox(i, 1) Then
        kb = kb & SecForm(i).Name & " ComboBox required=" & SecComboBox(i, 0) & vbCrLf
        Call WriteForm(SecForm(i), TmpFileCh)
        Call WriteComboBox(SecForm(i), TmpFileCh)
    End If
    If SecOptionButton(i, 0) <> SecOptionButton(i, 1) Then
        kb = kb & SecForm(i).Name & " OptionButton required=" & SecOptionButton(i, 0) & vbCrLf
        Call WriteForm(SecForm(i), TmpFileCh)
        Call WriteOptionButton(SecForm(i), TmpFileCh)
    End If
    If SecTextLabel(i, 0) <> SecTextLabel(i, 1) Then
        kb = kb & SecForm(i).Name & " TextLabel required=" & SecTextLabel(i, 0) & vbCrLf
'        Call WriteForm(SecForm(i), TmpFileCh)
        Call WriteTextLabel(TmpFileCh)  'no form
    End If
    If SecFlexGrid(i, 0) <> SecFlexGrid(i, 1) Then
        kb = kb & SecForm(i).Name & " FlexGrid required=" & SecFlexGrid(i, 0) & vbCrLf
        Call WriteForm(SecForm(i), TmpFileCh)
        Call WriteFlexGrid(SecForm(i), TmpFileCh)
    End If
    If SecFlexgrid1(i, 0) <> SecFlexgrid1(i, 1) Then
        kb = kb & SecForm(i).Name & " Flexgrid1 required=" & SecFlexgrid1(i, 0) & vbCrLf
        Call WriteForm(SecForm(i), TmpFileCh)
        Call WriteFlexGrid1(SecForm(i), TmpFileCh)
    End If
Next i
Close #IniFileCh  'old file being used as template
Close #TmpFileCh
TmpFileCh = 0
FileCopy NewFileName & ".tmp", NewFileName
If ExitLogFileCh <> 0 Then Print #ExitLogFileCh, Now() & vbTab & "File *.tmp copied to .ini"
Kill NewFileName & ".tmp"
If ExitLogFileCh <> 0 Then Print #ExitLogFileCh, Now() & vbTab & "File: "; NewFileName & ".tmp - deleted"
Exit Sub

nofil:
    MsgBox "Can't Find initialisation File (" & IniFileName & ")" _
    & vbCrLf & "Not Saved"

End Sub

Private Sub WriteForm(NewForm As Form, ch As Long)
'Is Not Nothing doesn't work with forms (don't know why)
If NewForm Is Nothing Then
Else
    Print #ch, "[FORM=""" & NewForm.Name & """]"
End If
End Sub

Private Sub WriteNodes(och As Long)
Dim i As Long
Dim arry() As String
Print #och, "[NODES]"
For i = 1 To TreeView1.Nodes.Count
    ReDim arry(6)
    arry(0) = ""
    On Error Resume Next    'no parent
    arry(0) = TreeView1.Nodes(i).Parent.Key
    On Error GoTo 0
    arry(1) = tvwChild
    arry(2) = TreeView1.Nodes(i).Key
    arry(3) = TreeView1.Nodes(i).Text
    arry(4) = TreeView1.Nodes(i).Tag    'Filter name
'convert to true/false - international issue
    arry(5) = CInt(TreeView1.Nodes(i).Checked)
    Print #och, Join(arry, ",")
Next i
Print #och, "[/NODES]"
End Sub

Private Sub WriteList(och As Long)
Dim i As Long
Dim j As Integer
Dim arry() As String
Dim ListFilter As Variant

Print #och, "[LIST]"
For Each ListFilter In ListFilters
    For i = 0 To ListFilter.List1.ListCount - 1
        ReDim arry(3)
        If ListFilter.List1.List(i) <> "" Then
            arry(0) = j
            arry(1) = ListFilter.List1.List(i)
            arry(2) = ListFilter.Tag        'filter name
            Print #och, Join(arry, ",")
        End If
    Next i
j = j + 1
Next
Print #och, "[/LIST]"
End Sub

Private Sub WriteFile(och As Long)
Dim i As Long
Dim j As Integer
Dim arry() As String
Dim ListFilter As Variant


Print #och, "[FILE]"
ReDim arry(3)
arry(0) = "1"
arry(1) = DefaultUserFilename(TagTemplateReadFile)
arry(2) = "TagTemplateReadFile"
Print #och, Join(arry, ",")
arry(0) = "2"
arry(1) = DefaultUserFilename(OutputFileNameNmea)
arry(2) = "OutputFileNameNmea"
Print #och, Join(arry, ",")
arry(0) = "3"
arry(1) = DefaultUserFilename(OutputFileNameCsv)
arry(2) = "OutputFileNameCsv"
Print #och, Join(arry, ",")
arry(0) = "4"
arry(1) = DefaultUserFilename(OutputFileNameTagged)
arry(2) = "OutputFileNameTagged"
Print #och, Join(arry, ",")
arry(0) = "5"
arry(1) = DefaultUserFilename(OutputFileName)
arry(2) = "OutputFileName"
Print #och, Join(arry, ",")
arry(0) = "6"
arry(1) = DefaultUserFilename(ShellFileName)
arry(2) = "ShellFileName"
Print #och, Join(arry, ",")
arry(0) = "7"
arry(1) = DefaultUserFilename(NmeaReadFile)
arry(2) = "NmeaReadFile"
Print #och, Join(arry, ",")
Print #och, "[/FILE]"
End Sub

Function DefaultUserFilename(FileName As String) As String
'dont keep if user defaults files
If InStr(1, FileName, "\jna\") Then
    Select Case NameFromFullPath(IniFileName)
    Case Is = "default.ini", "udptagsrange.ini", "GoogleEarth.ini", "GoogleMaps.ini"
        DefaultUserFilename = NameFromFullPath(FileName)
    Case Else
        DefaultUserFilename = FileName
    End Select
Else
    DefaultUserFilename = FileName
End If
End Function
Sub WriteCheckBox(CurrentForm As Form, och As Long)
Dim i As Long
Dim j As Integer
Dim arry() As String
Dim kb As String

Print #och, "[CHECKBOX]"
On Error GoTo Skip_form
With CurrentForm
    For i = .Check1().LBound To .Check1().UBound
        ReDim arry(6)
        arry(0) = i
'if a control has been deleted on the form its index will be missing
        On Error GoTo Missing   'this (index) is missing
        arry(1) = .Check1(i).Value
        arry(2) = .Check1(i).Caption
'File will be different        If .Check1(i).Visible = False Then arry(3) = "Hide"
        arry(4) = .Check1(i).Tag
        arry(5) = .Check1(i).Name
'Force ~TCP name as may have been changed to Servr or Client and if not changed back
'will force Save Changes to be asked on exit
        If CurrentForm.Name = "NmeaRcv" And i = 5 Then .Check1(i).Caption = "TCP"
        Print #och, Join(arry, ",")

Missing:
    Next i
End With
Skip_form:
Print #och, "[/CHECKBOX]"
End Sub

Sub WriteComboBox(CurrentForm As Form, och As Long)
Dim i As Long
Dim j As Integer
Dim arry() As String
Dim kb As String

Print #och, "[COMBOBOX]"
On Error GoTo Skip_form
With CurrentForm
    For i = .Combo1().LBound To .Combo1().UBound
        ReDim arry(6)
        arry(0) = i
'if a control has been deleted on the form its index will be missing
        On Error GoTo Missing   'this (index) is missing
'If no ports on PC dont create blank combo box otherwise it will fail when loading
'with invalid port
        If .Combo1(i).Text <> "" Then   'No value don't create v137
            arry(1) = .Combo1(i).Text
'        Arry(2) = .Combo1(i).Caption 'not a property
'file will be diff        If .Combo1(i).Visible = False Then arry(3) = "Hide"
            arry(4) = .Combo1(i).Tag
            arry(5) = .Combo1(i).Name
            Print #och, Join(arry, ",")
        End If
Missing:
    Next i
End With
Skip_form:
Print #och, "[/COMBOBOX]"
End Sub

Sub WriteOptionButton(CurrentForm As Form, och As Long)
Dim i As Long
Dim j As Integer
Dim arry() As String

Print #och, "[OPTIONBUTTON]"
On Error GoTo Skip_form
    With CurrentForm
    For i = .Option1().LBound To .Option1().UBound
        ReDim arry(6)
        arry(0) = i
        On Error Resume Next
'convert to true/false - international issue
        arry(1) = CInt(.Option1(i).Value)
        arry(2) = .Option1(i).Caption
'file will be diff        If .Option1(i).Visible = False Then arry(3) = "Hide"
        arry(4) = .Option1(i).Tag
        arry(5) = .Option1(i).Name
        On Error GoTo Skip_form
        Print #och, Join(arry, ",")
    Next i
End With
Skip_form:
Print #och, "[/OPTIONBUTTON]"
End Sub

Sub WriteFlexGrid(CurrentForm As Form, och As Long)
Dim i As Long
Dim j As Integer
Dim arry() As String
Dim objFlexGrid As Object

Print #och, "[FLEXGRID]"
On Error GoTo Skip_form
Set objFlexGrid = CurrentForm.FieldList
    With objFlexGrid
    For i = 1 To .Rows - 1
        ReDim arry(.Cols)
        For j = 0 To .Cols - 1
            arry(j) = .TextMatrix(i, j)
        Next j
        Print #och, Join(arry, ",")
    Next i
End With
Skip_form:
Print #och, "[/FLEXGRID]"
End Sub

Sub WriteFlexGrid1(CurrentForm As Form, och As Long)
Dim i As Long
Dim j As Integer
Dim arry() As String
Dim objFlexGrid As Object

Print #och, "[FLEXGRID1]"
On Error GoTo Skip_form
Set objFlexGrid = CurrentForm.TagList
    With objFlexGrid
    For i = 1 To .Rows - 1
        ReDim arry(.Cols)
        For j = 0 To .Cols - 1
            arry(j) = .TextMatrix(i, j)
        Next j
        Print #och, Join(arry, ",")
    Next i
End With
Skip_form:
Print #och, "[/FLEXGRID1]"
End Sub

Sub WriteTextLabel(och As Long)
Dim i As Long
Dim j As Integer
Dim arry() As String
#If False Then
    Dim oFTP As New FTP
#End If

Print #och, "[TEXTLABEL]"
On Error GoTo Skip_form
    With TreeFilter
    For i = .Text1().LBound To .Text1().UBound
        ReDim arry(4)
        arry(0) = i
'force tag name for compatibility with earlier .ini files
        If Text1(i).Text = "," Then Text1(i).Tag = "Char"
        If Text1(i).Text = "," Then 'cant have , within field
            arry(1) = ""
        Else
            arry(1) = .Text1(i).Text
        End If
'if french replace comma with dot
        If .Text1(i).Tag = "Numeric" Then
            If GetDecimalSep() = "," Then
                arry(1) = Replace(arry(1), ",", ".")
            End If
        End If
        If .Text1(i).Tag = "encrypt" Then
'MsgBox .Text1(i).Text
'Because we are writing back the ini file (may be a Save)
'Replace the Ini password
'            If FtpPasswordEncrypted <> FtpIniPasswordEncrypted Then
                FtpIniPasswordEncrypted = FtpPasswordEncrypted
                FtpIniPasswordDecrypted = FtpPassword
'            End If
            arry(1) = FtpPasswordEncrypted   'Written back to the .ini file
        End If
        If HasIndex(.Label1, i) Then
            arry(2) = .Label1(i).Caption
        Else
            arry(2) = "No Label"           'may have no caption (TCP host)
        End If
        arry(3) = .Text1(i).Tag
'DONT use join as encrypt over 100 cgrs wraps encrypted string
'            Print #och, Join(arry, ",")
        Print #och, arry(0) & "," & arry(1) & "," _
        & arry(2) & "," & arry(3) & ","
    Next i
End With
Skip_form:
Print #och, "[/TEXTLABEL]"
#If False Then
    Set oFTP = Nothing
#End If

End Sub

Private Sub cbOpenNew_Click()    'read
If InputState <> 0 Then NmeaRcv.cbStop.Value = True
IniFileName = FileSelect.AskFileName("IniFileName", True)   'file must exist
Call HideTreeFilter 'also Hides ListFilters
Call ReadIniFile    'to prevent reporting OutputOptions errors
Call SetOutputOptions   'loads the options when the file is fully loafed
TreeFilter.Show     'allow reopting error
Call CheckOutputOptions
NmeaRcv.Caption = App.EXEName & " - Control/Stats [" & NameFromFullPath(IniFileName) & "]"
End Sub

Private Sub Combo1_Change(Index As Integer)
If Visible Then
    Call SetOutputOptions
    Call CheckOutputOptions
End If

End Sub

Private Sub FieldList_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim i As Integer
If Button = vbLeftButton Then
    With FieldList
        .RowSel = .Row    'force selection to one row
        .ColSel = .col
        
        If .TextMatrix(.Row, 12) <> "" Then
                For i = 0 To .Cols - 1
                .col = i
                .CellBackColor = vbRed
            Next
'           .ColSel = .Col
            If MsgBox("Remove field from Output List ? ", vbOKCancel, "Delete") = vbOK Then
                If .Rows = 2 Then    'add blank row
                    .AddItem ""
                End If
                .RemoveItem .Row
                Call ResetTags("FieldList")
            Else        'remove red
            For i = 0 To .Cols - 1
                .col = i
                .CellBackColor = vbWhite
            Next
            End If
            .ColSel = .col
        End If
    End With
End If
Call SetOutputOptions   'to check if lat lon ok
Call CheckOutputOptions
End Sub

Private Sub Form_Load()
Dim reply As Long
Dim i As Long
Dim kb As String

With FieldList
'    .Cols = 9
    .FormatString = "<MsgKey|<Source|<Member|^From|^ReqBits|^Arg|^Arg1|^Column|^ Msg |^DAC|^FI|^ID|<Tag|<Name|<Value"
    .ColWidth(0) = 0 '   1500    'msgkey
    .ColWidth(1) = 0 '   500    'source
    .ColWidth(2) = 0 '   1500   'member
    .ColWidth(3) = 0 '   500   'from
    .ColWidth(4) = 0 '   500   'reqbits
    .ColWidth(5) = 0 '   500   'arg
    .ColWidth(6) = 0 '   500   'arg1
    .ColWidth(7) = 0 '   500    'column

    .ColWidth(8) = 800    'msgtype   'displayed fields
'    .ColWidth(9) = 0    'dac
'    .ColWidth(10) = 0   'fi
'    .ColWidth(11) = 0   'fiid
    .ColWidth(12) = 1500 '800    'TAG
    .ColWidth(13) = 3000    'valdes - name
    .ColWidth(14) = 0    'value probably not required

End With

With TagList
    .FormatString = "<No|<Tag|>Min|>Max|<Name"
    .ColWidth(0) = 300
    .ColWidth(1) = 1200
    .ColWidth(2) = 560
    .ColWidth(3) = 500
    .ColWidth(4) = 1300
End With

Call ScanCommPorts 'Available Ports and Speeds into Combo1 on TreeFilter before
                    'trying to load the .ini file - also causes NmeaRcv to be loaded

Call ResetSections

Call ReadIniFile        'NmeaRcv is already loaded by ScanCommPorts

'Call DisplayForms("After Reading Ini in TreeFilter")
Call SetOutputOptions   'loads the options when the file is fully loafed
'Call DisplayForms("After SetOutputOptions in TreeFilter")
Call CheckOutputOptions
'Call DisplayForms("After CheckOutputOptions in TreeFilter")
'done when TreeFilter is checked If TaggedOutput(0) = True Then Call NmeaRcv.ReadTagTemplate(TagTemplateReadFile)  'must be after inifile is read
'When the tag template is read the Overlay file name is established
'so refresh the files list
'Call Files.RefreshData
List.Hide
End Sub

'Called when AisDecoder is started, to set up the available Ports and speeds
'Called again when cbOptions is selected, incase serail device has been added since
'AisDecoder was started
Public Sub ScanCommPorts()
Dim Ports() As String
Dim k As Long
Dim PortName As String
Dim PortIndex As Long
Dim SpeedIndex As Long

'WriteLog "ScanCommPorts", LogForm

'You have to keep the existing index if called when tree filter is called
    If SerialPortCount > 0 Then
        On Error Resume Next
        PortName = Combo1(0).Text
        PortIndex = Combo1(0).ListIndex
        SpeedIndex = Combo1(1).ListIndex
        On Error GoTo 0
    End If
'WriteLog "ScanCommPorts-Erase", LogForm
    
    Erase Ports
    With Combo1(0)
        Do While .ListCount > 0
            .RemoveItem (.ListCount - 1) 'zero based
        Loop
    End With
    
    With Combo1(1)
        Do While .ListCount > 0
            .RemoveItem (.ListCount - 1) 'zero based
        Loop
    End With

'WriteLog "ScanCommPorts=GetSerialPorts", LogForm

    Ports = GetSerialPorts      'registry method
    SerialPortCount = GetSerialPortCount(Ports)
'May be no ports
'    On Error GoTo GotaPort
'    SerialPortCount = 1 + UBound(Ports) 'zero based
'GotaPort:     On Error GoTo 0
    Call WriteStartUpLog(SerialPortCount & " Serial ports found")
'0 = NO portS
    If SerialPortCount = 0 Then
WriteLog "No PC serial ports found"

'These must be set false to stop AddComboBox trying to Add an nonexistant
'Setting
        Combo1(0).Enabled = False  'serial port
        Combo1(0).BackColor = Frame9.BackColor  'change from white
        lblCombo1(0).Enabled = False
        Combo1(1).Enabled = False  'serial speed
        Combo1(1).BackColor = Frame9.BackColor
        lblCombo1(1).Enabled = False
        Frame19.Enabled = False
'Must prevent user ticking Serial Input
        NmeaRcv.Check1(1).Value = 0
        NmeaRcv.Check1(1).Enabled = False
        Combo1(0).ListIndex = -1    'None
        Combo1(1).ListIndex = -1    'None
    
    Else
'WriteLog SerialPortCount & " serial ports found", LogForm
'Me.Visible = True   'Debug
        For k = 0 To UBound(Ports)
            Combo1(0).AddItem Ports(k)
'Debug.Print Combo1(0).List(k)  'debug
            If Combo1(0).List(k) = PortName Then
                Combo1(0).ListIndex = k  'Comm port still exists
            End If
        Next k
'WriteLog "ScanCommPorts-SerialPortCount, chk 1", LogForm
        
        If Combo1(0).ListIndex = -1 Then    'Previous selection not found
            Combo1(0).ListIndex = Combo1(0).ListCount - 1 'Last serial port added
        End If
'WriteLog "ScanCommPorts-SerialPortCount, chk 2", LogForm
        Combo1(0).Enabled = True  'serial port
        Combo1(0).BackColor = vbWhite  'change to white
        lblCombo1(0).Enabled = True
'WriteLog "ScanCommPorts-SerialPortCount, chk 3", LogForm
        
        Combo1(1).Enabled = True  'serial speed
        Combo1(1).BackColor = vbWhite
        lblCombo1(1).Enabled = True
        Frame19.Enabled = True
'Dont change users previous selection
'WriteLog "ScanCommPorts-SerialPortCount, chk 4", LogForm
        
        NmeaRcv.Check1(1).Enabled = True
        With Combo1(1)
            .AddItem "1200"
            .AddItem "2400"
            .AddItem "4800"
            .AddItem "9600"
            .AddItem "19200"
            .AddItem "38400"
            If SpeedIndex <> -1 Then
                .ListIndex = SpeedIndex  'previous selection
            End If
        End With

    End If
'WriteLog "ScanCommPorts-SerialPortCount, end if", LogForm

'After setting available ports on ComboBox then set the default port/speed
'The argument is the cboObject as it could be Commcfg(NmeaRouter) or treefilter(AisDecoder)
'which have different combo boxes
    Call SetCommComboBoxDefault(Combo1(0), Combo1(1))  'Form can differ eg Treefilter

End Sub

Sub ResetSections()
Dim i As Long
Dim j As Long

'load if whether or not various sections are required
    Set SecForm(0) = Nothing
    Set SecForm(1) = TreeFilter
    Set SecForm(2) = NmeaRcv
    Erase SecNodes
    Erase SecList
    Erase SecCheckBox
    Erase SecComboBox
    Erase SecOptionButton
    Erase SecTextLabel
    Erase SecFlexGrid
    Erase SecFlexgrid1
    
'(SecFrom, 0=Required,1=Found)
    SecFile(0, 0) = True
    SecNodes(0, 0) = True   'force treefilter
    SecList(0, 0) = True    'force treefilter
    SecCheckBox(1, 0) = True
    SecCheckBox(2, 0) = True
    SecComboBox(1, 0) = True
    SecComboBox(2, 0) = True
    SecOptionButton(1, 0) = True
    SecOptionButton(2, 0) = True
    SecTextLabel(1, 0) = True
    SecFlexGrid(1, 0) = True
    SecFlexgrid1(1, 0) = True

End Sub

Private Function ReadIniFile()
Dim i As Long
Dim kb As String
Dim InputNodesOn As Boolean
Dim MmsisOn As Boolean
Dim CheckBoxOn As Boolean
Dim ComboBoxOn As Boolean
Dim OptionButtonOn As Boolean
Dim TextLabelOn As Boolean
Dim FlexGridOn As Boolean
Dim FlexGrid1On As Boolean
Dim FormOn As Boolean
Dim FileOn As Boolean
Dim TemplateOn As Boolean
Dim arry() As String
Dim colFld As New Collection
Dim colAisMsg As New Collection
Dim colView As New Collection
Dim vFld As Variant
Dim vAisMsg As Variant
Dim vView As Variant
Dim AisMsgArry() As String
Dim ViewArry() As String
Dim Control As String
Dim ListFilter As Variant
Dim CurrentForm As Form        'the current form that contains this control
Dim CurrentFormNo As Long   '0=none,1=TreeFilter,2=NmeaRcv
'This is done so that the Add and Write routines can use the same code
'all you need to do is to set the form in the ini file
'prior to defining the controls
BadIniFile = False
BadModule = "ReadIniFile"
TreeView1.Nodes.Clear
Call ResetSections  'reset the required sections of the .ini file

With TagList    'clear all rows in existinf Grids
    .AddItem "", .FixedRows 'add a blank row as the first non-fixed
    For i = .FixedRows + 1 To .Rows - 1
        .RemoveItem (.FixedRows + 1)
    Next i
End With
With FieldList    'clear all rows in existinf Grids
    .AddItem "", .FixedRows 'add a blank row as the first non-fixed
    For i = .FixedRows + 1 To .Rows - 1
        .RemoveItem (.FixedRows + 1)
    Next i
End With

For Each ListFilter In ListFilters
    If Not ListFilter Is Nothing Then
        kb = ListFilter.List1.Name
        ListFilter.List1.Clear
    End If
Next
ListFilters(0).Caption = "From MMSI Filter"
ListFilters(0).Label1 = "MMSI's of vessels sending the messages you wish to accept"
ListFilters(0).Tag = "AisMsgFromMmsi"
ListFilters(1).Caption = "To MMSI Filter"
ListFilters(1).Label1 = "MMSI's of vessels receiving the messages you wish to accept"
ListFilters(1).Tag = "AisMsgToMmsi"
ListFilters(2).Caption = "DAC Filter"
ListFilters(2).Label1 = "Enter DAC"
ListFilters(2).Tag = "DacList"
ListFilters(3).Caption = "FI Filter"
ListFilters(3).Label1 = "Enter FI"
ListFilters(3).Tag = "FiList"
ListFilters(4).Caption = "FIID Filter"
ListFilters(4).Label1 = "Enter FIID"
ListFilters(4).Tag = "AisMsgDacFiId"


' Open the file & create nodes in the edit collection
If GetAttr(IniFileName) And vbReadOnly Then
    TreeFilter.Caption = "Options [" & NameFromFullPath(IniFileName) & " - Read Only]"
    TreeFilter.cbSave.Enabled = False
Else
    TreeFilter.Caption = "Options [" & NameFromFullPath(IniFileName) & "]"
    TreeFilter.cbSave.Enabled = True
End If
'open the current ini file which were using as the template
'Visible = True  'debug display as tree filter is constructed
IniFileCh = FreeFile
Open IniFileName For Input As #IniFileCh
Do Until EOF(IniFileCh)
    Line Input #IniFileCh, kb
'replace !AIVDM with AIS v90
    If InputNodesOn Then
        kb = Replace(kb, "!AIVDM", "AIS")
    End If
'ignore blank line
    If kb <> "" Then
        arry() = Split(kb, ",")
'If arry(0) = "NMEA" Then Stop
        Select Case arry(0)
        Case Is = "[NODES]"
            InputNodesOn = True
            SecNodes(CurrentFormNo, 1) = True
        Case Is = "[/NODES]"
            InputNodesOn = False
        Case Is = "[LIST]"
            MmsisOn = True
            SecList(CurrentFormNo, 1) = True
        Case Is = "[/LIST]"
            MmsisOn = False
        Case Is = "[FILE]"
            FileOn = True
            SecFile(CurrentFormNo, 1) = True
        Case Is = "[/FILE]"
            FileOn = False
        Case Is = "[TEMPLATE]"
            NewTemplate = True
        Case Is = "[CHECKBOX]"
            CheckBoxOn = True
            SecCheckBox(CurrentFormNo, 1) = True
        Case Is = "[/CHECKBOX]"
            CheckBoxOn = False
        Case Is = "[COMBOBOX]"
            ComboBoxOn = True
            SecComboBox(CurrentFormNo, 1) = True
        Case Is = "[/COMBOBOX]"
            ComboBoxOn = False
        Case Is = "[OPTIONBUTTON]"
            OptionButtonOn = True
            SecOptionButton(CurrentFormNo, 1) = True
        Case Is = "[/OPTIONBUTTON]"
            OptionButtonOn = False
        Case Is = "[TEXTLABEL]"
            TextLabelOn = True
            SecTextLabel(CurrentFormNo, 1) = True
        Case Is = "[/TEXTLABEL]"
            TextLabelOn = False
        Case Is = "[FLEXGRID]"
            FlexGridOn = True
            SecFlexGrid(CurrentFormNo, 1) = True
        Case Is = "[/FLEXGRID]"
            FlexGridOn = False
        Case Is = "[FLEXGRID1]"
            FlexGrid1On = True
            SecFlexgrid1(CurrentFormNo, 1) = True
        Case Is = "[/FLEXGRID1]"
            FlexGrid1On = False
        Case Else
            If Left$(kb, 1) <> "[" Then
                If InputNodesOn = True Then
                    Call AddNode(arry())
                End If
                If MmsisOn = True Then Call AddMmsi(arry())
                If FileOn = True Then Call AddFile(arry())
                If CheckBoxOn = True Then Call AddCheckBox(CurrentForm, arry())
                If ComboBoxOn = True Then Call AddComboBox(CurrentForm, arry())
                If OptionButtonOn = True Then Call AddOptionButton(CurrentForm, arry())
                If TextLabelOn = True Then Call AddTextLabel(arry())
                If FlexGridOn = True Then Call AddFlexGrid(CurrentForm, arry())
                If FlexGrid1On = True Then Call AddFlexGrid1(CurrentForm, arry())
            Else
                Control = Mid$(kb, 1, InStr(1, kb, "]"))
                Select Case Control
                Case Is = "[STACK Fld]"
                    colFld.Add (Mid$(kb, 12, Len(kb) - 11))
                Case Is = "[STACK AisMsg]"
                    colAisMsg.Add (Mid$(kb, 15, Len(kb) - 14))
                Case Is = "[STACK View]"
                    colView.Add (Mid$(kb, 13, Len(kb) - 12))
                Case Is = "[STACKOUT_old]"
                    For Each vFld In colFld
                        arry() = Split(vFld, ",")
                        arry(2) = "Fld" & arry(4)
                        Call AddNode(arry)
                        For Each vAisMsg In colAisMsg
                            AisMsgArry() = Split(vFld, ",")
                            AisMsgArry(0) = arry(2)
'                           AisMsgArry(2)=
                        Next
                    Next
                    Case Else
'Trap the form instruction which sets the parent form until reset
'with another form
                    If Left$(kb, 6) = "[FORM=" Then Call CheckForm(kb, CurrentForm, CurrentFormNo)
                End Select
            End If
        End Select
    Else    'Is a blank line
'Stop
    End If  'Not a blank line
Loop
Close #IniFileCh
IniFileCh = 0
SetKeyValue HKEY_CURRENT_USER, "Software\Arundale\" & App.EXEName & "\Settings", "InitialisationFile", IniFileName, REG_SZ
'finished adding to edit forms

'Check weve got all required sections defined - debug purposes only
'(SecFrom, 0=Required,1=Found)
kb = ""
For i = 0 To 2
    If SecFile(i, 0) <> SecFile(i, 1) Then kb = kb & SecForm(i).Name & " File required=" & SecFile(i, 0) & vbCrLf
    If SecNodes(i, 0) <> SecNodes(i, 1) Then kb = kb & SecForm(i).Name & " Nodes required=" & SecNodes(i, 0) & vbCrLf
    If SecList(i, 0) <> SecList(i, 1) Then kb = kb & SecForm(i).Name & " List required=" & SecList(i, 0) & vbCrLf
    If SecCheckBox(i, 0) <> SecCheckBox(i, 1) Then kb = kb & SecForm(i).Name & " CheckBox required=" & SecCheckBox(i, 0) & vbCrLf
    If SecComboBox(i, 0) <> SecComboBox(i, 1) Then kb = kb & SecForm(i).Name & " ComboBox required=" & SecComboBox(i, 0) & vbCrLf
    If SecOptionButton(i, 0) <> SecOptionButton(i, 1) Then kb = kb & SecForm(i).Name & " OptionButton required=" & SecOptionButton(i, 0) & vbCrLf
    If SecTextLabel(i, 0) <> SecTextLabel(i, 1) Then kb = kb & SecForm(i).Name & " TextLabel required=" & SecTextLabel(i, 0) & vbCrLf
    If SecFlexGrid(i, 0) <> SecFlexGrid(i, 1) Then kb = kb & SecForm(i).Name & " FlexGrid required=" & SecFlexGrid(i, 0) & vbCrLf
    If SecFlexgrid1(i, 0) <> SecFlexgrid1(i, 1) Then kb = kb & SecForm(i).Name & " Flexgrid1 required=" & SecFlexgrid1(i, 0) & vbCrLf
Next i

If BadIniFile = False Then BadModule = "ReadIniFile"
On Error GoTo BadFil
'Check if IniFile is valid
Call CheckAll(TreeView1.Nodes("InputFilter"), "ReadIniFile")
 'Debug.Print "--End--"
'debug NmeaRcv.Show
On Error GoTo 0
'Call DisplayForms("After CheckAll in ReadIniFile in TreeFilter")
Call ResetTags("FieldList") 'set TagList to FieldList
Call ClearMyShip(MyShip)
'reading the ini file caused these forms to be shown
'bug with udp not binding if done when server is started
ReceivedData.Hide
List.Hide
If BadIniFile Then GoTo BadFil

'Is there an embedded template in the .ini file
'DONT THINK THIS IS USED ANY MORE (NEWTEMPLATE NOTIN ANY INI FILES)
If NewTemplate = True Then
End If
IniFileCh = 0   'Finished loading (used to supress synchronise tags when loading)
'Must be called after form is loaded (will not synchronise if not checked)
Call TreeFilter.SynchroniseTags

'Call DisplayTree("ReadIniFile")    'debug
Exit Function

BadFil:
    MsgBox "Initialisation File (" & IniFileName & ")" & vbCrLf _
    & "has an old or incorrect format " _
    & "in module " & BadModule & vbCrLf & vbCrLf _
    & "Please SAVE this file (Options > Save), " _
    & "Exit the Decoder & Re-start." & vbCrLf
    IniFileCh = 0
End Function

Sub CheckForm(kb As String, CurrentForm As Form, CurrentFormNo As Long)
                If Left$(kb, 6) = "[FORM=" Then
                    Select Case Mid$(kb, 8, Len(kb) - 9)
                    Case Is = "TreeFilter"
                        Set CurrentForm = TreeFilter
                        CurrentFormNo = 1
                    Case Is = "NmeaRcv"
                        Set CurrentForm = NmeaRcv
                        CurrentFormNo = 2
                    End Select
'CurrentForm.Hide
'CurrentForm.Show   'debug only
                End If

End Sub
Sub AddNode(arry() As String)
Dim nodX As Node
'If arry(0) = "AIS" Then Stop
        On Error GoTo DupKey:
        Select Case arry(0)
        Case Is = ""        'root
            Set nodX = TreeView1.Nodes.Add(, , arry(2), arry(3))
            TreeView1.SingleSel = True ' to view the children
            nodX.EnsureVisible
       Case Else
            Set nodX = TreeView1.Nodes.Add(arry(0), arry(1), arry(2), arry(3))
        End Select
        nodX.Tag = arry(4)
'internationalisation cstr(true) returns actual language
        If StringToTrueFalse(arry(5)) = True Then
            nodX.Checked = True
        Else
            nodX.Checked = False
        End If
Exit Sub

DupKey:
On Error GoTo 0
'MsgBox "Duplicate Key" & vbCrLf & arry(2) & vbCrLf & "Will be ignored" & vbCrLf
End Sub

Function StringToTrueFalse(kb As String) As Boolean
If kb = CStr(True) Then StringToTrueFalse = True
'initial files distributed are in english
'and may not have been used/converted
'check if works in German
'If kb = "Wahr" Then StringToTrueFalse = True
If IsNumeric(kb) Then
    If CBool(kb) = True Then StringToTrueFalse = True
End If
End Function

Sub AddMmsi(arry() As String)
Dim i As Integer
BadModule = "AddMmsi"
On Error GoTo Exit_err
ListFilters(arry(0)).List1.AddItem arry(1)
Exit Sub

Exit_err:
BadIniFile = True
End Sub

Sub AddFile(arry() As String)
BadModule = "AddFile"
On Error GoTo Exit_err
Select Case arry(2)
Case Is = "TagTemplateReadFile"
    TagTemplateReadFile = arry(1)
Case Is = "OutputFileNameNmea"
    OutputFileNameNmea = arry(1)
Case Is = "OutputFileNameCsv"
    OutputFileNameCsv = arry(1)
Case Is = "OutputFileNameTagged"
    OutputFileNameTagged = arry(1)
Case Is = "OutputFileName"
    OutputFileName = arry(1)
Case Is = "ShellFileName"
    ShellFileName = arry(1)
Case Is = "NmeaReadFile"
    NmeaReadFile = arry(1)
End Select
Exit Sub

Exit_err:
BadIniFile = True
End Sub

Sub AddCheckBox(CurrentForm As Form, arry() As String)
Dim i As Integer
'When check1(i).value is set it fires the Check1_Click event on the current form
'eg CurrentForm=NmeaRcv then NmeaRcv.Check1(0)_Click is fired, this sets
'ReceivedData.Show or ReceivedData.Hide
'This in turn fires the Load event ReceivedData.Form_Load
'which sets ReceivedData.Hide hence Received Data is never shown on first load
'as it is overridden by the load event.
'But because the form is never unloaded, any subsequent change to
'NmeaRcv.Check1(0) will still force the ReceveidData.Show or Hide to be actioned.
'The same proceedure is used to show/hide List and Detail
i = arry(0)
On Error GoTo Exit_err
With CurrentForm
    If i >= .Check1().LBound And i <= .Check1().UBound Then
        .Check1(i).Value = arry(1)
'        .Check1(i).Caption = Arry(2)
        If arry(3) = "Hide" Then .Check1(i).Visible = False
        .Check1(i).Tag = arry(4)
    End If
End With
'Call DisplayForms("After AddCheckBox in ReadIniFile in TreeFilter")
Exit Sub

Exit_err:
BadModule = "AddCheckBox (" & arry(5) & ")"
BadIniFile = True
End Sub

Sub AddComboBox(CurrentForm As Form, arry() As String)
Dim i As Integer
'When COMBO1(i).value is set it fires the COMBO1_Click event on the current form
'eg CurrentForm=NmeaRcv then NmeaRcv.COMBO1(0)_Click is fired, this sets
'ReceivedData.Show or ReceivedData.Hide
'This in turn fires the Load event ReceivedData.Form_Load
'which sets ReceivedData.Hide hence Received Data is never shown on first load
'as it is overridden by the load event.
'But because the form is never unloaded, any subsequent change to
'NmeaRcv.COMBO1(0) will still force the ReceveidData.Show or Hide to be actioned.
'The same proceedure is used to show/hide List and Detail
i = arry(0)
On Error GoTo Exit_err
With CurrentForm
    If i >= .Combo1().LBound And i <= .Combo1().UBound Then
        On Error GoTo SkipChoice
        If .Combo1(i).Enabled = True Then
            .Combo1(i).Text = arry(1)   'error if we try to pre-select a choice which does
                                        'not exists. Eg .ini file specifies com port which
                                        'does not exist on this PC
'            .Check1(i).Caption = Arry(2)
            If arry(3) = "Hide" Then
                .Combo1(i).Visible = False
            End If
            .Combo1(i).Tag = arry(4)
        End If
SkipChoice:
        On Error GoTo Exit_err
    End If
End With
'Call DisplayForms("After AddComboBox in ReadIniFile in TreeFilter")
Exit Sub

Exit_err:
BadModule = "AddComboBox (" & arry(5) & ")"
BadIniFile = True
On Error GoTo 0
End Sub

Sub AddOptionButton(CurrentForm As Form, arry() As String)
Dim i As Integer
'see comments on AddCheckBox
On Error GoTo Exit_err
i = arry(0)
With CurrentForm
    If i >= .Option1().LBound And i <= .Option1().UBound Then
        If StringToTrueFalse(arry(1)) = True Then
            .Option1(i).Value = True
        Else
            .Option1(i).Value = False
        End If
'        .Option1(i).Caption = Arry(2)
        If arry(3) = "Hide" Then .Option1(i).Visible = False
        If arry(3) = "Show" Then .Option1(i).Visible = True
        .Option1(i).Tag = arry(4)
    End If
End With
'If CurrentForm.Visible Then Call DisplayForms("After AddOptionButton (" & CurrentForm.Name & ") in ReadIniFile in TreeFilter")

Exit Sub

Exit_err:
BadModule = "AddOptionButton"
BadIniFile = True
End Sub

Sub AddFlexGrid(CurrentForm As Form, arry() As String)
Dim MsgType As String      'From FieldKey
Dim Dac As String           'diito
Dim Fi As String           'diito
Dim Fiid As String           'diito
Dim MsgKey As String   'arry(0)
Dim i As Long
On Error GoTo Exit_err
With CurrentForm.FieldList
'reformat key   lhjustify problem
'MsgBox "[" & TreeFilter.FieldList.TextMatrix(1, 0) & "]"
'if a null key ("") is replaced with a blank key ("   ")
'the first blank line will not be removed
    If arry(0) <> "" Then   'dont replace Nul key with blanks
'        If Left$(arry(0), 2) <> "  " Then 'ais not nmea
'ONLY for compatibility with 2010 .ini file having "  $xxxxx"
'V117 on
'See if we need an AisSpecific Message, if not remove
'the MsgType No
        If NullToZero(arry(0)) <> 0 And arry(1) = "NmeaOut" Then
'Stop
            Select Case arry(2)
'18        ,NmeaOut,receivedtime,0,0,,,2,18,    ,  ,  ,nmea_time,Received Time UTC,,
'  $       ,SentenceOut,receivedtime,0,0,,,2,$,,,,nmea_time,Received Time UTC,,
            Case Is = "receivedtime"
                arry(0) = "  $       "
                arry(1) = "SentenceOut"
                arry(8) = "$"
            Case Is = "aisword"
'Dont think anyone will have this be using aisword
'If they are skip it, as I dont know what to do with it
'They'll have to set up the tag again
                Exit Sub
            Case Else
                arry(0) = "  ais     "
                arry(1) = "NmeaAisOut"
                arry(8) = "ais"
            End Select
        End If

        If NullToZero(arry(0)) <> 0 And arry(1) = "DetailOut" Then
'  ais     ,NmeaAisOut,mmsi,0,0,,,2,ais,    ,  ,  ,mmsi,MMSI,,
' 1        ,DetailOut,mmsi,9,30,,,2, 1,    ,  ,  ,mmsi,MMSI,,
'Stop
            Select Case arry(2)
            Case Is = "aismsgtype"
                If arry(3) = "1" And arry(4) = "6" Then
                    arry(0) = "  ais     "
                    arry(1) = "NmeaAisOut"
                    arry(3) = "0"
                    arry(4) = "0"
                    arry(8) = "ais"
                End If
            Case Is = "repeat"
                If arry(3) = "7" And arry(4) = "2" Then
                    arry(0) = "  ais     "
                    arry(1) = "NmeaAisOut"
                    arry(3) = "0"
                    arry(4) = "0"
                    arry(8) = "ais"
                End If
            Case Is = "mmsi", "mid"
' 1        ,DetailOut,mmsi,9,30,,,2, 1,    ,  ,  ,mmsi,MMSI,,
' 1        ,DetailOut,mid,9,30,,,3, 1,    ,  ,  ,mid_3,MID,,
'  ais     ,NmeaAisOut,mmsi,0,0,,,2,ais,    ,  ,  ,mmsi,MMSI,,
                If arry(3) = "9" And arry(4) = "30" Then
                    arry(0) = "  ais     "
                    arry(1) = "NmeaAisOut"
                    arry(3) = "0"
                    arry(4) = "0"
                    arry(8) = "ais"
                End If
            Case Else
            End Select
        End If

'        aMsgType = DisplayFieldMsgType(Left$(arry(0), 2), arry(1))
'        If arry(8) <> aMsgType Then
'Stop
'            arry(8) = aMsgType
'        End If
        If NullToZero(Left$(arry(0), 2)) > 0 Then
            MsgKey = arry(0)
            MsgType = Left$(MsgKey, 2)
            Dac = Mid$(MsgKey, 3, 4)
            Fi = Mid$(MsgKey, 7, 2)
            Fiid = Mid$(MsgKey, 9, 2)
'compatibility with new key format (no zeros) will require removing
            On Error Resume Next
            If CInt(Dac) = 0 Then Dac = ""  'blank if aismsg 24A
            If CInt(Fi) = 0 Then Fi = ""  'blank if aismsg 24A
'if FiId is non numberic (msg 24 part A or B) must leave it
            If IsNumeric(arry(11)) = True Then
                If CInt(Fiid) = 0 Then Fiid = ""
            End If
            If CInt(arry(8)) = 0 Then arry(8) = ""      'remove zeros from fieldlist
            If CInt(arry(9)) = 0 Then arry(9) = ""
            If CInt(arry(10)) = 0 Then arry(10) = ""
'if FiId is non numberic (msg 24 part A or B) must leave it
            If IsNumeric(arry(11)) = True Then
                If CInt(arry(11)) = 0 Then arry(11) = ""
            End If
            On Error GoTo Exit_err
'ensure correct format
            arry(9) = cKey(arry(9), 4)  'DAC
            arry(10) = cKey(arry(10), 2)    'FI
            arry(11) = cKey(arry(11), 2)    'FIID
    
            MsgKey = cKey(MsgType, 2) _
            & cKey(Dac, 4) _
            & cKey(Fi, 2) _
            & cKey(Fiid, 2)
            arry(0) = MsgKey
        Else
'Arry(0) is numeric (Not 2010 Format)
'pad out KEY to 10 characters (1st 2 will be spaces)
'Remove any leading spaces if not AIS (Compatibility of .ini files)
'Trim removes Leading spaces Trim$ doesnt
'            arry(0) = Trim(arry(0))
            arry(0) = cKey(arry(0), 10)
'            arry(8) = Trim(arry(8))
'Display in sentence type in Msg col on field list
'Arry(8)must be 0
'            arry(8) = Trim(arry(0)) 'NMEA sentence type
'Display 1st argument as FIID on field list
'            arry(11) = arry(4)
        End If
'MsgBox Join(arry, vbTab)   'debug message key
'If arry(0) = "NMEA      " Then Stop

'From Nov13 we have changed the ini file so NmeaOut processes
'Nmea fields for all AisMessage types and we do not want to
'duplicate the fields for each individual message type
                
        arry(8) = DisplayFieldMsgType(MsgType, arry(1))
        If IsDuplicateFlexGridRow(arry) = False Then
            .AddItem Join(arry, vbTab)
            If .TextMatrix(1, 0) = "" Then .RemoveItem 1
        End If
    End If
End With
'MsgBox "[" & TreeFilter.FieldList.TextMatrix(1, 0) & "]"
'Call DisplayForms("After AddFlexGrid in ReadIniFile in TreeFilter")
'For i = 1 To 13
' Debug.Print i & ":" & CurrentForm.FieldList.TextMatrix(1, i)
'Next i
Exit Sub

Exit_err:
BadModule = "AddFlexGrid"
BadIniFile = True
End Sub

Public Function IsDuplicateFlexGridRow(PassedArry() As String) As Boolean
Dim i As Long
Dim j As Integer
Dim arry() As String
Dim myFlexGrid As MSFlexGrid
Dim Passed As String
Dim Existing As String

    Passed = Join(PassedArry, ",")
    Set myFlexGrid = FieldList
'    With objFlexGrid
    With myFlexGrid
    For i = 1 To .Rows - 1
        ReDim arry(.Cols)
        For j = 0 To .Cols - 1
            arry(j) = .TextMatrix(i, j)
        Next j
    Existing = Join(arry, ",")
    If Existing = Passed Then
        IsDuplicateFlexGridRow = True
        Exit Function
    End If
    Next i
End With
    Set myFlexGrid = Nothing
'Stop
End Function


Sub AddFlexGrid1(CurrentForm As Form, arry() As String)

On Error GoTo Exit_err
With CurrentForm.TagList
    .AddItem Join(arry, vbTab)
    If .TextMatrix(1, 0) = "" Then .RemoveItem 1

End With
'Call DisplayForms("After AddFlexGrid1 in ReadIniFile in TreeFilter")

Exit Sub

Exit_err:
BadModule = "AddFlexGrid1"
BadIniFile = True
End Sub

Sub AddTextLabel(arry() As String)
Dim i As Long
#If False Then
    Dim oFTP As New FTP
#End If

On Error GoTo Exit_err

#If False Then      'This was when I was debugging encrypt
'and it was splitting the password
Dim kb As String
Dim j As Long
Dim k As Long
Dim SplitLine() As String
Do Until UBound(arry) = 3
    j = UBound(arry)
    k = 0
    Line Input #IniFileCh, kb
    SplitLine() = Split(kb, ",")
    arry(j) = arry(j) & SplitLine(k)
    Do Until UBound(SplitLine) = k
        k = k + 1
        j = j + 1
        ReDim Preserve arry(j)
        arry(j) = SplitLine(k)
    Loop
Loop
#End If

i = arry(0)
With TreeFilter
        If i >= .Text1().LBound And i <= .Text1().UBound Then
            .Text1(i).Text = arry(1)
'            .Label1(i).Caption = Arry(2)
            .Text1(i).Tag = arry(3)
'force tag name for compatibility with earlier .ini files
            If i = 6 And .Text1(i).Text = "" Then .Text1(i).Text = ","
            If .Text1(i).Text = "," Then .Text1(i).Tag = "Char"
'if delimiter is blank force comma (cant use , because its a csv file thats been read
            If .Text1(i).Tag = "Char" Then
                If .Text1(i).Text = "" Then
                    .Text1(i).Text = ","
                    .Label1(i).Caption = "Delimiter"
                End If
            End If
'if french replace dot with comma 'was commented out in v131 (not sure why)
       If .Text1(i).Tag = "Numeric" Then
            If GetDecimalSep() = "," Then
                .Text1(i).Text = Replace(Text1(i).Text, ".", ",")
            End If
        End If
        If Text1(i).Tag = "encrypt" Then
'Change to the decrypted password and save both
'            MsgBox .Text1(i).Text
                        
            FtpIniPasswordEncrypted = .Text1(i).Text
            FtpIniPasswordDecrypted = DecryptString(.Text1(i).Text)
            .Text1(i).Text = FtpIniPasswordDecrypted
'Set the current Encrypted password
            FtpPasswordEncrypted = FtpIniPasswordEncrypted
            FtpPassword = FtpIniPasswordDecrypted
'If the Password has not been encrypted (in the .ini file) encrypt it
            If FtpPasswordEncrypted = FtpPassword Then
                FtpPasswordEncrypted = EncryptString(FtpPassword)
            End If
        End If
    End If
End With
#If False Then
    Set oFTP = Nothing
#End If
'Call DisplayForms("After AddTextLabel in ReadIniFile in TreeFilter")
Exit Sub

Exit_err:
BadModule = "AddTextLabel.(" & arry(0) & ")"
BadIniFile = True
'Resume Next
End Sub


'see http://visualbasic.freetutes.com/learn-vb6/8.3.causevalidation-property-validate-event.html
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
' You can't close this form without validating all the fields on it.
'Include line below if you need to exit with a cancel button
'If UnloadMode = vbFormControlMenu Then
Dim ctrl As Control

    On Error Resume Next
    If UnloadMode = vbFormControlMenu Then  'User clicked (X)
'If TreeFilter.Visible = True Then
' Give the focus to each control on the form, and then
' validate it.
        For Each ctrl In Controls
'only check if tag is numberic (set up in .ini file)
            If ctrl.Tag = "Numeric" Or ctrl.Tag = "Char" Or ctrl.Tag = "encrypt" Then
                err.Clear
                ctrl.SetFocus
                If err = 0 Then
' Don't validate controls that can't receive input focus.
                    ValidateControls
                    If err = 380 Then
' Validation failed, refuse to close.
                        Cancel = True: Exit Sub
                    End If
                End If
            End If
        Next
'Include line below if you need to exit with a cancel button
'    End If
        Call HideTreeFilter 'also Hides ListFilters
        Cancel = True   'dont actually close the form (just hide it)
    End If
End Sub

'We also have to hide the List Windows
Public Sub HideTreeFilter()
Dim ListFilter As Variant
    
    TreeFilter.Hide
    For Each ListFilter In ListFilters
        ListFilter.Hide
'When we Hide the tree filter we can no longer edit the Tree filter
'If the List filter contains no entries then by unticking it the user will have to
'tick if they wish to edit the list filter when the Options Window is next shown
'This caused the list filter to be displayed again (by ListFilterVisibility)
'ListFilter.Tag is [Nodes] Field 3 on /inin file
'eg "AisMsgFromMmsi"
'on AIS,4,AisMsgFromMmsi,Specify from MMSI's,AisMsgFromMmsi,-1,
        
        If NodeExists(ListFilter.Tag) = True Then   'Not in .ini file
            If ListFilter.List1.ListCount <= 0 Then
                TreeView1.Nodes(ListFilter.Tag).Checked = False
'Untick any parents with no children and re-create NmeaRcv Filter display
'I think this caused a node display problem - but not sure what ?
'Call DisplayTree("HideTreeFilter")
'                Call CheckAll(TreeFilter.TreeView1.Nodes("InputFilter"), "HideTreeFilter", True)
'Call DisplayTree("HideTreeFilter")
            End If
        End If
    Next ListFilter
End Sub

Private Sub Option1_Click(Index As Integer)
If Visible Then
    Call SetOutputOptions
    Call CheckOutputOptions
End If
End Sub

Private Sub TagList_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
Dim reply As Integer
Dim i As Long
Dim OldRowHeight As Long

If Button = vbLeftButton Then
    TreeFilter.Show     'display edit form
    With TagList
'.BackColor = vbRed
'if fixed row or column is selected selection is forced to next row/col
        .RowSel = .Row    'force selection to one cell
        .ColSel = .col
        .CellBackColor = vbRed
'trap a invalid cell
'If Arry(0) = "" Then           'blank column"
'            MsgBox "There is no source for this selection"
'            .CellBackColor = vbWhite
'            Exit Sub
'        End If
        
        Select Case .col
        Case Is = 1         'tag column
'Colour all fields to be deleted
            With FieldList
                OldRowHeight = .RowHeight(1)
                For i = 1 To .Rows - 1
                    .Row = i
                    If .TextMatrix(i, 12) = TagList.TextMatrix(TagList.Row, 1) Then
                        .col = 12
                        .CellBackColor = vbRed
                    Else
                        .RowHeight(.Row) = 0
                    End If
                Next i
            End With
            reply = MsgBox("This will delete all Output Fields with a " & .TextMatrix(.Row, .col) & " tag", vbOKCancel)
            If reply <> vbCancel Then
                If .Rows = 2 Then   'last Tag being removed
                    .AddItem ""
                End If
                .RemoveItem .Row
                Call ResetTags("TagList")
            End If
            With FieldList
                For i = 1 To .Rows - 1
                    .Row = i
                    .col = 12
                    .CellBackColor = vbWhite
                    .RowHeight(.Row) = OldRowHeight
                Next i
            End With
        Case Is = 2, 3
            TreeFilter.SetFocus
            TagInput.Label1 = .TextMatrix(0, .col) & " value"
            TagInput.Label2 = .TextMatrix(.Row, 1) & " - " & .TextMatrix(0, .col)
            TagInput.Text1.Text = .TextMatrix(.Row, .col)
            TagInput.Show vbModal

'MsgBox "return from field input " & FieldInput.Cancel & " " & FieldInput.Command2.Value
            If TagInput.Cancel = False Then
                If TagInput.Text1.Text <> "" Then
'check if numeric
                    If IsNumeric(TagInput.Text1.Text) = False Then
                        reply = MsgBox("Value entered is not numeric." _
                        & vbCrLf & "String value entered must be the same" _
                        & vbCrLf & "as decoded value of " & .TextMatrix(.Row, 1) _
                        & " for output to be valid.", vbOKCancel)
                    Else
'check if mmsi
                        If InStr(1, .TextMatrix(.Row, 1), "mmsi", vbTextCompare) <> 0 Then
                            If Len(TagInput.Text1.Text) <> 9 Then
                                reply = MsgBox("MMSI entered is not 9 digits." _
                        & vbCrLf & "This could cause unpredictable results" _
                        & vbCrLf & .TextMatrix(.Row, 1) & " entered is " _
                        & Len(TagInput.Text1.Text) & " digits.", vbOKCancel)
                            End If
                        End If
                    End If
                End If
                If reply <> vbCancel Then
                    .TextMatrix(.Row, .col) = TagInput.Text1.Text
                    Call ResetTags("FieldList")
                End If
            End If
        End Select
        .CellBackColor = vbWhite
    End With
End If
TreeFilter.SetFocus
Call SetOutputOptions   'to check lat lon
Call CheckOutputOptions
End Sub


'leave only rightmost char
Private Sub Text1_Change(Index As Integer)
Dim kb
    Select Case Index
    Case Is = 2, 6
        kb = "[" & Text1(Index) & "]" & Len(Text1(Index))   'debug
        If Len(Text1(Index)) > 1 Then
            Text1(Index) = Right$(Text1(Index), 1)
        End If
        kb = Text1(Index)   'debug
    End Select
End Sub
'select all the text in the box (if user presses del it will be cleared)
Private Sub Text1_GotFocus(Index As Integer)
Select Case Index
Case Is = 2, 6
Text1(Index).SelStart = 0
Text1(Index).SelLength = Len(Text1(Index))
End Select

End Sub

'This is called when form is minimised (see QueryUnload)
'and only if Tag is specified in QueryUnload
Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
    Select Case Text1(Index).Tag
    Case Is = "Numeric"
'Validate the input data. Highlight the offending text.
        If Not IsDataValid(Text1(Index).Text) Then
      'Set the focus back and highlight the text.
            Text1(Index).SelStart = 0
            Text1(Index).SelLength = Len(Text1(Index).Text)
            Cancel = True
        End If
'change caption for Delimiter-Comma to remove Comma
    Case Is = "Char"
        Label1(Index).Caption = "Delimiter"
    Case Is = "encrypt"
'MsgBox "Text1(" & Index & ") " & Text1(Index).Text & " " & Text1(Index).Tag
'unencrypted password has changed
        If Text1(Index).Text <> FtpPassword Then
            
            FtpPassword = Text1(Index).Text
'We change this so the new encrypted password is written outr to the .ini file
            FtpPasswordEncrypted = EncryptString(FtpPassword)
        End If
    End Select
If Visible Then
    Call SetOutputOptions
    Call CheckOutputOptions
End If

End Sub

Private Sub TreeView1_Click()
'this event occurs after the node_check event
'this is required to remove the tick from the selected
'node if the parent is not checked
Dim nSel As Node
Dim nParent As Node
Dim nR As Node
    
    Set nSel = TreeView1.SelectedItem
    Set nParent = nSel.Parent
    If IsObject(nSel.Parent) = False Then MsgBox "no object"
    If nParent Is Nothing Then  'bug is not nothing doesnt work !
    Else
        If nParent.Checked = False Then nSel.Checked = False
    End If
    Call CheckAll(nR, "TreeView1_Click")
'Display Lists if Ticked/Hide list if unticked
    Call ListFilterVisibility

End Sub

#If False Then
'Makes the form visible/invisible , if a list filter is selected and Treefilter is visible
Private Sub ListFilterVisibilityFromNodes()
Dim FilterNo As Integer

Select Case Node.Key
Case Is = "AisMsgFromMmsi"
    FilterNo = 0
Case Is = "AisMsgToMmsi"
    FilterNo = 1
Case Is = "DacList"
    FilterNo = 2
Case Is = "FiList"
    FilterNo = 3
Case Is = "ListFiId"
    FilterNo = 4
Case Else
    FilterNo = -1
End Select

If FilterNo > -1 Then
    If Node.Checked = True Then
        ListFilters(FilterNo).Show
    Else
        ListFilters(FilterNo).Hide
    End If
End If


End Sub
#End If
'Occures before Click event
Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim aNodes As String
Dim lLevel As Long

'CheckAll cannot be done in NodeCheck because if the checkbox
'is changed to false, when the even is exited it is re-set to true
'set the selected item so that it can be recovered with the click event
'keep a note of the clicked node, so that after checkall the
'clicked node is in view
Set ClickedNode = Node
TreeView1.SelectedItem = Node
End Sub

'Unchecks parent if no children checked
'If Copy is true then we create the Input and Output filter trees
'by transferring nodes selected on edit view, to the displayed view
'it is the displayed view that is checked to see if a node is filtered
'The Root node (nR) is ignored when copying to NmeaRcv

Public Sub CheckAll(nR As Node, Optional Caller As String)
Dim lLevel As Long
Dim nodX As Node
Dim i As Long
Dim kb As String

    On Error GoTo Proc_err
'    Debug.Print "#CheckAll Caller=" & Caller
    
    Set nR = TreeFilter.TreeView1.Nodes("InputFilter")
    
'Now clear the NmeaRcv tree (This is the tree used by the InputFilter)
    NmeaRcv.TreeView1.Nodes.Clear
'put Root node "Input Filter" on Input Filter display (NmeaRcv)
    Set nodX = NmeaRcv.TreeView1.Nodes.Add(, , nR.Key, nR.Text)
    NmeaRcv.TreeView1.SingleSel = True ' to view the children
    nodX.Tag = nR.Tag
    nodX.EnsureVisible
   
'NmeaRcv is the form containing "Input Filter"
'Copy TreeFilter ticked nodes to NmeaRcv
    ValidateTreeFilter nR, True, NmeaRcv
    Set nR = Nothing
    Set nodX = Nothing
    
'If we have clicked a node make sure we can see it in the display
    If ClickedNode Is Nothing Then
        TreeFilter.TreeView1.Nodes(1).EnsureVisible
    Else
        ClickedNode.EnsureVisible
    End If

Exit Sub

Proc_err:
    Select Case err.Number
    Case Is = 35603     'invalid key
        err.Clear
    Case Is = 35605      'Control has been deleted
        err.Clear
    Case Else
        MsgBox "Error " & err.Number & vbCrLf & err.Description, , "CheckAll"
    End Select
    Resume Next
End Sub

'walks the whole tree
'creates the NmeaRcv Tree display if copy=true
'Unchecks Parents with no ticked children
'Synchronise Filter to Tags also calls this routine
Sub ValidateTreeFilter(nR As Node, _
    Optional Copy As Boolean, Optional objFormName As Object)

Dim tNode As Node   'The Current Node in TreeFilter
Dim nodX As Node    'The DisplayOnly Node in NmeaRcv
Dim i As Long
Dim j As Long

'Call frmDebug.DisplayDebug("start ValidateTreeFilter")
'Call DisplayTree    'debug
'    Call SynchroniseListFilters 'Ticks/unticks parent if entries exist
'Call frmDebug.DisplayDebug("start ValidateTreeFilter")
'Call DisplayTree    'debug
    
    Set tNode = nR
    Do Until tNode Is Nothing
 'Debug.Print tNode
'tNode.Key is arry(0) in [nodes] in .ini file
    
'If Parent is unchecked, uncheck this child
        If Not tNode.Parent Is Nothing Then     'Has a Parent (Root wont)
            If tNode.Parent.Checked = False Then
'If tNode.Checked = True Then Stop
                tNode.Checked = False
                tNode.Parent.Expanded = False
            Else
                tNode.Parent.Expanded = True
            End If
    
            If Copy = True Then
                If tNode.Checked Then
'set up any ListFilter's (no parent required Just the List Entry)
'else just set up the node - with a unique Key by adding the ListIndex
                    Select Case tNode.Key
                    Case Is = "AisMsgFromMmsi"
                        For i = 0 To ListFilters(0).List1.ListCount - 1
                            Set nodX = _
objFormName.TreeView1.Nodes.Add(tNode.Parent.Key, tvwChild, "AisMsgFromMmsi" & ListFilters(0).List1.List(i), "From MMSI " & ListFilters(0).List1.List(i))
                            nodX.Tag = "AisMsgFromMmsi" 'filter name
                            nodX.EnsureVisible
'TreeView1.Nodes(ListFilter.Tag).Parent.Tag = True   'AIS sentences
                        Next i
                    Case Is = "AisMsgToMmsi"
                        For i = 0 To ListFilters(1).List1.ListCount - 1
                            Set nodX = _
objFormName.TreeView1.Nodes.Add(tNode.Parent.Key, tvwChild, "AisMsgToMmsi" & ListFilters(1).List1.List(i), "To MMSI " & ListFilters(1).List1.List(i))
                            nodX.Tag = "AisMsgToMmsi"
                            nodX.EnsureVisible
'TreeView1.Nodes(ListFilter.Tag).Parent.Tag = True   'AIS sentences
                        Next i
                    Case Is = "DacList"
                        For i = 0 To ListFilters(2).List1.ListCount - 1
                         Set nodX = _
objFormName.TreeView1.Nodes.Add(tNode.Parent.Key, tvwChild, "DacList" & ListFilters(2).List1.List(i), "DAC " & ListFilters(2).List1.List(i))
                            nodX.Tag = "DacList"  'filtername
                            nodX.EnsureVisible
                        Next i
                    Case Is = "FiList"
                        If tNode.Parent.Key = "DacList" Then
                            For i = 0 To ListFilters(2).List1.ListCount - 1
                                For j = 0 To ListFilters(3).List1.ListCount - 1
'                                Set nodX = _
'objFormName.TreeView1.Nodes.Add("Dac" & ListFilters(2).List1.List(i), tvwChild, "DacFi" & ListFilters(2).List1.List(i) & "-" & ListFilters(3).List1.List(j), "Function " & ListFilters(3).List1.List(j))
                                Set nodX = _
objFormName.TreeView1.Nodes.Add("Dac" & ListFilters(2).List1.List(i), tvwChild, "DacFi" & ListFilters(2).List1.List(i) & "-" & ListFilters(3).List1.List(j), "Function " & ListFilters(3).List1.List(j))
                                nodX.Tag = "FiList"  'filtername
                                nodX.EnsureVisible
                                Next j
                            Next i
                        Else
'All Dacs
                          For i = 0 To ListFilters(3).List1.ListCount - 1
'objFormName.TreeView1.Nodes.Add(tNode.Parent.Key, tvwChild, "DacList" & ListFilters(2).List1.List(i), "DAC " & ListFilters(2).List1.List(i))
                             Set nodX = _
objFormName.TreeView1.Nodes.Add(tNode.Parent.Key, tvwChild, "FiList" & ListFilters(3).List1.List(i), "FI " & ListFilters(3).List1.List(i))
                                nodX.Tag = "FiList"  'filtername
                                nodX.EnsureVisible
                            Next i
                        
                        End If
#If False Then
                    Case Is = "FiList"
'if the parent is a list we have to set child up for each parent in the list
                        If tNode.Parent.Key = "FiList" Then
                            For i = 0 To ListFilters(2).List1.ListCount - 1
                                For j = 0 To ListFilters(3).List1.ListCount - 1
'                                Set nodX = _
'objFormName.TreeView1.Nodes.Add("Dac" & ListFilters(2).List1.List(i), tvwChild, "DacFi" & ListFilters(2).List1.List(i) & "-" & ListFilters(3).List1.List(j), "Function " & ListFilters(3).List1.List(j))
                                Set nodX = _
objFormName.TreeView1.Nodes.Add("Fi" & ListFilters(2).List1.List(i), tvwChild, "DacFi" & ListFilters(2).List1.List(i) & "-" & ListFilters(3).List1.List(j), "Function " & ListFilters(3).List1.List(j))
                                nodX.Tag = "ListFi"  'filtername
                                nodX.EnsureVisible
                                Next j
                            Next i
                        End If
#End If
                    Case Else
'Put next node on display list (NmeaRcv)
                        Set nodX = objFormName.TreeView1.Nodes.Add(tNode.Parent.Key, tvwChild, tNode.Key, tNode.Text)
                        nodX.Tag = tNode.Tag
                        nodX.EnsureVisible
'NmeaRcv.Show
                    End Select
                End If
            End If
        End If
    
    
    
        If tNode.Children Then ' has child nodes; move to 1st child
            Set tNode = tNode.Child.FirstSibling
        ElseIf tNode.Next Is Nothing Then ' gotta move up level(s)
            Set tNode = tNode.Parent
            Do Until tNode Is Nothing   ' if Nothing, then done
                If Not tNode.Next Is Nothing Then
                    Set tNode = tNode.Next  'Next sibling
                    Exit Do
                End If
                Set tNode = tNode.Parent ' move up again
            Loop
        Else
            Set tNode = tNode.Next ' move to next sibling
        End If
    Loop

End Sub

'walks the whole tree & displays in msgbox - can be used for debugging
Sub DisplayTree(Optional Caller As String)
Static LastCaller As String
Static CallerCount As Long
Dim tNode As Node
Dim strText As String
Dim Level As Long

If DebugTreeFilter = False Then
    Exit Sub
End If

If Caller <> LastCaller Then CallerCount = 0
LastCaller = Caller
CallerCount = CallerCount + 1
Call frmDebug.Display(Caller & "(" & CallerCount & ")")
On Error GoTo NoNode
strText = ""
Level = 0
'Set tNode = TreeFilter.TreeView1.Nodes("InputFilter")
Set tNode = NmeaRcv.TreeView1.Nodes("InputFilter")
Do Until tNode Is Nothing

'Debug Display of specific nodes
'If Level < 3 Or tNode = "Specify from MMSI's" Then
    strText = strText & Level & String$(Level + 1, vbTab) & tNode & ":" & tNode.Checked & vbCrLf
'End If
    
    If tNode.Children Then ' has child nodes; move to 1st child
        Set tNode = tNode.Child.FirstSibling
        Level = Level + 1
'    strText = strText & String$(Level, vbTab) & tNode & ":" & tNode.Checked & vbCrLf
    ElseIf tNode.Next Is Nothing Then ' gotta move up level(s)
        Set tNode = tNode.Parent
        Level = Level - 1
        Do Until tNode Is Nothing   ' if Nothing, then done
            If Not tNode.Next Is Nothing Then
                Set tNode = tNode.Next
                Exit Do
            End If
            Set tNode = tNode.Parent ' move up again
            Level = Level - 1
        Loop
    Else
        Set tNode = tNode.Next ' move to next sibling
    End If
Loop
Call frmDebug.Display(strText)
Exit Sub

NoNode:
    Resume Next
End Sub

Private Function IsDataValid(InputData As String) As Boolean
'Return true if the input data is numeric.
'Actually must be a positive integer, for all current numeric fields
If IsNumeric(InputData) Then
      'Data is numeric.
        If InputData < 0 Then
            Exit Function   'Data is negative.
        End If
        If Int(CSng(InputData)) <> CSng(InputData) Then
'            Exit Function   'not an integer
        End If
'we check for decimal separators as 0.0 would otherwise ve accepted
'this could then appear as valid if the text string is checked
        If InStr(InputData, ".") Then
'            Exit Function
        End If
        If InStr(InputData, ",") Then
'            Exit Function
        End If
            'Data is not negative.
        IsDataValid = True
   End If
End Function

Public Sub SynchroniseTags()
Dim i As Long
Dim n As Integer
Dim nR As Node
Dim kb As String
Dim arry() As String
Dim ParentNode As String
Dim ReqNode As String
Dim FilterText As String
Dim TreeText As String
Dim FilterName As String
Dim AisMsgNo As String
Dim Dac As String
Dim Fi As String
Dim Fiid As String

'Dont synchronise the tags when loading a new profile until inifile is closed because
'if there are no tags any
'filter settings (in the ini file) will be overwritten
    If IniFileCh > 0 Then   'TagArray not initialised
        Exit Sub
    End If

'No tags to synchronise
    If TagArray(0, 0) = "" Then
        Exit Sub
    End If

'Debug.Print "#SynchroniseTags"

'need to clear all checked first, walk the tree clearing checkboxes
'this ensures both first chilren are not set
    If TreeFilter.Check1(1).Value = vbChecked Then   'checked
        Set nR = TreeView1.Nodes("InputFilter") 'V129
        TreeView1.Nodes("InputFilter").Checked = True
        TreeView1.Nodes("CRCerror").Checked = False
        TreeView1.Nodes("NMEA").Checked = False
'Check the Field Tags
        With TreeFilter.FieldList
            For i = 1 To .Rows - 1
                .Row = i
'NMEA must be ticked (it should always exist in the treefilter)
                TreeView1.Nodes("NMEA").Checked = True
'If IEC exists - It must also be ticked (Splits NMEA into VDO and VDM)
                On Error Resume Next
                TreeView1.Nodes("IEC").Checked = True
                On Error GoTo 0
                Select Case .TextMatrix(.Row, 1)
Case Is = "SentenceOut"
'Stop
Case Is = "MyShipOut"
'Stop
Case Is = "NmeaAisOut"
'Stop
Case Is = "CommentOut"
'Stop
                Case Is = "NmeaOut"
'If you cant find the node in the tree skip it
'as the next highest will be checked
On Error GoTo No_Node:  'v 3.2.135 (Cunningham)Element not found err 35601

kb = .TextMatrix(.Row, 0)
                    Select Case .TextMatrix(.Row, 2)
                    Case Is = "nmeacrc"
                        TreeView1.Nodes("CRCerror").Checked = True
                    End Select
                    TreeView1.Nodes("InputFilter").Checked = True
'the nmea code can be smartened up a bit
'Trim (no $) removes leading spaces
'kb = Trim(.TextMatrix(.Row, 0))
'                    If Left$(Trim(.TextMatrix(.Row, 0)), 1) = "$" Then
'                        TreeView1.Nodes("$").Checked = True
'                        If Left$(Trim(.TextMatrix(.Row, 0)), 3) = "$GP" Then
'                            TreeView1.Nodes("$GP").Checked = True
'                        End If
'                    End If
                    Select Case Trim(.TextMatrix(.Row, 0))
                    Case Is = "$GPZDA"
'                        TreeView1.Nodes("$GPZDA").Checked = True    'v130
                        TreeView1.Nodes("NMEA$-$GPZDA").Checked = True  'v135
                        TreeView1.Nodes("NMEA$").Checked = True         'v135
                    Case Is = "$GPGGA"
'                        TreeView1.Nodes("$GPGGA").Checked = True
                        TreeView1.Nodes("NMEA$-$GPGGA").Checked = True
                        TreeView1.Nodes("NMEA$").Checked = True
                    Case Is = "$GPRMC"
'                        TreeView1.Nodes("$GPRMC").Checked = True
                        TreeView1.Nodes("NMEA$-$GPRMC").Checked = True
                        TreeView1.Nodes("NMEA$").Checked = True
                    Case Is = "$PGHP"
'                        TreeView1.Nodes("$PGHP").Checked = True
                        TreeView1.Nodes("NMEA$-$PGHP").Checked = True
                        TreeView1.Nodes("NMEA$").Checked = True
                    Case Is = "$AITXT"
                        TreeView1.Nodes("$AITXT").Checked = True
                    Case Is = "$AIALR"
                        TreeView1.Nodes("NMEA$-$AIALR").Checked = True
                        TreeView1.Nodes("$AIALR").Checked = True
                    Case Is = "$AIBRM"
                        TreeView1.Nodes("NMEA$-$AIBRM").Checked = True
                        TreeView1.Nodes("$AIBRM").Checked = True
                    Case Else
'Node not set up on input filter (new sentence TAG has been
'created by user clicking on detail display)
'$,4,$PGHP,$PGHP GH Internal Message,$,0,
'Same structure as ini file
                        ReDim arry(6)
                        arry(0) = "NMEA"
                        arry(1) = tvwChild
                        arry(2) = Trim(.TextMatrix(.Row, 0))
                        arry(3) = arry(2)       'set description same as sentence
                        arry(4) = "$"    'Filter name
                        arry(5) = -1    'True
                        Call AddNode(arry)
                        TreeView1.Nodes(arry(2)).Checked = True
                    
                    End Select
            
                Case Is = "DetailOut"
kb = .TextMatrix(.Row, 1)
'V90                    TreeView1.Nodes("!AIVDM").Checked = True
                    TreeView1.Nodes("AIS").Checked = True
                    If SplitFieldKey(.TextMatrix(.Row, 0), AisMsgNo, Dac, Fi, Fiid) = True Then
                        FilterName = "AisMsg"
                        ReqNode = "AisMsg" & AisMsgNo
'If you cant find the node in the tree skip it
'as the next highest will be checked
On Error GoTo No_Node:
                        TreeView1.Nodes(ReqNode).Checked = True
                        If Dac <> "" Then
ParentNode = ReqNode
FilterName = "AisMsgDac"
TreeText = "DAC " & Dac
ReqNode = "AisMsg" & AisMsgNo & "Dac-" & Dac
                            TreeView1.Nodes(ReqNode).Checked = True
                        End If
                        If Fi <> "" Then
ParentNode = ReqNode
FilterName = "AisMsgDacFi"
TreeText = "Function " & Fi
ReqNode = "AisMsg" & AisMsgNo & "Dac" & Dac & "Fi-" & Fi
                            TreeView1.Nodes(ReqNode).Checked = True
                        End If
                        If Fiid <> "" Then
                            ParentNode = ReqNode
                            FilterName = "AisMsgDacFiID"
                            If Fi <> "" Then    'has Fi and Fiid
                                TreeText = "ID " & Fiid
ReqNode = "AisMsg" & AisMsgNo & "Dac" & Dac & "Fi" & Fi & "Fiid-" & Fi
                                TreeView1.Nodes(ReqNode).Checked = True
                            Else    'Only Fiid (Msg 24 A or B)
'AisMsg24,4,AisMsg24Id-A,Type A,AisMsgFiID,0,
                                TreeText = "Type " & Fiid
ReqNode = "AisMsg" & AisMsgNo & "Id-" & Fiid
                            End If
                            TreeView1.Nodes(ReqNode).Checked = True
                        End If
On Error GoTo 0
                    End If
                End Select
            Next i
        End With

'If we are synchronising to tags and there are no AisSentences ticked
'we must tick AisSentences if any MmsiFrom are set up on ListFilter
're-construct List Filter on nmearcv
        Call CheckAll(nR, "SynchroniseTags")
    End If

Exit Sub
No_Node:
'This will create a node from the Field List
'Same structure as ini file
    With TreeFilter.FieldList
                        ReDim arry(6)
                        arry(0) = ParentNode
                        arry(1) = tvwChild
kb = Trim(.TextMatrix(.Row, 8))
                        arry(2) = ReqNode   'Trim(.TextMatrix(.Row, 0))
                        arry(3) = TreeText  'Trim(.TextMatrix(.Row, 8)) '   arry(2)       'set description same as sentence
                        arry(4) = FilterName    'Filter name
                        arry(5) = -1    'True
'v135 node already exists
On Error Resume Next
                        Call AddNode(arry)
On Error GoTo 0
                        TreeView1.Nodes(arry(2)).Checked = True
    End With
Resume Next
End Sub

'If List Filter has entries tick parents(s)
'If no entries dont untick parent because you will be unable to enter the first entry
'If there are List entries but no parent node these will not be shown
'in the list as the form will not get displayed
'Don't show list Filters unless parent Node is ticked
'even if there are entries in the list
'To show these the user has to untick then tick the parent node
'Otherwise the list will be shown when the user has MMSI's in the list (& doesnt want to
'have to keep re-entering them) but does not what them actioning
Public Function ListFilterVisibility()
Dim ListFilter As Variant
Dim kb As String
    
    
    On Error GoTo NoNode
    For Each ListFilter In ListFilters
kb = ListFilter.Caption & ":" & ListFilter.List1.ListCount & ":" & ListFilter.Tag
            
'ListFilter.Tag is [Nodes] Field 3 on /inin file
'eg "AisMsgFromMmsi"
'on AIS,4,AisMsgFromMmsi,Specify from MMSI's,AisMsgFromMmsi,-1,
        If NodeExists(ListFilter.Tag) = True Then
'           If ListFilter.List1.ListCount > 0 Then
'               TreeView1.Nodes(ListFilter.Tag).Checked = True
'               TreeView1.Nodes(ListFilter.Tag).Parent.Checked = True   'AIS sentences
'               TreeView1.Nodes(ListFilter.Tag).Parent.Parent.Checked = True   'NMEA Sentences
'           End If
            If TreeView1.Nodes(ListFilter.Tag).Checked = True Then
                ListFilter.Visible = TreeFilter.Visible
                ListFilter.WindowState = TreeFilter.WindowState
            Else
               ListFilter.Hide
            End If
           
        Else    'This Node does not exits (not defined in .ini file)
            ListFilter.Hide
        End If
    Next ListFilter
'Call frmDebug.DisplayDebug("after SynchroniseListFilters")
'Call DisplayTree    'debug
Exit Function

NoNode:
    Select Case err.Number
    Case Is = 35601     'element doesnt exist (Filter not set up on .ini file)
        err.Clear
    Case Is = 35603     'invalid key (Tag incorrect)
        err.Clear
    Case Else
        MsgBox "Error " & err.Number & vbCrLf & err.Description, , "SynchroniseListFilters"
    End Select
    Resume Next
End Function

'Checks if this node exists in Treeview1
Private Function NodeExists(NodeKey As String) As Boolean
    On Error GoTo NoNode
    If TreeView1.Nodes(NodeKey).Key <> "" Then NodeExists = True
NoNode:
End Function

Private Function HasIndex(ControlArray As Object, ByVal Index As Integer) As Boolean
    HasIndex = (VarType(ControlArray(Index)) <> vbObject)
End Function

Public Function EnableDisableOptions()
Dim Setting As Boolean
    
    If InputState <> 0 Then
        Setting = False
    Else
        Setting = True
    End If
    Text1(3).Enabled = Setting   'client udp port
    Combo1(0).Enabled = Setting  'serial port
    Combo1(1).Enabled = Setting  'serial speed
    Text1(12).Enabled = Setting   'client tdp host
    Text1(13).Enabled = Setting   'client tdp port
    Text1(0).Enabled = Setting   'Scheduled Mins
     
'stop changing UDP output
    Text1(4).Enabled = Setting
    Text1(5).Enabled = Setting

'Stopchanging FileNames
    cbNewFile(0).Enabled = Setting  'OutputFile
    cbNewFile(1).Enabled = Setting  'TagTemplateFile
    cbNewFile(2).Enabled = Setting  'Shell Command File
    
'Stop changing log file
    Option1(0).Enabled = Setting
    Option1(1).Enabled = Setting
    Option1(2).Enabled = Setting
    Option1(3).Enabled = Setting
'File name rollover
    Check1(3).Enabled = Setting    'Input log file rollover
    Check1(22).Enabled = Setting   'output file rollover
    Check1(12).Enabled = Setting      'Shell Enabled
'FTP Server
    Text1(7).Enabled = Setting
    Text1(8).Enabled = Setting
    Text1(9).Enabled = Setting
    Text1(10).Enabled = Setting
    
'required to reset FTP server options if they have changes
'otherwise changes are not written into global public variables
    Call SetOutputOptions
    Call CheckOutputOptions
    
End Function
