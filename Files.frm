VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.ocx"
Begin VB.Form Files 
   Caption         =   "Files"
   ClientHeight    =   5355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4725
   Icon            =   "Files.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   4725
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5055
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   8916
      _Version        =   393216
   End
End
Attribute VB_Name = "Files"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
'? scroll bars count as a col or row
Dim i As Long
Dim tot As Long

With MSFlexGrid1
'shift grid to top left
    .Move 0, 0
'allow user to resize column width
    .AllowUserResizing = flexResizeColumns
'allow scroll bars both
    .ScrollBars = flexScrollBarVertical
'set headings
    .FormatString = "<Name |<Variable|<File Name|<Location"
'    .FormatString = "<Variable       |Element 0|Element 1|Element 2|<Variable       |Element 0|Element 1|Element 2"
'set initial column widths
    .ColWidth(0) = 2000
    .ColWidth(1) = 0
    .ColWidth(2) = 2000
    .ColWidth(3) = 7000
'    .Cols = 2  'create cols if reqd
 'width - Scale width is diff between internal and external size of form
   For i = 0 To .Cols - 1
'        .ColWidth(i) = 2000  'not actually set to 500 but 495
        tot = tot + .ColWidth(i)
    Next i
End With
'resize the form to the grid NOT the other way round
'width - Scale width is diff between internal and external size of form
Width = Width - ScaleWidth + tot
'resize event is fired when form has loaded

Call RefreshData

End Sub

Public Sub RefreshData()
Dim kb As String
Dim i As Long
Dim FileName As String

Dim Description As String
Dim BriefFileName As String
Dim Directory As String
Dim Variable As String

With MSFlexGrid1
'    .Clear
    For i = 2 To .Rows - 1
        .RemoveItem 1
    Next i
    
    Description = "Command Initialisation File"
    Variable = "cmdIniFileName"
    FileName = cmdIniFileName
    If NameFromFullPath(FileName) <> "" Then
        BriefFileName = NameFromFullPath(FileName)
    Else
        BriefFileName = "Not yet defined"
    End If
    If PathFromFullName(FileName) <> "" Then
        Directory = PathFromFullName(FileName)
    Else
        Directory = "Not yet defined"
    End If
    kb = Description & vbTab & Variable & vbTab & BriefFileName & vbTab & Directory
    .AddItem kb
    .RemoveItem 1    'first blank line
    
    Description = "Initialisation File"
    Variable = "IniFileName"
    FileName = IniFileName
    If NameFromFullPath(FileName) <> "" Then
        BriefFileName = NameFromFullPath(FileName)
    Else
        BriefFileName = "Not yet defined"
    End If
    If PathFromFullName(FileName) <> "" Then
        Directory = PathFromFullName(FileName)
    Else
        Directory = "Not yet defined"
    End If
    kb = Description & vbTab & Variable & vbTab & BriefFileName & vbTab & Directory
    .AddItem kb
    
    #If False Then
    Description = "Web Install File"
    Variable = "WebCurrentInstallFileName"
    FileName = WebCurrentInstallFileName
    If NameFromFullPath(FileName) <> "" Then
        BriefFileName = NameFromFullPath(FileName)
    Else
        BriefFileName = "Not yet defined"
    End If
    If PathFromFullName(FileName) <> "" Then
        Directory = PathFromFullName(FileName)
    Else
        Directory = "Not yet defined"
    End If
    kb = Description & vbTab & Variable & vbTab & BriefFileName & vbTab & Directory
    .AddItem kb
    #End If
    
    Description = "Error Log"
    Variable = "ErrorLogFile"
    FileName = ErrorLogFile
    If NameFromFullPath(FileName) <> "" Then
        BriefFileName = NameFromFullPath(FileName)
    Else
        BriefFileName = "Not yet defined"
    End If
    If PathFromFullName(FileName) <> "" Then
        Directory = PathFromFullName(FileName)
    Else
        Directory = "Not yet defined"
    End If
    kb = Description & vbTab & Variable & vbTab & BriefFileName & vbTab & Directory
    .AddItem kb
    
    Description = "NMEA Input Log"
    Variable = "NmeaLogFile"
    FileName = NmeaLogFile
    If NameFromFullPath(FileName) <> "" Then
        BriefFileName = NameFromFullPath(FileName)
    Else
        BriefFileName = "Not yet defined"
    End If
    If PathFromFullName(FileName) <> "" Then
        Directory = PathFromFullName(FileName)
    Else
        Directory = "Not yet defined"
    End If
    kb = Description & vbTab & Variable & vbTab & BriefFileName & vbTab & Directory
    .AddItem kb

    Description = "Output File"
    Variable = "OutputFileName"
    FileName = OutputFileName
    If NameFromFullPath(FileName) <> "" Then
        BriefFileName = NameFromFullPath(FileName)
    Else
        BriefFileName = "Not yet defined"
    End If
    If PathFromFullName(FileName) <> "" Then
        Directory = PathFromFullName(FileName)
    Else
        Directory = "Not yet defined"
    End If
    kb = Description & vbTab & Variable & vbTab & BriefFileName & vbTab & Directory
    .AddItem kb

    Description = "NMEA Output File"
    Variable = "OutputFileNameNmea"
    FileName = OutputFileNameNmea
    If NameFromFullPath(FileName) <> "" Then
        BriefFileName = NameFromFullPath(FileName)
    Else
        BriefFileName = "Not yet defined"
    End If
    If PathFromFullName(FileName) <> "" Then
        Directory = PathFromFullName(FileName)
    Else
        Directory = "Not yet defined"
    End If
    kb = Description & vbTab & Variable & vbTab & BriefFileName & vbTab & Directory
    .AddItem kb

    Description = "CSV Output File"
    Variable = "OutputFileNameCsv"
    FileName = OutputFileNameCsv
    If NameFromFullPath(FileName) <> "" Then
        BriefFileName = NameFromFullPath(FileName)
    Else
        BriefFileName = "Not yet defined"
    End If
    If PathFromFullName(FileName) <> "" Then
        Directory = PathFromFullName(FileName)
    Else
        Directory = "Not yet defined"
    End If
    kb = Description & vbTab & Variable & vbTab & BriefFileName & vbTab & Directory
    .AddItem kb

    Description = "Tagged Output File"
    Variable = "OutputFileNameTagged"
    FileName = OutputFileNameTagged
    If NameFromFullPath(FileName) <> "" Then
        BriefFileName = NameFromFullPath(FileName)
    Else
        BriefFileName = "Not yet defined"
    End If
    If PathFromFullName(FileName) <> "" Then
        Directory = PathFromFullName(FileName)
    Else
        Directory = "Not yet defined"
    End If
    kb = Description & vbTab & Variable & vbTab & BriefFileName & vbTab & Directory
    .AddItem kb

    Description = "Vessels File"
    Variable = "VesselsFileName"
    FileName = VesselsFileName
    If NameFromFullPath(FileName) <> "" Then
        BriefFileName = NameFromFullPath(FileName)
    Else
        BriefFileName = "Not yet defined"
    End If
    If PathFromFullName(FileName) <> "" Then
        Directory = PathFromFullName(FileName)
    Else
        Directory = "Not yet defined"
    End If
    kb = Description & vbTab & Variable & vbTab & BriefFileName & vbTab & Directory
    .AddItem kb

    Description = "TrappedMsgs File"
    Variable = "TrappedMsgsFileName"
    FileName = TrappedMsgsFileName
    If NameFromFullPath(FileName) <> "" Then
        BriefFileName = NameFromFullPath(FileName)
    Else
        BriefFileName = "Not yet defined"
    End If
    If PathFromFullName(FileName) <> "" Then
        Directory = PathFromFullName(FileName)
    Else
        Directory = "Not yet defined"
    End If
    kb = Description & vbTab & Variable & vbTab & BriefFileName & vbTab & Directory
    .AddItem kb

    Description = "Template File"
    Variable = "TagTemplateReadFile"
    FileName = TagTemplateReadFile
    If NameFromFullPath(FileName) <> "" Then
        BriefFileName = NameFromFullPath(FileName)
    Else
        BriefFileName = "Not yet defined"
    End If
    If PathFromFullName(FileName) <> "" Then
        Directory = PathFromFullName(FileName)
    Else
        Directory = "Not yet defined"
    End If
    kb = Description & vbTab & Variable & vbTab & BriefFileName & vbTab & Directory
    .AddItem kb

    Description = "Overlay Template File"
    Variable = "OverlayTemplateReadFile"
    FileName = OverlayTemplateReadFile
    If NameFromFullPath(FileName) <> "" Then
        BriefFileName = NameFromFullPath(FileName)
    Else
        BriefFileName = "Not yet defined"
    End If
    If PathFromFullName(FileName) <> "" Then
        Directory = PathFromFullName(FileName)
    Else
        Directory = "Not yet defined"
    End If
    kb = Description & vbTab & Variable & vbTab & BriefFileName & vbTab & Directory
    .AddItem kb

    Description = "Overlay Output File"
    Variable = "OverlayOutputFileName"
    FileName = OverlayOutputFileName
    If NameFromFullPath(FileName) <> "" Then
        BriefFileName = NameFromFullPath(FileName)
    Else
        BriefFileName = "Not yet defined"
    End If
    If PathFromFullName(FileName) <> "" Then
        Directory = PathFromFullName(FileName)
    Else
        Directory = "Not yet defined"
    End If
    kb = Description & vbTab & Variable & vbTab & BriefFileName & vbTab & Directory
    .AddItem kb

    Description = "NMEA Input File"
    Variable = "NmeaReadFile"
    FileName = NmeaReadFile
    If NameFromFullPath(FileName) <> "" Then
        BriefFileName = NameFromFullPath(FileName)
    Else
        BriefFileName = "Not yet defined"
    End If
    If PathFromFullName(FileName) <> "" Then
        Directory = PathFromFullName(FileName)
    Else
        Directory = "Not yet defined"
    End If
    kb = Description & vbTab & Variable & vbTab & BriefFileName & vbTab & Directory
    .AddItem kb

    Description = "Startup Log"
    Variable = "StartupLogFile"
    FileName = StartupLogFile
    If NameFromFullPath(FileName) <> "" Then
        BriefFileName = NameFromFullPath(FileName)
    Else
        BriefFileName = "Not yet defined"
    End If
    If PathFromFullName(FileName) <> "" Then
        Directory = PathFromFullName(FileName)
    Else
        Directory = "Not yet defined"
    End If
    kb = Description & vbTab & Variable & vbTab & BriefFileName & vbTab & Directory
    .AddItem kb

    Description = "Exit Log"
    Variable = "ExitLogFile"
    FileName = ExitLogFile
    If NameFromFullPath(FileName) <> "" Then
        BriefFileName = NameFromFullPath(FileName)
    Else
        BriefFileName = "Not yet defined"
    End If
    If PathFromFullName(FileName) <> "" Then
        Directory = PathFromFullName(FileName)
    Else
        Directory = "Not yet defined"
    End If
    kb = Description & vbTab & Variable & vbTab & BriefFileName & vbTab & Directory
    .AddItem kb

'FTP file name now includes path
    Description = "Local FTP File"
    Variable = "FtpLocalFileName"
    FileName = FtpLocalFileName
    If NameFromFullPath(FileName) <> "" Then
        BriefFileName = NameFromFullPath(FileName)
    Else
        BriefFileName = "Not yet defined"
    End If
    If PathFromFullName(FileName) <> "" Then
        Directory = PathFromFullName(FileName)
    Else
        Directory = "Not yet defined"
    End If
    kb = Description & vbTab & Variable & vbTab & BriefFileName & vbTab & Directory
    .AddItem kb

'FTP file name now includes path
    Description = "Remote FTP File"
    Variable = "FtpRemoteFileName"
    FileName = FtpRemoteFileName
    If NameFromFullPath(FileName, "/") <> "" Then
        BriefFileName = NameFromFullPath(FileName, "/")
    Else
        BriefFileName = "Not yet defined"
    End If
    If PathFromFullName(FileName, "/") <> "" Then
        Directory = PathFromFullName(FileName, "/")
    Else
        Directory = "Not yet defined"
    End If
    kb = Description & vbTab & Variable & vbTab & BriefFileName & vbTab & Directory
    .AddItem kb

    Description = "Shell File"
    Variable = "ShellFileName"
    FileName = ShellFileName
    If NameFromFullPath(FileName) <> "" Then
        BriefFileName = NameFromFullPath(FileName)
    Else
        BriefFileName = "Not yet defined"
    End If
    If PathFromFullName(FileName) <> "" Then
        Directory = PathFromFullName(FileName)
    Else
        Directory = "Not yet defined"
    End If
    kb = Description & vbTab & Variable & vbTab & BriefFileName & vbTab & Directory
    .AddItem kb

    Description = "Licence File"
    Variable = "UserLicenceFileName"
    FileName = UserLicenceFileName
    If NameFromFullPath(FileName) <> "" Then
        BriefFileName = NameFromFullPath(FileName)
    Else
        BriefFileName = "Not yet defined"
    End If
    If PathFromFullName(FileName) <> "" Then
        Directory = PathFromFullName(FileName)
    Else
        Directory = "Not yet defined"
    End If
    kb = Description & vbTab & Variable & vbTab & BriefFileName & vbTab & Directory
    .AddItem kb

    Description = "TCP Login File"
    Variable = "TcpLoginFileName"
    FileName = TcpLoginFileName
    If NameFromFullPath(FileName) <> "" Then
        BriefFileName = NameFromFullPath(FileName)
    Else
        BriefFileName = "Not yet defined"
    End If
    If PathFromFullName(FileName) <> "" Then
        Directory = PathFromFullName(FileName)
    Else
        Directory = "Not yet defined"
    End If
    kb = Description & vbTab & Variable & vbTab & BriefFileName & vbTab & Directory
    .AddItem kb


End With
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'V142 Call DisplayQueryUnload(Me.Name, Cancel, UnloadMode)

    If UnloadMode = vbFormControlMenu Then  'V3.4.143 User clicked (X)
        If FormLoaded("TreeFilter") Then   'to prevent NmeaRcv being reloaded
'    If TreeFilter.Visible = True Then      'Never keep show files ticked
            TreeFilter.Check1(13).Value = 0 'Show Files Unticked
            Cancel = True   'just hide
            Files.Hide
'    End If
        End If
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbNormal Then
        With MSFlexGrid1
            .Move 0, 0, ScaleWidth, ScaleHeight
            .ColWidth(.Cols - 1) = ScaleWidth
        End With
    End If
'If ScaleWidth <> 0 Then ResizeColWidth
End Sub

Sub ResizeColWidth()
Dim i As Long
Dim tot As Long
Dim LastVisCol As Long
Dim IntWidth As Long    'internal width if flexgrid
With MSFlexGrid1
'ScaleWidth is the same as the tot of the cols when loaded
'make the size if the grid relative to the form the same as when loaded
    .Width = ScaleWidth
'the external size of the flexgrid is a fixed 100 larger than the sum of the colwidths
    IntWidth = .Width - 100
'make the last column large enough to fit in the grid
'add up total width of columns
    For i = 0 To .Cols - 1
'if this column will fit in grid add it else make it 0
        If tot + .ColWidth(i) < IntWidth Then
            tot = tot + .ColWidth(i)
        Else
            If tot = IntWidth Then
                .ColWidth(i) = 0
            Else
                .ColWidth(i) = IntWidth - tot
                tot = IntWidth
            End If
        End If
        If .ColWidth(i) > 0 Then LastVisCol = i
'If tot > IntWidth Then Stop
 'debug width
'.TextMatrix(1, i) = .ColWidth(i)
    Next i
'if last column too small to fill grid, make it larger
    If tot < IntWidth Then
'last column
        i = LastVisCol
'remove previous column size
        tot = tot - .ColWidth(LastVisCol)
        .ColWidth(i) = IntWidth - tot
'add on new column size
        tot = tot + .ColWidth(i)
 'debug width
'.TextMatrix(1, i) = .ColWidth(i)
'check weve got it right
'        If tot <> intwidth Then Stop
    End If
End With
End Sub

