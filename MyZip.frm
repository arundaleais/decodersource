VERSION 5.00
Begin VB.Form MyZip 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "MyZip.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "MyZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_cZip As cZip
Attribute m_cZip.VB_VarHelpID = -1
'http://www.vbaccelerator.com/home/VB/Code/Libraries/Compression/Zipping_Files/article.asp

'Public Function ZipUp(ZippedFileName As String)
'    Call ZipMeUp(ZippedFileName)
'End Function

Public Function ZipUp(ZippedFilename As String)
'the UnZippedFileName is the ZippedFile with ".unzipped" appended
'eg GooglEarth.kmz.unzipped
Dim kb As String
Dim UnZippedFile As String  'Just the file name
Dim OverlayFile As String   'ditto
Dim ZippedFile As String    'ditto
Dim Path As String          'The Path extracted from the UnzippedFileName
Dim DefaultTemplate As Boolean   'True if GoogleEarth.kml

'infil is OutputFileName (.kmz) (when its actually a kml file)
'overlay output file is OverlayTemplateReadFile + _link (.kml)
'On Exitting this routine
'the OutputFile (.kmz) contains
'    OutputFilename (.kml) and OverlayFile(_link.kml)
'Google Earth requires the zipped up file to be a kml file
'even though its actually a ZIP archive - may contain multiple files
'MsgBox "Zip ZippedFileName =" & ZippedFileName
    If NameFromFullPath(TagTemplateReadFile) = "GoogleEarth.kml" Then
        DefaultTemplate = True
    End If
    Path = PathFromFullName(ZippedFilename) & "\"
Debug.Print "Path " & Path
    ZippedFile = NameFromFullPath(ZippedFilename)
    UnZippedFile = ZippedFile
    UnZippedFile = Replace(UnZippedFile, ".kmz", ".kml")
    UnZippedFile = Replace(UnZippedFile, ".zip", ".uzip")
'Debug.Print "#ZipUp " & ZippedFile
'Debug.Print "UnZipped " & UnZippedFile
'Stop
    If FileExists(Path & ZippedFile) Then
        Kill Path & ZippedFile  'delete any existing file
    End If
    
    OverlayFile = NameFromFullPath(OverlayOutputFileName)
'The above must all be set BEFORE we SET the new czip
'This is what was causing the problem with the output
'zipped file not being created.

' Set the zip file:
    Set m_cZip = New cZip
   ' Make sure any previously zipped files are cleared:
    m_cZip.ClearFileSpecs
'set the base path as the same as the ZippedFileName
    m_cZip.BasePath = Path
    m_cZip.ZipFile = ZippedFile  '.kmz output (Name same as InputFile)
    m_cZip.AddFileSpec UnZippedFile
    If OverlayReq(1) = True Then    'File Output
        m_cZip.AddFileSpec "triangle.png"
        m_cZip.AddFileSpec "square.png"
        m_cZip.AddFileSpec "ScreenOverlay.png"
        m_cZip.AddFileSpec OverlayFile  '(LINK file name)
    Else
        If DefaultTemplate Then
'Include Default Icon (otherwise (X) is displayed)
            m_cZip.AddFileSpec "ship1.png"
        End If
    End If

   m_cZip.Zip

   ' Check for success failure:
   If Not (m_cZip.Success) Then
'Stop
      ' Zip failed.  One of the notifications will have
      ' provided the reason.
      ' e.g. can't write output file, can't find any
      '      matching files
    End If
'delete the unzipped file
'    Kill path & "/" & UnZippedFile
'the OutputFile (.kmz) is now a zip archive containing both
'the output .kml file and the overlay .kml file
    Set m_cZip = Nothing
End Function

Private Sub Form_Load()
'   Set m_cZip = New cZip
End Sub
