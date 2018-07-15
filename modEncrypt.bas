Attribute VB_Name = "modEncrypt"
Option Explicit
'Requires CAPICOM V2.1 Project > Reference including

Private Type defLoadLicence
'If any invalid report to user (no), fail means default is used
    ConfigVersion As Long
    ProgramName As String
    MajorVersion As Long
    MinorVersion As Long
    RevisionFrom As Long
    RevisionTo As Long
    UpdatingValidTo As Date  'Last date Config can be Updated
'Valid to Use this File
    ComputerName As String
    UserName As String
    ExpiryDate As Date
'New Settings to use (if valid to use file)
'If 0 the defaults are used (must be Licence in program)
    MaxRcvPerMin As Long    '-1=Default, 0=Unlimited
    MaxFilePerMin As Long    '-1=Default, 0=Unlimited
    sMaxInputFileSize As String    '-1=Default, 0=Unlimited (Long isnt big enough)
    sMaxOutputFileSize As String    '-1=Default, 0=Unlimited
    ExpiryDays As Long     '-1=Default, 0=unlimited,>0 calc ConfigExpires when loaded
'Memo information (set up when file sent to user)
    DateIssued As Date
    IssuedTo As String  'email address
    FileBlockLen As Long    '""=Default, 0=Unlimited
End Type

Private FileLicence As defLoadLicence  'The Config we are using

Public Type defUserLicence   'Constructed when Licence.cfg is read
'Settings to Throttle input
    FileBlockSize As Long    '0=default
    sMaxInputFileSize As String    '"0"=Unlimited (can be > 2^32)
    sMaxOutputFileSize As String    '"0"=Unlimited (ditto)
    MaxRcvPerMin As Long    '0=Unlimited
    MaxFilePerMin As Long
    ExpiryDate As Date 'Report to user on exit
    DateIssued As Date
    UpdatingValidTo As Date
    IssuedTo As String
    ComputerName As String
    UserName As String
End Type

Public UserLicence As defUserLicence

Private EncryptedFileName As String
Private DecryptedFileName As String

Public Function Decrypt(kb As String) As String
    Dim Secret As EncryptedData
    
    Set Secret = New EncryptedData
    Secret.Algorithm.Name = CAPICOM_ENCRYPTION_ALGORITHM_AES
    Secret.Algorithm.KeyLength = CAPICOM_ENCRYPTION_KEY_LENGTH_256_BITS
    Secret.SetSecret "My Secret Encryption Key"
'    Secret.Content = "password" ' just so we know that this is being reset by decryption
    On Error Resume Next    'errror if .secret differs
    Secret.Decrypt kb
    On Error GoTo 0
    Decrypt = Secret.Content
    Set Secret = Nothing
'MsgBox Decrypt
End Function

Public Function Encrypt(kb As String) As String
Dim Secret As EncryptedData
    If kb = "" Then Exit Function
    Set Secret = New EncryptedData
    Secret.Algorithm.Name = CAPICOM_ENCRYPTION_ALGORITHM_AES
    Secret.Algorithm.KeyLength = CAPICOM_ENCRYPTION_KEY_LENGTH_256_BITS
    Secret.SetSecret "My Secret Encryption Key"
    Secret.Content = kb ' what we want to encrypt
    Encrypt = Secret.Encrypt
'For Password encryption (AisDecoder)
'we must remove the split lines secret.content includes
'    Encrypt = Replace(Secret.Encrypt, vbCrLf, "")
'    MsgBox Encrypt
    Set Secret = Nothing
End Function

Public Function DecryptFile(EncryptedFileName As String, DecryptedFileName As String)
Dim EncryptedLines As String
Dim DecryptedLines As String
Dim ch As Long

'MsgBox "Decrypting " & EncryptedFileName & vbCrLf & "to " & DecryptedFileName
    If FileExists(EncryptedFileName) Then
        ch = FreeFile
        Open EncryptedFileName For Input As #ch
        EncryptedLines = StrConv(InputB(LOF(ch), ch), vbUnicode)
        Close ch
        DecryptedLines = Decrypt(EncryptedLines)
        Open DecryptedFileName For Output As #ch Len = Len(DecryptedLines)
        Print #ch, DecryptedLines
        Close #ch
    End If
End Function

Public Function EncryptFile(DecryptedFileName As String, EncryptedFileName As String)
Dim EncryptedLines As String
Dim DecryptedLines As String
Dim ch As Long
Dim l As Integer

'MsgBox "Encrypting " & DecryptedFileName & vbCrLf & "to " & EncryptedFileName
    ch = FreeFile
    Open DecryptedFileName For Input As #ch
    DecryptedLines = StrConv(InputB(LOF(ch), ch), vbUnicode)
    Close ch
    EncryptedLines = Encrypt(DecryptedLines)
    Open EncryptedFileName For Output As #ch '    Len = Len(EncryptedLines)
    Print #ch, EncryptedLines
    Close #ch
End Function

Public Function EncryptFiles(FilePath As String, DecryptedExt As String, EncryptedExt As String)
Dim DecryptedFileName As String
Dim EncryptedFileName As String

'MsgBox FilePath
    DecryptedFileName = Dir$(FilePath & "*" & DecryptedExt)
    Do While DecryptedFileName > ""
        If Right$(DecryptedFileName, Len(DecryptedExt)) = DecryptedExt Then
            EncryptedFileName = Replace(DecryptedFileName, DecryptedExt, EncryptedExt)
            Call EncryptFile(FilePath & DecryptedFileName, FilePath & EncryptedFileName)
        End If
        DecryptedFileName = Dir$
    Loop
End Function

Public Sub LoadLicence()
Dim ch As Long
Dim chOut As Long
Dim kb As String
Dim arry() As String
Dim Fail As Boolean
Dim ChangedDecryptedFileName As String
Dim CopyFileName As String  'Keep a copy for jna in /Users/ folder
Dim i As Long
Dim CreatingUserLicenceFile As Boolean
Dim InstallingLicence As Boolean

    If FolderExists(Environ("AppData") & "\Arundale\Ais Decoder\Users\") Then
        CreatingUserLicenceFile = True
    End If
    
    EncryptedFileName = Environ("AppData") & "\Arundale\Ais Decoder\" _
    & "Licence" & ".cfg"
    UserLicenceFileName = EncryptedFileName 'used to display location in Show Files
'Create an encrypted version of my Licence file
'This will be created in
    If CreatingUserLicenceFile = True Then
Debug.Print "Creating User Config"
        DecryptedFileName = Environ("AppData") _
        & "\Arundale\Ais Decoder\Users\Licence" & ".txt"
        Call EncryptFile(DecryptedFileName, EncryptedFileName)
    End If
    
'Decrypt the users (or Jna's) Licence.ini file
    DecryptedFileName = Replace(EncryptedFileName, ".cfg", ".tmp")
    Call DecryptFile(EncryptedFileName, DecryptedFileName)
    
'Read the Decrypted file and create FileLicence (only in modEncrypt)
'Then Create UserLicence(Global) containing any data the Program requires
'If within UpdatingValidTo, update the config file
    With FileLicence
        If FileExists(DecryptedFileName) = False Then
Debug.Print "No User Config"
            Fail = True
        Else
Debug.Print "User Config Exists"
            ch = FreeFile
            Open DecryptedFileName For Input As #ch
            Do Until EOF(ch)
                Line Input #ch, kb
'remove tabs
                kb = Replace(kb, vbTab, "")
'Remove comments
                i = InStr(1, kb, "'")
                If i > 0 Then
                    kb = Left$(kb, i - 1)
                End If
                arry = Split(kb, ",")
                If kb <> "" And UBound(arry) >= 1 Then
                    Select Case arry(0)
'Valid to Use or Update file
'If any invalid report to user (no)
                    Case Is = "ConfigVersion"
                        .ConfigVersion = CLng(arry(1))
                    Case Is = "ProgramName"
                        .ProgramName = arry(1)
                    Case Is = "MajorVersion"
                        .MajorVersion = NullToZero(arry(1))
                    Case Is = "MinorVersion"
                        .MinorVersion = NullToZero(arry(1))
                    Case Is = "RevisionFrom"
                        .RevisionFrom = NullToZero(arry(1))
                    Case Is = "RevisionTo"
                        .RevisionTo = NullToZero(arry(1))
'Valid to Update Config file upto this date (set 3 days after config sent to user)
                    Case Is = "UpdatingValidTo"
                        If IsDate(arry(1)) Then
                            .UpdatingValidTo = arry(1)
'Valid to Use File
                        End If
                    Case Is = "ComputerName"
                        .ComputerName = arry(1)
                    Case Is = "UserName"
                        .UserName = arry(1)
'-1=Default, 0=unlimited,>0 calc ConfigExpires when loaded                    Case Is = "ExpiryDays"
                    Case Is = "ExpiryDays"
                        .ExpiryDays = arry(1)
'New Settings to use (if valid to use file)
                    Case Is = "ExpiryDate"
                        .ExpiryDate = arry(1)
                    Case Is = "MaxRcvPerMin"        '""=Default, "0"=Unlimited
                        .MaxRcvPerMin = arry(1)     '""=Default, "0"=Unlimited
                    Case Is = "MaxFilePerMin"        '""=Default, "0"=Unlimited
                        .MaxFilePerMin = arry(1)     '""=Default, "0"=Unlimited
                    Case Is = "MaxInputFileSize"    '""=Default, "0"=Unlimited
                        .sMaxInputFileSize = arry(1)
                    Case Is = "MaxOutputFileSize"
                        .sMaxOutputFileSize = arry(1)
'Memo info
                    Case Is = "DateIssued"
                        If IsDate(arry(1)) Then
                            .DateIssued = arry(1)
                        Else
                            .DateIssued = Date
                        End If
                    Case Is = "IssuedTo"    'email address
                        .IssuedTo = arry(1)
                    Case Is = "FileBlockLen"
                        .FileBlockLen = arry(1)
                    Case Else
MsgBox arry(0) & " is invalid"
                    End Select
                End If  'Got an argument
            Loop
            Close #ch
'Finished with using the Decrypted file so get rid of it
            Kill DecryptedFileName
                
            If App.EXEName <> .ProgramName Then Fail = True
            If App.Major <> .MajorVersion Then Fail = True
            If App.Minor <> .MinorVersion Then Fail = True
            If .RevisionFrom > 0 And App.Revision < .RevisionFrom Then
                Fail = True
            End If
            If .RevisionTo > 0 And App.Revision > .RevisionTo Then
                Fail = True
            End If

'Fail determines if we use the default settings or the Licence file
'Valid to Update file - were going to modify the User's Licence file
'If jna we also copy the users config to the Users folder
'The encrypted file we should email to the user, who needs to place it in
'%AppData%\Arundale\AisDecoder\Licence.cfg
'If the user's config file is valid we update it
            If Fail = False And Date <= .UpdatingValidTo Then
                If .ExpiryDate = "00:00:00" Then
                    InstallingLicence = True
                End If
                ChangedDecryptedFileName = DecryptedFileName & "_newuser"
                chOut = FreeFile
                Open ChangedDecryptedFileName For Output As #chOut
                kb = "ConfigVersion," & .ConfigVersion
                Print #chOut, kb
                kb = "ProgramName," & .ProgramName
                Print #chOut, kb
                kb = "MajorVersion," & .MajorVersion
                Print #chOut, kb
                kb = "MinorVersion," & .MinorVersion
                Print #chOut, kb
                kb = "RevisionFrom," & .RevisionFrom
                Print #chOut, kb
                kb = "RevisionTo," & .RevisionTo
                Print #chOut, kb
                kb = "UpdatingValidTo," & .UpdatingValidTo  'Last date used can be setup"
                Print #chOut, kb
'Valid to Use this File
                If CreatingUserLicenceFile = False Then
                    .ComputerName = Environ$("ComputerName")
                End If
                kb = "ComputerName," & .ComputerName
                Print #chOut, kb
                If CreatingUserLicenceFile = False Then
                    .UserName = Environ$("UserName")
                End If
                kb = "UserName," & .UserName
                Print #chOut, kb
'ConfigExpires ExpiryDays after the config file has been updated
'When first installed ConfigExpires will be
'                .ConfigExpires = DateAdd("d", .ExpiryDays, Date)
'                kb = "ConfigExpires," & .ConfigExpires   'Set to months/years after which licence expires"
'                Print #chOut, kb
'New Settings to use (if valid to use file)
'If 0 the defaults are used (must be setup in program)
                kb = "MaxRcvPerMin," & .MaxRcvPerMin
                Print #chOut, kb
                kb = "MaxFilePerMin," & .MaxFilePerMin
                Print #chOut, kb
                kb = "MaxInputFileSize," & .sMaxInputFileSize
                Print #chOut, kb
                kb = "MaxOutputFileSize," & .sMaxOutputFileSize
                Print #chOut, kb
                kb = "ExpiryDays," & .ExpiryDays
                Print #chOut, kb
'Memo information (set up when file sent to user)
                kb = "DateIssued," & .DateIssued
                Print #chOut, kb
                kb = "IssuedTo," & .IssuedTo  'email address"
                Print #chOut, kb
                kb = "FileBlockLen," & .FileBlockLen
                Print #chOut, kb
'-1=Default, 0=unlimited (ExpiryDays) (< -1 allows me to backdate for testing)
                If .ExpiryDays > 0 Or .ExpiryDays < -1 Then     'not default or unlimited
'00:00:00=unlimited (ExpiryDate)
                    FileLicence.ExpiryDate = DateAdd("d", .ExpiryDays, Date)
                    kb = "ExpiryDate," & .ExpiryDate
'The userLicencefile must not have an expiry date set as we use
'a blank date to determine if the user has used this Licence file
                    If CreatingUserLicenceFile = False Then
                        Print #chOut, kb
                    End If
                Else
'If it is default or unlimited we do not write out an expiry date
'because we use ExpiryDays
                End If
                Close #chOut
                Call EncryptFiles(Environ("AppData") & "\Arundale\Ais Decoder\", ".tmp_newuser", ".cfg")
'Finished with the ChangedDecryptedFile
'Stop   'Stop here to check decrypted file (Licence.tmp_newuser)
                Kill ChangedDecryptedFileName
            End If
                    
'If Users Folder exits (jna only)
'Keep a copy of the en(.cfg) and de(.txt) crypted files in the Users Folder
            If CreatingUserLicenceFile = True Then
'Copy as Licence.cfg
                CopyFileName = Replace(EncryptedFileName, "\Ais Decoder\", "\Ais Decoder\Users\")
                FileCopy EncryptedFileName, CopyFileName
'Copy as "issued to"
                CopyFileName = Replace(CopyFileName, "\Licence", "\" & .IssuedTo)
                FileCopy EncryptedFileName, CopyFileName
MsgBox "Creating Licence.cfg & .txt" & vbCrLf & .IssuedTo & ".cfg & .txt", , "LoadLicence"
'Keep a decrypted copy as well as issued to
                Call DecryptFile(CopyFileName, Replace(CopyFileName, ".cfg", ".txt"))
            
            End If  'Updating config file
        
'        If Environ$("ComputerName") <> .ComputerName Then Fail = True
'        If Environ$("UserName") <> .UserName Then Fail = True
        
        End If  'Using Config File
    End With

'Create UserLicence containing any data the Program requires Globally
'From Config(if valid) else use the defaults

    With UserLicence
'set defaults
        .FileBlockSize = 0
        .sMaxInputFileSize = CDec("50000000")  '50MB
        .MaxRcvPerMin = 200
        .MaxFilePerMin = 2000
'Check if the ConfigFile has expired
'-1=Default, 0=unlimited (ExpiryDays)
        If FileLicence.ExpiryDays <> 0 Then     'not default
'00:00:00=unlimited (ExpiryDate)
            .ExpiryDate = FileLicence.ExpiryDate
            If .ExpiryDate < Date Then
                Fail = True
Debug.Print "Licence Expired"
            Else
Debug.Print "Licence Expires " & .ExpiryDate
            End If
        Else
Debug.Print "Licence Never Expires"
        End If
        
        If Fail = False And .ComputerName <> "" And Environ$("ComputerName") <> .ComputerName Then
kb = "You have transferred this file from another PC" & vbCrLf
kb = kb & "To install your enhanced options on another PC please" & vbCrLf
kb = kb & "email myself at neal@arundale.com and I will send" & vbCrLf
kb = kb & "you another configuration File" & vbCrLf
        MsgBox kb, vbInformation, "Computer Name Mis-match"
            Fail = True
        End If
'        If Environ$("UserName") <> .UserName Then Fail = True
        
        .DateIssued = FileLicence.DateIssued
        .UpdatingValidTo = FileLicence.UpdatingValidTo
        .IssuedTo = FileLicence.IssuedTo
        .ComputerName = FileLicence.ComputerName
        .UserName = FileLicence.UserName
'If Fail we do not use the config file for the settings we use
        If Fail = False Then
            If FileLicence.FileBlockLen > 0 Then
'0=default
                .FileBlockSize = FileLicence.FileBlockLen
            End If
            If FileLicence.MaxRcvPerMin <> -1 Then   'not default
'0=unlimited
                .MaxRcvPerMin = FileLicence.MaxRcvPerMin
            End If
            If FileLicence.MaxFilePerMin <> -1 Then   'not default
'0=unlimited
                .MaxFilePerMin = FileLicence.MaxFilePerMin
            End If
            If FileLicence.sMaxInputFileSize <> "-1" Then
'"0"=unlimited
                .sMaxInputFileSize = CDec(FileLicence.sMaxInputFileSize)
            End If
            If FileLicence.sMaxOutputFileSize <> "-1" Then
'"0"=unlimited
                .sMaxOutputFileSize = CDec(FileLicence.sMaxOutputFileSize)
            End If
            If InstallingLicence = True Then
                MsgBox "Licence Installed", vbInformation, "Licence Update"
            End If
        End If
Debug.Print "MaxFileSize " & aByte(.sMaxInputFileSize)
    End With
'Stop    'Stop here to check UserLicence
End Sub
