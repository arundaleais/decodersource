VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HugeBinaryFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'http://www.vbforums.com/showthread.php?531321-VB6-Huge-%28-gt-2GB%29-File-I-O-Class
'See #12 http://www.vbforums.com/attachment.php?attachmentid=84522&d=1309214093

Option Explicit

Public Event PercentDone(ByRef Data As String, ByVal Sentences As Long, ByRef Cancel As Boolean)
' Storage for the debug ID.
Private mlngDebugID As Long
'
'HugeBinaryFile
'==============
'
'A class for doing simple binary I/O on very large disk files
'(well over the usual 2GB limit).  It only does I/O using Byte
'arrays, and makes use of Currency values that are scaled to
'whole numbers in places:
'
'   For a file of one byte the FileLen property returns 1.0000 as
'   its value.
'
'Operation is similar in many ways to native VB Get#/Put# I/O, for
'example the EOF property must be checked after a ReadBytes() call.
'You must also Dim/Redim buffers to desired sizes before calling
'ReadBytes() or WriteBytes().
'
'Short (signed Long) relative seeks and long (unsigned Currency)
'absolute seeks from 0 may be done.
'
'AutoFlush may be set True to force buffer flushes on every write.
'The Flush() method may be called explicitly if necessary.
'

Public Enum HBF_Errors
  HBF_UNKNOWN_ERROR = 45600
  HBF_FILE_ALREADY_OPEN
  HBF_OPEN_FAILURE
  HbF_SEEK_FAILURE
  HBF_FILELEN_FAILURE
  HBF_READ_FAILURE
  HBF_WRITE_FAILURE
  HBF_FILE_ALREADY_CLOSED
End Enum

Private Const HBF_SOURCE = "HugeBinaryFile"

Private Const GENERIC_WRITE As Long = &H40000000
Private Const GENERIC_READ As Long = &H80000000
Private Const FILE_ATTRIBUTE_NORMAL As Long = &H80&
Private Const CREATE_ALWAYS = 2
Private Const OPEN_ALWAYS = 4
Private Const INVALID_HANDLE_VALUE = -1
Private Const INVALID_SET_FILE_POINTER = -1
Private Const INVALID_FILE_SIZE = -1
Private Const FILE_SHARE_READ = &H1

Private Const FILE_BEGIN = 0, FILE_CURRENT = 1, FILE_END = 2

Private Type MungeCurr
    Value As Currency
End Type

Private Type Munge2Long
    LowVal As Long
    HighVal As Long
End Type

Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" ( _
    ByVal dwFlags As Long, _
    lpSource As Long, _
    ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, _
    ByVal lpBuffer As String, _
    ByVal nSize As Long, _
    Arguments As Any) As Long

Private Declare Function ReadFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToRead As Long, _
    lpNumberOfBytesRead As Long, _
    ByVal lpOverlapped As Long) As Long

Private Declare Function CloseHandle Lib "kernel32" ( _
    ByVal hObject As Long) As Long

Private Declare Function GetFileSize Lib "kernel32" ( _
    ByVal hFile As Long, _
    lpFileSizeHigh As Long) As Long

Private Declare Function WriteFile Lib "kernel32" ( _
    ByVal hFile As Long, _
    lpBuffer As Any, _
    ByVal nNumberOfBytesToWrite As Long, _
    lpNumberOfBytesWritten As Long, _
    ByVal lpOverlapped As Long) As Long

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" ( _
    ByVal lpFileName As String, _
    ByVal dwDesiredAccess As Long, _
    ByVal dwShareMode As Long, _
    ByVal lpSecurityAttributes As Long, _
    ByVal dwCreationDisposition As Long, _
    ByVal dwFlagsAndAttributes As Long, _
    ByVal hTemplateFile As Long) As Long

Private Declare Function SetFilePointer Lib "kernel32" ( _
    ByVal hFile As Long, _
    ByVal lDistanceToMove As Long, _
    lpDistanceToMoveHigh As Long, _
    ByVal dwMoveMethod As Long) As Long

Private Declare Function FlushFileBuffers Lib "kernel32" ( _
    ByVal hFile As Long) As Long

Private hFile As Long
Private sFName As String
Private fAutoFlush As Boolean
Private fEOF As Boolean
Private C As MungeCurr
Private l As Munge2Long
'Private iFullReads As Long
'Private cLeftBytes As Currency
'Private iLeftBytes As Long
Private bStopRun As Boolean
Private lBlocksRead As Long
Private bytBuf() As Byte    'Buffer for Getting Chunks of File
Private blockbytesread As Long 'No of bytes read into bytBuf

Private ascBuf As String  'Buffer same as bytBuf converted to Ascii
Private ascBufLines As Long  'No of LF in the buffer (determins when to read another Block)
Private dtStartTime As Date
Private dtStartDate As Date
Private cTotalBytesRead As Currency 'Total read from file
Private cTotalLinesRead As Currency 'Total read from file
Private fXoff As Boolean
Private skb As String

Private Const DEFAULT_BLOCK_SIZE = 2 ^ 16 '=65536    131,072=2^17   1,048,576=2^20
Private lBlockSize As Long  'Actual block size
Private blnCancel As Boolean    'stop input
Private blnPause As Boolean     'Pause input

'Public Property Get NmeaRcv() As NmeaRcv
'    Set NmeaRcv = mNmeaRcv
'End Property

'Public Property Set NmeaRcv(ByVal NewNmeaRcv As NmeaRcv)
'    Set mNmeaRcv = NewNmeaRcv
'End Property

Public Property Let LongTaskCancel(ByVal NewVal As Boolean)
    RaiseErrorIfClosed
    blnCancel = NewVal
End Property

Public Property Get AutoFlush() As Boolean
    RaiseErrorIfClosed
    AutoFlush = fAutoFlush
End Property

Public Property Get BlocksRead() As Long
    RaiseErrorIfClosed
    BlocksRead = lBlocksRead
End Property

Public Property Let AutoFlush(ByVal NewVal As Boolean)
    RaiseErrorIfClosed
    fAutoFlush = NewVal
End Property

Public Property Get FileHandle() As Long
    RaiseErrorIfClosed
    FileHandle = hFile
End Property

Public Property Get FileLen() As Currency
    RaiseErrorIfClosed
'Get File Size returns Low Value and High value
'But if file sise is too small to have a high value the high value is set to 0
    l.LowVal = GetFileSize(hFile, l.HighVal)
    If l.LowVal = INVALID_FILE_SIZE Then
        If err.LastDllError Then RaiseError HBF_FILELEN_FAILURE
    End If
    LSet C = l
    FileLen = C.Value * 10000@
End Property

Public Property Get FileName() As String
    RaiseErrorIfClosed
    FileName = sFName
End Property

Public Property Get EOF() As Boolean
    RaiseErrorIfClosed
    EOF = fEOF
End Property

Public Property Let EOF(ByVal NewVal As Boolean)
    RaiseErrorIfClosed
    fEOF = NewVal
End Property

Public Property Get IsOpen() As Boolean
    IsOpen = hFile <> INVALID_HANDLE_VALUE
End Property

Public Property Get RunSeconds() As Long
RaiseErrorIfClosed
    RunSeconds = DateDiff("s", dtStartTime, Now())
End Property

Public Property Get aRunTime() As String
Dim Seconds As Long
Dim d As Long, h As Long, m As Long, S As Long
Dim rtn As String
Dim dtStopTime As Date

RaiseErrorIfClosed
'   dtStartTime = "1-Sep-14 18:00"  'test
'   dtStartTime = Now()             'test
dtStopTime = Now()
Seconds = DateDiff("s", dtStartTime, dtStopTime)
   d = Seconds \ 86400           'days
   h = (Seconds \ 3600) Mod 24   'hours
   m = (Seconds \ 60) Mod 60     'minutes
   S = Seconds Mod 60            'seconds

   h = h Mod 24
   If d Then rtn = FormatUnits(d, "day") & ", "
   If (h <> 0) Or (Len(rtn) > 0) Then rtn = rtn & FormatUnits(h, "hour") & ", "
   If (m <> 0) Or (Len(rtn) > 0) Then rtn = rtn & FormatUnits(m, "minute") & ", "
   If (S <> 0) Or (Len(rtn) > 0) Then rtn = rtn & FormatUnits(S, "second")
    If rtn = "" Then rtn = "under 1 second"
    
    aRunTime = rtn
    
End Property

Private Function FormatUnits(ByVal Value As Long, ByVal UnitName As String, Optional Pad As Boolean = False) As String
   FormatUnits = _
      Format$(Value, String$(Abs(Pad) + 1, "0")) & _
      " " & UnitName & IIf(Value <> 1, "s", "")
End Function

Public Property Get TotalBytesRead() As String
    RaiseErrorIfClosed
    TotalBytesRead = CDec(cTotalBytesRead * 10000)
End Property

Public Property Get TotalLinesRead() As String
    RaiseErrorIfClosed
    TotalLinesRead = CDec(cTotalLinesRead) * 10000
End Property

Public Property Let Pause(ByVal NewVal As Boolean)
    RaiseErrorIfClosed
    blnPause = NewVal
Debug.Print "InputFile.Pause=" & blnPause
End Property


#If False Then
Public Property Let LetAscbuf(ByVal NewVal As String)
    RaiseErrorIfClosed
    ascBuf = NewVal
End Property

Public Function Get_Data() As String
    RaiseErrorIfClosed
    Get_Data = ascBuf
    ascBuf = ""
End Function
#End If

Public Function Xoff()  'Call Xoff to turn off getting data (set Xoff=true)
    RaiseErrorIfClosed
    If fXoff = False Then
'Must action immediately to force exit from within ReadSequential DO loop
        fXoff = True
Debug.Print "Xoff (ReadSequential)"
    End If
End Function

Public Function Xon()   'Call Xon to turn on getting data (set Xoff=false)
    
    RaiseErrorIfClosed
'Only take action if not currently torned off, otherwise call to readSequential
'could be re-entrant (called from itslf)
    If fXoff = True Then
        fXoff = False
Debug.Print "Xon (ReadSequential)"
'        Call NextFileBlock(ascBuf)
    End If
End Function
Public Sub CloseFile()
'HugeWaitableTimer must be nothing
    RaiseErrorIfClosed
    CloseHandle hFile
    sFName = ""
    fAutoFlush = False
    fEOF = False
    cTotalBytesRead = 0@
    cTotalLinesRead = 0@
    hFile = INVALID_HANDLE_VALUE
End Sub

Public Sub Flush()
    RaiseErrorIfClosed
    FlushFileBuffers hFile
End Sub

Public Sub OpenFile(ByVal OpenFileName As String, Optional ByVal Mode As String, Optional BlockSize As Long)
'Debug.Print "#OpenFile " & Mid$(OpenFileName, InStrRev(OpenFileName, "\") + 1)
    sFName = OpenFileName
    If hFile <> INVALID_HANDLE_VALUE Then
        RaiseError HBF_FILE_ALREADY_OPEN
    End If
    hFile = CreateFile(OpenFileName, GENERIC_WRITE Or GENERIC_READ, 0, _
                       0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
    If hFile = INVALID_HANDLE_VALUE Then
        RaiseError HBF_OPEN_FAILURE
    Else
Debug.Print "Size " & FileLen
'        iFullReads = Int(FileLen / DEFAULT_BLOCK_SIZE)
'        If iFullReads > 0 Then ReDim bytBuf(1 To DEFAULT_BLOCK_SIZE) As Byte
'        cLeftBytes = FileLen
'        iLeftBytes = Int(cLeftBytes / DEFAULT_BLOCK_SIZE)
'        While cLeftBytes >= 0
'            cLeftBytes = cLeftBytes - DEFAULT_BLOCK_SIZE
'        Wend
  
'        iLeftBytes = cLeftBytes + DEFAULT_BLOCK_SIZE  ' Restore the underflow to get the correct answer.
'        If iFullReads = 0 Then ReDim bytBuf(1 To iLeftBytes) As Byte
'Set default buffer size
        If BlockSize = 0 Then
            lBlockSize = DEFAULT_BLOCK_SIZE
        Else
            lBlockSize = BlockSize
        End If
        ReDim bytBuf(1 To lBlockSize) As Byte
        dtStartTime = Now()
'Debug.Print "Reading " & iFullReads & " chunks + " & iLeftBytes & " bytes"
        If Mode = "Sequential" Then
'            Call NextFileBlock(ascBuf)
        End If
    End If
End Sub

Public Sub OpenFileAppend(ByVal OpenFileName As String)
Debug.Print "Append " & Mid$(OpenFileName, InStrRev(OpenFileName, "\") + 1)
    
    sFName = OpenFileName
    If hFile <> INVALID_HANDLE_VALUE Then
        RaiseError HBF_FILE_ALREADY_OPEN
    End If
    hFile = CreateFile(OpenFileName, GENERIC_WRITE, FILE_SHARE_READ, _
                       0, OPEN_ALWAYS, FILE_ATTRIBUTE_NORMAL, 0)
        
    If hFile = INVALID_HANDLE_VALUE Then
        RaiseError HBF_OPEN_FAILURE
    End If
    SeekEnd
ProcExit:
Exit Sub

error_OpenFileAppend:
    Resume ProcExit
End Sub

Public Sub LongTask(MaxFilePerTimerCycle As Long, ByVal Duration As Single, Optional Blocks As Long)
Dim sngThreshold As Single
Dim sngStart As Single
Dim Data As String
Dim SentThisTimerCycle As Long
Dim GetDataCount As Long
Dim Elapsed As Double
Dim MyClock As clsTimer
'    Set MyClock = New clsTimer
'    Set HugeWaitableTimer = New clsWaitableTimer
    
'    MyClock.StartTimer
    Do
        If blnPause = False Then
            GetDataCount = GetData(Data, MaxFilePerTimerCycle - SentThisTimerCycle, Blocks)
'Returned will be either the Required or can be under the required if
'EOF, blnCancel (because stop pressed)
'Returned will be Less than Required if Required is over one FileReadBlock
'in this case the duration will probably be less than the cyclespeed
Debug.Print "LongTask " & GetDataCount
'This actually calls NmeaRcv.InputFile_PercentDone
            RaiseEvent PercentDone(Data, GetDataCount, blnCancel)
            SentThisTimerCycle = SentThisTimerCycle + GetDataCount
            If SentThisTimerCycle = MaxFilePerTimerCycle Then
                Exit Sub
            End If
'If blnCancel is set to true in InputFile_PercentDone it will be true
'Immediately it is set in the Event
        
            If blnCancel Then
                Exit Sub
            End If
        End If  'not paused
    Loop Until GetDataCount = 0 Or blnCancel
End Sub

Public Function GetData(Data As String, Optional Sentences As Long, Optional Blocks As Long) As Long
Dim strBlock As String
Dim strBlockLines As Long
Dim OutputLen As Long
Dim OutputSentences As Long
Dim i As Long
Dim j As Long
Dim BlockCount As Long

'RcvSpeedClock  use for timeout
Debug.Print LFCount(Data)
'If unlimited the Sentences=0 so read the next block
    Do Until (Sentences And ascBufLines >= Sentences) Or fEOF
        ReadBytes bytBuf       'Get next Chunk of file in bytes
'Convert to Vb AND remove any bytes that are on disk after the EOF
        BlockCount = BlockCount + 1
        strBlock = Left$(StrConv(bytBuf, vbUnicode), blockbytesread)
        For i = 1 To Len(strBlock)
            If Mid$(strBlock, i, 1) = vbLf Then
                strBlockLines = strBlockLines + 1
                OutputLen = i   'Will be changed if Sentences > 0
            End If
        Next i
        cTotalLinesRead = cTotalLinesRead + strBlockLines * 0.0001@
        ascBuf = ascBuf + strBlock
        ascBufLines = ascBufLines + strBlockLines
        strBlockLines = 0
Debug.Print Len(ascBuf) & "(" & ascBufLines & ")"
        If BlockCount > Blocks Then Exit Do 'Max 1 block
    Loop

'Get where to output to and the no of lines
    If Sentences > 0 Then
        j = 0
        For i = 1 To Len(ascBuf)
            If Mid$(ascBuf, i, 1) = vbLf Then
                j = j + 1
                If j = Sentences Then
                    OutputSentences = j
                    Exit For
                End If
            End If
        Next i
        OutputLen = i
        OutputSentences = j
    
'Stop
'if insufficient LF 's then return position of the last one (if any)
        If OutputSentences < Sentences Then
            If OutputLen = 0 And Len(ascBuf) > 201 Then
                OutputLen = 201
            End If
        End If
    Else
        OutputSentences = ascBufLines
    End If
    
    Data = Left$(ascBuf, OutputLen)
    GetData = OutputSentences
    ascBuf = Mid$(ascBuf, OutputLen + 1)
    ascBufLines = ascBufLines - OutputSentences
'Stop
End Function

'Retorn ascBuf to FileDataArrival

Public Function NextFileBlock(ByRef ascBuf As String) As Long
Dim cpos As Long
Dim strBlock As String

'If ascBuf <> "" Then Stop   'check if clear is required
    ascBuf = ""
    Do
        ReadBytes bytBuf       'Get next Chunk of file in bytes
'Convert to Vb AND remove any bytes that are on disk after the EOF
        strBlock = Left$(StrConv(bytBuf, vbUnicode), blockbytesread)
        cTotalLinesRead = cTotalLinesRead + LFCount(strBlock) * 0.0001@
        
        ascBuf = ascBuf + strBlock
        If blockbytesread Then
            Exit Do
        End If
    Loop Until fEOF
    NextFileBlock = Len(ascBuf)
End Function

#If False Then
Private Function ReadSequential_old()
Dim cpos As Long
Dim strBlock As String
    
'Stop
If ascBuf <> "" Then Stop   'check if clear is required
    ascBuf = ""
    Do
'        If fXoff = True Then
'            Call ProcessNmeaBuf
'            Exit Function
'        End If
        ReadBytes bytBuf       'Get next Chunk of file in bytes
'Convert to Vb AND remove ant bytes that are on disk after the EOF
'        ascBuf = ascBuf & Left$(StrConv(bytBuf, vbUnicode), BlockBytesRead)  'Convert to Ascii
        strBlock = Left$(StrConv(bytBuf, vbUnicode), blockbytesread)
        cTotalLinesRead = cTotalLinesRead + LFCount(strBlock) * 0.0001@
        
        ascBuf = ascBuf + strBlock
        If blockbytesread Then
'        cpos = InStrRev(ascBuf, vbCrLf)     'Get last CRLF (if any)
'        If cpos > 0 Then
'Send all complete sentences
'            ReadLines = ReadLines & Left$(ascBuf, cpos + 1)
'Leave incomplete sentence in ascBuf
'            ascBuf = Mid$(ascBuf, cpos + 2)
'15Sep            RaiseEvent DataArrival(BlockBytesRead)
            GoTo ExitWithLines
'        Else
            'Didn't find end, extract the part we have.
'Convert to Vb AND remove ant bytes that are on disk after the EOF
'            ascBuf = ascBuf & Left$(StrConv(bytBuf, vbUnicode), BlockBytesRead)  'Convert to Ascii
'            ReadNextBuffer
        End If
    Loop Until fEOF
ExitWithLines:
'Stop
    If Len(ascBuf) > 0 Then
'15Sep        RaiseEvent DataArrival(blockbytesread)
        DoEvents
    End If
End Function
#End If

Public Function ReadBytes(Buffer() As Byte) As Long
    RaiseErrorIfClosed
    If ReadFile(hFile, _
                Buffer(LBound(Buffer)), _
                UBound(Buffer) - LBound(Buffer) + 1, _
                ReadBytes, _
                0) Then
        blockbytesread = ReadBytes
        lBlocksRead = lBlocksRead + 1
        cTotalBytesRead = cTotalBytesRead + ReadBytes * 0.0001@
Debug.Print "Rd " & lBlocksRead & ", CumTot " & TotalBytesRead
        If ReadBytes <> lBlockSize Then
            fEOF = True
        End If
    Else
        RaiseError HBF_READ_FAILURE
    End If
End Function

#If False Then
'ReadLines returns the Block of Data read off disk, after conversion from Bytes
'to a string, to FileDataArrival (which transfers the data to NmeaBuf)
Private Function ReadLines() As String
Dim cpos As Long
Dim strBlock As String

    ReadLines = ""
    Do
        ReadBytes bytBuf       'Get next Chunk of file in bytes
'Convert to Vb AND remove ant bytes that are on disk after the EOF
'        ascBuf = ascBuf & Left$(StrConv(bytBuf, vbUnicode), BlockBytesRead)  'Convert to Ascii
        strBlock = Left$(StrConv(bytBuf, vbUnicode), blockbytesread)
        cTotalLinesRead = cTotalLinesRead + LFCount(strBlock) * 0.0001@
        
        ascBuf = ascBuf + strBlock
        
        cpos = InStrRev(ascBuf, vbCrLf)     'Get last CRLF (if any)
        If cpos > 0 Then
'Send all complete sentences
            ReadLines = ReadLines & Left$(ascBuf, cpos + 1)
'Leave incomplete sentence in ascBuf
            ascBuf = Mid$(ascBuf, cpos + 2)
'            ReadNextBuffer
            GoTo ExitWithLines
        Else
            'Didn't find end, extract the part we have.
'Convert to Vb AND remove ant bytes that are on disk after the EOF
            ascBuf = ascBuf & Left$(StrConv(bytBuf, vbUnicode), blockbytesread)  'Convert to Ascii
            ReadNextBuffer
        End If
    Loop Until fEOF

ExitWithLines:
End Function
#End If

'Called when bytBuf
Private Sub ReadNextBuffer()
    
'Reuquired if Null File or no CRLF before EOF
    If fEOF Then    'Last Read has EOF set
        Exit Sub
    End If
    
'Check weve not read too many blocks
    If lBlocksRead > Int(FileLen / lBlockSize) Then
'Stop
    End If

    ReadBytes bytBuf
    If lBlocksRead Mod 10 = 0 Then
        DoEvents
    End If
End Sub

Public Sub SeekAbsolute(ByVal Position As Currency)
    RaiseErrorIfClosed
    C.Value = Position / 10000@
    LSet l = C
    If SetFilePointer(hFile, l.LowVal, l.HighVal, FILE_BEGIN) _
        = INVALID_SET_FILE_POINTER Then
            If err.LastDllError Then RaiseError HbF_SEEK_FAILURE
    End If
End Sub


Private Sub SeekEnd()   'From HugeTextDemo
    Dim lpDistanceToMoveHigh As Long
    
    If SetFilePointer(hFile, 0&, lpDistanceToMoveHigh, FILE_END) = INVALID_SET_FILE_POINTER Then
        RaiseError HbF_SEEK_FAILURE
    End If
End Sub

Public Sub SeekEnd_original()
    RaiseErrorIfClosed
    If SetFilePointer(hFile, 0&, ByVal 0&, FILE_END) _
        = INVALID_SET_FILE_POINTER Then
            RaiseError HbF_SEEK_FAILURE
    End If
End Sub

Public Sub SeekRelative(ByVal Offset As Long)
  'Offset is signed.
  RaiseErrorIfClosed
  If SetFilePointer(hFile, Offset, ByVal 0&, FILE_CURRENT) _
      = INVALID_SET_FILE_POINTER Then
          RaiseError HbF_SEEK_FAILURE
  End If
End Sub

Public Sub WriteString(kb As String)
    bytBuf = StrConv(kb, vbFromUnicode)
    Call WriteBytes(bytBuf)
End Sub

Public Function WriteBytes(Buffer() As Byte) As Long
    RaiseErrorIfClosed
    If WriteFile(hFile, _
                 Buffer(LBound(Buffer)), _
                 UBound(Buffer) - LBound(Buffer) + 1, _
                 WriteBytes, _
                 0) Then
'Debug.Print "Wr " & WriteBytes
        If fAutoFlush Then Flush
    Else
        RaiseError HBF_WRITE_FAILURE
    End If
End Function

Private Sub Class_Initialize()
   mlngDebugID = DebugSerial
   ' Add a string entry to the global collection.
   gcolDebug.Add "HugeBinaryFile; DebugID=" _
   & mlngDebugID, CStr(mlngDebugID)
    hFile = INVALID_HANDLE_VALUE
End Sub

Private Sub Class_Terminate()
    If hFile <> INVALID_HANDLE_VALUE Then CloseHandle hFile
   ' Remove the string entry, so you know the object
   ' isn't around any more.
   gcolDebug.Remove CStr(mlngDebugID)
End Sub

Private Sub RaiseError(ByVal ErrorCode As HBF_Errors)
    Dim Win32Err As Long, Win32Text As String

    Win32Err = err.LastDllError
    If Win32Err Then
        Win32Text = vbNewLine & "Error " & Win32Err & vbNewLine _
                  & DecodeAPIErrors(Win32Err)
    End If
    If IsOpen Then CloseFile
    Select Case ErrorCode
        Case HBF_FILE_ALREADY_OPEN
            err.Raise HBF_FILE_ALREADY_OPEN, HBF_SOURCE, _
                sFName & vbCrLf & "File already open."
        Case HBF_OPEN_FAILURE
            err.Raise HBF_OPEN_FAILURE, HBF_SOURCE, _
                sFName & vbCrLf & "Error opening file." & Win32Text
        Case HbF_SEEK_FAILURE
            err.Raise HbF_SEEK_FAILURE, HBF_SOURCE, _
                sFName & vbCrLf & "Seek Error." & Win32Text
        Case HBF_FILELEN_FAILURE
            err.Raise HBF_FILELEN_FAILURE, HBF_SOURCE, _
                sFName & vbCrLf & "GetFileSize Error." & Win32Text
        Case HBF_READ_FAILURE
            err.Raise HBF_READ_FAILURE, HBF_SOURCE, _
                sFName & vbCrLf & "Read failure." & Win32Text
        Case HBF_WRITE_FAILURE
            err.Raise HBF_WRITE_FAILURE, HBF_SOURCE, _
                sFName & vbCrLf & "Write failure." & Win32Text
        Case HBF_FILE_ALREADY_CLOSED
            err.Raise HBF_FILE_ALREADY_CLOSED, HBF_SOURCE, _
                sFName & vbCrLf & "File must be open for this operation."
        Case Else
            err.Raise HBF_UNKNOWN_ERROR, HBF_SOURCE, _
               sFName & vbCrLf & "Unknown error." & Win32Text
    End Select

End Sub

Private Sub RaiseErrorIfClosed()
    If hFile = INVALID_HANDLE_VALUE Then RaiseError HBF_FILE_ALREADY_CLOSED
End Sub

Private Function DecodeAPIErrors(ByVal ErrorCode As Long) As String
    Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000&
    Dim strMsg As String, lngMsgLen As Long

    strMsg = Space$(256)
    lngMsgLen = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0&, _
                              ErrorCode, 0&, strMsg, 256&, 0&)
    If lngMsgLen > 0 Then
        DecodeAPIErrors = Left(strMsg, lngMsgLen)
    Else
        DecodeAPIErrors = "Unknown Error."
    End If
End Function

'http://vbcity.com/forums/t/63086.aspx
'btw: The calling routine in your form should also have an "On Error Goto" statement to trap the error
'raised by this class method. Make sure in your VB environment you have the option checked to only break on unhandled errors.
#If False Then
Public Sub CalculateSomething()
    On Error GoTo err_ThisError 'trap the error in this module

'   ... do some coding ...

   If Condition Then
        ' raise your error by whatever means you currently do, it will be trapped by the
        ' error handler above and sent to the label "err_ThisError" below. No code after this
        ' err.Raise statement will get executed.
        err.Raise vbObjectError + 1015, "ThisModule.CalculateSomething", "An error has occurred in my class and this is the reason ..."
   End If

   Exit Sub

err_ThisError:
      ' now that you have trapped the error here you can pass it on to the calling routine (i.e. your form)
      ' by calling err.Raise again and use the error information set above
      err.Raise err.Number, err.Source, err.Description
      Exit Sub
End Sub
#End If


