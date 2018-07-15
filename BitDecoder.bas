Attribute VB_Name = "BitDecoder"
Option Explicit

Dim maskA(2) As Byte
Dim maskB(2) As Byte
Dim Shl(2) As Byte
Dim Shr(2) As Byte
'Dynamic so it can be cleared
Public Multipart() As String '  (9,9)    ID,fragment
'Created by PayloadReassembler from clssentence.nmeasentence (may be part sentence)
Private Type PayloadDef
    Bytes() As Byte 'Array of 8 bit bytes (last byte can have 0-7 fill bits at end)
    Len As Long     'Length of re-assembled payload excluding fill bits
End Type
Public Payload As PayloadDef

Public PayloadBytes() As Byte          'last payloadtobytes (includes fill bits)

Dim LastPayload As String
Dim LastBlock As Long          'last block read from payload


'gets long data type from given bit positions from payload
'PayloadReassembler must have been called (with no error) immediately before
Function pLong( _
ByVal From As Long, _
ByVal reqbits As Long) As Long

Dim b() As Byte
Dim rLong As Long
Dim i As Long
Dim inBlock As Long
Dim StuffBits As Long    'no of spare buts in last block
'Dim count As Long
Dim kb As String

'For count = 1 To 100000
b = AisToBits(From, reqbits)
'ReDim b(3)
'b(0) = 127
'b(1) = 255
'b(2) = 255
'b(3) = 0
inBlock = Int((reqbits - 1) / 8) + 1
'Next count

StuffBits = inBlock * 8 - reqbits

'NOTE it must be done like this to prevent overflow errors
'if we can move all 8 bits then do it simply
'otherwise shl rlong not of bits to move and
'extract same no of MSbits from b() and add

For i = 0 To (inBlock - 1)
'kb = BytesToBits(b(i)) 'debug only
    If i = inBlock - 1 And StuffBits Then
        rLong = rLong * 2 ^ (8 - StuffBits) + (b(i) And (Not (2 ^ StuffBits) - 1)) / 2 ^ (StuffBits)
    Else
        rLong = rLong * 2 ^ 8 + b(i)
    End If
'kb = ItoBits(rLong) 'debug only
Next i
pLong = rLong
End Function

'gets signed long data type from given bit positions from AisSentence
'PayloadReassembler must have been called (with no error) immediately before
Function pSi( _
ByVal From As Long, _
ByVal reqbits As Long) As Long
Dim rSi As Long
'Dim kb As String
'if rsi=2^(reqbits-1)   'not available  = 1000
rSi = pLong(From, reqbits)
'MsgBox ItoBits(rSi)
If rSi <= 2 ^ (reqbits - 1) - 1 Then
    pSi = rSi
Else
    pSi = rSi - 2 ^ reqbits
End If

End Function

'gets 6bit ascii data type from given bit positions from AisSentence
'PayloadReassembler must have been called (with no error) immediately before
Function p6bit( _
ByVal From As Long, _
ByVal reqbits As Long) As String
Dim NxtBit As Long
Dim Chrno As Long
Dim OutStr As String

NxtBit = From
'+6 because must have at least 6 bits available
'ie there may be fill bits at the end (msg 21)
Do Until NxtBit + 6 > From + reqbits
    Chrno = pLong(NxtBit, 6)
    If Chrno < 32 Then Chrno = Chrno + 64
    OutStr = OutStr + Chr$(Chrno)
    NxtBit = NxtBit + 6
Loop
'Debug.Print "p6bit=" & OutStr
'v136 OutStr = Replace(OutStr, "@", " ")
p6bit = OutStr
End Function
'get bits off as a hex string
Function pHex( _
ByVal From As Long, _
ByVal reqbits As Long) As String
Dim NxtBit As Long
Dim Chrno As Long
Dim OutStr As String
 'Debug.Print "Hex(" & From & "," & reqbits & ")" & (UBound(PayloadBytes) + 1) * 8

NxtBit = From
Do Until NxtBit >= From + reqbits
    Chrno = pLong(NxtBit, 4)
    OutStr = OutStr + Hex$(Chrno)
    NxtBit = NxtBit + 4
Loop
 'Debug.Print "Hex(" & From & "," & reqbits & ")" & (UBound(PayloadBytes) + 1) * 8
pHex = OutStr
End Function

'get bits as a string
Function pbits( _
ByVal From As Long, _
ByVal reqbits As Long) As String
Dim NxtBit As Long
Dim Chrno As Long
Dim OutStr As String

NxtBit = From
Do Until NxtBit >= From + reqbits
    OutStr = OutStr & pLong(NxtBit, 1)
    NxtBit = NxtBit + 1
Loop
pbits = OutStr
End Function


'Converts AisSentence AisSentence into byte array with only
'PayloadReassembler err must be 0 otherwise there could be no word(5)
'and routine will break
'containing only the data for the given bit positions
'if Nmea AisSentence is same as last time it was called, it does
'not unpack the data again
'runs at about 600k/min
'PayloadReassembler must have been called (with no error) immediately before
'A block is 8 bits of payload (1 6 bit ascii character)
Function AisToBits( _
ByVal From As Long, _
ByVal reqbits As Long) As Byte()
'Dim b() As Byte         'unpacked payload
Dim FstBlock As Long  'First block
Dim FstBitPos As Long 'First bit
Dim LstBlock As Long
Dim LstBitPos As Long
Dim arOut() As Byte
Dim ReqBlocks As Long    'no of blocks to output reqbits
Dim kb As String
Dim outBlock As Long     'Out block were currently constructing
Dim mask As Byte
'Dim PayloadBinaryBits As long
 'Debug.Print "aistobits(" & from & "," & reqbits & ")"
'if calculating from bit position back from end of payload
'and from is before start of buffer (because message is too short)
'force bits off end of buffer (0's are returned)

If From < 1 Then From = 1025

'extract a given number of bits from a byte array
'if were trying to access bits beyond the array (ie the
'max size of the payload) return 0's. This could happen
'if the data field is variable length
'Also if payload is zero length
    If reqbits > 0 Then
        ReqBlocks = Int((reqbits - 1) / 8)
    Else
        ReqBlocks = 0   'force zero output
    End If
    ReDim arOut(ReqBlocks)

    outBlock = 0
    FstBlock = Int((From - 1) / 8)   'base 0
    FstBitPos = From - (FstBlock) * 8
    mask = 2 ^ (8 - FstBitPos + 1) - 1
    LstBlock = Int((From + reqbits - 2) / 8) 'base 0
    LstBitPos = From + reqbits - 1 - (LstBlock) * 8
'This traps if payload is shorter than last required bit
'This can happen if a user is incorrectly coding the message
    If LstBlock > PayloadByteArraySize Then
        ReDim Preserve PayloadBytes(LstBlock)
       clsSentence.AisPayloadComments = "Payload too short"
    End If

'kb = ItoBits(CLng(mask))
'get LSBits off the first block and shift left by the number of bits missed
'if first block contains 1101 0111 ad we want last five bits
'mask will be 0001 1111, when AND with mask arout will contain
'0001 0111, which is then shifted 3 left (8-5)
    
    For outBlock = 0 To ReqBlocks
        arOut(outBlock) = (PayloadBytes(FstBlock + outBlock) And mask) * 2 ^ (FstBitPos - 1)  'shl

'arout now contains 1011 1000

'kb = ItoBits(CLng(arOut(outBlock)))

'we now want the remaining 3 MSbits from the next block
'the mask is the NOT of the last mask. Mask is 0001 1111 so
'the NOT of this mask is the inverse ie 1110 000
'if the next block contains 1101 0110 and is ANDed with the
'NOT mask the result will be the first 3 bits ie 1100 0000
'This result is shifted right by 5 bits and added to the result
'obtained from the last 5 bits of the first block.
'If the exctly 8 bits are obtained from the first block the
'second block must not be accessed as it may not exist

        If FstBlock + outBlock < LstBlock Then
            arOut(outBlock) = arOut(outBlock) _
            + (PayloadBytes(FstBlock + outBlock + 1) _
            And (Not mask)) / 2 ^ (8 - FstBitPos + 1) 'shr
        End If
        
        If FstBlock + outBlock = LstBlock Then
            mask = 255 - (2 ^ (8 - LstBitPos) - 1)
'kb = ItoBits(CLng(mask))
            arOut(outBlock) = arOut(outBlock) And mask
        End If
'kb = ItoBits(CLng(arOut(outBlock)))
    Next outBlock
AisToBits = arOut
End Function

'convert packed ais data into byte array
'runs at about 120k/min with debugging
'and about 300k/min without debugging
'Only called by PayloadReassembler, when payload is complete
'MaxBlock  is max no of 6 bit blocks were going to decode
Function PayloadToBytes(ByVal Payload As String, Optional ByVal MaxBlock As Long) As Byte()
Dim ar8Bit() As Byte
Dim ar6Bit() As Byte
Dim mask As Byte
Dim kb As String
Dim inBlock As Long
Dim outBlock As Long
Dim MoveBits As Long 'max number of bits we can move at one time
                        'from inByte buffer to OutByte
'convert 8 bit representation of 6 bit characters into 6 bits
Dim MoveCase As Long '0 to 2
'Dim arinBits() As String  'debug only
'Dim arOutBits() As String 'debug only
Dim i As Long

'set up the masks first time only
If maskA(0) = 0 Then    'masks not been set
    maskA(0) = 63   '00111111'
    maskA(1) = 15   '00001111'
    maskA(2) = 3    '00000011'
    maskB(0) = 48   '00110000'
    maskB(1) = 60   '00111100'
    maskB(2) = 63   '00111111'
    Shl(0) = 4      '2 bits shift left (multiply by 4)
    Shl(1) = 16     '4 bits
    Shl(2) = 64     '6 bits
    Shr(0) = 16     '4 bits shift right (divide by 16)
    Shr(1) = 4      '2 bits
    Shr(2) = 1      '0 bits
End If
If MaxBlock = 0 Or Len(Payload) < MaxBlock Then
    MaxBlock = Len(Payload)
End If
'payload must be at lease 2 characters long to get 8 bits
If (Payload <> LastPayload Or (Payload = LastPayload And LastBlock < MaxBlock)) And Len(Payload) > 1 Then

'we need to all bits if maxblock is missing or 0
'also if message type is 5 or 24 because we need the ships name
'for the list (5 is multipart (maxblock missing), 24 1st character
'of payload is "H", so force decode of all bits
'all above is to speed up processing
'for binary messages get enough to parse the AppID
'    If IsMissing(MaxBlock) Or Left$(Payload, 1) = "H" Then
'    If IsMissing(MaxBlock) Then
'        ReDim ar8Bit(Len(Payload))   '1 extra reqd if not on bit boundary
'    Else
#If False Then
    If MaxBlock And MaxBlock < Len(Payload) Then
            Select Case Left$(Payload, 1)
            Case Is = "6"       '88 bits
                ReDim ar6Bit(15)

'ReDim ar8Bit(BitsToChrno(88) - 1)   '0 based array
'above is 1 short & causes incorrect fi on list
'change maxblock to reqbits
            Case Is = "8"       '56 bits
                ReDim ar6Bit(10)
'ReDim ar8Bit(BitsToChrno(56) - 1)
'above is 1 short & causes incorrect fi on list
            Case Is = "H"   'msg 24
                ReDim ar6Bit(Len(Payload) - 1)
            Case Else
            ReDim ar6Bit(MaxBlock)
            End Select
        Else
            ReDim ar6Bit(Len(Payload) - 1)
        End If
'    End If
#End If
'maxblock is the no of payload characters we want to decode
ReDim ar6Bit(MaxBlock - 1)
LastBlock = MaxBlock

'ReDim arinBits(UBound(ar6Bit))    'debug only
If UBound(ar6Bit) + 1 > Len(Payload) Then
    kb = "Stop Encountered in PayloadTo Bytes" & vbCrLf _
    & "Ubound(ar6bit)=" & UBound(ar6Bit) & vbCrLf _
    & "Len(Payload)=" & Len(Payload) & vbCrLf
    Call WriteErrorLog(kb & vbCrLf & clsSentence.NmeaSentence)
'    Stop
End If
'first convert 8 bit representation of 6 bit ascii into 6 bit ascii
    For inBlock = 0 To UBound(ar6Bit)
        ar6Bit(inBlock) = AscB(Mid$(Payload, inBlock + 1, 1))
        If ar6Bit(inBlock) < 48 Then
            ar6Bit(inBlock) = 0     'error outside valid range
        Else
            ar6Bit(inBlock) = ar6Bit(inBlock) - 48      'convert to 6 bit
            If (ar6Bit(inBlock) > 40) Then ar6Bit(inBlock) = ar6Bit(inBlock) - 8
        End If
'this creates a bit array for debugging only
'arinBits(inBlock) = BytesToBits(ar6Bit(inBlock))    'debug only
    Next inBlock
    
'another way of loading the 6 bit array
#If False Then
    ar6Bit = StrConv(Payload, vbFromUnicode)
    For i = 0 To UBound(ar6Bit)
        If ar6Bit(i) < 48 Then
            ar6Bit(i) = 0     'error outside valid range
        Else
            ar6Bit(i) = ar6Bit(i) - 48      'convert to 6 bit
            If (ar6Bit(i) > 40) Then ar6Bit(i) = ar6Bit(i) - 8
        End If
    Next i
#End If

'no of bytes in payload is 1 more than 6bit array size (base is 0)
'size of 8 bit array is 1 less than payload bytes (base is also 0)
    ReDim PayloadBytes(ChrnoToWords(UBound(ar6Bit) + 1) - 1)    'Dec 14
    
'second move appropriate bits into output array
'ar8Bit has first 2 MSBits = 0, the last 6 LSBits must be
'concanated into 8 bits by removing 1st 2 bits (which are 0)
    
    For outBlock = 0 To PayloadByteArraySize
        inBlock = Int(outBlock * 8 / 6)
        MoveCase = outBlock - Int(outBlock / 3) * 3
        kb = outBlock & " " & inBlock & " (" & MoveCase & ")"
        PayloadBytes(outBlock) = _
        (ar6Bit(inBlock) And maskA(MoveCase)) * Shl(MoveCase)
        'last block exception
        If inBlock < UBound(ar6Bit) Then
            PayloadBytes(outBlock) = PayloadBytes(outBlock) _
            + (ar6Bit(inBlock + 1) And maskB(MoveCase)) / Shr(MoveCase)
        End If
'    arOutBits(outBlock) = BytesToBits(PayloadBytes(outBlock))  'debug only
    Next outBlock
    LastPayload = Payload
End If      'Different payload to when last called
PayloadToBytes = PayloadBytes   'Includes fill bits
clsSentence.AisPayload = Payload
End Function
'no of bytes required to hold given no of bits
'8 bit boundary
Function BitsToBytes(Bits As Long) As Long
If Bits Mod 8 = 0 Then
    BitsToBytes = Int(Bits / 8)
Else
    BitsToBytes = Int(Bits / 8) + 1
End If
End Function

'no of payload characters we need for a given no of decoded bits
Public Function BitsToChrno(Bits As Long) As Long
Dim b As Long
'for every 3 or part need 1 more bit
b = Int((Bits + 2) / 3) * 4
BitsToChrno = BitsToBytes(b)
End Function

'no of 8 bit words we require for each 6 bit payload character
Function ChrnoToWords(Chrno As Long) As Long
ChrnoToWords = Int(Chrno * 6 / 8)
If ChrnoToWords * 8 < Chrno * 6 Then ChrnoToWords = ChrnoToWords + 1 'Extra 8 bit word required
End Function

'No of fill bits for a given no of payload characters
'Wrong can return 6 and it should be 0 to 6
Function ChrnoToFillBits(Chrno As Long) As Long
ChrnoToFillBits = (8 - Chrno * 6 Mod 8) Mod 8 ' <> 0
End Function

Function BitsToPayloadChrno(Bits As Long) As Long
BitsToPayloadChrno = Bits / 6 + Bits Mod 8
End Function

Public Function ItoBits(ByVal A As Long) As String
Dim i As Long
Dim P1 As Long   'pointer
Dim Minus As Boolean

ItoBits = ""
i = A    'dont change argument
If i < 0 Then
    Minus = True
    i = i * (-1)
    ItoBits = "1"
End If
P1 = 31
Do Until P1 < 0
    If i >= 2 ^ P1 Then
        ItoBits = ItoBits + "1"
        i = i - (2 ^ P1)
    Else
        ItoBits = ItoBits + "0"
    End If
P1 = P1 - 1
Loop

End Function

Public Function BitstoI(A As String) As Long
Dim P1 As Integer   'pointer
Dim sLen As Integer
BitstoI = 0
P1 = 0
sLen = Len(A)
Do Until P1 = sLen
    If Mid$(A, sLen - P1, 1) = "1" Then BitstoI = BitstoI + 2 ^ (P1)
    P1 = P1 + 1
Loop
End Function

Public Function BytesToBits(ByVal A As Byte) As String
Dim i As Byte
Dim P1 As Long   'pointer
Dim Minus As Boolean

BytesToBits = ""
i = A    'dont change argument
P1 = 7
Do Until P1 < 0
    If i >= 2 ^ P1 Then
        BytesToBits = BytesToBits + "1"
        i = i - (2 ^ P1)
    Else
        BytesToBits = BytesToBits + "0"
    End If
P1 = P1 - 1
Loop

End Function

'returns error if PayloadReassembler message has earlier part missing
'if earlier part missing payload is still loaded to avoid a blank array
'payload is incrementally loaded so that mmsi, fi etc can be obtained
'after the first part received.
'If OK sets up last message bits & nmea sentence

'multipart id cant be cleared when last complete fragment is output as
'another call could be made to PayloadReassembler for more blocks than the last one
'even though its the same id

'Only called if IsAisSentece = true
'Sets AisMsgPartsComplete
Public Function PayloadReassembler(ByVal NmeaSentence As String, Comments As String, Optional ByVal MaxBlocks As Long) As Boolean
''Dim NmeaWords() As String
Dim Payload As String
Dim i As Long
Dim err As Boolean
Dim FillBits As Long
Dim WordNo As Long

'NmeaWords = Split(NmeaSentence, ",")  '350k/min

If UBound(NmeaWords) < 6 Then
    Comments = "Incomplete AIS sentence"
    GoTo Error_PayloadReassembler
End If
    
For WordNo = 1 To 3
    Comments = AisWordCheck(WordNo)
    If Comments <> "" Then
        GoTo Error_PayloadReassembler
    End If
Next WordNo

    If NmeaWords(1) = "1" Then  'single part message
        Payload = NmeaWords(5)
'        Call PayloadToBytes(NmeaWords(5), NmeaWords(6), MaxBlocks)   'creates clsSentence.AisPayload and PayloadBytes()
'        clsSentence.AisMsgPartsComplete = True
    Else    'Multi-Part
'Multipart(Sequence ID,PartNo)=Payload
        Multipart(NmeaWords(3), NmeaWords(2)) = NmeaWords(5)
'round bracket to trap
        Comments = " Msg ID = " & NmeaWords(3) & ", part " & NmeaWords(2) & " of " & NmeaWords(1)
'build payload
        For i = 1 To NmeaWords(1)   'No of Fragments in this sentence
'error if gt 9
'Sentence has no Fragment ID
            If Multipart(NmeaWords(3), i) = "" Then 'No Fragment
 'Scan Previously received fragments for sentences with this Sequence ID for any missing fragments
                If i < NmeaWords(2) Then '16/2
                    Comments = "Msg ID = " & NmeaWords(3) & ", Missing fragment " & i
                     err = True
                End If
                err = True  'Fragment missing up to this fragment
             End If
'error if gt 9
'Accumulate Payload fragments in order into Payload
            Payload = Payload + Multipart(NmeaWords(3), i)
        Next i  'next fragment
'this check must be done in PayloadReassembler as its the only place we can tell
'if the first part is missing & hence we cannot get the AisMsgType
        If Multipart(NmeaWords(3), 1) = "" Then clsSentence.AisMsgPart1Missing = True
    End If

'note if earlier missing fragment payload will be missing this fragment
' Debug.Print Len(Payload) & ",Fill=" & NmeaWords(6)
    Call PayloadToBytes(Payload, MaxBlocks)   'creates clsSentence.AisPayload and PayloadBytes()

'If len(payload)*6 = 426 Then Stop
        
'If last fragment
    If NmeaWords(1) = NullToZero(NmeaWords(2)) Then
'v137        clsSentence.AisPayloadFillBits = NmeaWords(6)   'valid only when last part
        clsSentence.AisPayloadFillBits = NmeaWords(clsSentence.NmeaCrcWord)   'valid only when last part
'last fragment AND no missing fragment
        If err = False Then
            clsSentence.AisMsgPartsComplete = True
'Trim payload length (ie bits) by payload fill bits
            clsSentence.AisPayloadBits = Len(Payload) * 6 - clsSentence.AisPayloadFillBits
'Finished with these fragments
            If NmeaWords(1) > 1 Then
                Call ClearPayloadFragments  'V129
            End If
        End If
    Else    'Not the last fragment
'If we have the first part we can use the PayloadBits to get Payload details
'eg MsgType
        If clsSentence.AisMsgPart1Missing = False Then
            clsSentence.AisPayloadBits = Len(Payload) * 6   'Dont know fill bits
        End If
    End If
    
'v133 not sure when the code below is actioned (if ever ?)
    clsSentence.PayloadReassemblerComments = Comments
    PayloadReassembler = err
    If clsSentence.AisMsgPartsComplete = True Then
        If NmeaWords(1) <> NullToZero(NmeaWords(2)) Then
'        Stop  'testv129
            clsSentence.AisMsgPartsComplete = False
        End If
    End If
Exit Function

Error_PayloadReassembler:
    PayloadReassembler = True
    clsSentence.PayloadReassemblerComments = Comments
    clsSentence.AisMsgPartsComplete = False
End Function

#If False Then
Public Function PayloadReassembler_old(ByVal NmeaSentence As String, Comments As String, Optional ByVal MaxBlocks As Long) As Boolean
''Dim NmeaWords() As String
Dim Payload As String
Dim i As Long
Dim err As Boolean
Dim FillBits As Long
Dim WordNo As Long

'NmeaWords = Split(NmeaSentence, ",")  '350k/min

If UBound(NmeaWords) < 6 Then
    Comments = "Incomplete AIS sentence"
    GoTo Error_PayloadReassembler
End If
    
For WordNo = 1 To 3
    Comments = AisWordCheck(WordNo)
    If Comments <> "" Then
        GoTo Error_PayloadReassembler
    End If
Next WordNo

    If NmeaWords(1) = "1" Then  'single part message
Payload = NmeaWords(5)
        Call PayloadToBytes(NmeaWords(5), NmeaWords(6), MaxBlocks)   'creates clsSentence.AisPayload and PayloadBytes()
        clsSentence.AisMsgPartsComplete = True
    Else
'testv129   Debug.Print NmeaSentence

'Multipart(Sequence ID,PartNo)=Payload
        Multipart(NmeaWords(3), NmeaWords(2)) = NmeaWords(5)
'round bracket to trap
        Comments = " Msg ID = " & NmeaWords(3) & ", part " & NmeaWords(2) & " of " & NmeaWords(1)
'build payload
        For i = 1 To NmeaWords(1)   'No of Fragments in this sentence
'error if gt 9
'Sentence has no Fragment ID
            If Multipart(NmeaWords(3), i) = "" Then 'No Fragment
 'Scan Previously received fragments for sentences with this Sequence ID for any missing fragments
                If i < NmeaWords(2) Then '16/2
                    Comments = "Msg ID = " & NmeaWords(3) & ", Missing fragment " & i
                     err = True
                End If
                err = True  'Fragment missing up to this fragment
             End If
'error if gt 9
'Accumulate Payload fragments in order into Payload
            Payload = Payload + Multipart(NmeaWords(3), i)
        Next i  'next fragment

'note if earlier missing fragment payload will be missing this fragment
 Debug.Print Len(Payload) & ",Fill=" & NmeaWords(6)
         Call PayloadToBytes(Payload, NmeaWords(6), MaxBlocks)   'creates clsSentence.AisPayload and PayloadBytes()
If PayloadBits = 426 Then Stop

'this check must be done in PayloadReassembler as its the only place we can tell
'if the first part is missing & hence we cannot get the AisMsgType
        If Multipart(NmeaWords(3), 1) = "" Then clsSentence.AisMsgPart1Missing = True
        
'If last fragment AND no missing fragment
        If NmeaWords(1) = NullToZero(NmeaWords(2)) And err = False Then
            clsSentence.AisMsgPartsComplete = True
'Finished with these fragments
            Call ClearPayloadFragments  'V129
        End If
    
'testv129   Debug.Print clsSentence.AisMsgPartsComplete 'testv129
    End If  'msg id check on multipart

    clsSentence.PayloadReassemblerComments = Comments
    PayloadReassembler = err
    If clsSentence.AisMsgPartsComplete = True Then
        If NmeaWords(1) <> NullToZero(NmeaWords(2)) Then
'        Stop  'testv129
            clsSentence.AisMsgPartsComplete = False
        End If
    End If
Exit Function

Error_PayloadReassembler:
    PayloadReassembler = True
    clsSentence.PayloadReassemblerComments = Comments
    clsSentence.AisMsgPartsComplete = False
End Function
#End If

'Clears Payload fragments in Multipart
'Requires clsSentence and NmeaWords
'Called before clsSentence is cleared
Public Function ClearPayloadFragments()
Dim i As Long
    On Error GoTo Error
'Check if Multipart with valid ID
                        
    For i = 0 To 9
        If Multipart(NmeaWords(3), i) <> "" Then
'testv129 Debug.Print "clear " & NmeaWords(3) & "," & i
            Multipart(CLng(NmeaWords(3)), i) = ""
        End If
    Next i  'next fragment
Exit Function

Error:
    NmeaRcv.StatusBar.Panels(1).Text = "ClearPayloadFragments, Error " & err.Number & " - " & clsSentence.NmeaSentence ' err.Description
    NmeaRcv.ClearStatusBarTimer.Enabled = True
End Function

'Returns -1 is no size
Public Function PayloadByteArraySize() As Long
    PayloadByteArraySize = -1
    On Error Resume Next
    PayloadByteArraySize = UBound(PayloadBytes) '0 based
End Function

'Returns a message for any AISWord in error & Bits for word 0 & 5
'In NmeaWords
Public Function AisWordCheck(WordNo As Long) As String
#If False Then
Dim Payload6Bits As Long    'used to calculate 8 bits
Dim Payload8Bits As Long
Dim Message As String
'Should not happen but here to prevent a subscript error
    If WordNo > UBound(NmeaWords) Then Exit Function
    
    Select Case WordNo
    Case Is = 0
        Message = TalkerDes(clsSentence.IecTalkerID)
        If UBound(NmeaWords) >= 5 Then
            If IsNumeric(NmeaWords(5)) Then
                Payload6Bits = Len(NmeaWords(5)) * 6
                Payload8Bits = Int(Payload6Bits / 8) * 8
                Message = Message & ", " & Payload8Bits & " bits (" & Len(NmeaWords(5)) & " 6-bit words)"
            End If
        End If
    End Select
    AisWordCheck = Message
#End If
'Check part no
    If WordNo = 1 Or WordNo = 2 Then
'Always 1 to 9
        If IsNumeric(NmeaWords(WordNo)) = False Then
            AisWordCheck = "Invalid - not numeric (" & NmeaWords(WordNo) & ")"
        Else
            If NmeaWords(WordNo) < 1 Or NmeaWords(WordNo) > 9 Then
                AisWordCheck = "Invalid - Out of range (" & NmeaWords(WordNo) & ")"
            End If
        End If
    End If

'Must be less than total no of fragments
    If WordNo = 2 Then
        If NullToZero(NmeaWords(2)) > NullToZero(NmeaWords(1)) Then
            AisWordCheck = "Invalid Fragment No (" & NmeaWords(WordNo) & ")"
        End If
    End If


'If single part Fragment ID must be blank
    If WordNo = 3 Then
'        If NmeaWords(WordNo) = "" Then
'            If NullToZero(NmeaWords(1)) <> 1 Then
'                AisWordCheck = "Invalid Msg ID (" & NmeaWords(WordNo) & ")"
'            End If
'        Else    'ID is not ""

'If single part Sequential ID MAY be specified, if Multipart MUST be specified
        If (NmeaWords(1) = 1 And NmeaWords(WordNo) <> "") Or NmeaWords(1) > 1 Then
'Can be 0
            If IsNumeric(NmeaWords(WordNo)) = False Then
                AisWordCheck = "Invalid - not numeric (" & NmeaWords(WordNo) & ")"
            Else
                If NmeaWords(WordNo) < 0 Or NmeaWords(WordNo) > 9 Then
                    AisWordCheck = "Invalid - Out of range (" & NmeaWords(WordNo) & ")"
                End If
            End If
'If ID is not blank the message should be Multi Part
'            If NullToZero(NmeaWords(1)) = 1 Then
'                AisWordCheck = "Invalid Msg ID (" & NmeaWords(WordNo) & ")"
'            End If
        End If
    End If
End Function

Public Function IecEncapsulationCheck(ByVal NmeaSentence As String) As Boolean
    
    With clsSentence
'Radio Code is an optional field
        If .NmeaCrcWord < 5 Then
            .IecEncapsulationComments = "(Encapsulation Invalid)"
            Exit Function
        End If
        
        If IsNumeric(NmeaWords(1)) = False Then
        .IecEncapsulationComments = "(No of Parts invalid)"
            Exit Function
        End If
'Elliott No of parts=1.0000, convert to integer (v139)
        NmeaWords(1) = CInt(NmeaWords(1))
        If NmeaWords(1) < 1 Or NmeaWords(1) > 9 Then
            .IecEncapsulationComments = "(No of Parts invalid)"
            Exit Function
        End If
    
        If IsNumeric(NmeaWords(2)) = False Then
            .IecEncapsulationComments = "(Part no invalid)"
            Exit Function
        End If
        If NmeaWords(2) > NmeaWords(1) Then
            .IecEncapsulationComments = "(Part no invalid)"
            Exit Function
        End If

'        If NmeaWords(1) = 1 Then    'Single part
'            If NmeaWords(3) <> "" Then
'            End If
'        Else                        'Multipart
'If single part Sequential ID MAY be specified, if Multipart MUST be specified
        If (NmeaWords(1) = 1 And NmeaWords(3) <> "") Or NmeaWords(1) > 1 Then
            If IsNumeric(NmeaWords(3)) = False Then
                .IecEncapsulationComments = "(Sequential ID invalid)"
                Exit Function
            End If
            If NmeaWords(3) < 0 Or NmeaWords(1) > 9 Then
                .IecEncapsulationComments = "(Sequential ID invalid)"
                Exit Function
            End If
        End If

'IEC 61162-1 can have a number of fields before the data
            If NmeaWords(.NmeaCrcWord - 1) = "" Then
                .IecEncapsulationComments = "(No Data)"
                Exit Function
            End If
    
            If IsNumeric(NmeaWords(.NmeaCrcWord)) = False Then
                .IecEncapsulationComments = "(No NMEA Fill bits)"
                Exit Function
            End If
            If NmeaWords(.NmeaCrcWord) > 5 Then
                .IecEncapsulationComments = "(" & NmeaWords(.NmeaCrcWord) & " NMEA Fill is bits invalid)"
'22/10/16 jean-yvon
                If (DisableNmeaFillBitsError = False) Then Exit Function
            End If
    End With
        
    IecEncapsulationCheck = True
End Function

Public Function NullToZero(kb As String) As Long
    On Error Resume Next
    NullToZero = CLng(kb)
End Function

