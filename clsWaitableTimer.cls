VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWaitableTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Storage for the debug ID.
Private mlngDebugID As Long
'**************************************
' Name:        clsWaitableTimer
'
' Description: This class encapsulate the WaitableTimer API functions to
'              put the thread of your application to Sleep for a period of time.
'              The benefit of a Waitable timer to the Sleep API is that your
'              application will still be responsive to events, where Sleep
'              will freeze your application for the set interval.
'
' Example:     'This is an example for idling your application
'              Private mobjWaitTimer As clsWaitableTimer
'              Private Sub RunProcess()
'                Set mobjWaitTimer = New clsWaitableTimer
'                Do
'                 If mbWorkToDo Then
'                   Call ProcessWork()
'                 Else
'                   mobjWaitTimer.Wait(5000) 'Wait for 5 seconds
'                 End If
'                Loop Until Not mbStop
'                Set mobjWaitTimer = nothing
'              End Sub
'
' Revision History:

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Const WAIT_ABANDONED& = &H80&
Private Const WAIT_ABANDONED_0& = &H80&
Private Const WAIT_FAILED& = -1&
Private Const WAIT_IO_COMPLETION& = &HC0&
Private Const WAIT_OBJECT_0& = 0
Private Const WAIT_OBJECT_1& = 1
Private Const WAIT_TIMEOUT& = &H102&
Private Const INFINITE = &HFFFF
Private Const ERROR_ALREADY_EXISTS = 183&
Private Const QS_HOTKEY& = &H80
Private Const QS_KEY& = &H1
Private Const QS_MOUSEBUTTON& = &H4
Private Const QS_MOUSEMOVE& = &H2
Private Const QS_PAINT& = &H20
Private Const QS_POSTMESSAGE& = &H8
Private Const QS_SENDMESSAGE& = &H40
Private Const QS_TIMER& = &H10
Private Const QS_MOUSE& = (QS_MOUSEMOVE Or QS_MOUSEBUTTON)
Private Const QS_INPUT& = (QS_MOUSE Or QS_KEY)
Private Const QS_ALLEVENTS& = (QS_INPUT Or QS_POSTMESSAGE Or QS_TIMER Or QS_PAINT Or QS_HOTKEY)
Private Const QS_ALLINPUT& = (QS_SENDMESSAGE Or QS_PAINT Or QS_TIMER Or QS_POSTMESSAGE Or QS_MOUSEBUTTON Or QS_MOUSEMOVE Or QS_HOTKEY Or QS_KEY)

Private Const UNITS = 4294967296#
Private Const MAX_LONG = -2147483648#

Private Declare Function CreateWaitableTimer Lib "kernel32" Alias "CreateWaitableTimerA" (ByVal lpSemaphoreAttributes As Long, ByVal bManualReset As Long, ByVal lpName As String) As Long
Private Declare Function OpenWaitableTimer Lib "kernel32" Alias "OpenWaitableTimerA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Private Declare Function SetWaitableTimer Lib "kernel32" (ByVal hTimer As Long, lpDueTime As FILETIME, ByVal lPeriod As Long, ByVal pfnCompletionRoutine As Long, ByVal lpArgToCompletionRoutine As Long, ByVal fResume As Long) As Long
Private Declare Function CancelWaitableTimer Lib "kernel32" (ByVal hTimer As Long)
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function MsgWaitForMultipleObjects Lib "user32" (ByVal nCount As Long, pHandles As Long, ByVal fWaitAll As Long, ByVal dwMilliseconds As Long, ByVal dwWakeMask As Long) As Long

Private mlTimer As Long
Private mCancel As Boolean

Private Sub Class_Initialize()
   mlngDebugID = DebugSerial
   ' Add a string entry to the global collection.
   gcolDebug.Add "clsWaitableTimer; DebugID=" _
   & mlngDebugID, CStr(mlngDebugID)

End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    If mlTimer <> 0 Then CloseHandle mlTimer
   ' Remove the string entry, so you know the object
   ' isn't around any more.
   gcolDebug.Remove CStr(mlngDebugID)
End Sub

'JNA added to cancel while in the wait
Public Sub Cancel()
    mCancel = True
End Sub

Public Sub Wait(MilliSeconds As Long)
    On Error GoTo errHandler
    Dim ft As FILETIME
    Dim lBusy As Long
    Dim lRet As Long
    Dim dblDelay As Double
    Dim dblDelayLow As Double
    
    mCancel = False
    mlTimer = CreateWaitableTimer(0, True, App.EXEName & "Timer" & Format$(Now(), "NNSS"))
    
    If err.LastDllError <> ERROR_ALREADY_EXISTS Then
        ft.dwLowDateTime = -1
        ft.dwHighDateTime = -1
        lRet = SetWaitableTimer(mlTimer, ft, 0, 0, 0, 0)
    End If
    
    ' Convert the Units to nanoseconds.
    dblDelay = CDbl(MilliSeconds) * 10000#
    
    ' By setting the high/low time to a negative number, it tells
    ' the Wait (in SetWaitableTimer) to use an offset time as
    ' opposed to a hardcoded time. If it were positive, it would
    ' try to convert the value to GMT.
    ft.dwHighDateTime = -CLng(dblDelay / UNITS) - 1
    dblDelayLow = -UNITS * (dblDelay / UNITS - Fix(CStr(dblDelay / UNITS)))
    
    If dblDelayLow < MAX_LONG Then dblDelayLow = UNITS + dblDelayLow
    
    ft.dwLowDateTime = CLng(dblDelayLow)
    lRet = SetWaitableTimer(mlTimer, ft, 0, 0, 0, False)
    
    Do
        ' QS_ALLINPUT means that MsgWaitForMultipleObjects will
        ' return every time the thread in which it is running gets
        ' a message. If you wanted to handle messages in here you could,
        ' but by calling Doevents you are letting DefWindowProc
        ' do its normal windows message handling---Like DDE, etc.
        lBusy = MsgWaitForMultipleObjects(1, mlTimer, False, INFINITE, QS_ALLINPUT&)
        DoEvents
    Loop Until lBusy = WAIT_OBJECT_0 Or mCancel
    
    ' Close the handles when you are done with them.
    CloseHandle mlTimer
    mlTimer = 0
    Exit Sub
    
errHandler:
    err.Raise err.Number, err.Source, "[clsWaitableTimer.Wait]" & err.Description
End Sub

