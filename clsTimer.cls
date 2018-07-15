VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'http://www.tek-tips.com/viewthread.cfm?qid=274248
Option Explicit
' Storage for the debug ID.
Private mlngDebugID As Long

Private Declare Function QueryPerformanceFrequencyAny Lib "kernel32" Alias _
    "QueryPerformanceFrequency" (lpFrequency As Any) As Long
Private Declare Function QueryPerformanceCounterAny Lib "kernel32" Alias _
    "QueryPerformanceCounter" (lpPerformanceCount As Any) As Long

Private mStartTime As Currency
Private mEndTime As Currency
Private mTotalTime As Double
Private mFrequency As Currency
Private mRunning As Boolean
Private mDuration As Double

Private Sub Class_Initialize()
' measure the frequency of the counter-timer
' ensure the computer is compatible
   mlngDebugID = DebugSerial
   ' Add a string entry to the global collection.
   gcolDebug.Add "clsTimer; DebugID=" _
   & mlngDebugID, CStr(mlngDebugID)
If QueryPerformanceFrequencyAny(mFrequency) = 0 Then
    err.Raise 1001, , "Hardware out of date!"
    mFrequency = 0
    Exit Sub
End If
End Sub

Public Sub StartTimer()
QueryPerformanceCounterAny mStartTime
mRunning = True        'used in Duration property
End Sub


Public Sub StopTimer()
mTotalTime = mDuration       'transfer cumulative time so far to module-level variable
mRunning = False             'no longer running
End Sub


Public Sub ResetTimer()
If mRunning Then
    StartTimer
End If
mTotalTime = 0#    'only used when in stopped condition
End Sub


Public Property Get Duration() As Double
Dim Interval As Double

If mFrequency = 0 Then
    err.Raise 1002, , "Accurate timer object not initialised"
    Duration = 0
    Exit Property
End If

If mRunning Then
    QueryPerformanceCounterAny mEndTime   'get a new endtime
    Interval = CDbl(mEndTime - mStartTime) / mFrequency
    mDuration = mTotalTime + Interval     'the stored total so far plus the new addition
Else
    mDuration = mTotalTime  'not running - return cumulative
End If

Duration = mDuration
End Property


Private Sub Class_Terminate()
   ' Remove the string entry, so you know the object
   ' isn't around any more.
   gcolDebug.Remove CStr(mlngDebugID)

End Sub
