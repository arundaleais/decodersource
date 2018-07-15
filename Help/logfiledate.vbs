'The output file (created by AisDecoder) is assumed to be output.csv
'This uses the WSH scripting object to reformat the date
'In Windows the problem with trying to do this any other way is
'the current date format uses the PC's date format and locale to
'return the current date/time.
'for example if you just try to append the system date, it will
'return 1/11/2016 in Europe & 11/1/2016 in USA
'forward slash(/) is also not permitted within a file name
'If you look a examples on the web of how to fix this issue
'THEY DONT WORK !! if the program is to be used internationally

Set FSO = CreateObject("Scripting.FileSystemObject")
Set wshShell = CreateObject( "WScript.Shell" )

'* Get current (DOS) directory (batch file default) vbscript p143
'wScript.Echo "CurrentDirectory is " & wshShell.CurrentDirectory
strFolder = wshShell.CurrentDirectory
'WScript.Echo "strFolder is " & strfolder
Set objFolder = FSO.GetFolder(strFolder)

'* Reformat Date YYMMDD_hhmmss
current = Now
'WScript.Echo "current is " & current
mth = Month(current)
d = Day(current)
yr = Year(current)
If Len(mth) <2 Then
    mth="0"&mth
End If
If Len(d) < 2 Then
    d = "0"&d
End If
hh = Hour (current)
mm = Minute (current)
ss = Second (current)
If Len(hh) <2 Then
    hh="0"&hh
End If
If Len(mm) <2 Then
    mm="0"&mm
End If
If Len(ss) <2 Then
    ss="0"&ss
End If
timestamp=yr & mth & d & "_" & hh & mm & ss
'WScript.Echo "timestamp is " & timestamp

'* Copy output.csv to time stamped file
FSO.CopyFile strfolder & "\output.csv", strfolder & "\output_" & timestamp & ".csv"
set FSO=nothing
	