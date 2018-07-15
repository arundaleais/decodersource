reg delete hkcr\typelib\{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}\2.0 /f
if exist %systemroot%\SysWOW64\cscript.exe goto 64 
%systemroot%\system32\regsvr32 /u mscomctl.ocx
%systemroot%\system32\regsvr32 mscomctl.ocx
exit

:64 
%systemroot%\sysWOW64\regsvr32 /u mscomctl.ocx
%systemroot%\sysWOW64\regsvr32 mscomctl.ocx
exit