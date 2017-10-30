echo off
set INSTDIR=%WINDIR%\..\btbootcmd\
set FROMDIR=%~dp0

if not EXIST %INSTDIR%  (
	mkdir %INSTDIR%
)

copy /Y "%FROMDIR%\fileop.vbs" "%INSTDIR%\fileop.vbs"
copy /Y "%FROMDIR%\baseop.vbs" "%INSTDIR%\baseop.vbs"
copy /Y "%FROMDIR%\cmdexec.vbs" "%INSTDIR%\cmdexec.vbs"
copy /Y "%FROMDIR%\vbsjson.vbs" "%INSTDIR%\vbsjson.vbs"
copy /Y "%FROMDIR%\changeip.vbs" "%INSTDIR%\changeip.vbs"
copy /Y "%FROMDIR%\runchangeip.bat" "%INSTDIR%\runchangeip.bat"

REM wevtutil set-log Microsoft-Windows-TaskScheduler/Operational /enabled:true
schtasks.exe /Create /TN btbootcmd /TR "\"%INSTDIR%\runchangeip.bat\"" /RU system /SC ONSTART
echo on