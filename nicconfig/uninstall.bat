echo off
set INSTDIR=%WINDIR%\..\btbootcmd\

schtasks.exe /Delete /TN btbootcmd /F


if EXIST %INSTDIR% (
	rmdir /s /q %INSTDIR%
)

if EXIST %WINDIR%\..\iplog.txt (
	del %WINDIR%\..\iplog.txt
)

if EXIST %WINDIR%\..\btcmd (
	del %WINDIR%\..\btcmd
)
echo on