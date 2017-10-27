echo off
set INSTDIR=%WINDIR%\..\btbootcmd\

schtasks.exe /Delete /TN btbootcmd /F


if EXIST "%INSTDIR%" (
	rmdir /s /q "%INSTDIR%"
)

if EXIST "%WINDIR%\..\iplog.txt" (
	del "%WINDIR%\..\iplog.txt"
)

if EXIST "%WINDIR%\..\changeip.log" (
	del "%WINDIR%\..\changeip.log"
)

if EXIST "%WINDIR%\..\changeip.err" (
	del "%WINDIR%\..\changeip.err"
)

if EXIST "%WINDIR%\..\btcmd" (
	del "%WINDIR%\..\btcmd"
)
echo on