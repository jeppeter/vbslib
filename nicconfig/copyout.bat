echo off
set OUTDIR=..\..\btbootcmd
set FROMDIR=%~dp0

if  [%1] == []   GOTO pass_1
set OUTDIR=%1
:pass_1

if not EXIST %OUTDIR% (
	mkdir %OUTDIR%
)

copy /Y "%FROMDIR%\fileop.vbs" %OUTDIR%\fileop.vbs
copy /Y "%FROMDIR%\baseop.vbs" %OUTDIR%\baseop.vbs
copy /Y "%FROMDIR%\cmdexec.vbs" %OUTDIR%\cmdexec.vbs
copy /Y "%FROMDIR%\vbsjson.vbs" %OUTDIR%\vbsjson.vbs
copy /Y "%FROMDIR%\changeip.vbs" %OUTDIR%\changeip.vbs
copy /Y "%FROMDIR%\btcmd" %OUTDIR%\btcmd
copy /Y "%FROMDIR%\install.bat" %OUTDIR%\install.bat
copy /Y "%FROMDIR%\uninstall.bat" %OUTDIR%\uninstall.bat
copy /Y "%FROMDIR%\runchangeip.bat" %OUTDIR%\runchangeip.bat

echo on