#include "helpusage.au3"
#include "log4auto.au3"
#include <File.au3>
#include "procapi.au3"


_Parse_Command_Line()

Local $netplwiz
Local $hWnds
Local $hwnd
Local $title
Local $dtitle
Local $pids
Local $i

While True
	$pids = _GetAllPids("Netplwiz.exe")
	if Ubound($pids) = 1 And $pids[0] = 0 Then
		ExitLoop
	EndIf
	_log4auto_DEBUG(StringFormat("Process [%d]", UBound($pids)),"main.au3")
	For $i=0 To UBound($pids)-1
		_log4auto_TRACE(StringFormat("Kill [%d]", $pids[$i]), "main.au3")
		ProcessClose($pids[$i])
	Next
	; sleep for a while
	Sleep(500)
Wend

$netplwiz = _PathFull(@SystemDir & "\\Netplwiz.exe")

_log4auto_DEBUG(StringFormat("will run[%s]", $netplwiz),"main.au3")

Run("cmd /c " & $netplwiz, "", @SW_SHOW)
WinWaitActive("[CLASS:#32770]", "", 10)
SetError(0)
$hWnds = _GetAllWnds("Netplwiz.exe")

if UBound($hWnds) = 1 And $hWnds[0] = 0 Then
	_log4auto_FATAL(StringFormat("can not find [%s] window [%d]", $netplwiz, $hWnds[0]),"main.au3")
	Exit 5
EndIf

$hwnd = $hWnds[0]
$title = WinGetTitle($hwnd)
_log4auto_DEBUG(StringFormat("get wind title [%s]", $title),"main.au3")

; now we shoule give the 
