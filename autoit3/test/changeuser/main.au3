#include "helpusage.au3"
#include "log4auto.au3"
#include <File.au3>
#include "procapi.au3"


_Parse_Command_Line()


Func _GetRunNetPlWiz()
	Local $netplwiz
	Local $hWnds
	Local $hwnd = 0
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
	$i = 0
	While $i < $tries
		WinWaitActive("[CLASS:#32770]", "", 1)
		SetError(0)
		$hWnds = _GetAllWnds("Netplwiz.exe")
		If UBound($hWnds) >0  And $hWnds[0] <> 0 Then
			$hwnd = $hWnds[0]
			ExitLoop
		EndIf
		_log4auto_FATAL(StringFormat("can not find [%s] window [%d]", $netplwiz, $hWnds[0]),"main.au3")
		$i = $i + 1
	Wend

	If $hwnd = 0 Then
		return $hwnd
	EndIf

	$title = WinGetTitle($hwnd)
	_log4auto_DEBUG(StringFormat("get wind title [%s]", $title),"main.au3")
	return $hwnd
EndFunc


Local $hwnd

$hwnd = _GetRunNetPlWiz()
If $hwnd = 0 Then
	Exit 5
EndIf
; now we shoule give the 
