#include-once

#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/mo /SCI=1

#include <WinAPIHObj.au3>
#include <WinAPIProc.au3>


Func _Usage($ec=0,$fmt="", $prog=$CmdLine[0])
	Local $fh = _WinAPI_GetStdHandle(2)

	If $ec = 0 Then
		$fh = _WinAPI_GetStdHandle(1)
	EndIf
	if $fmt <> "" Then
		_WinAPI_WriteConsole($fh, $fmt & @CRLF)
	EndIf

	_WinAPI_WriteConsole($fh, StringFormat("%s [OPTIONS]", $prog) & @CRLF)
	_WinAPI_WriteConsole($fh, @CRLF)
	_WinAPI_WriteConsole($fh, @CRLF)
	_WinAPI_WriteConsole($fh, @CRLF)
	_WinAPI_WriteConsole($fh, @CRLF)
	_WinAPI_WriteConsole($fh, @CRLF)

	ExitApp $ec
	return
EndFunc

Func _Parse_Command_Line()
	Local $i
	For 
EndFunc