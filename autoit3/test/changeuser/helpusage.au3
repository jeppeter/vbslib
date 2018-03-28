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
	_WinAPI_WriteConsole($fh, StringFormat("-help|-h                 to display this help information") & @CRLF)
	_WinAPI_WriteConsole($fh, StringFormat("-user|-u user            to specified the user name") & @CRLF)
	_WinAPI_WriteConsole($fh, StringFormat("-new|-n newuser          to specified the new user name") & @CRLF)
	_WinAPI_WriteConsole($fh, StringFormat("-admin|-a admin          to specified the admin name") & @CRLF)
	_WinAPI_WriteConsole($fh, StringFormat("-passwd|-p password      to specified the password") & @CRLF)
	_WinAPI_WriteConsole($fh, StringFormat("-try|-d tries            to specified the tries default 3") & @CRLF)
	_WinAPI_WriteConsole($fh, StringFormat("-timeout|-t timeout      to specified the timeout to wait default 500 millseconds") & @CRLF)

	ExitApp $ec
	return
EndFunc

Func _Parse_Command_Line()
	Local $i
EndFunc