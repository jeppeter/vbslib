#include-once

#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/mo /SCI=1

#include <WinAPIHObj.au3>
#include <WinAPIProc.au3>
#include "ParseCmdLine.au3"


AutoItSetOption("MustDeclareVars",1)
AutoItSetOption("ExpandVarStrings",1)


Global $helpmode = False
Global $user 
Global $newuser
Global $admin
Global $password
Global $tries
Global $timeout

Func _Usage($ec=0,$fmt="", $prog=@ScriptFullPath)
	Local $fh = _WinAPI_GetStdHandle(2)
	Local $curuser

	If $ec = 0 Then
		$fh = _WinAPI_GetStdHandle(1)
	EndIf
	if $fmt <> "" Then
		_WinAPI_WriteConsole($fh, $fmt & @CRLF)
	EndIf

	_WinAPI_WriteConsole($fh, StringFormat("%s [OPTIONS]", $prog) & @CRLF)
	_WinAPI_WriteConsole($fh, StringFormat("-help|-h                 to display this help information") & @CRLF)
	_WinAPI_WriteConsole($fh, StringFormat("-user|-u user            to specified the user name default[%s]",  @UserName) & @CRLF)
	_WinAPI_WriteConsole($fh, StringFormat("-new|-n newuser          to specified the new user name default[TN%s]",  @UserName) & @CRLF)
	_WinAPI_WriteConsole($fh, StringFormat("-admin|-a admin          to specified the admin name default Administrator") & @CRLF)
	_WinAPI_WriteConsole($fh, StringFormat("-passwd|-p password      to specified the password default []") & @CRLF)
	_WinAPI_WriteConsole($fh, StringFormat("-try|-d tries            to specified the tries default 3") & @CRLF)
	_WinAPI_WriteConsole($fh, StringFormat("-timeout|-t timeout      to specified the timeout to wait default 500 millseconds") & @CRLF)

	Exit $ec
	return
EndFunc

Func _Parse_Command_Line()
	Local $directives = [ _
		"b|help|helpmode|False", _
		"b|h|helpmode|False", _
		StringFormat("s|user|user|%s", @UserName), _
		StringFormat("s|u|user|%s", @UserName), _
		StringFormat("s|new|newuser|TN%s", @UserName), _
		StringFormat("s|n|newuser|TN%s", @UserName), _
		"s|try|tries|3", _
		"s|d|tries|3", _
		"s|timeout|timeout|500", _
		"s|t|timeout|500"]
	Local $cl = $CmdLine
	Local $ret
	$ret = _ParseCmdLine($cl,$directives)
	If $ret = 0 Then
		_Usage(3,"")
	EndIf

	If $helpmode Then
		_Usage(0,"")
	EndIf
	return
EndFunc