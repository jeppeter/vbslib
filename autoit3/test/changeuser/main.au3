#include "helpusage.au3"
#include "log4auto.au3"
#include <File.au3>
#include "procapi.au3"
#include <GuiButton.au3>



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
		Sleep(2000)
		$i = $i + 1
	Wend

	If $hwnd = 0 Then
		return $hwnd
	EndIf

	$title = WinGetTitle($hwnd)
	_log4auto_DEBUG(StringFormat("get wind title [%s]", $title),"main.au3")
	return $hwnd
EndFunc

Const $CHECK_CONTROL_ID = "[ID:1022]"

Func _GetUserName($hwnd,$inuser)
	Local $retuser=""
	Local $chkctrl
	Local $chked
	Local $i
	$chkctrl = ControlGetHandle($hwnd,"",$CHECK_CONTROL_ID)
	If $chkctrl = 0 Then
		_log4auto_ERROR(StringFormat("can not find %s error[%d]", $CHECK_CONTROL_ID, @Error), "main.au3")
		return $retuser
	EndIf

	_log4auto_DEBUG(StringFormat("get %s [%s]", $CHECK_CONTROL_ID,$chkctrl), "main.au3")

	$chked=_GUICtrlButton_GetCheck($chkctrl)
	_log4auto_TRACE(StringFormat("%s state [%s]", $CHECK_CONTROL_ID,$chked), "main.au3")
	If BitAND($chked , $BST_CHECKED) = 0  Then
		_log4auto_TRACE(StringFormat("%s not checked", $CHECK_CONTROL_ID), "main.au3")
		For $i=0 To $tries
			;ControlClick($hwnd, "", $CHECK_CONTROL_ID,"left",1,2,2)
			_GUICtrlButton_SetCheck($chkctrl)
			Sleep(500)
			$chked=_GUICtrlButton_GetCheck($chkctrl)
			_log4auto_TRACE(StringFormat("%s state [%s]", $CHECK_CONTROL_ID, $chked), "main.au3")
			If BitAND($chked , $BST_CHECKED) = $BST_CHECKED Then
				ExitLoop
			EndIf
			Sleep(1000)
			_log4auto_DEBUG(StringFormat("not checked %s", $CHECK_CONTROL_ID), "main.au3")
		Next

		If BitAND($chked , $BST_CHECKED) = 0 Then
			_log4auto_ERROR(StringFormat("can not make %s checked", $CHECK_CONTROL_ID), "main.au3")
			return $retuser
		EndIf
	EndIf

	return $retuser
EndFunc


Local $hwnd
Local $retuser

$hwnd = _GetRunNetPlWiz()
If $hwnd = 0 Then
	Exit 5
EndIf
; now we shoule give the 

$retuser = _GetUserName($hwnd,$user)