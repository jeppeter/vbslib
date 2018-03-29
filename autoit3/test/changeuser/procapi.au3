#include-once
#include <array.au3>
#include "log4auto.au3"

;Function for getting HWND from PID
Func _GetHwndFromPID($PID)
	Local $winlist
	Local $hWnd
	Local $i
	Local $iPID2
	$hWnd = 0
	$winlist = WinList()
	For $i = 1 To $winlist[0][0]
		If $winlist[$i][0] <> "" Then
			$iPID2 = WinGetProcess($winlist[$i][1])
			If $iPID2 = $PID Then
				$hWnd = $winlist[$i][1]
				ExitLoop
			EndIf
		EndIf
	Next
	Return $hWnd
EndFunc;==>_GetHwndFromPID

;=====================================================
; Function : _GetAllWnds
; input : 
;         $name the name to filter if "" will no filter
;
; output :
;         array of hwnds
;         
;=====================================================
Func _GetAllWnds($name="")
	Local $hwnds = []
	Local $pids = []
	Local $i
	Local $curhwnd

	$pids = ProcessList($name)
	_log4auto_TRACE(StringFormat("number [%d]", $pids[0][0]),"procapi.au3")
	For $i=1 To $pids[0][0]
		_log4auto_TRACE(StringFormat("[%d]=[%s][%d]", $i,$pids[$i][0],$pids[$i][1]),"procapi.au3")
		SetError(0)
		$curhwnd = _GetHwndFromPID($pids[$i][1])
		if $curhwnd <> 0 Then
			_log4auto_TRACE(StringFormat("get[%d][%d]=[%s]",$i,$pids[$i][1],$curhwnd), "procapi.au3")
			_ArrayPush($hwnds,$curhwnd)
		Else
			_log4auto_TRACE(StringFormat("Error for [%d]",$pids[$i][1]), "procapi.au3")
		EndIf
	Next
	return $hwnds
EndFunc

;=====================================================
; Function : _GetAllPids
; input : 
;         $name the name to filter if "" will no filter
;
; output :
;         array of pids
;         
;=====================================================
Func _GetAllPids($name="")
	Local $retpids = []
	Local $pids = []
	Local $i
	Local $curhwnd

	$pids = ProcessList($name)
	_log4auto_TRACE(StringFormat("number [%d]", $pids[0][0]),"procapi.au3")
	For $i=1 To $pids[0][0]
		_log4auto_TRACE(StringFormat("[%d]=[%s][%d]", $i,$pids[$i][0],$pids[$i][1]),"procapi.au3")
		_ArrayPush($retpids, $pids[$i][1])
	Next
	return $retpids
EndFunc