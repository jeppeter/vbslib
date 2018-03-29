#include "helpusage.au3"
#include "log4auto.au3"
#include <File.au3>

_Parse_Command_Line()

Local $netplwiz
Local $hWnd
Local $title
$netplwiz = _PathFull(@SystemDir & "\\Netplwiz.exe")

_log4auto_DEBUG(StringFormat("will run[%s]", $netplwiz))

Run("cmd /c " & $netplwiz, "", @SW_SHOW)
WinWaitActive("[CLASS:#32770]", "", 1)
SetError(0)
$hWnd = WinGetHandle("[CLASS:#32770]")
If @Error <> 0 Then
      _log4auto_FATAL(StringFormat("can not find [%s] window", $netplwiz))
      Exit 5
EndIf

$title = WinGetTitle($hWnd)

_log4auto_DEBUG(StringFormat("find window [%s] title [%s]", $hWnd, $title))
