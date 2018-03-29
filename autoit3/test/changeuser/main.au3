#include "helpusage.au3"
#include "log4auto.au3"
#include <File.au3>

_Parse_Command_Line()

Local $netplwiz
$netplwiz = _PathFull(@SystemDir & "\\Netplwiz.exe")

_log4auto_DEBUG(StringFormat("will run[%s]", $netplwiz))
