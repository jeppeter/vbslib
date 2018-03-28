#include "helpusage.au3"
#include "log4auto.au3"

_Parse_Command_Line()

_log4auto_Fatal("1", "main.au3")
_log4auto_Error("1", "main.au3")
_log4auto_Warn("1", "main.au3")
_log4auto_Info("1", "main.au3")
_log4auto_Debug("1", "main.au3")
_log4auto_Trace("1", "main.au3")
