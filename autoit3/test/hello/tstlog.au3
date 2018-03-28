#AutoIt3Wrapper_Run_Au3Stripper=y
#Au3Stripper_Parameters=/mo
#include "log4auto.au3"

_log4auto_SetLogLevel($LOG4AUTO_LEVEL_TRACE)
_log4auto_AddLogFile("cc.log",False)
_log4auto_Trace("hello world", "dd.au3")
_log4auto_Debug("hello world", "dd.au3")
_log4auto_Info("hello world", "dd.au3")
_log4auto_Warn("hello world", "dd.au3")
_log4auto_Error("hello world", "dd.au3")
_log4auto_Fatal("hello world", "dd.au3")
_log4auto_RmLogFile("cc.log")
_log4auto_Trace("hello world2", "dd.au3")
_log4auto_Debug("hello world2", "dd.au3")
_log4auto_Info("hello world2", "dd.au3")
_log4auto_Warn("hello world2", "dd.au3")
_log4auto_Error("hello world2", "dd.au3")
_log4auto_Fatal("hello world2", "dd.au3")
