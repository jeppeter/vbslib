#include-once

#include <WinAPIHObj.au3>
#include <WinAPIProc.au3>


; #INDEX# =======================================================================================================================
; Title .........: log4auto
; AutoIt Version : 3.2
; Language ......: English
; Description ...: Functions that assist with logging.
; Author(s) .....: jeppeter
; ===============================================================================================================================

; #CURRENT# =====================================================================================================================
;_log4auto_Debug
;_log4auto_Error
;_log4auto_Fatal
;_log4auto_Info
;_log4auto_Message
;_log4auto_SetFormat
;_log4auto_AddLogFile
;_log4auto_RmLogFile
;_log4auto_Trace
;_log4auto_Warn
;_log4auto_SetLogLevel
; ===============================================================================================================================

; Logging Enumerations
Global Enum $LOG4AUTO_LEVEL_TRACE,$LOG4AUTO_LEVEL_DEBUG,$LOG4AUTO_LEVEL_INFO, _
	$LOG4AUTO_LEVEL_WARN,$LOG4AUTO_LEVEL_ERROR,$LOG4AUTO_LEVEL_FATAL

; Internal variables
Global Const $__aLog4autoLevels[6] = ["Trace","Debug","Info","Warning","Error","Fatal"]
Global Const $__LOG4AUTO_VERSION = "0.1"

; Internal - Default configuration options
Global $__LOG4AUTO_FORMAT = "[${date}:${level}:${file}:${line}] ${message}" _
	, $__LOG4AUTO_LEVEL = $LOG4AUTO_LEVEL_ERROR

Global $__log4auto_oFiles = ObjCreate("Scripting.Dictionary")
Global $__log4auto_errhd = _WinAPI_GetStdHandle(2)
Global $__log4auto_backout = True


#region Configuration Functions


; #FUNCTION# ====================================================================================================================
; Name...........: _log4auto_SetFormat
; Description ...: Configures the format of logging messages (Default: "${date} ${level} ${message}").
; Syntax.........: _log4auto_SetFormat($sFormat)
; Parameters ....: $sFormat - A string representing the log message format. Formats can include the following macros:
;                  |${date} = Long date (i.e. MM/DD/YYYY HH:MM:SS)
;                  |${host} = Hostname of local machine
;                  |${level} = The current log level
;                  |${message} = The log message
;                  |${newline} = Insert a newline
;                  |${shortdate} = Short date (i.e. MM/DD/YYYY)
; Return values .: Success - Returns 0
;                  Failure - Returns 0
;                  @Error  - 0 = No error.
; Author ........: jeppeter
; Modified.......:
; Remarks .......:
; ===============================================================================================================================
Func _log4auto_SetFormat($sFormat)
	$__LOG4AUTO_FORMAT = $sFormat
EndFunc   ;==>_log4auto_SetFormat

; #FUNCTION# ====================================================================================================================
; Name...........: _log4auto_AddLogFile
; Description ...: Sets file to add
; Syntax.........: _log4a_AddLogFile($sFile, $isappend=True)
; Parameters ....: $sFile - A string specifying the path of the log file.
;                  | $isappend for the file is whether appending or not
; Return values .: Success - Returns 0
;                  Failure - Returns 1
;                  @Error  - 0 = No error.
; Author ........: jeppeter
; Modified.......:
; Remarks .......:
; ===============================================================================================================================
Func _log4auto_AddLogFile($sFile,$isappend=True)
	Local $hd
	Local $mode
	if $__log4auto_oFiles.Exists($sFile) Then
		return SetError(0,0,0)
	EndIf

	$mode = $FO_CREATEPATH 
	If $isappend Then
		$mode = $mode + $FO_APPEND
	Else
		$mode = $mode + $FO_overwrite
	EndIf

	$hd = FileOpen($sFile, $mode)

	if $hd = -1 Then
		return SetError(1,0,-1)
	EndIf

	$__log4auto_oFiles.Add($sFile, $hd)
	return SetError(0,0,0)
EndFunc   ;==>_log4auto_AddLogFile

; #FUNCTION# ====================================================================================================================
; Name...........: _log4auto_RmLogFile
; Description ...: Sets file to remove
; Syntax.........: _log4a_RmLogFile($sFile)
; Parameters ....: $sFile - A string specifying the path of the log file.
; Return values .: Success - Returns 0
;                  Failure - Returns 1
;                  @Error  - 0 = No error.
; Author ........: jeppeter
; Modified.......:
; Remarks .......:
; ===============================================================================================================================
Func _log4auto_RmLogFile($sFile)
	Local $hd
	if Not $__log4auto_oFiles.Exists($sFile) Then
		return SetError(1,0,1)
	EndIf

	$hd = $__log4auto_oFiles.Item($sFile)
	FileClose($hd)
	$__log4auto_oFiles.Remove($sFile)
	return SetError(0,0,0)
EndFunc   ;==>_log4auto_RmLogFile


; #FUNCTION# ====================================================================================================================
; Name...........: _log4auto_SetLogLevel
; Description ...: Sets LogLevel
; Syntax.........: _log4a_SetLogLevel($level)
; Parameters ....: $level can be $LOG4AUTO_LEVEL_ERROR , $LOG4AUTO_LEVEL_FATAL , $LOG4AUTO_LEVEL_WARN, _
;                  $LOG4AUTO_LEVEL_INFO , $LOG4AUTO_LEVEL_DEBUG , $LOG4AUTO_LEVEL_TRACE
; Return values .: old log level
;                  Errors set error 1
; Author ........: jeppeter
; Modified.......:
; Remarks .......:
; ===============================================================================================================================
Func _log4auto_SetLogLevel($level=$LOG4AUTO_LEVEL_ERROR)
	Local $oldlevel
	$oldlevel = $__LOG4AUTO_LEVEL
	If $level < $LOG4AUTO_LEVEL_TRACE And $level > $LOG4AUTO_LEVEL_FATAL Then
		return SetError(1,0, $oldlevel)
	EndIf
	$__LOG4AUTO_LEVEL = $level
	return SetError(0,0,$oldlevel)
EndFunc

#endregion Configuration Functions

#region Message Functions


; #FUNCTION# ====================================================================================================================
; Name...........: _log4auto_Message
; Description ...: Logs a message to the configured outputs.
; Syntax.........: _log4a_Message($sMessage, $eLevel)
; Parameters ....: $sMessage - A string containing the message to log.
;                  $eLevel - An integer specifying a level enumeration value. Must be one of the following:
;                  |$LOG4A_LEVEL_TRACE
;                  |$LOG4A_LEVEL_DEBUG
;                  |$LOG4A_LEVEL_INFO
;                  |$LOG4A_LEVEL_WARN
;                  |$LOG4A_LEVEL_ERROR
;                  |$LOG4A_LEVEL_FATAL
;                  $bOverride - A boolean specifying to override log filters. If true, messages will be logged regardless of
;                               enabled status or log level range (min/max values).
; Return values .: Success - Returns 0
;                  Failure - Returns 0
;                  @Error  - 0 = No error.
;                  |1 = Invalid $eLevel parameter
;                  |2 = Level is out of range
; Author ........: jeppeter
; Modified.......:
; Remarks .......:
; ===============================================================================================================================
Func _log4auto_Message($sMessage, $eLevel = $LOG4A_LEVEL_INFO, $file = @ScriptFullPath , $line=@ScriptLineNumber)
	If $eLevel < $LOG4AUTO_LEVEL_TRACE Or $eLevel > $LOG4AUTO_LEVEL_FATAL Then Return SetError(1)

	If $eLevel < $__LOG4AUTO_LEVEL Then
		return SetError(0)
	EndIf

	Local $f
	Local $fh
	Local $formatted = __log4auto_FormatMessage($sMessage,$__aLog4autoLevels[$eLevel],$file,$line)

	For $f in $__log4auto_oFiles
		$fh = $__log4auto_oFiles.Item($f)
		FileWrite($fh, $formatted & @CRLF)
	Next


	If $__log4auto_errhd <> -1 Then
		_WinAPI_WriteConsole($__log4auto_errhd, $formatted & @CRLF)
	EndIf

	if $__log4auto_backout Then
		DllCall("kernel32.dll", "none", "OutputDebugString", "str", $formatted & @CRLF)
	EndIf
	return SetError(0)
EndFunc   ;==>_log4auto_Message


; #FUNCTION# ====================================================================================================================
; Name...........: _log4auto_Trace
; Description ...: Logs a message at the trace level.
; Syntax.........: _log4auto_Trace($sMessage,$file,$line)
; Parameters ....: $sMessage - A string containing the message to log.
;                  $file file to formatted 
;                  $line line to formatted
; Return values .: Success - Returns 0
;                  Failure - Returns 0
;                  @Error  - 0 = No error.
; Author ........: jeppeter
; Modified.......:
; Remarks .......:
; ===============================================================================================================================
Func _log4auto_Trace($sMessage, $file=@ScriptFullPath,$line=@ScriptLineNumber)
	_log4auto_Message($sMessage, $LOG4AUTO_LEVEL_TRACE, $file,$line)
EndFunc   ;==>_log4auto_Trace

; #FUNCTION# ====================================================================================================================
; Name...........: _log4auto_Debug
; Description ...: Logs a message at the debug level.
; Syntax.........: _log4auto_Debug($sMessage, $file,$line)
; Parameters ....: $sMessage - A string containing the message to log.
;                  $file file to formatted 
;                  $line line to formatted
; Return values .: Success - Returns 0
;                  Failure - Returns 0
;                  @Error  - 0 = No error.
; Author ........: jeppeter
; Modified.......:
; Remarks .......:
; ===============================================================================================================================
Func _log4auto_Debug($sMessage, $file=@ScriptFullPath,$line=@ScriptLineNumber)
	_log4auto_Message($sMessage, $LOG4AUTO_LEVEL_DEBUG, $file,$line)
EndFunc   ;==>_log4auto_Debug

; #FUNCTION# ====================================================================================================================
; Name...........: _log4auto_Info
; Description ...: Logs a message at the info level.
; Syntax.........: _log4auto_Info($sMessage, $file,$line)
; Parameters ....: $sMessage - A string containing the message to log.
;                  $file file to formatted 
;                  $line line to formatted
; Return values .: Success - Returns 0
;                  Failure - Returns 0
;                  @Error  - 0 = No error.
; Author ........: jeppeter
; Modified.......:
; Remarks .......:
; ===============================================================================================================================
Func _log4auto_Info($sMessage,$file=@ScriptFullPath,$line=@ScriptLineNumber)
	_log4auto_Message($sMessage, $LOG4AUTO_LEVEL_INFO, $file,$line)
EndFunc   ;==>_log4auto_Info

; #FUNCTION# ====================================================================================================================
; Name...........: _log4auto_Warn
; Description ...: Logs a message at the warn level.
; Syntax.........: _log4auto_Warn($sMessage, $file,$line)
; Parameters ....: $sMessage - A string containing the message to log.
;                  $file file to formatted 
;                  $line line to formatted
; Return values .: Success - Returns 0
;                  Failure - Returns 0
;                  @Error  - 0 = No error.
; Author ........: jeppeter
; Modified.......:
; Remarks .......:
; ===============================================================================================================================
Func _log4auto_Warn($sMessage, $file=@ScriptFullPath,$line=@ScriptLineNumber)
	_log4auto_Message($sMessage, $LOG4AUTO_LEVEL_WARN, $file,$line)
EndFunc   ;==>_log4auto_Warn

; #FUNCTION# ====================================================================================================================
; Name...........: _log4auto_Error
; Description ...: Logs a message at the error level.
; Syntax.........: _log4auto_Error($sMessage, $file,$line)
; Parameters ....: $sMessage - A string containing the message to log.
;                  $file file to formatted 
;                  $line line to formatted
; Return values .: Success - Returns 0
;                  Failure - Returns 0
;                  @Error  - 0 = No error.
; Author ........: jeppeter
; Modified.......:
; Remarks .......:
; ===============================================================================================================================
Func _log4auto_Error($sMessage, $file=@ScriptFullPath,$line=@ScriptLineNumber)
	_log4auto_Message($sMessage, $LOG4AUTO_LEVEL_ERROR, $file,$line)
EndFunc   ;==>_log4auto_Error

; #FUNCTION# ====================================================================================================================
; Name...........: _log4auto_Fatal
; Description ...: Logs a message at the fatal level.
; Syntax.........: _log4auto_Fatal($sMessage, $file,$line)
; Parameters ....: $sMessage - A string containing the message to log.
;                  $file file to formatted 
;                  $line line to formatted
; Return values .: Success - Returns 0
;                  Failure - Returns 0
;                  @Error  - 0 = No error.
; Author ........: Michael Mims (zorphnog)
; Modified.......:
; Remarks .......:
; ===============================================================================================================================
Func _log4auto_Fatal($sMessage, $file=@ScriptFullPath,$line=@ScriptLineNumber)
	_log4auto_Message($sMessage, $LOG4AUTO_LEVEL_FATAL, $file,$line)
EndFunc   ;==>_log4auto_Fatal

#endregion Message Functions

#region Internal Functions

Func __log4auto_FormatMessage($sMessage, $sLevel,$file=@ScriptFullPath,$line=@ScriptLineNumber)
	Local $sFormatted = $__LOG4AUTO_FORMAT

	$sFormatted = StringReplace($sFormatted, "${file}", $file)
	$sFormatted = StringReplace($sFormatted, "${line}", $line)
	$sFormatted = StringReplace($sFormatted, "${date}", _
		StringFormat("%04d/%02d/%02d %02d:%02d:%02d", @YEAR, @MON, @MDAY,  @HOUR, @MIN, @SEC))
	$sFormatted = StringReplace($sFormatted, "${host}", @ComputerName)
	$sFormatted = StringReplace($sFormatted, "${level}", $sLevel)
	$sFormatted = StringReplace($sFormatted, "${message}", $sMessage)
	$sFormatted = StringReplace($sFormatted, "${newline}", @CRLF)
	$sFormatted = StringReplace($sFormatted, "${shortdate}", _
		StringFormat("%02d\\%02d\\%04d", @MON, @MDAY, @YEAR))

	Return $sFormatted
EndFunc   ;==>__log4auto_FormatMessage

#endregion Internal Functions