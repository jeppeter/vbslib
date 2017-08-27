
Option Explicit

Function GetEnv(varname)
	dim wsh,val,key
	set wsh = WScript.CreateObject("WScript.Shell")
	key = "%" & varname & "%"
	val = wsh.ExpandEnvironmentStrings(key)
	if val  = key Then
		GetEnv=null
	else
		GetEnv=val
	end if
End Function
