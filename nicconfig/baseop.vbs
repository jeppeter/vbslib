
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

Function FormatArray(arr)
	dim retstr
	dim icnt
	dim curitem
	retstr = ""
	if VarType(arr) = vbObject or (VarType(arr) & vbArray) = vbArray  Then
		retstr= retstr & "["
		icnt = 0
		For Each curitem in arr
			if icnt > 0 Then
				retstr = retstr & ","
			End If
			retstr = retstr & curitem
			icnt = icnt + 1
		Next
		retstr = retstr & "]"
	Elseif VarType(arr) = vbString Then
		retstr = chr(34) & arr & chr(34)
	Else
		retstr = arr
	End If

	FormatArray=retstr
End Function
