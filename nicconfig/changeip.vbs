
Option Explicit

Sub includeFile(fSpec)
    With CreateObject("Scripting.FileSystemObject")
       executeGlobal .openTextFile(fSpec).readAll()
    End With
End Sub

Function GetScriptDir()
	dim fso ,scriptpath
	Set fso = CreateObject("Scripting.FileSystemObject") 
	GetScriptDir=fso.GetParentFolderName(Wscript.ScriptFullName)
End Function


call includeFile( GetScriptDir() & "\baseop.vbs")
call includeFile( GetScriptDir() & "\fileop.vbs")
call includeFile( GetScriptDir() & "\cmdexec.vbs")
call includeFile( GetScriptDir() & "\vbsjson.vbs")

dim ipfile
dim windir 
dim logf
dim existfile
dim content
dim jsondec
dim runok
dim index
dim retval
dim json

windir = GetEnv("WINDIR")

If IsNull(windir) Then
	WScript.Stderr.WriteLine("can not find WINDIR")
	Wscript.Quit(3)
End If

logf = windir & "\..\iplog.txt"
ipfile = windir & "\..\btcmd"

If WScript.Arguments.Count > 0 Then
	ipfile = WScript.Arguments(0)
End If

If WScript.Arguments.Count > 1 Then
	logf = WScript.Arguments(1)
End If

existfile = FileExists(ipfile)
if not existfile Then
	'  nothing to handle
	WScript.Quit(0)
End If

content = ReadFileAll(ipfile)
if Len(content) = 0 Then
	' nothing to handle
	DeleteFileSafe ipfile
	Wscript.Quit(0)
End If


set json = new VbsJson
set jsondec = json.decode(content)
if IsNull(jsondec) or IsEmpty(jsondec) Then
	LogFile logf,"can not parse [" & content & "]"
	Wscript.Quit(3)
End If

runok=False

if jsondec.Exists("ipconfig") Then
	if jsondec("ipconfig").Exists("ipaddr") and _
		jsondec("ipconfig").Exists("netmask") and _
		jsondec("ipconfig").Exists("gateway") and _
		jsondec("ipconfig").Exists("dns") Then

		if jsondec("ipconfig").Exists("macaddr") Then
			index=GetInterfaceIndexByMac(jsondec("ipconfig")("macaddr"))
		Else
			index=GetInterfaceIndexByFirst()
		End If
		if index <> "-1" Then
			retval = SetIpNetMask(index,jsondec("ipconfig")("ipaddr"), jsondec("ipconfig")("netmask"))
			if retval = 0 Then
				retval = SetGateWay(index,jsondec("ipconfig")("gateway"))
				if retval = 0 Then
					retval = SetDns(index,jsondec("ipconfig")("dns"))
					if retval = 0 Then
						runok = True
					Else
						LogFile logf,"can not set dns"
					End If
				Else
					LogFile logf,"can not set gateway"
				End If
			Else
				LogFile logf,"can not set SetIpNetMask"
			End If
		End If

	Else
		LogFile logf,"not valid in ipconfig for ipaddr or netmask or gateway or dns"
	End If
End If


if runok Then
	DeleteFileSafe(ipfile)
	LogFile logf,"Run Ip Change succ"
	WScript.Quit(0)
End If

WScript.Quit(3)