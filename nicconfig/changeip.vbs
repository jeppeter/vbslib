
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
dim logfile
dim exists
dim content
dim jsondec
dim runok
dim index
dim retval

windir = GetEnv("WINDIR")

If IsNull(windir) Then
	WScript.Stderr.WriteLine("can not find WINDIR")
	Wscript.Quit(3)
End If

logfile = windir & "\..\iplog.txt"
ipfile = windir & "\..\btcmd"




exists = FileExists ipfile
if not exists Then
	'  nothing to handle
	WScript.Quit(0)
End If

content = ReadFileAll ipfile
if Len(content) = 0 Then
	' nothing to handle
	DeleteFileSafe ipfile
	Wscript.Quit(0)
End If

jsondec = json.decode(content)
if IsNull(jsondec) or IsEmpty(jsondec) Then
	AppendFile logfile,"can not parse [" & content & "]"
	Wscript.Quit(3)
End If

runok=False

if jsondec.Exists("ipconfig") Then
	if jsondec("ipconfig").Exists("ipaddr") and _
		jsondec("ipconfig").Exists("netmask") and _
		jsondec("ipconfig").Exists("gateway") and _
		jsondec("ipconfig").Exists("dns") Then

		index=GetInterfaceIndex()
		if index <> "-1" Then
			retval = SetIpNetMask index,jsondec("ipconfig")("ipaddr"), jsondec("ipconfig")("netmask")
			if retval = 0 Then
				retval = SetGateWay index,jsondec("ipconfig")("gateway")
				if retval = 0 Then
					retval = SetDns index,jsondec("ipconfig")("dns")
					if retval = 0 Then
						runok = True
					Else
						AppendFile logfile,"can not set dns"
					End If
				Else
					AppendFile logfile,"can not set gateway"
				End If
			Else
				AppendFile logfile,"can not set SetIpNetMask"
			End If
		End If

	Else
		AppendFile logfile,"not valid in ipconfig for ipaddr or netmask or gateway or dns"
	End If
End If


if runok Then
	DeleteFileSafe ipfile
	WScript.Quit(0)
End If

WScript.Quit(3)