
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
dim countmax
dim icnt
dim waitsleep
dim lasterrmsg

countmax = 5
windir = GetEnv("WINDIR")

If IsNull(windir) Then
	WScript.Echo "can not find WINDIR"
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

If jsondec.Exists("ipconfig") Then
	icnt = 0
	waitsleep = False
	lasterrmsg = "first start"
	countmax=60
	Do While True
		icnt = icnt + 1
		If waitsleep Then
			WScript.Sleep 5000
		End If
		waitsleep=True

		If icnt >= countmax Then
			LogFile logf,lasterrmsg
			WScript.Echo lasterrmsg
			WScript.Quit(3)
		End If

		If jsondec("ipconfig").Exists("macaddr") Then
			index=GetInterfaceIndexByMac(jsondec("ipconfig")("macaddr"))
		Else
			index=GetInterfaceIndexByFirst()
		End If

		If index = "-1" Then
			if jsondec("ipconfig").Exists("macaddr") Then
				lasterrmsg = "[" & icnt & "] can not find [" & jsondec("ipconfig").Exists("macaddr") &"]"
			Else
				lasterrmsg = "[" & icnt & "] can not find first ipenabled ethernet card"
			End If
			WScript.Echo lasterrmsg
			continue
		End If

		if jsondec("ipconfig").Exists("ipaddr") and _
			jsondec("ipconfig").Exists("netmask") and _
			jsondec("ipconfig").Exists("gateway") Then

			retval = SetIpNetMask(index,jsondec("ipconfig")("ipaddr"), jsondec("ipconfig")("netmask"))
			If not retval Then
				lasterrmsg = "[" & icnt & "] can not set ipaddr[" & jsondec("ipconfig")("ipaddr") & "] netmask[" & jsondec("ipconfig")("netmask") & "]"
				WScript.Echo lasterrmsg
				continue
			End If

			retval = retval = SetGateWay(index,jsondec("ipconfig")("gateway"))
			If not retval Then
				lasterrmsg = "[" & icnt & "] can not set gateway[" & jsondec("ipconfig")("gateway") & "]"
				WScript.Echo lasterrmsg
				continue
			End If
		Else
			retval = SetIpDhcp(index)
			If not retval Then
				lasterrmsg = "[" & icnt & "] can not set dhcp"
				Wscript.Echo lasterrmsg
				continue
			End If
		End If

		If jsondec("ipconfig").Exists("dns") Then
			retval = SetDns(index,jsondec("ipconfig")("dns"))
			If not retval Then
				lasterrmsg = "[" & icnt & "] can not set dns [" & jsondec("ipconfig")("dns") & "]"
				WScript.Echo lasterrmsg
				continue
			End If
		Else
			If not jsondec("ipconfig").Exists("ipaddr") or  _
				not jsondec("ipconfig").Exists("netmask") or _
				not jsondec("ipconfig").Exists("gateway") Then
				retval = SetDnsDhcp(index)
				If not retval Then
					lasterrmsg = "[" & icnt & "] can not set dns dhcp"
					WScript.Echo lasterrmsg
					continue
				End If
			Else
				LogFile logf,"invalid ipconfig for dns dhcp"
				WScript.Quit(4)
			End If
		End If

		runok = True
		Exit Do
	Loop
End If


if runok Then
	DeleteFileSafe(ipfile)
	LogFile logf,"Run Ip Change succ"
	WScript.Quit(0)
End If

LogFile logf,"Failed Running"
WScript.Quit(3)