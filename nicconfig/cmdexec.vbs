
Option Explicit

Function RunCommand(cmd)
	dim objsh,res
	set objsh = wscript.CreateObject("WScript.Shell")
	res = objsh.Run(cmd,1,true)
	RunCommand=res
End Function


Function RuncmdOutput(cmd)
	dim outputlines
	dim runobj
	dim objshell
	dim reply

	outputlines=""
	set objshell = WScript.CreateObject("WScript.Shell")
	set runobj = objshell.Exec(cmd)
	Do While Not runobj.Stdout.AtEndOfStream
		reply = runobj.Stdout.ReadAll()
		outputlines = outputlines & reply
	Loop
	RuncmdOutput=outputlines
End Function

Function GetInterfaceIndexByFirst()
	dim outputlines
	dim runcmd
	dim sarr
	dim num
	dim re
	dim matches

	runcmd = "wmic nicconfig where "  & chr(34) & "IPEnabled=true" & chr(34) & " Get Index"
	outputlines=RuncmdOutput(runcmd)
	sarr=Split(outputlines,chr(10))
	If Ubound(sarr) >= 2 and Len(sarr(1)) > 0 Then
		num= sarr(1)
		set re =  new regexp
		re.Pattern = "[0-9]+"
		set matches = re.Execute(sarr(1))
		if matches.Count > 0 Then
			GetInterfaceIndexByFirst=matches(0)
		Else
			GetInterfaceIndexByFirst="-1"
		End If
	Else
		GetInterfaceIndexByFirst="-1"
	End If
End Function

Function GetInterfaceIndexByMac(macaddr)
	dim maclower
	dim outputlines
	dim cmd
	dim re
	dim sarr
	dim larr
	dim curline
	dim retval
	dim partstr
	dim matches
	maclower = LCase(macaddr)
	cmd = "wmic nicconfig get Index,MACAddress"
	outputlines = RuncmdOutput(cmd)

	larr =  Split(outputlines,chr(10))
	For Each curline in larr
		curline = LCase(curline)
		retval = InStr(curline,maclower)
		If not IsNull(retval) Then
			if retval > 0 Then
				partstr = Left(curline,(retval - 1))
				set re = New regexp
				re.Pattern = "[0-9]+"
				set matches = re.Execute(partstr)
				if matches.Count > 0 Then
					GetInterfaceIndexByMac=matches(0)
					Exit Function
				End If
			End If
		End If
	Next
	GetInterfaceIndexByMac="-1"	
End Function


Function GetWmicResult(outputs)
	dim re
	dim sarr
	dim curline
	dim matches
	dim retval
	dim lsarr
	dim partre
	dim partmatches
	set re = new regexp
	set partre = new regexp
	sarr = Split(outputs,chr(10))
	re.Pattern = "\s+ReturnValue\s+=\s+([\d]+);"
	partre.Pattern = "[\d]+"
	if Ubound(sarr) > 0 THen
		For Each curline in sarr
			curline = replace(curline,chr(13),"")
			set matches = re.Execute(curline)
			if matches.Count > 0 Then
				set partmatches = partre.Execute(curline)
				if partmatches.Count > 0 Then
					retval = CInt(partmatches(0))
					if retval = 0 Then
						GetWmicResult=True
					Else
						GetWmicResult=False
					End If
					Exit Function
				End If
			End If
		Next
	End If
	GetWmicResult=False
End Function

Function SetIpNetMask(index,ipaddr,netmask)
	dim cmd
	dim outputs
	cmd = "wmic nicconfig where index=" & index & " call enablestatic(" & chr(34) & ipaddr & chr(34) & "),(" & chr(34) & netmask & chr(34) & ")"
	WScript.Echo "run cmd [" & cmd & "]"
	outputs=RuncmdOutput(cmd)
	SetIpNetMask=GetWmicResult(outputs)
End Function

Function SetGateWay(index,gatewayip)
	dim cmd
	dim optip
	dim gatestr
	dim metricstr
	dim curip
	dim curnum
	dim outputs
	gatestr = ""
	metricstr = ""
	if VarType(gatewayip) & vbArray Then
		curnum = 1
		For Each curip in gatewayip
			If Len(gatestr) = 0 Then
				gatestr =   chr(34) & curip & chr(34)
				metricstr = curnum
			Else
				gatestr = gatestr & "," & chr(34) & curip & chr(34)
				metricstr = metricstr & "," & curnum
			End If
			curnum = curnum + 1
		Next
	Else
		gatestr = chr(34) & gatewayip & chr(34)
		metricstr = "1"
	End If
	cmd = "wmic nicconfig where index=" & index & " call setgateways("& gatestr & "),(" & metricstr & ")"
	WScript.Echo "run cmd [" & cmd & "]"
	outputs=RuncmdOutput(cmd)
	SetGateWay=GetWmicResult(outputs)
End Function

Function SetDns(index,dnsserver)
	dim cmd
	dim curip
	dim dnsstr
	dim outputs
	if VarType(dnsserver) & vbArray Then
		dnsstr = ""
		For Each curip in dnsserver
			If Len(dnsstr) = 0 Then
				dnsstr = chr(34) & curip & chr(34)				
			Else
				dnsstr = dnsstr & "," & chr(34) & curip & chr(34)
			End If
		Next		
	Else
		dnsstr = chr(34) & dnsserver & chr(34)
	End If
	cmd = "wmic nicconfig where index=" & index & " call SetDNSServerSearchOrder(" & dnsstr & ")"
	WScript.Echo "run cmd [" & cmd & "]"
	outputs=RuncmdOutput(cmd)
	SetDns=GetWmicResult(outputs)
End Function

Function SetIpDhcp(index)
	dim cmd
	dim outputs
	cmd = "wmic nicconfig where index=" & index & " call EnableDHCP"
	WScript.Echo "run cmd [" & cmd & "]"
	outputs=RuncmdOutput(cmd)
	SetIpDhcp=GetWmicResult(outputs)
End Function

Function SetDnsDhcp(index)
	dim cmd
	dim outputs
	cmd = "wmic nicconfig where index=" & index & " call SetDNSServerSearchOrder()"
	WScript.Echo "run cmd [" & cmd & "]"
	outputs=RuncmdOutput(cmd)
	SetDnsDhcp=GetWmicResult(outputs)
End Function

