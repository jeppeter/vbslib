
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


Function SetIpNetMask(index,ipaddr,netmask)
	dim cmd
	cmd = "wmic nicconfig where index=" & index & " call enablestatic(" & chr(34) & ipaddr & chr(34) & "),(" & chr(34) & netmask & chr(34) & ")"
	WScript.Echo "run cmd [" & cmd & "]"
	SetIpNetMask=RunCommand(cmd)
End Function

Function SetGateWay(index,gatewayip)
	dim cmd
	dim optip
	dim gatestr
	dim metricstr
	dim curip
	dim curnum
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
	SetGateWay=RunCommand(cmd)
End Function

Function SetDns(index,dnsserver)
	dim cmd
	dim curip
	dim dnsstr
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
	SetDns=RunCommand(cmd)
End Function
