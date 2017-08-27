
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

Function GetInterfaceIndex()
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
			GetInterfaceIndex=matches(0)
		Else
			GetInterfaceIndex="-1"
		End If
	Else
		GetInterfaceIndex="-1"
	End If
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
	if gatewayip = "192.168.0.1" Then
		optip = "192.168.0.2"
	Else
		optip = "192.168.0.1"
	End If
	cmd = "wmic nicconfig where index=" & index & " call setgateways(" & chr(34) & gatewayip & chr(34) & "," & chr(34) & optip & chr(34) & "),(1,2)"
	WScript.Echo "run cmd [" & cmd & "]"
	SetGateWay=RunCommand(cmd)
End Function

Function SetDns(index,dnsserver)
	dim cmd
	cmd = "wmic nicconfig where index=" & index & " call SetDNSServerSearchOrder(" & chr(34) & dnsserver & chr(34) & ")"
	WScript.Echo "run cmd [" & cmd & "]"
	SetDns=RunCommand(cmd)
End Function
