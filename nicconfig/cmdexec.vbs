
Option Explicit

Function RuncmdOutput(cmd)
	dim outputlines
	dim runobj
	dim objshell
	dim reply

	outputlines=""
	set objshell = WScript.CreateObject("WScript.Shell")
	set runobj = objshell.Exec(cmd)
	Do While Not runobj.Stdout.AtEndOfStream
		reply = runobj.Stdout.ReadLine()
		outputlines = outputlines & reply & chr(10)
	Loop
	RuncmdOutput=outputlines
End Function

Function GetInterfaceIndex()
	dim outputlines
	dim runcmd
	dim sarr

	runcmd = "wmic nicconfig where "  & chr(34) & "IPEnabled=true" & chr(34) & " Get Index"
	outputlines=RuncmdOutput(runcmd)
	sarr=Split(outputlines,chr(10))
	If Ubound(sarr) >= 2 and Len(sarr(1)) > 0 Then
		GetInterfaceIndex=sarr(1)
	Else
		GetInterfaceIndex=-1
	End If
End Function