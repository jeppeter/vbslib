Option Explicit
'  this is the file to collection of set ip address and get ip address

Function GetAllInterfaces(svrname)
	dim objwmi
	dim interfaces
	set objwmi = GetObject("winmgmts:\\"& svrname & "\root\cimv2")
	set interfaces=objwmi.ExecQuery ("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
	GetAllInterfaces=interfaces
Function End