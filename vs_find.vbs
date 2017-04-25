
Option Explicit


Function GetVsVersion()
	dim regk ,regv,regobj,regarr,c
	regk = "HKEY_CLASSES_ROOT\VisualStudio.DTE\CurVer\"
	regv = ReadReg(regk)
	if IsEmpty(regv) Then
		GetVsVersion=""
		Exit Function
	End If
	set regobj = CreateObject("VBScript.RegExp")
	regobj.Global = True
	regobj.IgnoreCase = True	
	regobj.Pattern = "[\d]+\.[\d]+"	
	set regarr = regobj.Execute(regv)
	for each c in regarr
		set regarr = Nothing
		set regobj = Nothing
		GetVsVersion=c
		Exit Function
	Next
	GetVsVersion=""
End Function

Function GetDevenvCom()
	dim regkey,regval
	dim regobj
	regkey="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\devenv.exe\"
	regval=ReadReg(regkey)
	if IsEmpty(regval) Then
		GetDevenvCom=Empty
		Exit Function
	End If
	set regobj = CreateObject("VBScript.RegExp")
	regobj.Global = True
	regobj.IgnoreCase = True	
	regobj.Pattern = "^" & chr(34) 
	regval = regobj.Replace(regval,"")
	regobj.Pattern = chr(34) & "$"
	regval = regobj.Replace(regval,"")
	regobj.Pattern = "\.exe$"
	regval=regobj.Replace(regval,".com")
	set regobj = Nothing
	GetDevenvCom=regval
End Function



Function CheckInstallBasedir(directory)
	dim narr,i,j,num,sp,ret,clexe
	narr = Split(directory,"\")
	if IsArray(narr) Then
	   num=Ubound(narr)
	   i=num
	   j=1
	   sp=narr(0)
	   While j < (i+1)
	        clexe=sp & "\Common7\IDE\devenv.exe"
	        ret=FileExists(clexe)
	        if ret Then
	            CheckInstallBasedir=sp
	            Exit Function
	        End If
	        sp = sp & "\" & narr(j)
	        j = j+1
	   Wend
	End If
	CheckInstallBasedir=Empty
End Function

Function IsInstallVisualStudio(version)
	dim curval
	dim curver
	dim instdir
	dim versions
	dim values 
	dim idx
	dim max
	dim devcom
	dim getversion
	versions=Array(15.0,14.0,12.0)
	values=Array("15.0","14.0","12.0")
	devcom=GetDevenvCom()
	if IsEmpty(devcom) Then
		' nothing to get ,so we return error
		IsInstallVisualStudio=Null
		Exit Function
	End If
	'WScript.Stderr.writeline("devcom[" & devcom &"]")

	' now get the version 
	getversion = GetVsVersion()
	if Len(getversion) <= 0 Then
		IsInstallVisualStudio=Null
		Exit Function
	End If
	getversion=CDbl(getversion)


	idx = 0
	max = Ubound(versions) + 1
	Do while idx < max
		curver = versions(idx)
		curval = values(idx)
		If curver >= (getversion-0.01) and curver <= (getversion + 0.01) and (version) <= (getversion) Then
			instdir=CheckInstallBasedir(devcom)
			if not IsEmpty(instdir) Then
				IsInstallVisualStudio=curval
				Exit Function
			End If
		End If
		idx = idx + 1
	Loop
	IsInstallVisualStudio=Null
End Function

Function GetVisualStudioInstdir(version)
	dim curval
	dim curver
	dim instdir
	dim versions
	dim values 
	dim idx
	dim max
	dim devcom
	dim getversion
	versions=Array(15.0,14.0,12.0)
	devcom=GetDevenvCom()
	if IsEmpty(devcom) Then
		' nothing to get ,so we return error
		GetVisualStudioInstdir=Null
		Exit Function
	End If

	instdir=CheckInstallBasedir(devcom)
	if not IsEmpty(instdir) Then
		GetVisualStudioInstdir=instdir
		Exit Function
	End If
	GetVisualStudioInstdir=Null
End Function

Function VsVerifyCL(vsver)
	dim getver
	getver=IsInstallVisualStudio(vsver)
	if not IsNull(getver) Then
		VsVerifyCL=1
	Else
		VsVerifyCL=0
	End If	
End Function

Function GetNmake(basedir,vsver)
	dim msvcdir
	dim curdir
	dim fso
	dim i,max
	dim reads
	dim arrdir
	dim curfile
	if vsver = "12.0" or vsver = "14.0" Then
		curfile = basedir & "\VC\bin\nmake.exe"
		Set fso = CreateObject("Scripting.FileSystemObject")
		if fso.FileExists(curfile) Then
			GetNmake=curfile
			set fso = Nothing
			Exit Function
		End If
		GetNmake=Empty
	ElseIf vsver = "15.0" Then
		msvcdir= basedir & "\VC\Tools\MSVC"
		reads = ReadDir(msvcdir)
		arrdir = Split(reads,";")
		i = 0
		max = Ubound(arrdir) + 1
		Set fso = CreateObject("Scripting.FileSystemObject")
		Do While i < max
			curdir = arrdir(i)
			curfile = curdir & "\bin\Hostx64\x64\nmake.exe"
			if fso.FileExists(curfile) Then
				set fso = Nothing
				GetNmake=curfile
				Exit Function
			End If
			i = i + 1
		Loop
		set fso = Nothing
		GetNmake=empty
	Else
		GetNmake=empty		
	End If	
End Function

