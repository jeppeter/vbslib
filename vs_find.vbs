
Option Explicit

Class VsFilterVersion
    Private m_version
    Public Function FilterVersion(line)
        dim re,result,num,a,resa,b
        set re = new regexp
        '  to get the gox.x.x version number
        re.Pattern = "^Microsoft Visual Studio\s+.*\s+([0-9]+((\.[0-9]*)+))"
        set result = re.Execute(line)   
        num = 0
        if Not IsEmpty(result) Then
            for Each a in result
            	' we take two version
                re.Pattern = "[\d]+\.[\d]+"
                set resa = re.Execute(a)
                if Not IsEmpty(resa) Then
                    for Each b in resa
                        b = Trim(b)
                        m_version=b
                    Next
                End If
            Next
        End If
    End Function
    Public Function  GetVersion()
        GetVersion=m_version
    End Function
End Class


Function FilterVsversion(line,filterctx)
    filterctx.FilterVersion(line)
    FilterVsversion=true
End Function

dim vscmdversion

Function GetVsVersion(devenvcom)
	dim cmd
	set vscmdversion = new VsFilterVersion
	call GetRunOut(devenvcom,"/?","FilterVsversion","vscmdversion")
	GetVsVersion=vscmdversion.GetVersion()
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

Function FindoutInstallBasedir(directory)
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
	            FindoutInstallBasedir=sp
	            Exit Function
	        End If
	        sp = sp & "\" & narr(j)
	        j = j+1
	   Wend
	End If
	FindoutInstallBasedir=Empty
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
	devcom=GetDevenvCom()
	if IsEmpty(devcom) Then
		' nothing to get ,so we return error
		IsInstallVisualStudio=Null
		Exit Function
	End If
	WScript.Stderr.writeline("devcom[" & devcom &"]")

	' now get the version 
	getversion = GetVsVersion(devcom)
	if Len(getversion) <= 0 Then
		IsInstallVisualStudio=Null
		Exit Function
	End If
	getversion=CDbl(getversion)


	idx = 0
	max = Ubound(versions) + 1
	Do while idx < max
		curver = versions(idx)
		If curver >= (getversion-0.01) and curver <= (getversion + 0.01) and (version) <= (getversion) Then
			instdir=CheckInstallBasedir(devcom)
			if not IsEmpty(instdir) Then
				IsInstallVisualStudio=curver
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

