
Option Explicit

dim vsversion


class VSWhereVersion
	Private m_version
	Private m_instdir
	Public Function FilterVersion(line)
        dim rever,repath,result,num,a,resa,b
        dim setv
        dim mathed
        set rever = new regexp
        set repath = new regexp

        '  to get the gox.x.x version number
        rever.Pattern = "installationVersion:\s+([0-9]+((\.[0-9]*)+))$"
        repath.Pattern = "installationPath:\s+(.*)$"

        set result = rever.Execute(line)   
        num = 0
        setv=false
        mathed=false
        if Not IsEmpty(result) Then
            for Each a in result
                rever.Pattern = "[0-9]+((\.[0-9]*)+)"
                set resa = rever.Execute(a)
                if Not IsEmpty(resa) Then
                    for Each b in resa
                        b = Trim(b)
                        m_version=b                        
                        mathed=true
                    Next
                End If
            Next
        End If
        if not mathed Then
        	set result = repath.Execute(line)
            if Not IsEmpty(result) Then
                repath.Pattern = "installationPath:\s+"
                for each a in result
                    resa =repath.Replace(a,"")
                    m_instdir=resa
                    mathed=true
                Next
            End IF
        End If
        FilterVersion=setv
	End Function
    Public Function GetVersion()
         GetVersion=m_version
    End Function

    Public Function GetInstPath()
        GetInstPath=m_instdir
    End Function

End Class


Function GetVsVersion_New()
    dim patharr
    dim pathval
    dim curpath
    dim curcmd
    dim curver
    dim re,result,a
    Dim process_architecture
    pathval = GetEnv("PATH")
    If IsNull(pathval) Then
        GetVsVersion_New=""
        Exit Function
    End If

    patharr = Split(pathval & ";.",";")
    set vsversion = new VSWhereVersion
    For Each curpath in patharr
        If FileExists( curpath & "\" & "vswhere.exe") Then
            curcmd = curpath & "\" & "vswhere.exe"
            set vsversion = new VSWhereVersion
            call GetRunOut(curcmd,"","FilterText","vsversion")
            WScript.Stdout.Writeline("vsversion version " & vsversion.GetVersion())
        End If
    Next

    GetVsVersion_New=vsversion.GetVersion()
End Function

Function GetVsInstdir_New()
    dim patharr
    dim pathval
    dim curpath
    dim curcmd
    dim curver
    dim re,result,a
    Dim process_architecture
    pathval = GetEnv("PATH")
    If IsNull(pathval) Then
        GetVsInstdir_New=""
        Exit Function
    End If

    patharr = Split(pathval & ";.",";")
    set vsversion = new VSWhereVersion
    For Each curpath in patharr
        If FileExists( curpath & "\" & "vswhere.exe") Then
            curcmd = curpath & "\" & "vswhere.exe"
            set vsversion = new VSWhereVersion
            call GetRunOut(curcmd,"","FilterText","vsversion")
            WScript.Stdout.Writeline("vsversion version " & vsversion.GetVersion())
        End If
    Next

    GetVsInstdir_New=vsversion.GetInstPath()
End Function


Function GetVsVersion_Old()
	dim regk ,regv,regobj,regarr,c
	regk = "HKEY_CLASSES_ROOT\VisualStudio.DTE\CurVer\"
	regv = ReadReg(regk)
	if IsEmpty(regv) Then
		GetVsVersion_Old=""
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
		GetVsVersion_Old=c
		Exit Function
	Next
	GetVsVersion_Old=""
End Function

Function FilterVsVersion_New(version)
    dim values
    dim verpat
    dim result
    dim a
    values=Array("16.10")
    for each a in values
        set verpat = new regexp
        verpat.Pattern = "^" & a
        set result = verpat.Execute(version)
        if Not IsEmpty(result) Then
            FilterVsVersion_New=a
            Exit Function
        End If
    Next
    FilterVsVersion_New=version
End Function


Function GetVsVersion()
    dim version
    version=GetVsVersion_New()
    if Len(version) > 0 Then
    	GetVsVersion=FilterVsVersion_New(version)
    	Exit Function
    ENd If
    version=GetVsVersion_Old()
    GetVsVersion=version
End Function


Function GetDevenvCom_Old()
	dim regkey,regval
	dim regobj
	regkey="HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\devenv.exe\"
	regval=ReadReg(regkey)
	if IsEmpty(regval) Then
		GetDevenvCom_Old=Empty
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
	GetDevenvCom_Old=regval
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
	dim verget
	versions=Array(16.0,15.0,14.0,12.0)
	values=Array("16.0","15.0","14.0","12.0")
	verget=GetVsVersion_New()
	if len(verget) > 0 Then
		IsInstallVisualStudio=FilterVsVersion_New(verget)
		Exit Function
	End If

	devcom=GetDevenvCom_Old()
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
	instdir=GetVsInstdir_New()
	if len(instdir) > 0 Then
		GetVisualStudioInstdir=instdir
		Exit Function
	End If

	versions=Array(16.0,15.0,14.0,12.0)
	devcom=GetDevenvCom_Old()
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
	ElseIf vsver = "15.0" or vsver = "16.0" or vsver = "16.10" Then
	    WScript.Stderr.writeline("get vsver "& vsver)
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

Function GetVsAllBatchCall(vsver,basedir,compiletarget)
	dim cmdline
	cmdline=""
	if vsver = "12.0" or vsver = "14.0" Then
		cmdline="Call " & chr(34) & basedir & "\VC\vcvarsall.bat" & chr(34) & " " & compiletarget & " >NUL"
	Else
		if compiletarget = "amd64" Then
			cmdline="call " & chr(34) & basedir & "\VC\Auxiliary\Build\vcvarsall.bat" & chr(34) & " x64" & " >NUL"
		ElseIf compiletarget = "amd64_x86" Then
			cmdline = "call " & chr(34) & basedir & "\VC\Auxiliary\Build\vcvarsall.bat" & chr(34) & " x64_x86" & " >NUL"
		Else
			WScript.Stderr.writeline("not supported compiletarget[" & compiletarget & "]")
		End If
	End If

	GetVsAllBatchCall=cmdline
End Function


Function GetDevenvSlnRun(vsver,slnfile,basedir,target)
	dim cmdline
	cmdline = chr(34) & basedir & "\Common7\IDE\devenv.exe" & chr(34) & " " & chr(34) & slnfile & chr(34) & " /useenv /build "  & chr(34) & target & chr(34) 
	GetDevenvSlnRun=cmdline
End Function
