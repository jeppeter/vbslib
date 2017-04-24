
Option Explicit

Function CheckInstallBasedir(directory,vsver)
	dim narr,i,j,num,sp,ret,clexe
	If vsver < 15.0 Then
		'  this is version 12.0 or 14.0 for vs2013 or vs2015
		narr = Split(directory,"\")
		if IsArray(narr) Then
		   num=Ubound(narr)
		   i=num
		   j=1
		   sp=narr(0)
		   While j < (i+1)
		        clexe=sp & "\VC\bin\cl.exe"
		        ret=FileExists(clexe)
		        if ret Then
		        	Wscript.Stdout.writeline("check["&vsver&"]="&sp)
		            CheckInstallBasedir=sp
		            Exit Function
		        End If
		        sp = sp & "\" & narr(j)
		        j = j+1
		   Wend
		End If
	Else
	    '  this is version 15.0 for vs2017
	    narr = Split(directory,"\")
	    if IsArray(narr) Then
	        num=Ubound(narr)
	        i=num
	        j=1
	        sp=narr(0)
	        While j < (i+1)
	        	clexe = sp & "\VC\Auxiliary\Build\vcvars32.bat"
	        	ret=FileExists(clexe)
	        	if ret Then
	        	     CheckInstallBasedir=sp
	        	     Exit Function
	        	End If
	        	sp = sp & "\" & narr(j)
	        	j = j + 1
	        Wend
	    End If
	End If
	CheckInstallBasedir=Empty
End Function

Function FindoutInstallBasedir(directory,vsver)    
	dim narr,i,j,num,sp,ret,clexe
	If vsver < 15.0 Then
		'  this is version 12.0 or 14.0 for vs2013 or vs2015
		narr = Split(directory,"\")
		if IsArray(narr) Then
		   num=Ubound(narr)
		   i=num
		   j=1
		   sp=narr(0)
		   While j < (i+1)
		        clexe=sp & "\VC\bin\cl.exe"
		        ret=FileExists(clexe)
		        if ret Then
		            FindoutInstallBasedir=sp
		            Exit Function
		        End If
		        sp = sp & "\" & narr(j)
		        j = j+1
		   Wend
		End If
	Else
	    '  this is version 15.0 for vs2017
	    narr = Split(directory,"\")
	    if IsArray(narr) Then
	        num=Ubound(narr)
	        i=num
	        j=1
	        sp=narr(0)
	        While j < (i+1)
	        	clexe = sp & "\VC\Auxiliary\Build\vcvars32.bat"
	        	ret=FileExists(clexe)
	        	if ret Then
	        	     FindoutInstallBasedir=sp
	        	     Exit Function
	        	End If
	        	sp = sp & "\" & narr(j)
	        	j = j + 1
	        Wend
	    End If
	End If
	FindoutInstallBasedir=""
End Function

Function IsInstallVisualStudio(version)
	dim curval
	dim curver
	dim envvar
	dim envkey
	dim instdir
	dim versions
	dim values 
	dim idx
	dim max
	versions=Array(15.0,14.0,12.0)
	values=Array("150","140","120")
	idx = 0
	max = Ubound(versions) + 1
	Do while idx < max
		curver = versions(idx)
		curval=values(idx)
		Wscript.Stdout.writeline("check["&curver&"]")
		If curver >= version Then
			envkey = "VS" & curval & "COMNTOOLS"
			envvar = GetEnv(envkey)
			If not IsNull(envvar) Then
				instdir=CheckInstallBasedir(envvar,curver)
				if not IsEmpty(instdir) Then
					IsInstallVisualStudio=curver
					Exit Function
				End If
			End If
		End If
		idx = idx + 1
	Loop
	IsInstallVisualStudio=Null
End Function

Function GetVisualStudioInstdir(version)
	dim curval
	dim curver
	dim envvar
	dim envkey
	dim instdir
	dim versions
	dim values 
	dim idx
	dim max
	versions=Array(15.0,14.0,12.0)
	values=Array("150","140","120")
	idx = 0
	max = Ubound(versions) + 1
	Do while idx < max
		curver = versions(idx)
		curval=values(idx)
		If curver >= version Then
			envkey = "VS" & curval & "COMNTOOLS"
			envvar = GetEnv(envkey)
			If not IsNull(envvar) Then
				instdir=CheckInstallBasedir(envvar,curver)
				if not IsEmpty(instdir) Then
					Wscript.Stdout.write("get["&curver&"]="&instdir)
					GetVisualStudioInstdir=instdir
					Exit Function
				End If
			End If
		End If
		idx = idx + 1
	Loop
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

