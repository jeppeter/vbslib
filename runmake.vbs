

Option Explicit

Sub includeFile(fSpec)
    With CreateObject("Scripting.FileSystemObject")
       executeGlobal .openTextFile(fSpec).readAll()
    End With
End Sub

Function GetAppVersion(fname)
	dim objFSO,objReadFile,content
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objReadFile = objFSO.OpenTextFile(fname, 1, False)
	content = objReadFile.ReadAll
	content = Replace(content, vbCr, "")
	content = Replace(content, vbLf, "")	
	GetAppVersion=content
End Function


dim basedir,vsver,vspdir,iidx,jidx

Function GetScriptDir()
	dim fso ,scriptpath
	Set fso = CreateObject("Scripting.FileSystemObject") 
	GetScriptDir=fso.GetParentFolderName(Wscript.ScriptFullName)
End Function


call includeFile( GetScriptDir() & "\reg_op.vbs")
call includeFile( GetScriptDir() & "\vs_find.vbs")
call includeFile( GetScriptDir() & "\base_func.vbs")

Function Usage(ec,fmt)
    dim fh
    set fh = WScript.Stderr
    if ec = 0 Then
        set fh = WScript.Stdout
    End if

    if fmt <> "" Then
        fh.Writeline(fmt)
    End if
    fh.Writeline(WScript.ScriptName & " [OPTIONS] [FILE] [TARGETS]...")
    fh.Writeline(chr(9) &"-h|--help                    to display this information")
    fh.Writeline(chr(9) &"-t|--timestamp ENVVAR        to set timestamp in the environment value")
    fh.Writeline(chr(9) &"-c|--check ENVVAR            to check environment value set")
    fh.Writeline(chr(9) &"-V|--versionvar ENVVAR       to set version variable")
    fh.Writeline(chr(9) &"[FILE]                       file of makefile")
    fh.Writeline(chr(9) &"[TARGETS]...                 target to compile")
    WScript.Quit(ec)
End Function

Function ParseArgs(args)
	dim i
	dim max
	dim argobj
	max = Ubound(args)
	set argobj = new DictObject
	i = 0
	Do While i < max
		if args(i) = "-h" or args(i) = "--help" Then
			Usage 0, ""
		Elseif args(i) = "-t" or args(i) = "--timestamp" Then
			If (i+1) >= max Then
				Usage 3,args(i) & " need arg"
			End If
			argobj.Add "timestamp",args(i+1)
			i = i + 1
		Elseif args(i) = "-c" or args(i) = "--check" Then
			If (i+1) >= max Then
				Usage 3,args(i) & " need arg"
			End If
			argobj.Append "check",args(i+1)
			i = i + 1
		Elseif args(i) = "-V" or args(i) = "--versionvar" Then
			If (i+1) >= max Then
				Usage 3,args(i) & " need arg"
			End If
			argobj.Add "versionvar",args(i+1)
			i = i + 1
		Else
			Exit Do
		End If
		i = i + 1
	Loop

	do While i < max
		argobj.Append "args",args(i)
		i = i + 1
	Loop
	set ParseArgs=argobj
End Function


dim num,i,argobj
num = WScript.Arguments.Count()

if num = 0 Then
    Usage 3,"need args"
End if

redim args(num)

for i=0 to (num - 1)
    args(i) = WScript.Arguments.Item(i)
next

set argobj = ParseArgs(args)


vsver=IsInstallVisualStudio(10.0,"SOFTWARE\Microsoft\VisualStudio")
if IsEmpty(vsver) Then
	wscript.stderr.writeline("Please Install visual studio new version than 14.0")
	WScript.Quit(3)
End If

vspdir=ReadReg("HKEY_CURRENT_USER\SOFTWARE\Microsoft\VisualStudio\"& vsver &"_Config\InstallDir")
if IsEmpty(vspdir) Then
	wscript.stderr.writeline("can not find visual studio install directory")
	wscript.quit(4)
End If

basedir=FindoutInstallBasedir(vspdir,vsver)
if basedir = "" Then
	wscript.stderr.writeline("can not find visual studio install directory on " & vspdir)
	wscript.quit(5)
End If

dim nmakeexe,cmd,makefile,makedep
dim dt ,timestamp,version


nmakeexe=basedir+"\VC\bin\nmake.exe"
wscript.echo ("basedir (" & basedir & ") nmake (" & nmakeexe & ")")


If argobj.Exists("check") Then
	dim arrobj
	set arrobj = argobj.Value("check")
	i = 0
	num = arrobj.Size()
	Do While i < num
		call CheckVariable(arrobj.GetItem(i))
		i = i + 1
	Loop
End If

If argobj.Exists("timestamp") Then
	dt=now
	timestamp = year(dt)
	if len(month(dt)) < 2  Then
		timestamp = timestamp & "0" & month(dt)
	else
		timestamp = timestamp & month(dt)
	End IF

	if len(day(dt)) < 2 Then
		timestamp = timestamp & "0" & day(dt)
	else
		timestamp = timestamp & day(dt)
	End If

	call SetEnv(argobj.Value("timestamp"),timestamp)
End If

if argobj.Exists("versionvar") Then
	version=GetAppVersion("VERSION")
	call SetEnv(argobj.Value("versionvar"),version)
End If

if argobj.Exists("args") Then
	set arrobj = argobj.Value("args")
	makefile=arrobj.GetItem(0)
	if arrobj.Size() >= 2 Then
		i = 1
		Do While i < arrobj.Size()
			makedep=arrobj.GetItem(i)
			cmd =  chr(34) & nmakeexe & chr(34) & " /f " & chr(34) & makefile  & chr(34) & " " & chr(34) & makedep & chr(34)
			RunCommand(cmd)
			i = i + 1
		Loop
	Else
		cmd= chr(34) & nmakeexe & chr(34) & " /f " & chr(34) & makefile  & chr(34)
		RunCommand(cmd)
	End If
Else
	cmd=chr(34) & nmakeexe & chr(34) 
	RunCommand(cmd)
End If