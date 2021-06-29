Option Explicit

Sub includeFile(fSpec)
    With CreateObject("Scripting.FileSystemObject")
       executeGlobal .openTextFile(fSpec).readAll()
    End With
End Sub

Function GetScriptDir()
    dim fso ,scriptpath
    Set fso = CreateObject("Scripting.FileSystemObject") 
    GetScriptDir=fso.GetParentFolderName(Wscript.ScriptFullName)
End Function

call includeFile( GetScriptDir() & "\..\base_func.vbs")
call includeFile( GetScriptDir() & "\..\reg_op.vbs")
call includeFile( GetScriptDir() & "\..\vs_find.vbs")



dim vsver
dim nmakeexe
dim basedir

vsver=IsInstallVisualStudio(10.0)
basedir=GetVisualStudioInstdir(10.0)


nmakeexe = GetNmake(basedir,vsver)
Wscript.stderr.writeline("version " & vsver & " basedir" & basedir & " nmakeexe" & nmakeexe)
