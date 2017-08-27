Option Explicit

Function FileExists(infile)
	dim fso

	set fso = WScript.CreateObject("Scripting.FileSystemObject")
	if fso.FileExists(infile) Then
		FileExists=True
	Else
		FileExists=False
	End If
End Function

Function ReadFileAll(infile)
	dim fso
	dim fin
	dim ftext

	set fso = WScript.CreateObject("Scripting.FileSystemObject")
	set fin = fso.OpenTextFile(infile,1)
	ftext = fin.ReadAll()
	fin.Close()
	ReadFileAll=ftext
End Function

Function DeleteFileSafe(infile)
	dim exists
	dim retval
	dim fso
	retval = False
	exists = FileExists(infile)
	if exists Then
		set fso = WScript.CreateObject("Scripting.FileSystemObject")
		fso.DeleteFile(infile)
		retval = True
	End If
	DeleteFileSafe=retval
End Function