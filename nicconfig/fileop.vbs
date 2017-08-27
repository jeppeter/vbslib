Option Explicit


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