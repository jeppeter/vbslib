

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


call includeFile( GetScriptDir() & "\vbsjson.vbs")


'Author: Demon
'Date: 2012/5/3
'Website: http://demon.tw
Dim fso, json, str, o, i
Set json = New VbsJson
Set fso = WScript.CreateObject("Scripting.Filesystemobject")
str = fso.OpenTextFile("json.txt").ReadAll
Set o = json.Decode(str)

if o.Exists("Images") Then
	WScript.Echo "exists Images"
else
	Wscript.Echo "not exists Images"
End If

WScript.Echo o("Image")("Width")
WScript.Echo o("Image")("Height")
WScript.Echo o("Image")("Title")
WScript.Echo o("Image")("Thumbnail")("Url")
For Each i In o("Image")("IDs")
    WScript.Echo i
Next