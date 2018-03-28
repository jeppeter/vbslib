
#include <ParseCmdLine.au3>

AutoItSetOption("MustDeclareVars",1)
AutoItSetOption("ExpandVarStrings",1)

; $args has 5 entries for 5 switches: -nodebug -w -d... -u... -t...
; each entry has 4 parts separated by a | (last part is optional): "type|switch|var|default"
; "type" is either b for a boolean or s for a string switch
; "switch" is what the user has to type on the command line
; "var" is the AU3 variable (w/o the $) which will hold the result (needs not be declared)
; "default" is the default initialiser for "var"
; "default" is optional: if it's mssing a boolean (ie "type" b) comes back false, a string (s) empty
;
; CAUTION: all "switch" strings should be unambiguous; however, a longer "switch" before
; a shorter "switch" in $args should work (ie first "nodebug" and later "n"). The order on
; the command line is irrelevant in any case: -nodebug will never be interpreted as -n"odebug"
; as long as -nodebug is defined before -n
;
Global $args[5]=[ _ ; Watch out: the initialisers for -d and -t require 'AutoItSetOption("ExpandVarStrings",1)')!
    "b|nodebug|DebugMode|True", _   ; boolean -nodebug: $DebugMode is True if switch is not specified
    "b|w|Write|False", _            ; boolean -w: $Write is False if switch is not specified
    "s|d|Dir|@SCRIPTDIR@\", _       ; string -d...: $Dir is Scriptdir if switch is not specified
    "s|u|User|St. Jacques", _       ; string -u...: $User is St. Jacques if switch is not specified
    "s|t|Time|@HOUR@:@MIN@:@SEC@" _ ; string -t...: $Time is current time if switch is not specified
]

; Another example:
; the "switch" part can also include a separator or other special characters:
; the following means -nodebug -w -dir=... -name=... -n
; using -nodebug, -name and -n in one command line is probably not a good idea (see comment above)
;Global $args[5]=["b|nodebug|DebugMode|True","b|w|Write","s|dir=|Dir|c:\test","s|name=|Name|Holy Cow!","b|n|NoFiles"]

; as $CmdLine[] is a read-only array (strange idea, that... and something
; the AU3 documentation should mention in passing), parsing has to be done with a copy
Global $cl=$CmdLine
If _ParseCmdLine($cl,$args)=0 Then
    If @Error=2 Then
        ConsoleWrite("error in $args[] while parsing: "&$args[@extended]&@CRLF)
    ElseIf @Error=3 Then
        ConsoleWrite("error in $cl[] while parsing: "&$cl[@extended]&@CRLF)
    EndIf
    Exit
EndIf

; so let's see the results after parsing:
WriteVar("$Dir")
WriteVar("$User")
WriteVar("$Time")
If $DebugMode Then ConsoleWrite("Debugging on"&@CRLF)
If $Write Then ConsoleWrite("Writing enabled"&@CRLF)
;If $NoFiles Then ConsoleWrite("No files"&@CRLF)

ConsoleWrite($cl[0]&" command line argument(s) after parsing:"&@CRLF)
For $i=1 To $cl[0]
    WriteVar("$cl["&$i&"]")
Next

Func WriteVar($v)
    ConsoleWrite($v&"=!>"&Execute($v)&"<!"&@CRLF)
EndFunc