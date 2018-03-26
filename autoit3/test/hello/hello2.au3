#include <MsgBoxConstants.au3>

MsgBox($MB_SYSTEMMODAL, "My second script!", "Hello from the main script!")

Example_Func()

Func Example_Func()
    Return MsgBox($MB_SYSTEMMODAL, "My Second Script!", "Hello from the functions!")
EndFunc   ;==>Example_Func 