
#If  (vMoncnt=1)
!#Esc::
:*:s.AS::
#If  (vMoncnt>1)
:*:s.AS::
F24::
f_LoadDevicesOut()
f_LoadDevicesIn()
if WinActive("AS - Audio-Switcher")
    AS_Escape()
Else
    f_ShowGUI()
return
#if WinActive("AS - Audio-Switcher")
WheelDown::Volume_Down
WheelUp::Volume_Up
Left::
Right::
SC029::
f_SwitchSets()
return
!Esc::reload
