/*
 
    If you Start\Run `mmsys.cpl` you'll see the `Name` of your devices.
 
    As for the `DeviceNumber`, you need to run the script in the `SoundSet` command.
 
    What I do when I change the sound output is to first set the volume then unmute.
 
    Why?
 
    I have sensible ears and often forget to unmute, so if I were to have the volume
    loud and change the output I get a nasty blast. This way is better.
 
    The hotkey is: if mouse overt the taskbar, Ctrl+Side buttons changes output.
 
*/
 
return ; End of auto-execute
 
#If MouseOver("ahk_class Shell_TrayWnd")
    ^XButton1:: ; Switch audio output
    ^XButton2::SoundDevice(SubStr(A_ThisHotkey, 0))
#If
 
SoundDevice(Output) {
    devices := []
    devices[1] := {Name:"Speakers"  , DeviceNumber:3, Volume:10}
    devices[2] := {Name:"Headphones", DeviceNumber:5, Volume: 5}
    SoundSet % devices[Output, "Volume"],,, % Device.DeviceNumber ; Volume
    SoundSet 0,, MUTE, % devices[Output, "DeviceNumber"]          ; Unmute
    loop 3 ; Change output
    {
        ; VA_SetDefaultEndpoint(devices.Name, A_Index - 1)
    }
    ; OSD(devices[Output, "Name"]) ; Notification
}   
MouseOver(Control)
{
    WinGet, OutputVar [, Cmd, WinTitle, WinText, ExcludeTitle, ExcludeText]
    MouseGetPos,,,, OutputVarControl
    MouseGetPos, , , , CurrControl, 3
    return Instr(CurrControl,Control)
}