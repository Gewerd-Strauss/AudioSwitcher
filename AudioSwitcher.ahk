#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance, Force
SysGet, vMoncnt, MonitorCount
Menu, Tray, Icon, C:\windows\system32\shell32.dll,138 ;Set custom Script icon
#Include <scriptObj/scriptObj>
CreditsRaw=
(LTRIM
author1   -		 snippetName1		   		  			-	URL1
Flipeador   -		 original Script		   		  			-	https://www.autohotkey.com/boards/viewtopic.php?p=221751#p221751
Gewerd Strauss		- snippetName2|SnippetName3 (both at the same URL)								-	/
XMCQCX      - DeviceIDPnP   - https://www.autohotkey.com/boards/viewtopic.php?f=6&t=108930
)
FileGetTime, ModDate,%A_ScriptFullPath%,M
FileGetTime, CrtDate,%A_ScriptFullPath%,C
CrtDate:=SubStr(CrtDate,7,  2) "." SubStr(CrtDate,5,2) "." SubStr(CrtDate,1,4)
ModDate:=SubStr(ModDate,7,  2) "." SubStr(ModDate,5,2) "." SubStr(ModDate,1,4)
global script := {   base         : script
                    ,name         : regexreplace(A_ScriptName, "\.\w+")
                    ,version      : FileOpen(A_ScriptDir "\version.ini","r").Read()
                    ,author       : "Gewerd Strauss"
					,authorID	  : "Laptop-C"
					,authorlink   : ""
                    ,email        : ""
                    ,credits      : CreditsRaw
					,creditslink  : ""
                    ,crtdate      : CrtDate
                    ,moddate      : ModDate
                    ,homepagetext : ""
                    ,homepagelink : ""
                    ,ghtext 	  : "GH-Repo"
                    ,ghlink       : "https://github.com/Gewerd-Strauss/AudioSwitcher"
                    ,doctext	  : ""
                    ,doclink	  : ""
                    ,forumtext	  : ""
                    ,forumlink	  : ""
                    ,donateLink	  : ""
                    ,resfolder    : A_ScriptDir "\res"
                    ,iconfile	  : ""
                    ,reqInternet   : false
					,rfile  	  : "https://github.com/Gewerd-Strauss/AudioSwitcher/archive/refs/heads/MAIN.zip"
					,vfile_raw	  : "https://raw.githubusercontent.com/Gewerd-Strauss/AudioSwitcher/main/version.ini" 
					,vfile 		  : "https://raw.githubusercontent.com/Gewerd-Strauss/AudioSwitcher/main/version.ini" 
					,vfile_local  : A_ScriptDir "\version.ini" 
                    ,config:		[]
					,configfile   : A_ScriptDir "\INI-Files\" regexreplace(A_ScriptName, "\.\w+") ".ini"
                    ,configfolder : A_ScriptDir "\INI-Files"}

global bStartOnOutPut:=true
script.Update(,,1) ;DO NOT ACTIVATE THISLINE UNTIL YOU DUMBO HAS FIXED THE DAMN METHOD. God damn it.
f_CreateTrayMenu()
 


 oMyDevices := {}
f_LoadDevicesOut()
f_LoadDevicesIn()
f_CreateGUI()
oMyDevices.Push({"DeviceName":"Kopfhörer (WH-1000XM3)", "DeviceID":"SWD\MMDEVAPI\{0.0.0.00000000}.{2DA0C039-7454-45FD-BFCA-4656F85C1384}"})
 
 
 
 oDevicesConnected := {}
 For Device in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_PnPEntity")
     oDevicesConnected.Push({"DeviceName":Device.Name, "DeviceID":Device.DeviceID, "DevicePNPClass":Device.PNPClass, "DeviceStatus":Device.Status})
 
 ;=============================================================================================
 
 ; Establish the status of the devices in oMyDevices
 For Index, MyDevice in oMyDevices
 {
     DeviceFound := ""
     For Index, DeviceConnected in oDevicesConnected
     {
         If (MyDevice.DeviceID = DeviceConnected.DeviceID)
         {
             If (DeviceConnected.DeviceStatus = "OK"), DeviceFound := "Yes"
                 MyDevice.DeviceStatus := "Connected"
 
             If (DeviceConnected.DeviceStatus = "Unknown"), DeviceFound := "Yes"
                 MyDevice.DeviceStatus := "Disconnected"
         }
     }
     If !DeviceFound
         MyDevice.DeviceStatus := "Disconnected"
 }
 
 ;=============================================================================================
 
 ; Run or close scripts/programs if the devices are connected/disconnected when the script start.
 Loop % oMyDevices.Count()
 {
     DeviceStatusAtStartup := oMyDevices[A_Index].DeviceName A_Space oMyDevices[A_Index].DeviceStatus
     DevicesActions(DeviceStatusAtStartup)
     DeviceStatusAtStartup := StrReplace(DeviceStatusAtStartup, "Disconnected", "Not connected")
     strTooltip .= DeviceStatusAtStartup "`n"
 }
;  If strTooltip
;      strTooltip := RTrim(strTooltip, "`n")
        ;  Tooltip, % strTooltip, 0, 0
        ;      SetTimer, RemoveToolTipDeviceStatus, -6000
 
 ;=============================================================================================
 
 OnMessage(0x219, "WM_DEVICECHANGE") 
 WM_DEVICECHANGE(wParam, lParam, msg, hwnd)
 {
     SetTimer, CheckDevicesStatus , -1250
 }

 Return





 DevicesActions(ThisDeviceStatusHasChanged) {
 static count:=0
   
    If (ThisDeviceStatusHasChanged = "Kopfhörer (WH-1000XM3) Connected") && count>0
        reload
    If (ThisDeviceStatusHasChanged = "Kopfhörer (WH-1000XM3) Disconnected") && count>0
        reload
    count++
}
 ;=============================================================================================
 
 CheckDevicesStatus:
 
     ;=============================================================================================
     
     ; Check devices connected
     oDevicesConnected.Delete(1, oDevicesConnected.Length())
     For Device in ComObjGet("winmgmts:").ExecQuery("Select * from Win32_PnPEntity")
         oDevicesConnected.Push({"DeviceName":Device.Name, "DeviceID":Device.DeviceID, "DevicePNPClass":Device.PNPClass, "DeviceStatus":Device.Status})
 
     ;=============================================================================================
 
     ; Find which devices status has changed in oMyDevices
     oMyDevicesStatusHasChanged := []
     For Index, MyDevice in oMyDevices
     {
         DeviceFound := ""
         For Index, DeviceConnected in oDevicesConnected
         {
             If (MyDevice.DeviceID = DeviceConnected.DeviceID)
                 If (DeviceConnected.DeviceStatus = "OK")
                     If (MyDevice.DeviceStatus = "Disconnected"), MyDevice.DeviceStatus := "Connected", DeviceFound := "Yes"
                             oMyDevicesStatusHasChanged.Push(MyDevice.DeviceName " Connected")
                 
                 If (DeviceConnected.DeviceStatus = "Unknown")
                     If (MyDevice.DeviceStatus = "Connected"), MyDevice.DeviceStatus := "Disconnected", DeviceFound := "Yes"
                             oMyDevicesStatusHasChanged.Push(MyDevice.DeviceName " Disconnected")
         }
         If !DeviceFound
             If (MyDevice.DeviceStatus = "Connected"), MyDevice.DeviceStatus := "Disconnected"
                     oMyDevicesStatusHasChanged.Push(MyDevice.DeviceName " Disconnected")
     }
 
     ;=============================================================================================
 
     ; If devices in oMyDevices status has changed go to DevicesActions()
     If (oMyDevicesStatusHasChanged)
     {
         strTooltip := ""
         Loop % oMyDevicesStatusHasChanged.Count()
         {
             DevicesActions(oMyDevicesStatusHasChanged[1])
             strTooltip .= oMyDevicesStatusHasChanged[1] "`n"
             oMyDevicesStatusHasChanged.RemoveAt(1)
         }
         If strTooltip
             strTooltip := RTrim(strTooltip, "`n")
                ;  Tooltip, % strTooltip, 0, 0
                    ;  SetTimer, RemoveToolTipDeviceStatus, -6000
     }
     
     ;=============================================================================================
 
 return
 
;=============================================================================================

RemoveToolTipDeviceStatus:
ToolTip
return

f_CreateGUI()
{ ;; create the GUI
    global
    Gui, AS: new, +AlwaysOnTop -SysMenu -ToolWindow -caption +Border +LabelAS_ ;+Resize +MinSize1000x
    gui, AS: default
    gui, add, text,vTitleString,%  "Choose Audio-Out Device"
    gui, font, s8
    Gui, Margin, 16,0
    gui, +hwndASGUI
    ; DevicesIn.Test:="A"
    MaxIn:=DevicesIn.Count()
    MaxOut:=DevicesOut.Count()
    Matches:={}
    Ind:=0

    if ((MaxIn>MaxOut) || (MaxIn==Max))
    {
        bDevicesOutFirst:=!bDevicesInFirst:=1
        for k,v in DevicesIn
        {
            Matches.push(k)
            if !Instr(Matches[Matches.MaxIndex()],"||||")
            {
                MatchMax:=Matches.MaxIndex()
                for s,w in DevicesOut
                {
                    LastAdd:=Matches[Matches.MaxIndex()-1]
                    if Instr(LastAdd,"|||" s)
                        continue
                    if (OldInsert==s)
                        continue
                    if Instr(History,s "||||" Insert)
                        continue
                    Ind++
                    if (Ind==MatchMax)
                    {
                        Insert:=s
                        History.="||||" Insert
                        Matches[Matches.MaxIndex()]:=Matches[Matches.MaxIndex()] "||||" Insert
                        OldInsert:=Insert
                        break
                    }
                }
                Insert:=DevicesOut[A_Index] ;; I have NO clue why this is required, but the script breaks without it. No touch this
            }
        }
    }
    Else
    {   ;; MaxOut larger MaxIn
        bDevicesInFirst:=!bDevicesOutFirst:=1
        for k,v in DevicesOut
        {
            Matches.push(k)
            if !Instr(Matches[Matches.MaxIndex()],"||||")
            {
                MatchMax:=Matches.MaxIndex()
                for s,w in DevicesIn
                {
                    LastAdd:=Matches[Matches.MaxIndex()-1]
                    if Instr(LastAdd,"|||" s)
                        continue
                    if (OldInsert==s)
                        continue
                    if Instr(History,s "||||" Insert)
                        continue
                    Ind++
                    if (Ind==MatchMax)
                    {
                        Insert:=s
                        History.="||||" Insert
                        Matches[Matches.MaxIndex()]:=Matches[Matches.MaxIndex()] "||||" Insert
                        OldInsert:=Insert
                        break
                    }
                }
                Insert:=DevicesOut[A_Index] ;; I have NO clue why this is required, but the script breaks without it. No touch this
            }
        }
    }
    ; Clipboard:=Obj2Str(Matches)
    ActiveDevices:=strsplit(getInfo(),"`n")
    for k,v in Matches
    {
        C:=strsplit(v,"||||")
        if bDevicesInFirst
        {
            ButtonFaceIn:=C.1
            ButtonFaceOut:=C.2
            gui, add, button,Default r2 w200 gChooseAudioIn hidden disabled vValIn%A_Index%  ,% "&" A_Index " " trim(C.1)
            gui, add, button,yp Default r2 w200  gChooseAudioOut vVal%A_Index%,% "&" A_Index " " trim(C.2)
            if (trim(C.2)==ActiveDevices[2])
                guicontrol, disable, % "ValIn" A_Index
            if (trim(C.1)==ActiveDevices[1])
                guicontrol, disable, % "Val" A_Index
        }
        else
        { ;; audioOUT First
            ButtonFaceIn:=C.1
            ButtonFaceOut:=C.2
            ; if (c.MaxIndex()!=2) ;; removed until solution found
            gui, add, button,Default r2 w200 gChooseAudioIn hidden disabled vValIn%A_Index%  ,% "&" A_Index " " trim(C.2)
            gui, add, button,yp Default r2 w200  gChooseAudioOut vVal%A_Index%,% "&" A_Index " " trim(C.1)
            if (trim(C.2)==ActiveDevices[2]) 
                guicontrol, disable, % "ValIn" A_Index
            if (trim(C.1)==ActiveDevices[1])
                guicontrol, disable, % "Val" A_Index
            if (c.MaxIndex()!=2) ;; append credits-string to the end of the 
            gui, add, button, yp+35  h20 glCredits vCredits2, Credits 

        }

    }

    return
}
lCredits:
script.About()
return
AS_Escape()
{ ;; close the GUI
    gui, AS: hide
}
f_ShowGUI()
{ ;; show the GUI
    GUI, AS: show, AutoSize , AS - Audio-Switcher
    sleep, 300 ;; make sure accidental spacebar-inputs or other don't immediately trigger a button.
}
ChooseAudioOut() 
{ ;; chooses a value to continue with.
    global
    gui, Submit, 
    GuiControlGet, ChosenDevice,, % A_GuiControl
    ChosenDevice2:=regexreplace(ChosenDevice,"\&\d*\s*")

    for k,v in DevicesOut
    {
        ; Clipboard:=Obj2Str(DevicesOut)
            k:=(bDevicesOutFirst?StrSplit(k, "||||" ).1:StrSplit(k, "||||" ).2)
        A:=strsplit(k," - ")
        Value:="&" A_Index " " trim(A[A.MaxIndex()])
        guicontrol, enable, % "Val" A_Index
        k2:=strreplace(k," - "," ")
        ChosenDevice:=strreplace(ChosenDevice," - "," ")
        if Instr(ChosenDevice,k2)
            Submit:=[ChosenDevice,k,v]
    }
    guicontrol, disable, % ChosenValueOut:=A_GuiControl
    f_SelectAudio(Submit[3])
    return
} ; ChosenValueOut ChosenValueIn

ChooseAudioIn() 
{ ;; chooses a value to continue with.
    global
    gui, Submit, 
    GuiControlGet, ChosenDevice,, % A_GuiControl
    ChosenDevice2:=regexreplace(ChosenDevice,"\&\d*\s*")
    for k,v in DevicesIn
    {
        A:=strsplit(k," - ")
        Value:="&" A_Index " " trim(A[A.MaxIndex()])
        guicontrol, enable, % "ValIn" A_Index
        k2:=strreplace(k," - "," ")
        if Instr(ChosenDevice,k2)
            Submit:=[ChosenDevice,k,v]
    }
    guicontrol, disable, % ChosenValueIn:=A_GuiControl
    f_SelectAudio(Submit[3])
    return
}
f_LoadDevicesOut()
{ ;; retrieve all output audio DevicesOut
    global
    DevicesOut := {}
    IMMDeviceEnumerator := ComObjCreate("{BCDE0395-E52F-467C-8E3D-C4579291692E}", "{A95664D2-9614-4F35-A746-DE8DB63617E6}")

    ; IMMDeviceEnumerator::EnumAudioEndpoints
    ; eRender = 0, eCapture, eAll
    ; 0x1 = DEVICE_STATE_ACTIVE
    DllCall(NumGet(NumGet(IMMDeviceEnumerator+0)+3*A_PtrSize), "UPtr", IMMDeviceEnumerator, "UInt", 0, "UInt", 0x1, "UPtrP", IMMDeviceCollection, "UInt")   ;; output DevicesOut
    ; DllCall(NumGet(NumGet(IMMDeviceEnumerator+0)+3*A_PtrSize), "UPtr", IMMDeviceEnumerator, "UInt", 1, "UInt", 0x1, "UPtrP", IMMDeviceCollection, "UInt") ;; input DevicesOut
    ObjRelease(IMMDeviceEnumerator)

    ; IMMDeviceCollection::GetCount
    DllCall(NumGet(NumGet(IMMDeviceCollection+0)+3*A_PtrSize), "UPtr", IMMDeviceCollection, "UIntP", Count, "UInt")
    Loop % (Count)
    {
        ; IMMDeviceCollection::Item
        DllCall(NumGet(NumGet(IMMDeviceCollection+0)+4*A_PtrSize), "UPtr", IMMDeviceCollection, "UInt", A_Index-1, "UPtrP", IMMDevice, "UInt")

        ; IMMDevice::GetId
        DllCall(NumGet(NumGet(IMMDevice+0)+5*A_PtrSize), "UPtr", IMMDevice, "UPtrP", pBuffer, "UInt")
        DeviceID := StrGet(pBuffer, "UTF-16"), DllCall("Ole32.dll\CoTaskMemFree", "UPtr", pBuffer)

        ; IMMDevice::OpenPropertyStore
        ; 0x0 = STGM_READ
        DllCall(NumGet(NumGet(IMMDevice+0)+4*A_PtrSize), "UPtr", IMMDevice, "UInt", 0x0, "UPtrP", IPropertyStore, "UInt")
        ObjRelease(IMMDevice)

        ; IPropertyStore::GetValue
        VarSetCapacity(PROPVARIANT, A_PtrSize == 4 ? 16 : 24)
        VarSetCapacity(PROPERTYKEY, 20)
        DllCall("Ole32.dll\CLSIDFromString", "Str", "{A45C254E-DF1C-4EFD-8020-67D146A850E0}", "UPtr", &PROPERTYKEY)
        NumPut(14, &PROPERTYKEY + 16, "UInt")
        DllCall(NumGet(NumGet(IPropertyStore+0)+5*A_PtrSize), "UPtr", IPropertyStore, "UPtr", &PROPERTYKEY, "UPtr", &PROPVARIANT, "UInt")
        DeviceName := StrGet(NumGet(&PROPVARIANT + 8), "UTF-16")    ; LPWSTR PROPVARIANT.pwszVal
        DllCall("Ole32.dll\CoTaskMemFree", "UPtr", NumGet(&PROPVARIANT + 8))    ; LPWSTR PROPVARIANT.pwszVal
        ObjRelease(IPropertyStore)
        
        ObjRawSet(DevicesOut, DeviceName, DeviceID)
    }
    DevicesOut2:={}
    Map:={}
    for DeviceName, DeviceID in DevicesOut
    {
        ObjRawSet(Map,DeviceName,A_Index)
        ObjRawSet(DevicesOut2,A_Index,DeviceID)
    }
    ObjRelease(IMMDeviceCollection)
    return
}
f_LoadDevicesIn()
{ ;; retrieve all output audio DevicesIn
    global
    DevicesIn := {}
    IMMDeviceEnumerator := ComObjCreate("{BCDE0395-E52F-467C-8E3D-C4579291692E}", "{A95664D2-9614-4F35-A746-DE8DB63617E6}")

    ; IMMDeviceEnumerator::EnumAudioEndpoints
    ; eRender = 0, eCapture, eAll
    ; 0x1 = DEVICE_STATE_ACTIVE
    ; DllCall(NumGet(NumGet(IMMDeviceEnumerator+0)+3*A_PtrSize), "UPtr", IMMDeviceEnumerator, "UInt", 0, "UInt", 0x1, "UPtrP", IMMDeviceCollection, "UInt")   ;; output DevicesOut
    DllCall(NumGet(NumGet(IMMDeviceEnumerator+0)+3*A_PtrSize), "UPtr", IMMDeviceEnumerator, "UInt", 1, "UInt", 0x1, "UPtrP", IMMDeviceCollection, "UInt") ;; input DevicesIn
    ObjRelease(IMMDeviceEnumerator)

    ; IMMDeviceCollection::GetCount
    DllCall(NumGet(NumGet(IMMDeviceCollection+0)+3*A_PtrSize), "UPtr", IMMDeviceCollection, "UIntP", Count, "UInt")
    Loop % (Count)
    {
        ; IMMDeviceCollection::Item
        DllCall(NumGet(NumGet(IMMDeviceCollection+0)+4*A_PtrSize), "UPtr", IMMDeviceCollection, "UInt", A_Index-1, "UPtrP", IMMDevice, "UInt")

        ; IMMDevice::GetId
        DllCall(NumGet(NumGet(IMMDevice+0)+5*A_PtrSize), "UPtr", IMMDevice, "UPtrP", pBuffer, "UInt")
        DeviceID := StrGet(pBuffer, "UTF-16"), DllCall("Ole32.dll\CoTaskMemFree", "UPtr", pBuffer)

        ; IMMDevice::OpenPropertyStore
        ; 0x0 = STGM_READ
        DllCall(NumGet(NumGet(IMMDevice+0)+4*A_PtrSize), "UPtr", IMMDevice, "UInt", 0x0, "UPtrP", IPropertyStore, "UInt")
        ObjRelease(IMMDevice)

        ; IPropertyStore::GetValue
        VarSetCapacity(PROPVARIANT, A_PtrSize == 4 ? 16 : 24)
        VarSetCapacity(PROPERTYKEY, 20)
        DllCall("Ole32.dll\CLSIDFromString", "Str", "{A45C254E-DF1C-4EFD-8020-67D146A850E0}", "UPtr", &PROPERTYKEY)
        NumPut(14, &PROPERTYKEY + 16, "UInt")
        DllCall(NumGet(NumGet(IPropertyStore+0)+5*A_PtrSize), "UPtr", IPropertyStore, "UPtr", &PROPERTYKEY, "UPtr", &PROPVARIANT, "UInt")
        DeviceName := StrGet(NumGet(&PROPVARIANT + 8), "UTF-16")    ; LPWSTR PROPVARIANT.pwszVal
        DllCall("Ole32.dll\CoTaskMemFree", "UPtr", NumGet(&PROPVARIANT + 8))    ; LPWSTR PROPVARIANT.pwszVal
        ObjRelease(IPropertyStore)
        
        ObjRawSet(DevicesIn, DeviceName, DeviceID)
    }
    DevicesIn2:={}
    Map:={}
    for DeviceName, DeviceID in DevicesIn
    {
        ; oMyDevices.Push({"DeviceName":"Kopfhörer (WH-1000XM3)", "DeviceID":"SWD\MMDEVAPI\{0.0.0.00000000}.{2DA0C039-7454-45FD-BFCA-4656F85C1384}"})
        ; if Instr(Devicename,"Kopfhörer")
        ;     oMyDevices.push({"DeviceName":DeviceName,"DeviceID":DeviceID})
        ObjRawSet(Map,DeviceName,A_Index)
        ObjRawSet(DevicesIn2,A_Index,DeviceID)
    }
    ObjRelease(IMMDeviceCollection)
    return
}

f_SwitchSets()
{
    global
    gui, AS: default
    gui, Submit, NoHide
    ActiveDevices:=strsplit(getInfo(),"`n")
    if bStartOnOutPut
    {
        ;; show In
        guicontrol,,TitleString, % "Choose Audio-In Device"
        for k,v in DevicesOut
        {
            A:=strsplit(k," - ")
            ; gui, add, button,Default w200 gChooseAudioIn vValIn%A_Index% disabled hidden ,% "&" A_Index " " trim(A[A.MaxIndex()])

            guicontrol,disable, Val%A_Index% 
            guicontrol,hide, Val%A_Index% 
            MaxValOut:="Val" A_Index
            if (trim(A.1)==ActiveDevices[1]) || (v==ActiveDevices[1])
                guicontrol, disable, % "Val" A_Index
        }
        for k,v in DevicesIn
        {
            A:=strsplit(k," - ")
            ; gui, add, button,Default w200 gChooseAudioIn vValIn%A_Index% disabled hidden ,% "&" A_Index " " trim(A[A.MaxIndex()])
            guicontrol,Enable, ValIn%A_Index% 
            guicontrol,Show, ValIn%A_Index% 
            MaxValIn:="ValIn" A_Index
            if (trim(A.1)==ActiveDevices[2]) || (v==ActiveDevices[2])
                guicontrol, disable, % "ValIn" A_Index

        }
    }
    else
    {
        ;; show Out
        guicontrol,,TitleString, % "Choose Audio-Out Device"
        for k,v in DevicesOut
        {
            A:=strsplit(k," - ")
            ; gui, add, button,Default w200 gChooseAudioIn vValIn%A_Index% disabled hidden ,% "&" A_Index " " trim(A[A.MaxIndex()])
            guicontrol,enable, Val%A_Index% 
            guicontrol,show, Val%A_Index% 
            MaxValOut:="Val" A_Index 
        }
        for k,v in DevicesOut
        {
            if (trim(k)==ActiveDevices[1])
                guicontrol, disable, % "Val" A_Index
        }
        for k,v in DevicesIn
        {
            A:=strsplit(k," - ")
            ; gui, add, button,Default w200 gChooseAudioIn vValIn%A_Index% disabled hidden ,% "&" A_Index " " trim(A[A.MaxIndex()])
            guicontrol,disable, ValIn%A_Index% 
            guicontrol,hide, ValIn%A_Index% 
            MaxValIn:="ValIn" A_Index

        }
        
        guicontrol,hide, % (bStartOnOutPut?"Credits2":"Credits1")
        ControlGetPos,  X, Y, Width, Height,% e:="Val" SubStr(MaxValOut,4,1)-1,% "AS - Audio-Switcher"
        ; ControlGetPos, [ X, Y, Width, Height, Control, WinTitle, WinText, ExcludeTitle
        ; guicontrol, move, % MaxValOut, % xp "y" d:=SubStr(MaxValOut,4,1)*33
        ; guicontrol, move, % Credits1, % xp "y" d:=(SubStr(MaxValOut,4,1)+1)*33^
        ; ChosenValueOut ChosenValueIn
    }
    guicontrol, disable, % ChosenValueIn
    guicontrol, disable, % ChosenValueOut
    bStartOnOutPut:= !bStartOnOutPut


    gui, Submit, NoHide
    f_ShowGUI()
    return
}


f_SelectAudio(Device)
{ ;; Select the audio device
    IPolicyConfig := ComObjCreate("{870af99c-171d-4f9e-af0d-e63df40c2bc9}", "{F8679F50-850A-41CF-9C72-430F290290C8}") ;00000102-0000-0000-C000-000000000046 00000000-0000-0000-C000-000000000046
    R := DllCall(NumGet(NumGet(IPolicyConfig+0)+13*A_PtrSize), "UPtr", IPolicyConfig, "Str", Device, "UInt", 0, "UInt")
    ; Clipboard:=Device
    ObjRelease(IPolicyConfig)
    ; MsgBox % Format("0x{:081X}", R)
}
f_CreateTrayMenu(IniObj:="")
{ ;; facilitates creation of the tray menu
	menu, tray, add,
	Menu, Misc, add, Open Script-folder, lOpenScriptFolder
	menu, Misc, Add, Reload, lReload
	menu, Misc, Add, About, Label_AboutFile
	SplitPath, A_ScriptName,,,, scriptname
	Menu, tray, add, Miscellaneous, :Misc
	menu, tray, add,
	return
}
lOpenScriptFolder:
run, % A_ScriptDir
return
lReload: 
reload
return
Label_AboutFile:
script.about()
return

#If  (vMoncnt=1)
!#Esc::
:*:s.AS::
#If  (vMoncnt>1)
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
Obj2Str(Obj,FullPath:=1,BottomBlank:=0){
	static String,Blank
	if(FullPath=1)
		String:=FullPath:=Blank:=""
	if(IsObject(Obj)){
		for a,b in Obj{
			if(IsObject(b))
				Obj2Str(b,FullPath "." a,BottomBlank)
			else{
				if(BottomBlank=0)
					String.=FullPath "." a " = " b "`n"
				else if(b!="")
					String.=FullPath "." a " = " b "`n"
				else
					Blank.=FullPath "." a " =`n"
			}
	}}
	return String Blank
}



getInfo() {
	device := VA_GetDevice("playback")

	deviceName := VA_GetDeviceName(device)
    device := VA_GetDevice("capture")
	deviceName .= "`n" VA_GetDeviceName(device)
	Menu, Tray, Tip, Audio Source Switcher - %deviceName%
    return deviceName
}



; --uID:3950707558
 ; Metadata:
  ; Snippet: VA  ;  (v.2.3)
  ; --------------------------------------------------------------
  ; Author: lexikos
  ; License: paste
  ; LicenseURL:  https://raw.githubusercontent.com/ahkscript/VistaAudio/master/LICENSE
  ; Source: https://github.com/ahkscript/VistaAudio
  ; 
  ; --------------------------------------------------------------
  ; Library: Libs
  ; Section: 24 - System/User/hardware
  ; Dependencies: /
  ; AHK_Version: /
  ; --------------------------------------------------------------
  ; Keywords: audio

 ;; Description:
  ;; Note: SoundSet and SoundGet on AutoHotkey v1.1.10 and later support Vista and later natively. You don't need VA.ahk unless you want to use advanced functions not supported by SoundSet/SoundGet.
  ;; 
  ;; VA provides alternatives to some SoundSet/SoundGet subcommands, as well as some additional features that SoundSet/SoundGet do not support. See the online documentation for a list of functions.
  ;; 
  ;; Note: This library depends entirely upon APIs present only in Windows Vista and later. Scripts using it should NOT be run in XP compatibility mode or on any version of Windows older than Vista.
  ;; 
  ;; Notes for v2.1 and later:
  ;; 
  ;;     Requires AutoHotkey v1.1.
  ;;     COM.ahk is NOT required. COM_Init() does NOT need to be called.
  ;; 
  ;; Notes for v2.0 (do not get this version if you have AutoHotkey v1.1):
  ;; 
  ;;     Requires Sean's Standard Library COM.ahk
  ;;     COM must be initialized prior to calling any VA functions: COM_Init().
  ;; 
  ;; Device Topology / Subunits
  ;; 
  ;; Subunit/component names are defined by the audio drivers, so will vary from PC to PC. Volume subunits often have the same names as shown on the Levels tab in the Properties of the sound device. Mute subunits might not have unique names; in those cases, use a numeric index instead of a name. 

 ; VA v2.3
 
 ;
 ; MASTER CONTROLS
 ;
 
 VA_GetMasterVolume(channel="", device_desc="playback")
 {
     if ! aev := VA_GetAudioEndpointVolume(device_desc)
         return
     if channel =
         VA_IAudioEndpointVolume_GetMasterVolumeLevelScalar(aev, vol)
     else
         VA_IAudioEndpointVolume_GetChannelVolumeLevelScalar(aev, channel-1, vol)
     ObjRelease(aev)
     return Round(vol*100,3)
 }
 
 VA_SetMasterVolume(vol, channel="", device_desc="playback")
 {
     vol := vol>100 ? 100 : vol<0 ? 0 : vol
     if ! aev := VA_GetAudioEndpointVolume(device_desc)
         return
     if channel =
         VA_IAudioEndpointVolume_SetMasterVolumeLevelScalar(aev, vol/100)
     else
         VA_IAudioEndpointVolume_SetChannelVolumeLevelScalar(aev, channel-1, vol/100)
     ObjRelease(aev)
 }
 
 VA_GetMasterChannelCount(device_desc="playback")
 {
     if ! aev := VA_GetAudioEndpointVolume(device_desc)
         return
     VA_IAudioEndpointVolume_GetChannelCount(aev, count)
     ObjRelease(aev)
     return count
 }
 
 VA_SetMasterMute(mute, device_desc="playback")
 {
     if ! aev := VA_GetAudioEndpointVolume(device_desc)
         return
     VA_IAudioEndpointVolume_SetMute(aev, mute)
     ObjRelease(aev)
 }
 
 VA_GetMasterMute(device_desc="playback")
 {
     if ! aev := VA_GetAudioEndpointVolume(device_desc)
         return
     VA_IAudioEndpointVolume_GetMute(aev, mute)
     ObjRelease(aev)
     return mute
 }
 
 ;
 ; SUBUNIT CONTROLS
 ;
 
 VA_GetVolume(subunit_desc="1", channel="", device_desc="playback")
 {
     if ! avl := VA_GetDeviceSubunit(device_desc, subunit_desc, "{7FB7B48F-531D-44A2-BCB3-5AD5A134B3DC}")
         return
     VA_IPerChannelDbLevel_GetChannelCount(avl, channel_count)
     if channel =
     {
         vol = 0
         
         Loop, %channel_count%
         {
             VA_IPerChannelDbLevel_GetLevelRange(avl, A_Index-1, min_dB, max_dB, step_dB)
             VA_IPerChannelDbLevel_GetLevel(avl, A_Index-1, this_vol)
             this_vol := VA_dB2Scalar(this_vol, min_dB, max_dB)
             
             ; "Speakers Properties" reports the highest channel as the volume.
             if (this_vol > vol)
                 vol := this_vol
         }
     }
     else if channel between 1 and channel_count
     {
         channel -= 1
         VA_IPerChannelDbLevel_GetLevelRange(avl, channel, min_dB, max_dB, step_dB)
         VA_IPerChannelDbLevel_GetLevel(avl, channel, vol)
         vol := VA_dB2Scalar(vol, min_dB, max_dB)
     }
     ObjRelease(avl)
     return vol
 }
 
 VA_SetVolume(vol, subunit_desc="1", channel="", device_desc="playback")
 {
     if ! avl := VA_GetDeviceSubunit(device_desc, subunit_desc, "{7FB7B48F-531D-44A2-BCB3-5AD5A134B3DC}")
         return
     
     vol := vol<0 ? 0 : vol>100 ? 100 : vol
     
     VA_IPerChannelDbLevel_GetChannelCount(avl, channel_count)
     
     if channel =
     {
         ; Simple method -- resets balance to "center":
         ;VA_IPerChannelDbLevel_SetLevelUniform(avl, vol)
         
         vol_max = 0
         
         Loop, %channel_count%
         {
             VA_IPerChannelDbLevel_GetLevelRange(avl, A_Index-1, min_dB, max_dB, step_dB)
             VA_IPerChannelDbLevel_GetLevel(avl, A_Index-1, this_vol)
             this_vol := VA_dB2Scalar(this_vol, min_dB, max_dB)
             
             channel%A_Index%vol := this_vol
             channel%A_Index%min := min_dB
             channel%A_Index%max := max_dB
             
             ; Scale all channels relative to the loudest channel.
             ; (This is how Vista's "Speakers Properties" dialog seems to work.)
             if (this_vol > vol_max)
                 vol_max := this_vol
         }
         
         Loop, %channel_count%
         {
             this_vol := vol_max ? channel%A_Index%vol / vol_max * vol : vol
             this_vol := VA_Scalar2dB(this_vol/100, channel%A_Index%min, channel%A_Index%max)            
             VA_IPerChannelDbLevel_SetLevel(avl, A_Index-1, this_vol)
         }
     }
     else if channel between 1 and %channel_count%
     {
         channel -= 1
         VA_IPerChannelDbLevel_GetLevelRange(avl, channel, min_dB, max_dB, step_dB)
         VA_IPerChannelDbLevel_SetLevel(avl, channel, VA_Scalar2dB(vol/100, min_dB, max_dB))
     }
     ObjRelease(avl)
 }
 
 VA_GetChannelCount(subunit_desc="1", device_desc="playback")
 {
     if ! avl := VA_GetDeviceSubunit(device_desc, subunit_desc, "{7FB7B48F-531D-44A2-BCB3-5AD5A134B3DC}")
         return
     VA_IPerChannelDbLevel_GetChannelCount(avl, channel_count)
     ObjRelease(avl)
     return channel_count
 }
 
 VA_SetMute(mute, subunit_desc="1", device_desc="playback")
 {
     if ! amute := VA_GetDeviceSubunit(device_desc, subunit_desc, "{DF45AEEA-B74A-4B6B-AFAD-2366B6AA012E}")
         return
     VA_IAudioMute_SetMute(amute, mute)
     ObjRelease(amute)
 }
 
 VA_GetMute(subunit_desc="1", device_desc="playback")
 {
     if ! amute := VA_GetDeviceSubunit(device_desc, subunit_desc, "{DF45AEEA-B74A-4B6B-AFAD-2366B6AA012E}")
         return
     VA_IAudioMute_GetMute(amute, muted)
     ObjRelease(amute)
     return muted
 }
 
 ;
 ; AUDIO METERING
 ;
 
 VA_GetAudioMeter(device_desc="playback")
 {
     if ! device := VA_GetDevice(device_desc)
         return 0
     VA_IMMDevice_Activate(device, "{C02216F6-8C67-4B5B-9D00-D008E73E0064}", 7, 0, audioMeter)
     ObjRelease(device)
     return audioMeter
 }
 
 VA_GetDevicePeriod(device_desc, ByRef default_period, ByRef minimum_period="")
 {
     defaultPeriod := minimumPeriod := 0
     if ! device := VA_GetDevice(device_desc)
         return false
     VA_IMMDevice_Activate(device, "{1CB9AD4C-DBFA-4c32-B178-C2F568A703B2}", 7, 0, audioClient)
     ObjRelease(device)
     ; IAudioClient::GetDevicePeriod
     DllCall(NumGet(NumGet(audioClient+0)+9*A_PtrSize), "ptr",audioClient, "int64*",default_period, "int64*",minimum_period)
     ; Convert 100-nanosecond units to milliseconds.
     default_period /= 10000
     minimum_period /= 10000    
     ObjRelease(audioClient)
     return true
 }
 
 VA_GetAudioEndpointVolume(device_desc="playback")
 {
     if ! device := VA_GetDevice(device_desc)
         return 0
     VA_IMMDevice_Activate(device, "{5CDF2C82-841E-4546-9722-0CF74078229A}", 7, 0, endpointVolume)
     ObjRelease(device)
     return endpointVolume
 }
 
 VA_GetDeviceSubunit(device_desc, subunit_desc, subunit_iid)
 {
     if ! device := VA_GetDevice(device_desc)
         return 0
     subunit := VA_FindSubunit(device, subunit_desc, subunit_iid)
     ObjRelease(device)
     return subunit
 }
 
 VA_FindSubunit(device, target_desc, target_iid)
 {
     if target_desc is integer
         target_index := target_desc
     else
         RegExMatch(target_desc, "(?<_name>.*?)(?::(?<_index>\d+))?$", target)
     ; v2.01: Since target_name is now a regular expression, default to case-insensitive mode if no options are specified.
     if !RegExMatch(target_name,"^[^\(]+\)")
         target_name := "i)" target_name
     r := VA_EnumSubunits(device, "VA_FindSubunitCallback", target_name, target_iid
             , Object(0, target_index ? target_index : 1, 1, 0))
     return r
 }
 
 VA_FindSubunitCallback(part, interface, index)
 {
     index[1] := index[1] + 1 ; current += 1
     if (index[0] == index[1]) ; target == current ?
     {
         ObjAddRef(interface)
         return interface
     }
 }
 
 VA_EnumSubunits(device, callback, target_name="", target_iid="", callback_param="")
 {
     VA_IMMDevice_Activate(device, "{2A07407E-6497-4A18-9787-32F79BD0D98F}", 7, 0, deviceTopology)
     VA_IDeviceTopology_GetConnector(deviceTopology, 0, conn)
     ObjRelease(deviceTopology)
     VA_IConnector_GetConnectedTo(conn, conn_to)
     VA_IConnector_GetDataFlow(conn, data_flow)
     ObjRelease(conn)
     if !conn_to
         return ; blank to indicate error
     part := ComObjQuery(conn_to, "{AE2DE0E4-5BCA-4F2D-AA46-5D13F8FDB3A9}") ; IID_IPart
     ObjRelease(conn_to)
     if !part
         return
     r := VA_EnumSubunitsEx(part, data_flow, callback, target_name, target_iid, callback_param)
     ObjRelease(part)
     return r ; value returned by callback, or zero.
 }
 
 VA_EnumSubunitsEx(part, data_flow, callback, target_name="", target_iid="", callback_param="")
 {
     r := 0
     
     VA_IPart_GetPartType(part, type)
    
     if type = 1 ; Subunit
     {
         VA_IPart_GetName(part, name)
         
         ; v2.01: target_name is now a regular expression.
         if RegExMatch(name, target_name)
         {
             if target_iid =
                 r := %callback%(part, 0, callback_param)
             else
                 if VA_IPart_Activate(part, 7, target_iid, interface) = 0
                 {
                     r := %callback%(part, interface, callback_param)
                     ; The callback is responsible for calling ObjAddRef()
                     ; if it intends to keep the interface pointer.
                     ObjRelease(interface)
                 }
 
             if r
                 return r ; early termination
         }
     }
     
     if data_flow = 0
         VA_IPart_EnumPartsIncoming(part, parts)
     else
         VA_IPart_EnumPartsOutgoing(part, parts)
     
     VA_IPartsList_GetCount(parts, count)
     Loop %count%
     {
         VA_IPartsList_GetPart(parts, A_Index-1, subpart)        
         r := VA_EnumSubunitsEx(subpart, data_flow, callback, target_name, target_iid, callback_param)
         ObjRelease(subpart)
         if r
             break ; early termination
     }
     ObjRelease(parts)
     return r ; continue/finished enumeration
 }
 
 ; device_desc = device_id
 ;               | ( friendly_name | 'playback' | 'capture' ) [ ':' index ]
 VA_GetDevice(device_desc="playback")
 {
     static CLSID_MMDeviceEnumerator := "{BCDE0395-E52F-467C-8E3D-C4579291692E}"
         , IID_IMMDeviceEnumerator := "{A95664D2-9614-4F35-A746-DE8DB63617E6}"
     if !(deviceEnumerator := ComObjCreate(CLSID_MMDeviceEnumerator, IID_IMMDeviceEnumerator))
         return 0
     
     device := 0
     
     if VA_IMMDeviceEnumerator_GetDevice(deviceEnumerator, device_desc, device) = 0
         goto VA_GetDevice_Return
     
     if device_desc is integer
     {
         m2 := device_desc
         if m2 >= 4096 ; Probably a device pointer, passed here indirectly via VA_GetAudioMeter or such.
         {
             ObjAddRef(device := m2)
             goto VA_GetDevice_Return
         }
     }
     else
         RegExMatch(device_desc, "(.*?)\s*(?::(\d+))?$", m)
     
     if m1 in playback,p
         m1 := "", flow := 0 ; eRender
     else if m1 in capture,c
         m1 := "", flow := 1 ; eCapture
     else if (m1 . m2) = ""  ; no name or number specified
         m1 := "", flow := 0 ; eRender (default)
     else
         flow := 2 ; eAll
     
     if (m1 . m2) = ""   ; no name or number (maybe "playback" or "capture")
     {
         VA_IMMDeviceEnumerator_GetDefaultAudioEndpoint(deviceEnumerator, flow, 0, device)
         goto VA_GetDevice_Return
     }
 
     VA_IMMDeviceEnumerator_EnumAudioEndpoints(deviceEnumerator, flow, 1, devices)
     
     if m1 =
     {
         VA_IMMDeviceCollection_Item(devices, m2-1, device)
         goto VA_GetDevice_Return
     }
     
     VA_IMMDeviceCollection_GetCount(devices, count)
     index := 0
     Loop % count
         if VA_IMMDeviceCollection_Item(devices, A_Index-1, device) = 0
             if InStr(VA_GetDeviceName(device), m1) && (m2 = "" || ++index = m2)
                 goto VA_GetDevice_Return
             else
                 ObjRelease(device), device:=0
 
 VA_GetDevice_Return:
     ObjRelease(deviceEnumerator)
     if devices
         ObjRelease(devices)
     
     return device ; may be 0
 }
 
 VA_GetDeviceName(device)
 {
     static PKEY_Device_FriendlyName
     if !VarSetCapacity(PKEY_Device_FriendlyName)
         VarSetCapacity(PKEY_Device_FriendlyName, 20)
         ,VA_GUID(PKEY_Device_FriendlyName :="{A45C254E-DF1C-4EFD-8020-67D146A850E0}")
         ,NumPut(14, PKEY_Device_FriendlyName, 16)
     VarSetCapacity(prop, 16)
     VA_IMMDevice_OpenPropertyStore(device, 0, store)
     ; store->GetValue(.., [out] prop)
     DllCall(NumGet(NumGet(store+0)+5*A_PtrSize), "ptr", store, "ptr", &PKEY_Device_FriendlyName, "ptr", &prop)
     ObjRelease(store)
     VA_WStrOut(deviceName := NumGet(prop,8))
     return deviceName
 }
 
 VA_SetDefaultEndpoint(device_desc, role)
 {
     /* Roles:
          eConsole        = 0  ; Default Device
          eMultimedia     = 1
          eCommunications = 2  ; Default Communications Device
     */
     if ! device := VA_GetDevice(device_desc)
         return 0
     if VA_IMMDevice_GetId(device, id) = 0
     {
         cfg := ComObjCreate("{294935CE-F637-4E7C-A41B-AB255460B862}"
                           , "{568b9108-44bf-40b4-9006-86afe5b5a620}")
         hr := VA_xIPolicyConfigVista_SetDefaultEndpoint(cfg, id, role)
         ObjRelease(cfg)
     }
     ObjRelease(device)
     return hr = 0
 }
 
 
 ;
 ; HELPERS
 ;
 
 ; Convert string to binary GUID structure.
 VA_GUID(ByRef guid_out, guid_in="%guid_out%") {
     if (guid_in == "%guid_out%")
         guid_in :=   guid_out
     if  guid_in is integer
         return guid_in
     VarSetCapacity(guid_out, 16, 0)
 	DllCall("ole32\CLSIDFromString", "wstr", guid_in, "ptr", &guid_out)
 	return &guid_out
 }
 
 ; Convert binary GUID structure to string.
 VA_GUIDOut(ByRef guid) {
     VarSetCapacity(buf, 78)
     DllCall("ole32\StringFromGUID2", "ptr", &guid, "ptr", &buf, "int", 39)
     guid := StrGet(&buf, "UTF-16")
 }
 
 ; Convert COM-allocated wide char string pointer to usable string.
 VA_WStrOut(ByRef str) {
     str := StrGet(ptr := str, "UTF-16")
     DllCall("ole32\CoTaskMemFree", "ptr", ptr)  ; FREES THE STRING.
 }
 
 VA_dB2Scalar(dB, min_dB, max_dB) {
     min_s := 10**(min_dB/20), max_s := 10**(max_dB/20)
     return ((10**(dB/20))-min_s)/(max_s-min_s)*100
 }
 
 VA_Scalar2dB(s, min_dB, max_dB) {
     min_s := 10**(min_dB/20), max_s := 10**(max_dB/20)
     return log((max_s-min_s)*s+min_s)*20
 }
 
 
 ;
 ; INTERFACE WRAPPERS
 ;   Reference: Core Audio APIs in Windows Vista -- Programming Reference
 ;       http://msdn2.microsoft.com/en-us/library/ms679156(VS.85).aspx
 ;
 
 ;
 ; IMMDevice : {D666063F-1587-4E43-81F1-B948E807363F}
 ;
 VA_IMMDevice_Activate(this, iid, ClsCtx, ActivationParams, ByRef Interface) {
     return DllCall(NumGet(NumGet(this+0)+3*A_PtrSize), "ptr", this, "ptr", VA_GUID(iid), "uint", ClsCtx, "uint", ActivationParams, "ptr*", Interface)
 }
 VA_IMMDevice_OpenPropertyStore(this, Access, ByRef Properties) {
     return DllCall(NumGet(NumGet(this+0)+4*A_PtrSize), "ptr", this, "uint", Access, "ptr*", Properties)
 }
 VA_IMMDevice_GetId(this, ByRef Id) {
     hr := DllCall(NumGet(NumGet(this+0)+5*A_PtrSize), "ptr", this, "uint*", Id)
     VA_WStrOut(Id)
     return hr
 }
 VA_IMMDevice_GetState(this, ByRef State) {
     return DllCall(NumGet(NumGet(this+0)+6*A_PtrSize), "ptr", this, "uint*", State)
 }
 
 ;
 ; IDeviceTopology : {2A07407E-6497-4A18-9787-32F79BD0D98F}
 ;
 VA_IDeviceTopology_GetConnectorCount(this, ByRef Count) {
     return DllCall(NumGet(NumGet(this+0)+3*A_PtrSize), "ptr", this, "uint*", Count)
 }
 VA_IDeviceTopology_GetConnector(this, Index, ByRef Connector) {
     return DllCall(NumGet(NumGet(this+0)+4*A_PtrSize), "ptr", this, "uint", Index, "ptr*", Connector)
 }
 VA_IDeviceTopology_GetSubunitCount(this, ByRef Count) {
     return DllCall(NumGet(NumGet(this+0)+5*A_PtrSize), "ptr", this, "uint*", Count)
 }
 VA_IDeviceTopology_GetSubunit(this, Index, ByRef Subunit) {
     return DllCall(NumGet(NumGet(this+0)+6*A_PtrSize), "ptr", this, "uint", Index, "ptr*", Subunit)
 }
 VA_IDeviceTopology_GetPartById(this, Id, ByRef Part) {
     return DllCall(NumGet(NumGet(this+0)+7*A_PtrSize), "ptr", this, "uint", Id, "ptr*", Part)
 }
 VA_IDeviceTopology_GetDeviceId(this, ByRef DeviceId) {
     hr := DllCall(NumGet(NumGet(this+0)+8*A_PtrSize), "ptr", this, "uint*", DeviceId)
     VA_WStrOut(DeviceId)
     return hr
 }
 VA_IDeviceTopology_GetSignalPath(this, PartFrom, PartTo, RejectMixedPaths, ByRef Parts) {
     return DllCall(NumGet(NumGet(this+0)+9*A_PtrSize), "ptr", this, "ptr", PartFrom, "ptr", PartTo, "int", RejectMixedPaths, "ptr*", Parts)
 }
 
 ;
 ; IConnector : {9c2c4058-23f5-41de-877a-df3af236a09e}
 ;
 VA_IConnector_GetType(this, ByRef Type) {
     return DllCall(NumGet(NumGet(this+0)+3*A_PtrSize), "ptr", this, "int*", Type)
 }
 VA_IConnector_GetDataFlow(this, ByRef Flow) {
     return DllCall(NumGet(NumGet(this+0)+4*A_PtrSize), "ptr", this, "int*", Flow)
 }
 VA_IConnector_ConnectTo(this, ConnectTo) {
     return DllCall(NumGet(NumGet(this+0)+5*A_PtrSize), "ptr", this, "ptr", ConnectTo)
 }
 VA_IConnector_Disconnect(this) {
     return DllCall(NumGet(NumGet(this+0)+6*A_PtrSize), "ptr", this)
 }
 VA_IConnector_IsConnected(this, ByRef Connected) {
     return DllCall(NumGet(NumGet(this+0)+7*A_PtrSize), "ptr", this, "int*", Connected)
 }
 VA_IConnector_GetConnectedTo(this, ByRef ConTo) {
     return DllCall(NumGet(NumGet(this+0)+8*A_PtrSize), "ptr", this, "ptr*", ConTo)
 }
 VA_IConnector_GetConnectorIdConnectedTo(this, ByRef ConnectorId) {
     hr := DllCall(NumGet(NumGet(this+0)+9*A_PtrSize), "ptr", this, "ptr*", ConnectorId)
     VA_WStrOut(ConnectorId)
     return hr
 }
 VA_IConnector_GetDeviceIdConnectedTo(this, ByRef DeviceId) {
     hr := DllCall(NumGet(NumGet(this+0)+10*A_PtrSize), "ptr", this, "ptr*", DeviceId)
     VA_WStrOut(DeviceId)
     return hr
 }
 
 ;
 ; IPart : {AE2DE0E4-5BCA-4F2D-AA46-5D13F8FDB3A9}
 ;
 VA_IPart_GetName(this, ByRef Name) {
     hr := DllCall(NumGet(NumGet(this+0)+3*A_PtrSize), "ptr", this, "ptr*", Name)
     VA_WStrOut(Name)
     return hr
 }
 VA_IPart_GetLocalId(this, ByRef Id) {
     return DllCall(NumGet(NumGet(this+0)+4*A_PtrSize), "ptr", this, "uint*", Id)
 }
 VA_IPart_GetGlobalId(this, ByRef GlobalId) {
     hr := DllCall(NumGet(NumGet(this+0)+5*A_PtrSize), "ptr", this, "ptr*", GlobalId)
     VA_WStrOut(GlobalId)
     return hr
 }
 VA_IPart_GetPartType(this, ByRef PartType) {
     return DllCall(NumGet(NumGet(this+0)+6*A_PtrSize), "ptr", this, "int*", PartType)
 }
 VA_IPart_GetSubType(this, ByRef SubType) {
     VarSetCapacity(SubType,16,0)
     hr := DllCall(NumGet(NumGet(this+0)+7*A_PtrSize), "ptr", this, "ptr", &SubType)
     VA_GUIDOut(SubType)
     return hr
 }
 VA_IPart_GetControlInterfaceCount(this, ByRef Count) {
     return DllCall(NumGet(NumGet(this+0)+8*A_PtrSize), "ptr", this, "uint*", Count)
 }
 VA_IPart_GetControlInterface(this, Index, ByRef InterfaceDesc) {
     return DllCall(NumGet(NumGet(this+0)+9*A_PtrSize), "ptr", this, "uint", Index, "ptr*", InterfaceDesc)
 }
 VA_IPart_EnumPartsIncoming(this, ByRef Parts) {
     return DllCall(NumGet(NumGet(this+0)+10*A_PtrSize), "ptr", this, "ptr*", Parts)
 }
 VA_IPart_EnumPartsOutgoing(this, ByRef Parts) {
     return DllCall(NumGet(NumGet(this+0)+11*A_PtrSize), "ptr", this, "ptr*", Parts)
 }
 VA_IPart_GetTopologyObject(this, ByRef Topology) {
     return DllCall(NumGet(NumGet(this+0)+12*A_PtrSize), "ptr", this, "ptr*", Topology)
 }
 VA_IPart_Activate(this, ClsContext, iid, ByRef Object) {
     return DllCall(NumGet(NumGet(this+0)+13*A_PtrSize), "ptr", this, "uint", ClsContext, "ptr", VA_GUID(iid), "ptr*", Object)
 }
 VA_IPart_RegisterControlChangeCallback(this, iid, Notify) {
     return DllCall(NumGet(NumGet(this+0)+14*A_PtrSize), "ptr", this, "ptr", VA_GUID(iid), "ptr", Notify)
 }
 VA_IPart_UnregisterControlChangeCallback(this, Notify) {
     return DllCall(NumGet(NumGet(this+0)+15*A_PtrSize), "ptr", this, "ptr", Notify)
 }
 
 ;
 ; IPartsList : {6DAA848C-5EB0-45CC-AEA5-998A2CDA1FFB}
 ;
 VA_IPartsList_GetCount(this, ByRef Count) {
     return DllCall(NumGet(NumGet(this+0)+3*A_PtrSize), "ptr", this, "uint*", Count)
 }
 VA_IPartsList_GetPart(this, INdex, ByRef Part) {
     return DllCall(NumGet(NumGet(this+0)+4*A_PtrSize), "ptr", this, "uint", Index, "ptr*", Part)
 }
 
 ;
 ; IAudioEndpointVolume : {5CDF2C82-841E-4546-9722-0CF74078229A}
 ;
 VA_IAudioEndpointVolume_RegisterControlChangeNotify(this, Notify) {
     return DllCall(NumGet(NumGet(this+0)+3*A_PtrSize), "ptr", this, "ptr", Notify)
 }
 VA_IAudioEndpointVolume_UnregisterControlChangeNotify(this, Notify) {
     return DllCall(NumGet(NumGet(this+0)+4*A_PtrSize), "ptr", this, "ptr", Notify)
 }
 VA_IAudioEndpointVolume_GetChannelCount(this, ByRef ChannelCount) {
     return DllCall(NumGet(NumGet(this+0)+5*A_PtrSize), "ptr", this, "uint*", ChannelCount)
 }
 VA_IAudioEndpointVolume_SetMasterVolumeLevel(this, LevelDB, GuidEventContext="") {
     return DllCall(NumGet(NumGet(this+0)+6*A_PtrSize), "ptr", this, "float", LevelDB, "ptr", VA_GUID(GuidEventContext))
 }
 VA_IAudioEndpointVolume_SetMasterVolumeLevelScalar(this, Level, GuidEventContext="") {
     return DllCall(NumGet(NumGet(this+0)+7*A_PtrSize), "ptr", this, "float", Level, "ptr", VA_GUID(GuidEventContext))
 }
 VA_IAudioEndpointVolume_GetMasterVolumeLevel(this, ByRef LevelDB) {
     return DllCall(NumGet(NumGet(this+0)+8*A_PtrSize), "ptr", this, "float*", LevelDB)
 }
 VA_IAudioEndpointVolume_GetMasterVolumeLevelScalar(this, ByRef Level) {
     return DllCall(NumGet(NumGet(this+0)+9*A_PtrSize), "ptr", this, "float*", Level)
 }
 VA_IAudioEndpointVolume_SetChannelVolumeLevel(this, Channel, LevelDB, GuidEventContext="") {
     return DllCall(NumGet(NumGet(this+0)+10*A_PtrSize), "ptr", this, "uint", Channel, "float", LevelDB, "ptr", VA_GUID(GuidEventContext))
 }
 VA_IAudioEndpointVolume_SetChannelVolumeLevelScalar(this, Channel, Level, GuidEventContext="") {
     return DllCall(NumGet(NumGet(this+0)+11*A_PtrSize), "ptr", this, "uint", Channel, "float", Level, "ptr", VA_GUID(GuidEventContext))
 }
 VA_IAudioEndpointVolume_GetChannelVolumeLevel(this, Channel, ByRef LevelDB) {
     return DllCall(NumGet(NumGet(this+0)+12*A_PtrSize), "ptr", this, "uint", Channel, "float*", LevelDB)
 }
 VA_IAudioEndpointVolume_GetChannelVolumeLevelScalar(this, Channel, ByRef Level) {
     return DllCall(NumGet(NumGet(this+0)+13*A_PtrSize), "ptr", this, "uint", Channel, "float*", Level)
 }
 VA_IAudioEndpointVolume_SetMute(this, Mute, GuidEventContext="") {
     return DllCall(NumGet(NumGet(this+0)+14*A_PtrSize), "ptr", this, "int", Mute, "ptr", VA_GUID(GuidEventContext))
 }
 VA_IAudioEndpointVolume_GetMute(this, ByRef Mute) {
     return DllCall(NumGet(NumGet(this+0)+15*A_PtrSize), "ptr", this, "int*", Mute)
 }
 VA_IAudioEndpointVolume_GetVolumeStepInfo(this, ByRef Step, ByRef StepCount) {
     return DllCall(NumGet(NumGet(this+0)+16*A_PtrSize), "ptr", this, "uint*", Step, "uint*", StepCount)
 }
 VA_IAudioEndpointVolume_VolumeStepUp(this, GuidEventContext="") {
     return DllCall(NumGet(NumGet(this+0)+17*A_PtrSize), "ptr", this, "ptr", VA_GUID(GuidEventContext))
 }
 VA_IAudioEndpointVolume_VolumeStepDown(this, GuidEventContext="") {
     return DllCall(NumGet(NumGet(this+0)+18*A_PtrSize), "ptr", this, "ptr", VA_GUID(GuidEventContext))
 }
 VA_IAudioEndpointVolume_QueryHardwareSupport(this, ByRef HardwareSupportMask) {
     return DllCall(NumGet(NumGet(this+0)+19*A_PtrSize), "ptr", this, "uint*", HardwareSupportMask)
 }
 VA_IAudioEndpointVolume_GetVolumeRange(this, ByRef MinDB, ByRef MaxDB, ByRef IncrementDB) {
     return DllCall(NumGet(NumGet(this+0)+20*A_PtrSize), "ptr", this, "float*", MinDB, "float*", MaxDB, "float*", IncrementDB)
 }
 
 ;
 ; IPerChannelDbLevel  : {C2F8E001-F205-4BC9-99BC-C13B1E048CCB}
 ;   IAudioVolumeLevel : {7FB7B48F-531D-44A2-BCB3-5AD5A134B3DC}
 ;   IAudioBass        : {A2B1A1D9-4DB3-425D-A2B2-BD335CB3E2E5}
 ;   IAudioMidrange    : {5E54B6D7-B44B-40D9-9A9E-E691D9CE6EDF}
 ;   IAudioTreble      : {0A717812-694E-4907-B74B-BAFA5CFDCA7B}
 ;
 VA_IPerChannelDbLevel_GetChannelCount(this, ByRef Channels) {
     return DllCall(NumGet(NumGet(this+0)+3*A_PtrSize), "ptr", this, "uint*", Channels)
 }
 VA_IPerChannelDbLevel_GetLevelRange(this, Channel, ByRef MinLevelDB, ByRef MaxLevelDB, ByRef Stepping) {
     return DllCall(NumGet(NumGet(this+0)+4*A_PtrSize), "ptr", this, "uint", Channel, "float*", MinLevelDB, "float*", MaxLevelDB, "float*", Stepping)
 }
 VA_IPerChannelDbLevel_GetLevel(this, Channel, ByRef LevelDB) {
     return DllCall(NumGet(NumGet(this+0)+5*A_PtrSize), "ptr", this, "uint", Channel, "float*", LevelDB)
 }
 VA_IPerChannelDbLevel_SetLevel(this, Channel, LevelDB, GuidEventContext="") {
     return DllCall(NumGet(NumGet(this+0)+6*A_PtrSize), "ptr", this, "uint", Channel, "float", LevelDB, "ptr", VA_GUID(GuidEventContext))
 }
 VA_IPerChannelDbLevel_SetLevelUniform(this, LevelDB, GuidEventContext="") {
     return DllCall(NumGet(NumGet(this+0)+7*A_PtrSize), "ptr", this, "float", LevelDB, "ptr", VA_GUID(GuidEventContext))
 }
 VA_IPerChannelDbLevel_SetLevelAllChannels(this, LevelsDB, ChannelCount, GuidEventContext="") {
     return DllCall(NumGet(NumGet(this+0)+8*A_PtrSize), "ptr", this, "uint", LevelsDB, "uint", ChannelCount, "ptr", VA_GUID(GuidEventContext))
 }
 
 ;
 ; IAudioMute : {DF45AEEA-B74A-4B6B-AFAD-2366B6AA012E}
 ;
 VA_IAudioMute_SetMute(this, Muted, GuidEventContext="") {
     return DllCall(NumGet(NumGet(this+0)+3*A_PtrSize), "ptr", this, "int", Muted, "ptr", VA_GUID(GuidEventContext))
 }
 VA_IAudioMute_GetMute(this, ByRef Muted) {
     return DllCall(NumGet(NumGet(this+0)+4*A_PtrSize), "ptr", this, "int*", Muted)
 }
 
 ;
 ; IAudioAutoGainControl : {85401FD4-6DE4-4b9d-9869-2D6753A82F3C}
 ;
 VA_IAudioAutoGainControl_GetEnabled(this, ByRef Enabled) {
     return DllCall(NumGet(NumGet(this+0)+3*A_PtrSize), "ptr", this, "int*", Enabled)
 }
 VA_IAudioAutoGainControl_SetEnabled(this, Enable, GuidEventContext="") {
     return DllCall(NumGet(NumGet(this+0)+4*A_PtrSize), "ptr", this, "int", Enable, "ptr", VA_GUID(GuidEventContext))
 }
 
 ;
 ; IAudioMeterInformation : {C02216F6-8C67-4B5B-9D00-D008E73E0064}
 ;
 VA_IAudioMeterInformation_GetPeakValue(this, ByRef Peak) {
     return DllCall(NumGet(NumGet(this+0)+3*A_PtrSize), "ptr", this, "float*", Peak)
 }
 VA_IAudioMeterInformation_GetMeteringChannelCount(this, ByRef ChannelCount) {
     return DllCall(NumGet(NumGet(this+0)+4*A_PtrSize), "ptr", this, "uint*", ChannelCount)
 }
 VA_IAudioMeterInformation_GetChannelsPeakValues(this, ChannelCount, PeakValues) {
     return DllCall(NumGet(NumGet(this+0)+5*A_PtrSize), "ptr", this, "uint", ChannelCount, "ptr", PeakValues)
 }
 VA_IAudioMeterInformation_QueryHardwareSupport(this, ByRef HardwareSupportMask) {
     return DllCall(NumGet(NumGet(this+0)+6*A_PtrSize), "ptr", this, "uint*", HardwareSupportMask)
 }
 
 ;
 ; IAudioClient : {1CB9AD4C-DBFA-4c32-B178-C2F568A703B2}
 ;
 VA_IAudioClient_Initialize(this, ShareMode, StreamFlags, BufferDuration, Periodicity, Format, AudioSessionGuid) {
     return DllCall(NumGet(NumGet(this+0)+3*A_PtrSize), "ptr", this, "int", ShareMode, "uint", StreamFlags, "int64", BufferDuration, "int64", Periodicity, "ptr", Format, "ptr", VA_GUID(AudioSessionGuid))
 }
 VA_IAudioClient_GetBufferSize(this, ByRef NumBufferFrames) {
     return DllCall(NumGet(NumGet(this+0)+4*A_PtrSize), "ptr", this, "uint*", NumBufferFrames)
 }
 VA_IAudioClient_GetStreamLatency(this, ByRef Latency) {
     return DllCall(NumGet(NumGet(this+0)+5*A_PtrSize), "ptr", this, "int64*", Latency)
 }
 VA_IAudioClient_GetCurrentPadding(this, ByRef NumPaddingFrames) {
     return DllCall(NumGet(NumGet(this+0)+6*A_PtrSize), "ptr", this, "uint*", NumPaddingFrames)
 }
 VA_IAudioClient_IsFormatSupported(this, ShareMode, Format, ByRef ClosestMatch) {
     return DllCall(NumGet(NumGet(this+0)+7*A_PtrSize), "ptr", this, "int", ShareMode, "ptr", Format, "ptr*", ClosestMatch)
 }
 VA_IAudioClient_GetMixFormat(this, ByRef Format) {
     return DllCall(NumGet(NumGet(this+0)+8*A_PtrSize), "ptr", this, "uint*", Format)
 }
 VA_IAudioClient_GetDevicePeriod(this, ByRef DefaultDevicePeriod, ByRef MinimumDevicePeriod) {
     return DllCall(NumGet(NumGet(this+0)+9*A_PtrSize), "ptr", this, "int64*", DefaultDevicePeriod, "int64*", MinimumDevicePeriod)
 }
 VA_IAudioClient_Start(this) {
     return DllCall(NumGet(NumGet(this+0)+10*A_PtrSize), "ptr", this)
 }
 VA_IAudioClient_Stop(this) {
     return DllCall(NumGet(NumGet(this+0)+11*A_PtrSize), "ptr", this)
 }
 VA_IAudioClient_Reset(this) {
     return DllCall(NumGet(NumGet(this+0)+12*A_PtrSize), "ptr", this)
 }
 VA_IAudioClient_SetEventHandle(this, eventHandle) {
     return DllCall(NumGet(NumGet(this+0)+13*A_PtrSize), "ptr", this, "ptr", eventHandle)
 }
 VA_IAudioClient_GetService(this, iid, ByRef Service) {
     return DllCall(NumGet(NumGet(this+0)+14*A_PtrSize), "ptr", this, "ptr", VA_GUID(iid), "ptr*", Service)
 }
 
 ;
 ; IAudioSessionControl : {F4B1A599-7266-4319-A8CA-E70ACB11E8CD}
 ;
 /*
 AudioSessionStateInactive = 0
 AudioSessionStateActive = 1
 AudioSessionStateExpired = 2
 */
 VA_IAudioSessionControl_GetState(this, ByRef State) {
     return DllCall(NumGet(NumGet(this+0)+3*A_PtrSize), "ptr", this, "int*", State)
 }
 VA_IAudioSessionControl_GetDisplayName(this, ByRef DisplayName) {
     hr := DllCall(NumGet(NumGet(this+0)+4*A_PtrSize), "ptr", this, "ptr*", DisplayName)
     VA_WStrOut(DisplayName)
     return hr
 }
 VA_IAudioSessionControl_SetDisplayName(this, DisplayName, EventContext) {
     return DllCall(NumGet(NumGet(this+0)+5*A_PtrSize), "ptr", this, "wstr", DisplayName, "ptr", VA_GUID(EventContext))
 }
 VA_IAudioSessionControl_GetIconPath(this, ByRef IconPath) {
     hr := DllCall(NumGet(NumGet(this+0)+6*A_PtrSize), "ptr", this, "ptr*", IconPath)
     VA_WStrOut(IconPath)
     return hr
 }
 VA_IAudioSessionControl_SetIconPath(this, IconPath) {
     return DllCall(NumGet(NumGet(this+0)+7*A_PtrSize), "ptr", this, "wstr", IconPath)
 }
 VA_IAudioSessionControl_GetGroupingParam(this, ByRef Param) {
     VarSetCapacity(Param,16,0)
     hr := DllCall(NumGet(NumGet(this+0)+8*A_PtrSize), "ptr", this, "ptr", &Param)
     VA_GUIDOut(Param)
     return hr
 }
 VA_IAudioSessionControl_SetGroupingParam(this, Param, EventContext) {
     return DllCall(NumGet(NumGet(this+0)+9*A_PtrSize), "ptr", this, "ptr", VA_GUID(Param), "ptr", VA_GUID(EventContext))
 }
 VA_IAudioSessionControl_RegisterAudioSessionNotification(this, NewNotifications) {
     return DllCall(NumGet(NumGet(this+0)+10*A_PtrSize), "ptr", this, "ptr", NewNotifications)
 }
 VA_IAudioSessionControl_UnregisterAudioSessionNotification(this, NewNotifications) {
     return DllCall(NumGet(NumGet(this+0)+11*A_PtrSize), "ptr", this, "ptr", NewNotifications)
 }
 
 ;
 ; IAudioSessionManager : {BFA971F1-4D5E-40BB-935E-967039BFBEE4}
 ;
 VA_IAudioSessionManager_GetAudioSessionControl(this, AudioSessionGuid) {
     return DllCall(NumGet(NumGet(this+0)+3*A_PtrSize), "ptr", this, "ptr", VA_GUID(AudioSessionGuid))
 }
 VA_IAudioSessionManager_GetSimpleAudioVolume(this, AudioSessionGuid, StreamFlags, ByRef AudioVolume) {
     return DllCall(NumGet(NumGet(this+0)+4*A_PtrSize), "ptr", this, "ptr", VA_GUID(AudioSessionGuid), "uint", StreamFlags, "uint*", AudioVolume)
 }
 
 ;
 ; IMMDeviceEnumerator
 ;
 VA_IMMDeviceEnumerator_EnumAudioEndpoints(this, DataFlow, StateMask, ByRef Devices) {
     return DllCall(NumGet(NumGet(this+0)+3*A_PtrSize), "ptr", this, "int", DataFlow, "uint", StateMask, "ptr*", Devices)
 }
 VA_IMMDeviceEnumerator_GetDefaultAudioEndpoint(this, DataFlow, Role, ByRef Endpoint) {
     return DllCall(NumGet(NumGet(this+0)+4*A_PtrSize), "ptr", this, "int", DataFlow, "int", Role, "ptr*", Endpoint)
 }
 VA_IMMDeviceEnumerator_GetDevice(this, id, ByRef Device) {
     return DllCall(NumGet(NumGet(this+0)+5*A_PtrSize), "ptr", this, "wstr", id, "ptr*", Device)
 }
 VA_IMMDeviceEnumerator_RegisterEndpointNotificationCallback(this, Client) {
     return DllCall(NumGet(NumGet(this+0)+6*A_PtrSize), "ptr", this, "ptr", Client)
 }
 VA_IMMDeviceEnumerator_UnregisterEndpointNotificationCallback(this, Client) {
     return DllCall(NumGet(NumGet(this+0)+7*A_PtrSize), "ptr", this, "ptr", Client)
 }
 
 ;
 ; IMMDeviceCollection
 ;
 VA_IMMDeviceCollection_GetCount(this, ByRef Count) {
     return DllCall(NumGet(NumGet(this+0)+3*A_PtrSize), "ptr", this, "uint*", Count)
 }
 VA_IMMDeviceCollection_Item(this, Index, ByRef Device) {
     return DllCall(NumGet(NumGet(this+0)+4*A_PtrSize), "ptr", this, "uint", Index, "ptr*", Device)
 }
 
 ;
 ; IControlInterface
 ;
 VA_IControlInterface_GetName(this, ByRef Name) {
     hr := DllCall(NumGet(NumGet(this+0)+3*A_PtrSize), "ptr", this, "ptr*", Name)
     VA_WStrOut(Name)
     return hr
 }
 VA_IControlInterface_GetIID(this, ByRef IID) {
     VarSetCapacity(IID,16,0)
     hr := DllCall(NumGet(NumGet(this+0)+4*A_PtrSize), "ptr", this, "ptr", &IID)
     VA_GUIDOut(IID)
     return hr
 }
 
 
 /*
     INTERFACES REQUIRING WINDOWS 7 / SERVER 2008 R2
 */
 
 ;
 ; IAudioSessionControl2 : {bfb7ff88-7239-4fc9-8fa2-07c950be9c6d}
 ;   extends IAudioSessionControl
 ;
 VA_IAudioSessionControl2_GetSessionIdentifier(this, ByRef id) {
     hr := DllCall(NumGet(NumGet(this+0)+12*A_PtrSize), "ptr", this, "ptr*", id)
     VA_WStrOut(id)
     return hr
 }
 VA_IAudioSessionControl2_GetSessionInstanceIdentifier(this, ByRef id) {
     hr := DllCall(NumGet(NumGet(this+0)+13*A_PtrSize), "ptr", this, "ptr*", id)
     VA_WStrOut(id)
     return hr
 }
 VA_IAudioSessionControl2_GetProcessId(this, ByRef pid) {
     return DllCall(NumGet(NumGet(this+0)+14*A_PtrSize), "ptr", this, "uint*", pid)
 }
 VA_IAudioSessionControl2_IsSystemSoundsSession(this) {
     return DllCall(NumGet(NumGet(this+0)+15*A_PtrSize), "ptr", this)
 }
 VA_IAudioSessionControl2_SetDuckingPreference(this, OptOut) {
     return DllCall(NumGet(NumGet(this+0)+16*A_PtrSize), "ptr", this, "int", OptOut)
 }
 
 ;
 ; IAudioSessionManager2 : {77AA99A0-1BD6-484F-8BC7-2C654C9A9B6F}
 ;   extends IAudioSessionManager
 ;
 VA_IAudioSessionManager2_GetSessionEnumerator(this, ByRef SessionEnum) {
     return DllCall(NumGet(NumGet(this+0)+5*A_PtrSize), "ptr", this, "ptr*", SessionEnum)
 }
 VA_IAudioSessionManager2_RegisterSessionNotification(this, SessionNotification) {
     return DllCall(NumGet(NumGet(this+0)+6*A_PtrSize), "ptr", this, "ptr", SessionNotification)
 }
 VA_IAudioSessionManager2_UnregisterSessionNotification(this, SessionNotification) {
     return DllCall(NumGet(NumGet(this+0)+7*A_PtrSize), "ptr", this, "ptr", SessionNotification)
 }
 VA_IAudioSessionManager2_RegisterDuckNotification(this, SessionNotification) {
     return DllCall(NumGet(NumGet(this+0)+8*A_PtrSize), "ptr", this, "ptr", SessionNotification)
 }
 VA_IAudioSessionManager2_UnregisterDuckNotification(this, SessionNotification) {
     return DllCall(NumGet(NumGet(this+0)+9*A_PtrSize), "ptr", this, "ptr", SessionNotification)
 }
 
 ;
 ; IAudioSessionEnumerator : {E2F5BB11-0570-40CA-ACDD-3AA01277DEE8}
 ;
 VA_IAudioSessionEnumerator_GetCount(this, ByRef SessionCount) {
     return DllCall(NumGet(NumGet(this+0)+3*A_PtrSize), "ptr", this, "int*", SessionCount)
 }
 VA_IAudioSessionEnumerator_GetSession(this, SessionCount, ByRef Session) {
     return DllCall(NumGet(NumGet(this+0)+4*A_PtrSize), "ptr", this, "int", SessionCount, "ptr*", Session)
 }
 
 
 /*
     UNDOCUMENTED INTERFACES
 */
 
 ; Thanks to Dave Amenta for publishing this interface - http://goo.gl/6L93L
 ; IID := "{568b9108-44bf-40b4-9006-86afe5b5a620}"
 ; CLSID := "{294935CE-F637-4E7C-A41B-AB255460B862}"
 VA_xIPolicyConfigVista_SetDefaultEndpoint(this, DeviceId, Role) {
     return DllCall(NumGet(NumGet(this+0)+12*A_PtrSize), "ptr", this, "wstr", DeviceId, "int", Role)
 }
 


 ; License:

  ; Copyright (c) 2013 Lexikos
  ; 
  ; This software is provided 'as-is', without any express or implied
  ; warranty. In no event will the authors be held liable for any damages
  ; arising from the use of this software.
  ; 
  ; Permission is granted to anyone to use this software for any purpose,
  ; including commercial applications, and to alter it and redistribute it
  ; freely, subject to the following restrictions:
  ; 
  ; 1. The origin of this software must not be misrepresented; you must not
  ;    claim that you wrote the original software. If you use this software
  ;    in a product, an acknowledgment in the product documentation would be
  ;    appreciated but is not required.
  ; 2. Altered source versions must be plainly marked as such, and must not be
  ;    misrepresented as being the original software.

; --uID:3950707558