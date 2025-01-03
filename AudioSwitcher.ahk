#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#Requires AutoHotkey v1.1+
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance Force
SysGet vMoncnt, MonitorCount
Menu Tray, Icon, C:\windows\system32\shell32.dll,138 ;Set custom Script icon

FileGetTime ModDate,%A_ScriptFullPath%,M
FileGetTime CrtDate,%A_ScriptFullPath%,C
CrtDate:=SubStr(CrtDate,7,  2) "." SubStr(CrtDate,5,2) "." SubStr(CrtDate,1,4)
ModDate:=SubStr(ModDate,7,  2) "." SubStr(ModDate,5,2) "." SubStr(ModDate,1,4)
global script := new script()

script := {base         : script.base
        , name         : regexreplace(A_ScriptName, "\.\w+")
        , crtdate      : CrtDate
        , moddate      : ModDate
        , resfolder    : A_ScriptDir "\res"
        , iconfile	  : ""
        , config:		[]
        , configfile   : A_ScriptDir "\INI-Files\" regexreplace(A_ScriptName, "\.\w+") ".ini"
        , configfolder : A_ScriptDir "\INI-Files"
        , aboutPath : A_ScriptDir "\res\About.html"
        , reqInternet   : false
        , authorID	  : "Laptop-C"
        , Computername : A_ComputerName
        , license : A_ScriptDir "\res\LICENSE.txt"
        , blank : "" }
; , rfile  	  : "https://github.com/Gewerd-Strauss/AudioSwitcher/archive/refs/heads/MAIN.zip"
; , vfile_raw	  : "https://raw.githubusercontent.com/Gewerd-Strauss/AudioSwitcher/main/version.ini" 
; , vfile 		  : "https://raw.githubusercontent.com/Gewerd-Strauss/AudioSwitcher/main/version.ini" 
; , vfile_local  : A_ScriptDir "\version.ini" 

global bStartOnOutPut:=true
FileRead version__, % A_ScriptDir "\version.ini"

script.version := StrReplace(version__,"`r`n")
script.Update(,,1) ;DO NOT ACTIVATE THISLINE UNTIL YOU DUMBO HAS FIXED THE DAMN METHOD. God damn it.
    , script.loadCredits(script.resfolder "\credits.txt")
    , script.loadMetadata(script.resfolder "\meta.txt")
f_CreateTrayMenu()


oMyDevices := {}
f_LoadDevicesOut()
f_LoadDevicesIn()
f_CreateGUI()
Onmessage(0x404,"f_showGUI2")
; oMyDevices.Push({"DeviceName":"Kopfhörer (WH-1000XM3)", "DeviceID":"SWD\MMDEVAPI\{0.0.0.00000000}.{2DA0C039-7454-45FD-BFCA-4656F85C1384}"})



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

OnMessage(0x0219, "WM_DEVICECHANGE") 
Return
WM_DEVICECHANGE(wParam, lParam, msg, hwnd)
{
    SetTimer CheckDevicesStatus , -1250
}





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

ChooseAudioOut() 
{ ;; chooses a value to continue with.
    global
    gui Submit,
    GuiControlGet ChosenDevice,, % A_GuiControl
    ChosenDevice2:=regexreplace(ChosenDevice,"\&\d*\s*")

    for k,v in DevicesOut
    {
        ; Clipboard:=Obj2Str(DevicesOut)
        k:=(bDevicesOutFirst?StrSplit(k, "||||" ).1:StrSplit(k, "||||" ).2)
        A:=strsplit(k," - ")
        Value:="&" A_Index " " trim(A[A.MaxIndex()])
        guicontrol enable, % "Val" A_Index
        k2:=strreplace(k," - "," ")
        ChosenDevice:=strreplace(ChosenDevice," - "," ")
        if Instr(ChosenDevice,k2)
            Submit:=[ChosenDevice,k,v]
    }
    guicontrol disable, % ChosenValueOut:=A_GuiControl
    f_SelectAudio(Submit[3])
    return
} ; ChosenValueOut ChosenValueIn

ChooseAudioIn() 
{ ;; chooses a value to continue with.
    global
    gui Submit,
    GuiControlGet ChosenDevice,, % A_GuiControl
    ChosenDevice2:=regexreplace(ChosenDevice,"\&\d*\s*")
    for k,v in DevicesIn
    {
        A:=strsplit(k," - ")
        Value:="&" A_Index " " trim(A[A.MaxIndex()])
        guicontrol enable, % "ValIn" A_Index
        k2:=strreplace(k," - "," ")
        if Instr(ChosenDevice,k2)
            Submit:=[ChosenDevice,k,v]
    }
    guicontrol disable, % ChosenValueIn:=A_GuiControl
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
    gui AS: default
    gui Submit, NoHide
    ActiveDevices:=strsplit(getInfo(),"`n")
    if bStartOnOutPut
    {
        ;; show In
        guicontrol,,TitleString, % "Choose Audio-In Device"
        for k,v in DevicesOut
        {
            A:=strsplit(k," - ")
            ; gui, add, button,Default w200 gChooseAudioIn vValIn%A_Index% disabled hidden ,% "&" A_Index " " trim(A[A.MaxIndex()])

            guicontrol disable, Val%A_Index%
            guicontrol hide, Val%A_Index%
            MaxValOut:="Val" A_Index
            if (trim(A.1)==ActiveDevices[1]) || (v==ActiveDevices[1])
                guicontrol disable, % "Val" A_Index
        }
        for k,v in DevicesIn
        {
            A:=strsplit(k," - ")
            ; gui, add, button,Default w200 gChooseAudioIn vValIn%A_Index% disabled hidden ,% "&" A_Index " " trim(A[A.MaxIndex()])
            guicontrol Enable, ValIn%A_Index%
            guicontrol Show, ValIn%A_Index%
            MaxValIn:="ValIn" A_Index
            if (trim(A.1)==ActiveDevices[2]) || (v==ActiveDevices[2])
                guicontrol disable, % "ValIn" A_Index

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
            guicontrol enable, Val%A_Index%
            guicontrol show, Val%A_Index%
            MaxValOut:="Val" A_Index 
        }
        for k,v in DevicesOut
        {
            if (trim(k)==ActiveDevices[1])
                guicontrol disable, % "Val" A_Index
        }
        for k,v in DevicesIn
        {
            A:=strsplit(k," - ")
            ; gui, add, button,Default w200 gChooseAudioIn vValIn%A_Index% disabled hidden ,% "&" A_Index " " trim(A[A.MaxIndex()])
            guicontrol disable, ValIn%A_Index%
            guicontrol hide, ValIn%A_Index%
            MaxValIn:="ValIn" A_Index

        }

        guicontrol hide, % (bStartOnOutPut?"Credits2":"Credits1")
        ControlGetPos X, Y, Width, Height,% e:="Val" SubStr(MaxValOut,4,1)-1,% "AS - Audio-Switcher"
        ; ControlGetPos, [ X, Y, Width, Height, Control, WinTitle, WinText, ExcludeTitle
        ; guicontrol, move, % MaxValOut, % xp "y" d:=SubStr(MaxValOut,4,1)*33
        ; guicontrol, move, % Credits1, % xp "y" d:=(SubStr(MaxValOut,4,1)+1)*33^
        ; ChosenValueOut ChosenValueIn
    }
    guicontrol disable, % ChosenValueIn
    guicontrol disable, % ChosenValueOut
    bStartOnOutPut:= !bStartOnOutPut


    gui Submit, NoHide
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
    menu tray, add,
    Menu Misc, add, Open Script-folder, lOpenScriptFolder
    menu Misc, Add, Reload, lReload
    menu Misc, Add, About, Label_AboutFile
    SplitPath A_ScriptName,,,, scriptname
    Menu tray, add, Miscellaneous, :Misc
    menu tray, add,
    return
}
lOpenScriptFolder:
run % A_ScriptDir
return
lReload: 
reload
return
Label_AboutFile:
FileDelete % script.AboutPath
script.about()
return


getInfo() {
    device := VA_GetDevice("playback")

    deviceName := VA_GetDeviceName(device)
    device := VA_GetDevice("capture")
    deviceName .= "`n" VA_GetDeviceName(device)
    Menu Tray, Tip, Audio Source Switcher - %deviceName%
    return deviceName
}


#Include <gui>
#Include <VA>
#Include <Obj2Str>
#Include <script>
#Include <isDebug>
#Include <hotkeys>
