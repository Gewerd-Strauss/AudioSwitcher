#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance, Force
OnMessage(0x44, "OnMsgBox")
MsgBox 0x10, NOT FINISHED, This file is an exact copy of AudioSwitcher.ahk`, as I have not found the codes required for adjustments`n
OnMessage(0x44, "")

OnMsgBox() {
    DetectHiddenWindows, On
    Process, Exist
    If (WinExist("ahk_class #32770 ahk_pid " . ErrorLevel)) {
        ControlSetText Button1, Exit
    }
}
Menu, Tray, Icon, C:\windows\system32\shell32.dll,138 ;Set custom Script icon
#Include <scriptObj/scriptObj>
CreditsRaw=
(LTRIM
author1   -		 snippetName1		   		  			-	URL1
Flipeador   -		 original Script		   		  			-	https://www.autohotkey.com/boards/viewtopic.php?p=221751#p221751
Gewerd Strauss		- snippetName2|SnippetName3 (both at the same URL)								-	/
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
                    ,ghlink       : "https://github.com/Gewerd-Strauss/REPOSITORY_NAME"
                    ,doctext	  : ""
                    ,doclink	  : ""
                    ,forumtext	  : ""
                    ,forumlink	  : ""
                    ,donateLink	  : ""
                    ,resfolder    : A_ScriptDir "\res"
                    ,iconfile	  : ""
					,rfile  	  : "https://github.com/Gewerd-Strauss/REPOSITORY_NAME/archive/refs/heads/BRANCH_NAME.zip"
					,vfile_raw	  : "https://raw.githubusercontent.com/Gewerd-Strauss/REPOSITORY_NAME/BRANCH_NAME/version.ini" 
					,vfile 		  : "https://raw.githubusercontent.com/Gewerd-Strauss/REPOSITORY_NAME/BRANCH_NAME/version.ini" 
					,vfile_local  : A_ScriptDir "\version.ini" 
                    ,config:		[]
					,configfile   : A_ScriptDir "\INI-Files\" regexreplace(A_ScriptName, "\.\w+") ".ini"
                    ,configfolder : A_ScriptDir "\INI-Files"}
/*	
	For throwing errors via script.debug
	script.Error:={	 Level		:""
					,Label		:""
					,Message	:""	
					,Error		:""		
					,Vars:		:[]
					,AddInfo:	:""}
	if script.error
		script.Debug(script.error.Level,script.error.Label,script.error.Message,script.error.AddInfo,script.error.Vars)
*/
; script.About()
script.Load()
, script.Update() ;DO NOT ACTIVATE THISLINE UNTIL YOU DUMBO HAS FIXED THE DAMN METHOD. God damn it.
f_CreateTrayMenu()





f_LoadDevices()
f_CreateGUI()
; f_ShowGUI()
return


f_CreateGUI()
{ ;; create the GUI
    global
    Gui, AS: new, +AlwaysOnTop -SysMenu -ToolWindow -caption +Border +LabelAS_ ;+Resize +MinSize1000x
    gui, default
    gui, add, text,,%  "Choose Audio-Out Device"
    gui, +hwndASGUI
    for k,v in Devices
    {
        A:=strsplit(k," - ")
        gui, add, button,Default w200 gChooseAudio vVal%A_Index%,% "&" A_Index " " trim(A[A.MaxIndex()])
    }
    return
}
AS_Escape()
{ ;; close the GUI
    gui, AS: hide
}
f_ShowGUI()
{ ;; show the GUI
    GUI, AS: show, AutoSize , AS - Audio-Switcher
}
ChooseAudio() 
{ ;; chooses a value to continue with.
    global
    gui, Submit, 
    GuiControlGet, ChosenDevice,, % A_GuiControl
    ChosenDevice2:=regexreplace(ChosenDevice,"\&\d*\s*")
    for k,v in Devices
    {
        A:=strsplit(k," - ")
        Value:="&" A_Index " " trim(A[A.MaxIndex()])
        guicontrol, enable, % "Val" A_Index
        k2:=strreplace(k," - "," ")
        if Instr(ChosenDevice,k2)
            Submit:=[ChosenDevice,k,v]
    }
    guicontrol, disable, % A_GuiControl
    f_SelectAudio(Submit[3])
    return
}
f_LoadDevices()
{ ;; retrieve all audio devices
    global
    Devices := {}
    IMMDeviceEnumerator := ComObjCreate("{BCDE0395-E52F-467C-8E3D-C4579291692E}", "{A95664D2-9614-4F35-A746-DE8DB63617E6}")

    ; IMMDeviceEnumerator::EnumAudioEndpoints
    ; eRender = 0, eCapture, eAll
    ; 0x1 = DEVICE_STATE_ACTIVE
    DllCall(NumGet(NumGet(IMMDeviceEnumerator+0)+3*A_PtrSize), "UPtr", IMMDeviceEnumerator, "UInt", 0, "UInt", 0x1, "UPtrP", IMMDeviceCollection, "UInt")
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
        
        ObjRawSet(Devices, DeviceName, DeviceID)
    }
    Devices2:={}
    Map:={}
    for DeviceName, DeviceID in Devices
    {
        ObjRawSet(Map,DeviceName,A_Index)
        ObjRawSet(Devices2,A_Index,DeviceID)
    }
    ObjRelease(IMMDeviceCollection)
    return
}
f_SelectAudio(Device)
{ ;; Select the audio device
    IPolicyConfig := ComObjCreate("{870af99c-171d-4f9e-af0d-e63df40c2bc9}", "{F8679F50-850A-41CF-9C72-430F290290C8}") ;00000102-0000-0000-C000-000000000046 00000000-0000-0000-C000-000000000046
    R := DllCall(NumGet(NumGet(IPolicyConfig+0)+13*A_PtrSize), "UPtr", IPolicyConfig, "Str", Device, "UInt", 0, "UInt")
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


#If 
F24::
f_LoadDevices()
if WinActive("AS - Audio-Switcher")
    AS_Escape()
Else
    f_ShowGUI()
return
#if WinActive("AS - Audio-Switcher")
WheelDown::Volume_Down
WheelUp::Volume_Up