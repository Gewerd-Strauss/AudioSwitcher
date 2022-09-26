#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
#SingleInstance, Force

Menu, Tray, Icon, C:\windows\system32\shell32.dll,138 ;Set custom Script icon
#Include <scriptObj/scriptObj>
CreditsRaw=
(LTRI
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
                    ,ghlink       : "https://github.com/Gewerd-Strauss/AudioSwitcher"
                    ,doctext	  : ""
                    ,doclink	  : ""
                    ,forumtext	  : ""
                    ,forumlink	  : ""
                    ,donateLink	  : ""
                    ,resfolder    : A_ScriptDir "\res"
                    ,iconfile	  : ""
					,rfile  	  : "https://github.com/Gewerd-Strauss/AudioSwitcher/archive/refs/heads/MAIN.zip"
					,vfile_raw	  : "https://raw.githubusercontent.com/Gewerd-Strauss/AudioSwitcher/main/version.ini" 
					,vfile 		  : "https://raw.githubusercontent.com/Gewerd-Strauss/AudioSwitcher/main/version.ini" 
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
global bStartOnOutPut:=true
script.Load()
, script.Update(,,1) ;DO NOT ACTIVATE THISLINE UNTIL YOU DUMBO HAS FIXED THE DAMN METHOD. God damn it.
f_CreateTrayMenu()




f_LoadDevicesOut()
f_LoadDevicesIn()
f_CreateGUI()
; f_ShowGUI()
return


f_CreateGUI()
{ ;; create the GUI
    global
    Gui, AS: new, +AlwaysOnTop -SysMenu -ToolWindow -caption +Border +LabelAS_ ;+Resize +MinSize1000x
    gui, AS: default
    gui, add, text,vTitleString,%  "Choose Audio-Out Device"
    gui, font, s8
    ; gui, add, text, yp+15  vCredits,% "v." script.version " - by " script.Author
    Gui, Margin, 16,0
    gui, +hwndASGUI
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
    for k,v in Matches
    {
        C:=strsplit(v,"||||")
        if bDevicesInFirst
        {
            ButtonFaceIn:=C.1
            ButtonFaceOut:=C.2
            gui, add, button,Default r2 w200 gChooseAudioIn hidden disabled vValIn%A_Index%  ,% "&" A_Index " " trim(C.1)
            gui, add, button,yp Default r2 w200  gChooseAudioOut vVal%A_Index%,% "&" A_Index " " trim(C.2)
        }
        else
        { ;; audioOUT First
            ButtonFaceIn:=C.1
            ButtonFaceOut:=C.2
            ; if (c.MaxIndex()!=2) ;; removed until solution found
            ;     gui, add, text, yp+35 vCredits1,% "v." script.version " - by " script.Author
            gui, add, button,Default r2 w200 gChooseAudioIn hidden disabled vValIn%A_Index%  ,% "&" A_Index " " trim(C.2)
            gui, add, button,yp Default r2 w200  gChooseAudioOut vVal%A_Index%,% "&" A_Index " " trim(C.1)
            if (c.MaxIndex()!=2) ;; append credits-string to the end of the 
            gui, add, button, yp+35  h20 glCredits vCredits2, Credits ;% "v." script.version " - by " script.Author
                ; gui, add, text, yp+15  vCredits,% "v." script.version " - by " script.Author

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
        }
        for k,v in DevicesIn
        {
            A:=strsplit(k," - ")
            ; gui, add, button,Default w200 gChooseAudioIn vValIn%A_Index% disabled hidden ,% "&" A_Index " " trim(A[A.MaxIndex()])
            guicontrol,Enable, ValIn%A_Index% 
            guicontrol,Show, ValIn%A_Index% 
            MaxValIn:="ValIn" A_Index

        }
        guicontrol,show, % (bStartOnOutPut?"Credits2":"Credits1")
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


#If 
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