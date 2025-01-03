f_CreateGUI()
{ ;; create the GUI
    global
    Gui AS: new, +AlwaysOnTop -SysMenu -ToolWindow -caption +Border +LabelAS_ ;+Resize +MinSize1000x
    gui AS: default
    gui add, text,vTitleString,%  "Choose Audio-Out Device"
    gui font, s8
    Gui Margin, 16,0
    gui +hwndASGUI
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
            gui add, button,Default r2 w200 gChooseAudioIn hidden disabled vValIn%A_Index%  ,% "&" A_Index " " trim(C.1)
            gui add, button,yp Default r2 w200  gChooseAudioOut vVal%A_Index%,% "&" A_Index " " trim(C.2)
            if (trim(C.2)==ActiveDevices[2])
                guicontrol disable, % "ValIn" A_Index
            if (trim(C.1)==ActiveDevices[1])
                guicontrol disable, % "Val" A_Index
        }
        else
        { ;; audioOUT First
            ButtonFaceIn:=C.1
            ButtonFaceOut:=C.2
            ; if (c.MaxIndex()!=2) ;; removed until solution found
            gui add, button,Default r2 w200 gChooseAudioIn hidden disabled vValIn%A_Index%  ,% "&" A_Index " " trim(C.2)
            gui add, button,yp Default r2 w200  gChooseAudioOut vVal%A_Index%,% "&" A_Index " " trim(C.1)
            if (trim(C.2)==ActiveDevices[2]) 
                guicontrol disable, % "ValIn" A_Index
            if (trim(C.1)==ActiveDevices[1])
                guicontrol disable, % "Val" A_Index
            ; if (c.MaxIndex()!=2) ;; append credits-string to the end of the 

        }

    }
    gui add, button, yp+35  h20 glCredits vCredits2, &Credits

    return
}
lCredits:
script.About()
return
AS_Escape()
{ ;; close the GUI
    gui AS: hide
}
f_showGUI2(wParam, lParam)
{
    if (lParam = 0x202)
    {
        f_ShowGUI()
        return 0
    }
}
f_ShowGUI()
{ ;; show the GUI
    GUI AS: show, AutoSize , AS - Audio-Switcher
    sleep 300 ;; make sure accidental spacebar-inputs or other don't immediately trigger a button.
}
