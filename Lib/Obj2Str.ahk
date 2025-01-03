; #region:Obj2Str (3878929696)

; #region:Metadata:
; Snippet: Obj2Str
; --------------------------------------------------------------
; Author: maestrith
; Source: https://www.autohotkey.com/boards/viewtopic.php?t=60522#p255613
; (16.06.2024)
; --------------------------------------------------------------
; Library: AHK-Rare
; Section: 11 - Objects
; Dependencies: /
; AHK_Version: v1, v2
; --------------------------------------------------------------

; #endregion:Metadata


; #region:Description:
; Convert an Object to a string
; #endregion:Description

; #region:Example
; associativeObj1:={A:1,B:2,C:[1,2,3]}
; simpleObj2:=[1,2,3,4]
; msgbox, % Obj2Str(associativeObj1)
; msgbox, % Obj2Str(simpleObj2)
; #endregion:Example


; #region:Code
Obj2Str(Obj, FullPath := 1, BottomBlank := 0) {
    static String, Blank
    if (FullPath = 1)
        String := FullPath := Blank := ""
    if (IsObject(Obj)) {
        for a, b in Obj {
            if (IsObject(b))
                Obj2Str(b, FullPath "." a, BottomBlank)
            else {
                if (BottomBlank = 0)
                    String .= FullPath "." a " = " b "`n"
                else if (b != "")
                    String .= FullPath "." a " = " b "`n"
                else
                    Blank .= FullPath "." a " =`n"
            }
        }
    }
    return String Blank
}
; #endregion:Code



; #endregion:Obj2Str (3878929696)
