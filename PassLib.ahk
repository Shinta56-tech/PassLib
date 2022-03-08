;*****************************************************************************************************************************************
; Standard Setting
;*****************************************************************************************************************************************

#Persistent
#SingleInstance, Force
#NoEnv
#UseHook
#InstallKeybdHook
#InstallMouseHook
#HotkeyInterval, 2000
#MaxHotkeysPerInterval, 200

Process, Priority, , Realtime
SendMode, Input
SetWorkingDir %A_ScriptDir%
SetTitleMatchMode, 2
SetMouseDelay -1
SetKeyDelay, -1
SetBatchLines, -1

;*****************************************************************************************************************************************
; Initialize
;*****************************************************************************************************************************************

initPassLib()

showGuiPasswordLibrary()

Return

^!Escape::ExitApp

;*****************************************************************************************************************************************
; Initialize
;*****************************************************************************************************************************************

initPassLib() {

    Gosub initPassLib

    Gui, PL_LIV:New
    Gui, PL_LIV:Add, Edit, gPSearch vPSearch w250 Section
    Gui, PL_LIV:Add, Button, gPBSend yp+0 x+10, Send
    Gui, PL_LIV:Add, Button, gPBAdd yp+0 x+20, Add
    Gui, PL_LIV:Add, Button, gPBUpdate yp+0 x+10, Update
    Gui, PL_LIV:Add, Button, gPBDelete yp+0 x+10, Delete
    Gui, PL_LIV:Add, Button, gPBSetfile yp+0 x+10, Setfile
    Gui, PL_LIV:Add, ListView, r14 w550 xs+0 ys+30 Sort vgPLv, Key|Value

    Gui, PL_REG:New
    Gui, PL_REG:Add, Text, Section, Key
    Gui, PL_REG:Add, Edit, xs+50 yp+0 w300 vgEditKey
    Gui, PL_REG:Add, Text, xs+0 y+10, Value
    Gui, PL_REG:Add, Edit, xs+50 yp+0 w300 vgEditValue
    Gui, PL_REG:Add, Button, w200 gPBOkAdd vgPBOkAdd xm+80 y+10 Section, Add
    Gui, PL_REG:Add, Button, w200 gPBOkUpd vgPBOkUpd xs+0 ys+0, Update
}

initPassLib:
    Global INI_PassLib := "PassLib.ini"

    Global iPasss, PSearch
    Global gPLv, gEditKey, gEditValue, gPBOkAdd, gPBOkUpd
    Global gBefoEditKey
Return

;*****************************************************************************************************************************************
; Function
;*****************************************************************************************************************************************

showGuiPasswordLibrary() {
    Gui, PL_LIV:Default
    Gui, PL_LIV:ListView, gPLv
    LV_Delete()
    IniRead, iPasss, %INI_PassLib%, PASS
    Loop, Parse, iPasss, `n
    {
        If RegExMatch(A_LoopField, "^(.*)=(.*)$", found)
        {
            LV_Add(, found1, found2)
        }
    }
    LV_ModifyCol()
    LV_Modify(1, "+Select")
    GuiControl, PL_LIV:Text, PSearch
    GuiControl, PL_LIV:Focus, PSearch
    Gui, PL_LIV:Show, , Passward Library
    WinSet, AlwaysOnTop, On, Passward Library
}

filteredHotStr(key) {
    Gui, PL_LIV:Default
    Gui, PL_LIV:ListView, gPLv
    LV_Delete()
    If (key == "") {
        key := ".*?"
    } Else {
        keyAry := StrSplit(key)
        key := ".*?"
        Loop % keyAry.MaxIndex()
        {
            key .= keyAry[A_Index] . ".*?"
        }
    }
    Loop, Parse, iPasss, `n
    {
        If (RegExMatch(A_LoopField, "^(" key ")=(.*)$", found)) {
            LV_Add("", found1, found2)
        }
    }
    LV_ModifyCol()
    LV_Modify(1, "+Select")
}

sendFunction() {
    Gui, PL_LIV:Default
    Gui, PL_LIV:Cancel
    Gui, PL_LIV:ListView, gPLv
    SelectedRow := LV_GetNext()
    If (SelectedRow != 0) {
        LV_GetText(RowText, SelectedRow, 2)
        Send, % RowText
    }
    ExitApp
}

;*****************************************************************************************************************************************
; Label
;*****************************************************************************************************************************************

PBSend:
    sendFunction()
Return

PBAdd:
    GuiControl, PL_REG:Show, gPBOkAdd
    GuiControl, PL_REG:Hide, gPBOkUpd
    Gui, PL_REG:Show, , Register Passward
    WinSet, AlwaysOnTop, On, Register Passward
Return

PBUpdate:
    Gui, PL_LIV:Default
    Gui, PL_LIV:ListView, gPLv
    LV_GetText(vKey, LV_GetNext(), 1)
    LV_GetText(vValue, LV_GetNext(), 2)
    Gui, PL_REG:Default
    gBefoEditKey := vKey
    GuiControl, PL_REG:Text, gEditKey, %vKey%
    GuiControl, PL_REG:Text, gEditValue, %vValue%
    GuiControl, PL_REG:Hide, gPBOkAdd
    GuiControl, PL_REG:Show, gPBOkUpd
    Gui, PL_REG:Show, , Register Passward
    WinSet, AlwaysOnTop, On, Register Passward
Return

PBDelete:
    Gui, PL_LIV:Default
    Gui, PL_LIV:ListView, gPLv
    LV_GetText(vKey, LV_GetNext())
    IniDelete, %INI_PassLib%, PASS, %vKey%
    IniRead, iPasss, %INI_PassLib%, PASS
    LV_Delete(LV_GetNext())
    LV_ModifyCol()
Return

PBSetfile:
    Run, %INI_PassLib%
Return

PSearch:
    Gui, PL_LIV:Submit, NoHide
    filteredHotStr(PSearch)
Return

PBOkAdd:
    Gui, PL_REG:Submit, NoHide
    Gui, PL_REG:Hide
    vValue := gEditKey . "=" . gEditValue
    If RegExMatch(vValue, "^(.+)=(.+)$")
    {
        Gui, PL_LIV:Default
        Gui, PL_LIV:ListView, gPLv
        LV_Add(, gEditKey, gEditValue)
        LV_ModifyCol()
        IniWrite, %gEditValue%, %INI_PassLib%, PASS, %gEditKey%
        IniRead, iPasss, %INI_PassLib%, PASS
    }
Return

PBOkUpd:
    Gui, PL_REG:Submit, NoHide
    Gui, PL_REG:Hide
    vValue := gEditKey . "=" . gEditValue
    If RegExMatch(vValue, "^(.+)=(.+)$")
    {
        Gui, PL_LIV:Default
        Gui, PL_LIV:ListView, gPLv
        vRowNum := LV_GetNext()
        LV_Modify(vRowNum, , gEditKey, gEditValue)
        LV_ModifyCol()
        IniDelete, %INI_PassLib%, PASS, %gBefoEditKey%
        IniWrite, %gEditValue%, %INI_PassLib%, PASS, %gEditKey%
        IniRead, iPasss, %INI_PassLib%, PASS
    }
Return

;*****************************************************************************************************************************************
; HotKey
;*****************************************************************************************************************************************

#IfWinActive, Passward Library

    *Tab::
        sendFunction()
    Return

#IfWinActive