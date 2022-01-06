/* Setup environment for optimized functioning
*/
#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
#InstallMouseHook ; Get physical mouse events
#InstallKeybdHook ; Get physical key events
#WinActivateForce ; Fore activate window
#SingleInstance force ; Skips the dialog box and replaces the old instance automatically
#MaxHotkeysPerInterval 300 ; maximize 300 hotkeys per 2s
SetWorkingDir %A_ScriptDir% ; Ensures a consistent starting directory.
SendMode Input ; Recommended for new scripts due to its superior speed and reliability.
SetBatchLines -1 ; The script run at maximum speed
SetMouseDelay, -1 ; No delay for mouse movement and click
SetDefaultMouseSpeed, 0 ; Move the mouse instantly
SetControlDelay, -1 ; No delay after each control-modifying command
SetKeyDelay, -1 ; No delay after each keystroke sent by Send and ControlSend
CoordMode, Mouse, Client ; independence from screen resolution
CoordMode, ToolTip, Screen
SetTitleMatchMode, 2   ; Detect a window's title can contain WinTitle anywhere inside it to be a match.
DetectHiddenWindows, On
SetNumLockState, AlwaysOn

Menu, TRAY, Icon, system_green.png, , 1




/* Variables & Options
-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
*/

; remove all traced .torrent files
FileDelete, %A_AppData%\uTorrent\*.torrent

; clean-up acestream cache
FileDelete, C:\_acestream_cache_\*

; reset SyncBackPro
; f__run_reset_syncbackpro()

; Brightness Variables
brightness_increment := 10 ; < lower for a more granular change, higher for larger jump in brightness 
current_brightness := f__get_current_brightness()




; * *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** ***
Return  ; At startup, only scripts above this will be processed *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** ***
; * *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** *** ***





/* Hotkeys
-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
*/

; Brightness Hot Keys
; F5::
;     f__change_brightness( current_brightness -= brightness_increment )
; return

; F6::
;     f__change_brightness( current_brightness += brightness_increment )
; return

; +F10::
;     if WinExist("ahk_exe Audials.exe") {
;         f__close_audials()
;     } else {
;         f__open_audials()
;     }
; Return

F11::
    f__run_csgo_helper_play()
return

; and WinExist("ahk_exe RazerCortex.exe")
+F11::
    if (WinExist("ahk_exe steam.exe")) {
        if (!is_qc35_connected()) {
            f__mute_system_volume()

            MsgBox, 4,, Bose QC35 not connected. Continue?
            IfMsgBox No
                return
        }
        f__play_csgo_game()
    } else {
    	f__open_steam()
        ; if (!WinExist("ahk_exe steam.exe")) {
        ;     f__open_steam()
        ; }
        ; if (!WinExist("ahk_exe RazerCortex.exe")) {
        ;     f__open_razer_cortex()
        ; }
    }
return

F12::
    f__close_csgo_helper()

    ; Mute system volume if Bose QC35 not connected
    if (!is_qc35_connected()) {
        f__mute_system_volume()
    }
return

+F12::
    f__quit_csgo_game()
return

!s::
    f__save_documents()
Return

^+Space::
    f__toggle_unikey()
return

; Volume_Mute::
;     f__mute_system_volume()
; Return

#Esc::
    f__reload()
return

NumpadDot::
	Send {.}
return

#Space::
    ; do nothing -> disable keyboard layout switching
return





/* Functions
-----------------------------------------------------------------------------------------------------------------------------------------------------------------------
*/

f__open_audials() {
    Run, "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Audials.lnk"
    SetTimer, close_audials_updater, 200
}

f__close_audials() {
    WinClose, ahk_exe Audials.exe
    SetTimer, close_audials_updater, 200
}

close_audials_updater:
    f__close_audials_updater()
Return

f__close_audials_updater() {
    if WinExist("Audials Update Center") {
        WinClose
        SetTimer, close_audials_updater, Off
    }
}

; f__run_reset_syncbackpro() {
;     Run, "D:\System\reset_SyncBack_Pro.cmd"
; }

f__toggle_unikey() {
    if WinExist("ahk_exe UniKeyNT.exe") {
        WinClose, ahk_exe UniKeyNT.exe
        WinClose, ahk_exe UniKeyNT.exe
    } else Run, "C:\Users\Kien Dang\Documents\system\uniKey\UniKeyNT.exe"
}

f__close_unikey() {
    if WinExist("ahk_exe UniKeyNT.exe") {
        WinClose, ahk_exe UniKeyNT.exe
        WinClose, ahk_exe UniKeyNT.exe
    }
}

; f__run_csgo_helper_release() {
;     Run, "D:\Devs\projects\csgo_helper\_release\v0.3.8.5\csgo_helper.exe"
; }

; f__run_csgo_helper_dev() {
;     Run, "D:\Devs\projects\csgo_helper\csgo_helper.ahk"
; }

f__run_csgo_helper_play() {
    Run, "C:\Users\Kien Dang\Documents\system\csgo_helper\csgo_helper.ahk"
}

f__close_csgo_helper() {
    ; close csgo_helper
    WinClose, csgo_helper.ahk ahk_class AutoHotkey
    WinClose, csgo_helper.exe
    ; bome timer
    WinClose, bomb_timer.ahk ahk_class AutoHotkey
    WinClose, bomb_timer.exe
    ; stop gsi_server
    WinClose, csgo_helper_server ahk_class ConsoleWindowClass
    WinClose, csgo_helper_server.exe
    ; remove ToolTip
    ToolTip
}

; f__open_razer_cortex() {
;     Run, "C:\Program Files (x86)\Razer\Razer Cortex\CortexLauncher.exe"
; }

f__open_steam() {
    Run, "C:\Program Files (x86)\Steam\steam.exe"
}

f__play_csgo_game() {
	f__run_csgo_helper_play()
    ; global csgo_helper_play_script
    ; run CSGO
    Run, "C:\Users\Kien Dang\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Steam\Counter-Strike Global Offensive.url"
    ; Razer Cotex boosting
    ; Send, {Ctrl Down}{Alt Down}{b Down}{b Up}{Alt Up}{Ctrl Up}
    ; run csgo_helper
    ; Switch csgo_helper_play_script
    ; {
    ; Case "dev":
    ;     f__run_csgo_helper_dev()
    ; Case "release":
    ;     f__run_csgo_helper_release()
    ; Case "play":
    ;     f__run_csgo_helper_play()
    ; }

    ; Close UniKeyNT
    f__close_unikey()

    ; Set Max Screen Brightness
    global current_brightness
    current_brightness := f__get_current_brightness()
    f__change_brightness(100)

    ; Mute system volume if Bose QC35 not connected
    if (!is_qc35_connected()) {
        f__mute_system_volume()
    }
}

f__quit_csgo_game() {
    ; close CSGO
    Process, Close, csgo.exe
    ; restore Razer Cotex boosting
    ; Send, {Ctrl Down}{Alt Down}{r Down}{r Up}{Alt Up}{Ctrl Up}
    ; close csgo_helper
    f__close_csgo_helper()
    ; Restore Brightness
    global current_brightness
    f__change_brightness(current_brightness)
}

is_qc35_connected() {
    RegRead, qc35_state, HKEY_LOCAL_MACHINE, SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\Capture\{39b5abda-5764-455f-a3ee-2a054abff9f4}, DeviceState
    if (qc35_state == 1) {
        return True
    } else {
        return False
    }
}

f__mute_system_volume() {
    SoundGet, master_volume
    ;
    if (master_volume) {
       SoundSet, 0
       SoundSet, 1, , mute
    }
}

f__save_documents() {
    ; save documents: words, excel
    if WinActive("ahk_exe WINWORD.EXE") || WinActive("ahk_exe EXCEL.EXE") {
        ; save Docs
        Send, {Ctrl Down}{s Down}{s Up}{Ctrl Up}
        Sleep, 5000
        ; run BackSync
        Send, {Ctrl Down}{F1 Down}{F1 Up}{Ctrl Up}
        Sleep, 1000
        Send, {Ctrl Down}{F2 Down}{F2 Up}{Ctrl Up}
    } else {
        ; simplely saving
        Send, {Ctrl Down}{s Down}{s Up}{Ctrl Up}
    }
}

f__list_all_processes() {
    global
    local strComputer, objWMIService, colProcesses
    strComputer := "."
    objWMIService := ComObjGet("winmgmts:\\" . strComputer . "\root\cimv2")
    colProcesses := objWMIService.ExecQuery("Select * From Win32_Process")._NewEnum
    Gui, Add, ListView, x2 y0 w500 h700 vMyLV,Process Name|WorkingSetSize|Priority|ThreadCount
    GuiControl, -Redraw, MyLV 
    While colProcesses[objProcess]
        LV_Add("",objProcess.Name , objProcess.WorkingSetSize , objProcess.Priority , objProcess.ThreadCount)
    GuiControl, +Redraw, MyLV 
    LV_ModifyCol(1,160)
    Gui, Show, w500 h700, Process List
}

f__toggle_suspend() {
    Suspend
    If (A_IsSuspended) {
        Menu, TRAY, Icon, system_red.png, , 1
        Tooltip, System - OFF
    } else {
        Menu, TRAY, Icon, system_green.png, , 1
        Tooltip, System - ON
    }
    Sleep, 1000
    ToolTip
}

f__reload() {
    Reload
}

; Brightness
f__change_brightness( ByRef brightness := 50, timeout = 1 ) {
    if ( brightness > 100 ) {
        brightness := 100
    } else if ( brightness < 0 ) {
        brightness := 0
    }

    For property in ComObjGet( "winmgmts:\\.\root\WMI" ).ExecQuery( "SELECT * FROM WmiMonitorBrightnessMethods" )
        property.WmiSetBrightness( timeout, brightness )    
}

f__get_current_brightness() {
    For property in ComObjGet( "winmgmts:\\.\root\WMI" ).ExecQuery( "SELECT * FROM WmiMonitorBrightness" )
        currentBrightness := property.CurrentBrightness 
    return currentBrightness
}

/*
~+::RapidHotkey("Plus")
~h::RapidHotkey("{Raw}Hello World!", 3) ;Press h 3 times rapidly to send Hello World!
~o::RapidHotkey("^o", 4, 0.2) ;be careful, if you use this hotkey, above will not work properly
~Esc::RapidHotkey("exit", 4, 0.2, 1) ;Press Esc 4 times rapidly to exit this script
~LControl::RapidHotkey("!{TAB}",2) ;Press LControl rapidly twice to AltTab
~RControl::RapidHotkey("+!{TAB}",2) ;Press RControl rapidly twice to ShiftAltTab
~LShift::RapidHotkey("^{TAB}", 2) ;Switch back in internal windows
~RShift::RapidHotkey("^+{TAB}", 2) ;Switch between internal windows
~e::RapidHotkey("#e""#r",3) ;Run Windows Explorer
~^!7::RapidHotkey("{{}{}}{Left}", 2)

~a::RapidHotkey("test_label", 2, 0.3, 1) ;You Can also specify a Label to be launched
test_label:
MsgBox, Test
Return
*/

RapidHotkey(keystroke, times="2", delay=0.2, IsLabel=0) {
    Pattern := Morse(delay*1000)
    If (StrLen(Pattern) < 2 and Chr(Asc(times)) != "1")
        Return
    If (times = "" and InStr(keystroke, """"))
    {
        Loop, Parse, keystroke,""   
            If (StrLen(Pattern) = A_Index+1)
                continue := A_Index, times := StrLen(Pattern)
    }
    Else if (RegExMatch(times, "^\d+$") and InStr(keystroke, """"))
    {
        Loop, Parse, keystroke,""
            If (StrLen(Pattern) = A_Index+times-1)
                times := StrLen(Pattern), continue := A_Index
    }
    Else if InStr(times, """")
    {
        Loop, Parse, times,""
            If (StrLen(Pattern) = A_LoopField)
                continue := A_Index, times := A_LoopField
    }
    Else if (times = "")
        continue := 1, times := 2
    Else if (times = StrLen(Pattern))
        continue = 1
    If !continue
        Return
    Loop, Parse, keystroke,""
        If (continue = A_Index)
            keystr := A_LoopField
    Loop, Parse, IsLabel,""
        If (continue = A_Index)
            IsLabel := A_LoopField
    hotkey := RegExReplace(A_ThisHotkey, "[\*\~\$\#\+\!\^]")
    IfInString, hotkey, %A_Space%
        StringTrimLeft, hotkey,hotkey,% InStr(hotkey,A_Space,1,0)
    backspace := "{BS " times "}"
    keywait = Ctrl|Alt|Shift|LWin|RWin
    Loop, Parse, keywait, |
        KeyWait, %A_LoopField%
    If ((!IsLabel or (IsLabel and IsLabel(keystr))) and InStr(A_ThisHotkey, "~") and !RegExMatch(A_ThisHotkey
    , "i)\^[^\!\d]|![^\d]|#|Control|Ctrl|LCtrl|RCtrl|Shift|RShift|LShift|RWin|LWin|Alt|LAlt|RAlt|Escape|BackSpace|F\d\d?|"
    . "Insert|Esc|Escape|BS|Delete|Home|End|PgDn|PgUp|Up|Down|Left|Right|ScrollLock|CapsLock|NumLock|AppsKey|"
    . "PrintScreen|CtrlDown|Pause|Break|Help|Sleep|Browser_Back|Browser_Forward|Browser_Refresh|Browser_Stop|"
    . "Browser_Search|Browser_Favorites|Browser_Home|Volume_Mute|Volume_Down|Volume_Up|MButton|RButton|LButton|"
    . "Media_Next|Media_Prev|Media_Stop|Media_Play_Pause|Launch_Mail|Launch_Media|Launch_App1|Launch_App2"))
        Send % backspace
    If (WinExist("AHK_class #32768") and hotkey = "RButton")
        WinClose, AHK_class #32768
    If !IsLabel
        Send % keystr
    else if IsLabel(keystr)
        Gosub, %keystr%
    Return
}

; Laszo -> http://www.autohotkey.com/forum/viewtopic.php?t=16951 (Modified to return: KeyWait %key%, T%tout%)
Morse(timeout = 400) {
   tout := timeout/1000
   key := RegExReplace(A_ThisHotKey,"[\*\~\$\#\+\!\^]")
   IfInString, key, %A_Space%
        StringTrimLeft, key, key,% InStr(key,A_Space,1,0)
    If Key in Shift,Win,Ctrl,Alt
        key1:="{L" key "}{R" key "}"
   Loop {
      t := A_TickCount
      KeyWait %key%, T%tout%
        Pattern .= A_TickCount-t > timeout
        If(ErrorLevel)
            Return Pattern
    If key in Capslock,LButton,RButton,MButton,ScrollLock,CapsLock,NumLock
      KeyWait,%key%,T%tout% D
    else if Asc(A_ThisHotkey)=36
        KeyWait,%key%,T%tout% D
    else
      Input,pressed,T%tout% L1 V,{%key%}%key1%
    If (ErrorLevel="Timeout" or ErrorLevel=1)
        Return Pattern
    else if (ErrorLevel="Max")
        Return
   }
}
