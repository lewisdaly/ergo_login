#SingleInstance Force

; Version: 9
; Revision: $Date: 2017/02/19 21:08:57 $
;
; Change Log:
; 2017/02/19 21:08:57 adapt to new ZwiftLauncher version (window name 'Zwift Launcher' in 1.0.39 instead of 'MainWindow' in 1.0.19)
; 2016/12/10 10:50:16 adapt to new zwift launcher screen
; 2016/12/03 20:07:19 added call of taskkill zwiftapp.exe after close of zwiftlauncher. Always close both before launching
; 2016/10/20 13:18:30 Set font size
; 2016/10/19 15:01:44 Always Force close zwiftlauncher.exe on start
; 2016/10/14 13:18:46 option to not launch zwift-hotkeys
; 2016/09/19 15:39:36 support for compiled scripts (instead of .ahk)
; 2016/03/01 12:04:22 Button to close zwiftlauncher.exe (useful in endless
; loop after an update of Zwift)
;
; AutoHotkey Version: 1.x
; Language:       English
; Platform:       Win
; Author:         Jesper Rosenlund Nielsen <jesper@rosenlundnielsen.net>
;
; Script Function:
;	Zwift launcher
;

#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

; -- Command Line Parsing

; defaults
ParamOnlyZwift := FALSE

; parse loop
Loop, %0%  ; For each parameter:
{
    param := %A_Index%  ; Fetch the contents of the variable whose name is contained in A_Index.
    if ( param = "/onlyzwift" ) {
		ParamOnlyZwift := TRUE
	}
}


; -- Configuration

IniFile = zwift-login.ini
HotkeyScript = zwift-hotkeys.ahk
HotkeyScriptExe = zwift-hotkeys.exe
PreferencesScript = zwift-preferences.ahk
PreferencesScriptExe = zwift-preferences.exe


; -- Globals

GlobalEml =
GlobalPwd =

; --

LoadFromIni()

Gui, New
Gui, +AlwaysOnTop

Gui, Font, s10


If (A_IsCompiled) {
	; If this script itself is compiled then use only compiled versions of other scripts
	If FileExist(PreferencesScriptExe) {
	  Gui, Add, Button,, Preferences
	}
} else {
	If ( FileExist(PreferencesScript) OR FileExist(PreferencesScriptExe) ){
	  Gui, Add, Button,, Preferences
	}
}

;TODO: load the user file from the dir
Gui, Add, Button,

Gui, Add, Button,  Default, Launch Zwift
Gui, Add, Button, yp xp+80, Restart ZwiftLauncher
Gui, Add, Button,xs, Store Password
Gui, Add, Text,, Email
Gui, Add, Edit, vZEmail w200
Gui, Add, Text,, Password
Gui, Add, Edit, vZPassword
Gui, Add, Button,, Save
Gui, Add, Button,, Do Not Save


Gui, Add, StatusBar
Cnt := 0

SetGuiInitialState()

Gui, Show

Return

; End of main routine


; ==========
; Button Handlers

;------

ButtonLaunchZwift:

;SB_SetText("Closes old launcher...")
;RunWait, % CloseLauncherFile(),  % ZwiftLauncherDir()

/* 2016-12-10
Run, %comspec% /c "taskkill /F /IM ZwiftLauncher.exe /T", , Hide
Run, %comspec% /c "taskkill /F /IM ZwiftApp.exe /T", , Hide
*/

SB_SetText("Launching...")


if (NOT ParamOnlyZwift) {
	If (A_IsCompiled) {
		; If this script itself is compiled then use only compiled versions of other scripts
		If FileExist(HotkeyScriptExe) {
		  Run, %HotkeyScriptExe% /dontlaunchlogin
		}
	} else {
		; .ahk script takes precedence over compiled scripts
		If FileExist(HotkeyScript) {
		  Run, %HotkeyScript% /dontlaunchlogin
		} else if FileExist(HotkeyScriptExe) {
		  Run, %HotkeyScriptExe% /dontlaunchlogin
		}
	}
}

;SetWorkingDir, % ZwiftLauncherDir()

;;zlf := zwiftlauncherfile()
;;msgbox , % zlf
Run, % ZwiftLauncherFile(),  % ZwiftLauncherDir()
; Run, % ZwiftLauncherDir() "\ZwiftApp.exe"
; SetWorkingDir %A_ScriptDir%

SB_SetText("Waiting for login screen...")

bLoginScreenLoaded := FALSE
cntOuterLoop := 0

While ( NOT bLoginScreenLoaded ) {

  SetTitleMatchMode RegEx
  ; T := "MainWindow ahk_class HwndWrapper.*ZwiftLauncher.exe"
  ; Fix to support both old 1.0.19 and new 1.0.39 version:
  T := "(MainWindow|Zwift Launcher) ahk_class HwndWrapper.*ZwiftLauncher.exe"

  WinWait, % T
  WinGet, active_id, ID, % T
  WinActivate, A

  ; WB := WBGet(T)
  WB := WBGet("ahk_id " active_id)

  SB_SetText("Login screen is loading ...")
  Sleep, 10

  cnt := 0

  ; wait for the page to load
  While ( wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ) && ( cnt < 100 ) {
   ; DEBUG ONLY ;
   ;SB_SetText("Login screen is loading ... " WB.document.readyState)
   SB_SetText("Login screen is loading ...... ")
   ; tooltip, % wb.document.readyState
   Sleep, 10
   cnt += 1
  }

  cntOuterLoop += 1

  If ( wb.readyState = 4 ) or ( wb.document.readyState = "complete" )  {
    bLoginScreenLoaded := TRUE
  }

  if ( cntOuterLoop > 20 AND NOT bLoginScreenLoaded ) {
    SB_SetText("You may have to restart ZwiftLauncher...")
    GuiControl, Enable, Restart ZwiftLauncher
    GuiControl, Show, Restart ZwiftLauncher
    GuiControl, Focus, Restart ZwiftLauncher
    Sleep, 2000
    cntOuterLoop := 0
  }

}

SB_SetText("Filling fields...")

 s := WB.document.documentElement.innerHTML
; Original login screen
; s := WB.document.getElementById("kc-login").value
; Old screen when user is remembered
; s := trim(WB.document.getElementById("kc-confirmuser").innerHTML)
; New (2016-12-10) screen when user is remembered
; s:= SubStr(trim(WB.document.getElementById("kc-watch").innerHTML, " `t`n`r"),1,4)
; listvars

If ( trim(WB.document.getElementById("kc-login").value) = "Log in" ) {
  ; Original login screen
  WB.document.getElementById("username").value := GlobalEml
  WB.document.getElementById("password").value := GlobalPwd

  SB_SetText("Pressing Log In button...")

  WB.document.getElementById("kc-login").click()

 ;; msgbox, NOW ON TO THE NEXT PAGE ....................................

  ; 2016-12-10 Added sleep, 1000 and wait for second screen to load
	sleep, 1000

  ; Get the next screen:

  WB := ""
  WB := WBGet("ahk_id " active_id)

 ; wait for the page to load
  While ( wb.readyState != 4 || wb.document.readyState != "complete" || wb.busy ) {
   SB_SetText("Welcome screen is loading ...... ")
   tooltip, % wb.document.readyState
   Sleep, 10

  }
  sleep, 1000



} else {
; SHOULD NEVER GET HERE ANYMORE !!  2016-12-10

  if ( trim(WB.document.getElementById("kc-confirmuser").innerHTML) <> "" ) {
    ; Old Screen when user is remembered

  SB_SetText("Pressing the Log In button...")

  WB.document.getElementById("kc-confirmuser").click()

  }
}


If ( SubStr(trim(WB.document.getElementById("kc-watch").innerHTML, " `t`n`r"),1,4) = "Ride" ) {
	; 2016-12-10 New Welcome / Ride screen

	; msgbox, WAITING !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

	SB_SetText("Pressing Ride button...")

	try {
		Loop {
			sleep, 2000
			WB.document.getElementsByTagName("FORM")[0].submit()
			;WB.document.getElementById("kc-watch").click()
		}
	} catch e {
		; just carry on
	}


}




SB_SetText("Ride On...")
Sleep, 1000
ExitApp

Return

; ------

ButtonStorePassword:
GuiControl, Disable, Launch Zwift
GuiControl, Disable, Store Password
GuiControl, Show, ZEmail
GuiControl, Show, ZPassword
GuiControl, Show, Save
GuiControl, Show, Do Not Save
GuiControl, Focus, ZEmail
Return

; --------

ButtonLoadFile:
=FileSelectFile, SelectedFile, 3, , Open a file, User File (*.ini;)
if SelectedFile =
    MsgBox, The user didn't select anything.
else
    MsgBox, The user selected the following:`n%SelectedFile%

;TODO: parse this file, and load into memory

Return


ButtonSave:
Gui, Submit, NoHide
IniWrite, %ZEmail%, %IniFile%, Login, email
IniWrite, %ZPassword%, %IniFile%, Login, password
GuiControl, Enable, Launch Zwift
GuiControl, Enable, Store Password
GuiControl, Hide, ZEmail
GuiControl, Hide, ZPassword
GuiControl, Hide, Save
GuiControl, Hide, Do Not Save
LoadFromIni()
SetGuiInitialState()
Return

ButtonDoNotSave:
GuiControl, Enable, Launch Zwift
GuiControl, Enable, Store Password
GuiControl, Hide, ZEmail
GuiControl, Hide, ZPassword
GuiControl, Hide, Save
GuiControl, Hide, Do Not Save
LoadFromIni()
SetGuiInitialState()
Return

ButtonRestartZwiftLauncher:
Process, WaitClose, ZwiftLauncher.exe, 4
if ( ErrorLevel > 0 ) {
  Run, %comspec% /c "taskkill.exe /F /IM ZwiftLauncher.exe /T", , Hide
  Run, %comspec% /c "taskkill.exe /F /IM ZwiftApp.exe /T", , Hide
  ; RunWait, % CloseLauncherFile(),  % ZwiftLauncherDir()
  ; Gui, -AlwaysOnTop
  ; msgbox, Could not close ZwiftLauncher programmatically.`n`nYou probably have to close ZwiftLauncher in the system tray manually (right click and choose 'Exit')`n`nAfter that is done just launch Zwift again.
}
Reload
Return

ButtonPreferences:
If (A_IsCompiled) {
	; If this script itself is compiled then use only compiled versions of other scripts
	If FileExist(PreferencesScriptExe) {
	  Gui, -AlwaysOnTop
	  Run, %PreferencesScriptExe%
	}
} else {
	; .ahk script takes precedence over compiled scripts
	If FileExist(PreferencesScript) {
	  Gui, -AlwaysOnTop
	  Run, %PreferencesScript%
	} else if FileExist(PreferencesScriptExe) {
	  Gui, -AlwaysOnTop
	  Run, %PreferencesScriptExe%
	}
}
Return


; =======================

LoadFromIni()
{
  global IniFile
  global GlobalEml
  global GlobalPwd
  IniRead, GlobalEml, %IniFile%, Login, email
  IniRead, GlobalPwd, %IniFile%, Login, password
  If ( GlobalEml = "ERROR" ) or ( GlobalPwd = "ERROR" ) {
    GlobalEML =
    GlobalPwd =
  }
}

SetGuiInitialState()
{
  global GlobalEml
  global GlobalPwd

  GuiControl, , ZEmail, %GlobalEml%
  GuiControl, , ZPassword, %GlobalPwd%

  GuiControl, Hide, Email
  GuiControl, Hide, ZEmail
  GuiControl, Hide, Password
  GuiControl, Hide, ZPassword
  GuiControl, Hide, Save
  GuiControl, Hide, Do Not Save

  GuiControl, Enable, Store Password

  If ( GlobalEml = "" ) or ( GlobalPwd = "" ) {
    GuiControl, Disable, Launch Zwift
    GuiControl, Focus, Store Password
  } else {
    GuiControl, Enable, Launch Zwift
    GuiControl, Focus, Launch Zwift
  }

  GuiControl, Disable, Restart ZwiftLauncher
  GuiControl, Hide, Restart ZwiftLauncher

  SB_SetText("")

}



CloseLauncherFile()
{
  EnvGet, pf, ProgramFiles(x86)
  Return pf "\Zwift\CloseLauncher.exe"
}


ZwiftLauncherFile()
{
  EnvGet, pf, ProgramFiles(x86)
  Return pf "\Zwift\ZwiftLauncher.exe"
}

ZwiftLauncherDir()
{
  EnvGet, pf, ProgramFiles(x86)
  Return pf "\Zwift"
}


WBGet(WinTitle="ahk_class IEFrame", Svr#=1) { ; based on ComObjQuery docs
	static	msg := DllCall("RegisterWindowMessage", "str", "WM_HTML_GETOBJECT")
	,	IID := "{0002DF05-0000-0000-C000-000000000046}" ; IID_IWebBrowserApp
;	,	IID := "{332C4427-26CB-11D0-B483-00C04FD90119}" ; IID_IHTMLWindow2

	SendMessage msg, 0, 0, Internet Explorer_Server%Svr#%, %WinTitle%
	if (ErrorLevel != "FAIL") {
		lResult:=ErrorLevel, VarSetCapacity(GUID,16,0)
		if DllCall("ole32\CLSIDFromString", "wstr","{332C4425-26CB-11D0-B483-00C04FD90119}", "ptr",&GUID) >= 0 {
			DllCall("oleacc\ObjectFromLresult", "ptr",lResult, "ptr",&GUID, "ptr",0, "ptr*",pdoc)
			return ComObj(9,ComObjQuery(pdoc,IID,IID),1), ObjRelease(pdoc)
		}
	}
}
