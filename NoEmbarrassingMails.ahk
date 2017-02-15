/*
  NoEmbarrassingMails.ahk
  
  Author:	Theoky
  Linzenz:	GPL 3.0
  ----------------------------------------------------------
  ChangeLog : 
     v0.0 initial version
	 v0.1 better adapt to changes in recipients during window in focus,
		  better customisation
  ----------------------------------------------------------
  Purpose:
	Check an open e-mail-Window for email recipients outside your company,
	thus preventing embarassing situationd when you send an email intended for
	internal recipients accidentally to external recipients.
	
	Currently works only when composing messages and meetings with Outlook and
	Exchange.
	
  Usage: 
	* install autohotkey (https://autohotkey.com/)
	* change MailDomain to your Company Name
	* Start this script
	* In case of undesired behaviour -> fix and submit patch ;)
*/

#SingleInstance force
#NoEnv
SetBatchLines, -1
ListLines, Off

if (A_AhkVersion < "1.1.23.00") {
    MsgBox, This script is not tested with your AutoHotkey version (%A_AhkVersion%).
	return
}

; Settings
	global TransN                := 150      ; 0~255
	global CheckTime			 := 1500     ; In milliseconds
	global WaitTime				 := 1.5
	global GuiPosition           := "Bottom" ; Top or Bottom
	global FontSize              := 16
	global GuiHeight             := 50
	global ExternalWarning		 := "External address!"
	
	global MailDomain			 := "i)/o=Your Company" ; <- customise this
	#Include *i Customisation.ahk
	
	global CheckOutlookCOM		 := true

CreateGUI()
SetTitleMatchMode, 2

#Persistent
SetTimer, WatchForEmail, %CheckTime%
return

isInternalMail(address) {
	if (StrLen(address) < 1) {
		return true
	}
	pos := RegExMatch(address, MailDomain)
	if (pos > 0) {
		return true
	}
	return false
}

WatchForEmail:

	warning := false

	if (CheckOutlookCom) {
		GroupAdd mail, Message
		GroupAdd mail, Meeting
		
		WinWaitActive, ahk_group mail
		WinGet, outlookWindow, ID, A
		WinGetTitle, title, A
		WinGet, pName, ProcessName, A
		WinGetClass, outlookWinClass, ahk_id %outlookWindow%
		
		; ahk_class rctrl_renwnd32 for outlook window (not for dialogs)
		if (pName == "OUTLOOK.EXE" and outlookWinClass == "rctrl_renwnd32" ) {
			try {
				ol := ComObjActive("Outlook.Application")
				
				try {
					rbn := ol.ActiveInspector.CurrentItem.ReceivedByName
					
					if (rbn <> "")
					{
						; ReceivedByName not empty - received mail, no check
						return
					}
				} catch {
					warning := false
				}
					
				Loop, % ol.ActiveInspector.CurrentItem.Recipients.Count
				{
					address := ol.ActiveInspector.CurrentItem.Recipients(A_Index).Address
					if (not isInternalMail(address))
					{
						warning := true
						break
					}
				}
			} catch {
				warning := false
			}
		}
		if (warning) {
			ShowWarning(ExternalWarning)
		}
		else {
			HideWarning()
		}
	}
	return

; ===================================================================================
CreateGUI() {
	global

	Gui, +AlwaysOnTop -Caption +Owner +LastFound +E0x20
	Gui, Margin, 0, 0
	Gui, Color,  e69900
	Gui, Font, cWhite s%FontSize% bold, Arial
	Gui, Add, Text, vIntMailText Center y20

	WinSet, Transparent, %TransN%
}

ShowWarning(warning) {
	WinGetPos, ActWin_X, ActWin_Y, ActWin_W, ActWin_H, A
	if !ActWin_W
		throw

	text_w := (ActWin_W > A_ScreenWidth) ? A_ScreenWidth : ActWin_W
	GuiControl,     , IntMailText, %warning%
	GuiControl, Move, IntMailText, w%text_w% Center

	if (GuiPosition = "Top")
		gui_y := ActWin_Y
	else
		gui_y := (ActWin_Y+ActWin_H) - GuiHeight - FontSize

	Gui, Show, NoActivate x%ActWin_X% y%gui_y% h%GuiHeight% w%text_w%
}


HideWarning() {
	Gui, Hide
}
