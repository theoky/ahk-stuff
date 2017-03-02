/*
  NoEmbarrassingMails.ahk
  
  Author:	Theoky
  Linzenz:	GPL 3.0
  ----------------------------------------------------------
  ChangeLog : 
     v0.0 initial version
	 v0.1 better adapt to changes in recipients during window in focus,
		  better customisation
	 v0.2 offline outlook handling, better focus management
	 v0.3 even more improved outlook handling (distribution lists)
	 v0.4 better handling of multiple Outlook windows
	 
  ----------------------------------------------------------
  Purpose:
	Check an open e-mail-Window for email recipients outside your company,
	thus preventing embarassing situations when you send an email intended for
	internal recipients accidentally to external recipients.
	
	Currently works only when composing messages and meetings with Outlook and
	Exchange.

	This script is for those cases when MailTips (since Outlook 2010) are not available
	or deactivated for the external recipients tip.	

  Usage: 
	* install autohotkey (https://autohotkey.com/)
	* copy customisation_template.ahk to customisation.ahk
	* change MailDomain to your Company Name in customisation
	* Start this script
	* In case of undesired behaviour -> fix it and submit patch ;)
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
	global warnTextExternal		 := "At least one recipient is external!"
	global warnTextUnknownAddr	 := "Address lookup failed for at least one recipient - Outlook offline?"
	
	global WarningNo		:= 0
	global WarningUnknown	:= 1
	global WarningExternal	:= 2
	
	global MailDomain			 := "i)/o=Your Company" ; <- customise this in Customisation.ahk
	#Include *i Customisation.ahk
	
	global CheckOutlookCOM		 := true

	; from https://msdn.microsoft.com/en-us/library/office/ff868214.aspx
	olExchangeAgentAddressEntry  = 3	; An address entry that is an Exchange agent.
	olExchangeDistributionListAddressEntry  = 1	; An address entry that is an Exchange distribution list.
	olExchangeOrganizationAddressEntry  = 4	; An address entry that is an Exchange organization.
	olExchangePublicFolderAddressEntry  = 2	; An address entry that is an Exchange public folder.
	olExchangeRemoteUserAddressEntry  = 5	; An Exchange user that belongs to a different Exchange forest.
	olExchangeUserAddressEntry  = 0	; An Exchange user that belongs to the same Exchange forest.
	olLdapAddressEntry  = 20	; An address entry that uses the Lightweight Directory Access Protocol (LDAP).
	olOtherAddressEntry  = 40	; A custom or some other type of address entry such as FAX.
	olOutlookContactAddressEntry  = 10	; An address entry in an Outlook Contacts folder.
	olOutlookDistributionListAddressEntry  = 11	; An address entry that is an Outlook distribution list.
	olSmtpAddressEntry  = 30	; An address entry that uses the Simple Mail Transfer Protocol (SMTP).

CreateGUI()
SetTitleMatchMode, 2

#Persistent
SetTimer, WatchForEmail, %CheckTime%
return

WatchForEmail:

	warning := WarningNo

	if (CheckOutlookCOM) {
		GroupAdd mail, Message
		GroupAdd mail, Meeting
		
		foundWindow := false
		while (not foundWindow) {
			WinWaitActive, ahk_group mail, , %WaitTime%
			if not ErrorLevel {
				foundWindow := true
			} else {
				HideWarning()
			}
		}

		WinGet, outlookWindow, ID, A
		WinGetTitle, title, ahk_id %outlookWindow%
		WinGet, pName, ProcessName, ahk_id %outlookWindow%
		WinGetClass, outlookWinClass, ahk_id %outlookWindow%
		
		warning := WarningNo
		
		; ahk_class rctrl_renwnd32 for outlook window (not for dialogs)
		if (pName == "OUTLOOK.EXE" and outlookWinClass == "rctrl_renwnd32" ) {
			try {
				ol := ComObjActive("Outlook.Application")
				
				rbn := ol.ActiveInspector.CurrentItem.ReceivedByName
				if (rbn <> "") {
					; ReceivedByName not empty - received mail, no check
					throw
				}
					
				Loop, % ol.ActiveInspector.CurrentItem.Recipients.Count {
					rec := ol.ActiveInspector.CurrentItem.Recipients(A_Index)
					addrEntry := rec.AddressEntry
					
					; OutputDebug, % 
					if (addrEntry.AddressEntryUserType == olExchangeDistributionListAddressEntry) {
						exchDL := addrEntry.GetExchangeDistributionList()
						addrEntries := exchDL.GetExchangeDistributionListMembers()
						if (addrEntries) {
							Loop, % addrEntries.Count {
								addrEntry := addrEntries.Item(A_Index)
								warning := checkAddressEntry(addrEntry, warning)
								
								if (warning == WarningExternal) {
									break
								}
								
;								OutputDebug, % addrEntries.Item(A_Index).Name
							}
						}
						
					} else if (addrEntry.AddressEntryUserType == olExchangeOrganizationAddressEntry
							or addrEntry.AddressEntryUserType == olOutlookContactAddressEntry
							or addrEntry.AddressEntryUserType == olExchangeRemoteUserAddressEntry) 
					{
						warning := checkAddressEntry(addrEntry, warning)
						if (warning == WarningExternal) {
							break
						}

					} else if (addrEntry.AddressEntryUserType == olSmtpAddressEntry) {
						address := rec.Address
						if (address and not isInternalMail(address)) {
							warning := WarningExternal
							break
						}
					} 
				}
			} catch {
				warning := WarningNo
			}
			
			if (warning == WarningExternal) {
				ShowWarning(warnTextExternal, outlookWindow)
				
			} else if (warning == WarningUnknown) {
				ShowWarning(warnTextUnknownAddr, outlookWindow)
			}
			else {
				HideWarning()
			}
		}
	}
	return

checkAddressEntry(addrEntry, curWarningLevel)
{
	warning := WarningNo
	exchangeUser := addrEntry.GetExchangeUser()
	address := ""
	if (exchangeUser) {
		address := exchangeUser.PrimarySMTPAddress 
		if (not address) {
			; e.g. outlook offline
			warning := WarningUnknown
		}
	} else {
		address := addrEntry.Address
	}
	
	if (address) {
		if  (not isInternalMail(address)) {
			warning := WarningExternal
		}
	} else { 
		warning := WarningUnknown
	}
		
	return warning
}

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

ShowWarning(warning, outlookWindow) {
	WinGetPos, ActWin_X, ActWin_Y, ActWin_W, ActWin_H, ahk_id %outlookWindow%
	if !ActWin_W
		return

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
