#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.
DetectHiddenWindows, On
;IniRead, PassControlID, vars.ini, var1

; Verstion 2.1
; 	- Added failure conditions to prevent script from getting too out of control, and hopefully allowing the script to recover from errors
!w::
	Run, Firefox.exe https://webadvisor.waketech.edu/WebAdvisor/WebAdvisor
Return
!b::
	Run, Firefox.exe https://dist-ed.waketech.edu/
Return
!c::
	Run, Firefox.exe "https://secure.waketech.edu/ResetPassword/"
Return

!i::

; Ask the user for an ID to lookup

	InputBox, InputID, User ID, Please Input a Student/Employee ID number or username
	
; Run the ID lookup on eaglesnest
; Note that chrome must be used, firefox does not select the text boxes correctly
	DetectHiddenWindows, On
	Run, Chrome.exe "https://secure.waketech.edu/eaglesnest/idlookup/"
	Sleep, 3500
	IfWinActive , ID Lookup Tools
	{
		WinMaximize, ID Lookup Tools
		Send, %InputID%
		Send, {ENTER}
	}
	Else IfWinActive, Eagles' Nest Home - Wake Technical Community College
	{
		Send, ^a
		Send, {BACKSPACE}
		Sleep, 100
		Send, Itshelpdesk
		Sleep, 100
		Send, {Tab}
		Sleep, 100
		Send, ^a
		Send, {BACKSPACE}
		Sleep, 100
		Send, P@ssW0rd
		Sleep, 100
		Send, {Tab}
		Sleep, 100
		Send, {ENTER}
		WinWaitActive, ID Lookup Tools
		WinActivate , ID Lookup Tools
		WinMaximize, ID Lookup Tools
		Send, %InputID%
		Send, {Enter}	
	}
	Else
	{
		MsgBox, ERROR - Login Unavailable, Login Unavailable, please navigate to eaglesnest login page
		loop
		{
			IfWinActive, ID Lookup Tools
			{
				WinMaximize, ID Lookup Tools
				Send, %InputID%
				Send, {ENTER}
				Break
			}
			Else IfWinActive, Eagles' Nest Home - Wake Technical Community College
			{
				Send, ^a
				Send, {BACKSPACE}
				Sleep, 100
				Send, Itshelpdesk
				Sleep, 100
				Send, {Tab}
				Sleep, 100
				Send, ^a
				Send, {BACKSPACE}
				Sleep, 100
				Send, P@ssW0rd
				Sleep, 100
				Send, {Tab}
				Sleep, 100
				Send, {ENTER}
				WinWaitActive, ID Lookup Tools
				WinActivate , ID Lookup Tools
				WinMaximize, ID Lookup Tools
				Send, %InputID%
				Send, {Enter}	
				Break
			}
			Sleep, 500
		}
	}
	Sleep, 300
	
;Retrieve some important info from the website
	Sleep, 200
	Send, ^a
	Sleep, 200
	Send, ^c
	Sleep, 200
	FileAppend, %clipboard%, site.txt
	Sleep, 30
	UserName = % TF_ReadLines("site.txt",21,21)
	Sleep, 30
	UserID = % TF_ReadLines("site.txt",22,22)
	Sleep, 30
	Name = % TF_ReadLines("site.txt",23,23)
	Sleep, 30
	StringTrimLeft, UserName, UserName, 21
	StringTrimRight, UserName, UserName, 1
	Sleep, 30
	StringTrimLeft, UserID, UserID, 21
	StringTrimRight, UserID, UserID, 1
	Sleep, 30
	StringTrimLeft, Name, Name, 14
	StringTrimRight, Name, Name, 1
	Sleep, 500
	FileDelete, site.txt



; Get logged into password control if neccesary, and then search the username, if it was provided
		

		IfWinExist, Password Control   (Waiting for Username...)
		{ 

			WinClose
		}
		Sleep, 50
		IfWinExist, Password Control   (Waiting for Password...)
		{
			WinClose
		}
		Sleep, 50
		Run, Password Control, , , PassControlID
		GroupAdd, PassControl
		Sleep, 1000
		Send, {alt}
		Sleep, 100
		Send, {ENTER}
		Sleep, 100
		Send, {ENTER}
		Sleep, 100
		SendInput, accountop
		Sleep, 500
		Send, {TAB}
		Sleep, 500
		Send, Stand@rd
		Sleep, 1000
		Send, {ENTER}
		Sleep, 1000
		
		;Enter the username we found through the lookup
		Send, %UserName%
		Sleep, 300
		Send, +`t

	;Append the inputted username/id to a txt file for later reference
	FileAppend, %A_YYYY%-%A_MM%-%A_DD%`,%A_Hour%:%A_Min%:%A_Sec%`,%Name%`,%UserID%`,%UserName%`n, log.csv
	Sleep, 100





Return



F12::
	WinGet, ActiveID, PID, A ;
	MsgBox, "%ActiveID%"
	clipboard = %ActiveID%
	Return

!d::
	Run, "C:\Program Files\DameWare Development\DameWare Mini Remote Control 7.5\DWRCC.exe"
	Return


!o:: ; Launch Outlook
	If WinExist
	Run, Outlook
	Return

!g::
	Run, "C:\Program Files (x86)\Citrix\GoToAssist Remote Support Expert\1092\g2ax_start.exe" "/Action Default" "/Trigger Shortcut" 
	Return
!n:: ; Staff member lookup
	InputBox, First, First Name, Input the employee's first name, or leave blank
	InputBox, Last, Last Name, Input the employee's last name, or leave blank
	Run, Chrome.exe "http://www.waketech.edu/directory-search"
	Sleep, 1500
	;Pressing tab 23 times gets to the last name field
	
	Loop, 22
	{
		Send, {TAB}
		Sleep, 100	
	}
	Send, %Last%
	Sleep, 100
	Send, {TAB}
	Sleep, 100
	Send, %First%
	Sleep, 100
	Send, {TAB}
	Sleep, 100
	Send, {ENTER}
	
Return



	

^+g::
	Send, ^c
	Run, Firefox.exe "http://www.google.com/search?q=%clipboard%"
	Return
	
Capslock::WinMinimizeAll


;-----------------------------End Of Code Written By Me------------------------------
;--------------------The Rest of this was written by other people--------------------

; Alt Click to drag window
~MButton & LButton::
Alt & LButton::
CoordMode, Mouse  ; Switch to screen/absolute coordinates.
MouseGetPos, EWD_MouseStartX, EWD_MouseStartY, EWD_MouseWin
WinGetPos, EWD_OriginalPosX, EWD_OriginalPosY,,, ahk_id %EWD_MouseWin%
WinGet, EWD_WinState, MinMax, ahk_id %EWD_MouseWin% 
if EWD_WinState = 0  ; Only if the window isn't maximized
    SetTimer, EWD_WatchMouse, 10 ; Track the mouse as the user drags it.
return

EWD_WatchMouse:
GetKeyState, EWD_LButtonState, LButton, P
if EWD_LButtonState = U  ; Button has been released, so drag is complete.
{
    SetTimer, EWD_WatchMouse, off
    return
}
GetKeyState, EWD_EscapeState, Escape, P
if EWD_EscapeState = D  ; Escape has been pressed, so drag is cancelled.
{
    SetTimer, EWD_WatchMouse, off
    WinMove, ahk_id %EWD_MouseWin%,, %EWD_OriginalPosX%, %EWD_OriginalPosY%
    return
}
; Otherwise, reposition the window to match the change in mouse coordinates
; caused by the user having dragged the mouse:
CoordMode, Mouse
MouseGetPos, EWD_MouseX, EWD_MouseY
WinGetPos, EWD_WinX, EWD_WinY,,, ahk_id %EWD_MouseWin%
SetWinDelay, -1   ; Makes the below move faster/smoother.
WinMove, ahk_id %EWD_MouseWin%,, EWD_WinX + EWD_MouseX - EWD_MouseStartX, EWD_WinY + EWD_MouseY - EWD_MouseStartY
EWD_MouseStartX := EWD_MouseX  ; Update for the next timer-call to this subroutine.
EWD_MouseStartY := EWD_MouseY
return


; Ctrl+Alt+c autocorrect selected text
^!c::
clipback := ClipboardAll
clipboard=
Send ^c
ClipWait, 0
UrlDownloadToFile % "https://www.google.com/search?q=" . clipboard, temp
FileRead, contents, temp
FileDelete temp
if (RegExMatch(contents, "(Showing results for|Did you mean:)</span>.*?>(.*?)</a>", match)) {
   StringReplace, clipboard, match2, <b><i>,, All
   StringReplace, clipboard, clipboard, </i></b>,, All
}
Send ^v
Sleep 500
clipboard := clipback
return

;Bakspace in file explorer
#IfWinActive, ahk_class CabinetWClass
`::
   ControlGet renamestatus,Visible,,Edit1,A
   ControlGetFocus focussed, A
   if(renamestatus!=1&&(focussed=”DirectUIHWND3″||focussed=SysTreeView321))
   {
    SendInput {Alt Down}{Up}{Alt Up}
  }else{
      Send {Backspace}
  }
#IfWinActive


