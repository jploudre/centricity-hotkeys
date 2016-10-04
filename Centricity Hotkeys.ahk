; Setup
FirstRun()
IniRead, Buddy, Settings.ini, Preferences, Buddy
ClinicalAssistantName = %Buddy%
CoordMode, Mouse, Window
#Persistent
SetKeyDelay, 30
SendMode Input
Menu, Tray, NoStandard
Menu, Tray, Add, Exit, ExitScript
return

ExitScript:
ExitApp

; Check if this is original time opened after a download
FirstRun(){
IfNotExist, Settings.ini
{
InputBox, BuddyName, Who's your Buddy?,
(

Who do you 'hold' things to most frequently
in Centricity?

Typically this might be your CAs last name...
), , 300, , , , , , 
if (Errorlevel= 0) {
IniWrite, %BuddyName%, Settings.ini, Preferences, Buddy
MsgBox, 64, Thanks,
(
If this is your first time:
    

   - Want keyboard stickers?
   - Print the help cheatsheet?


)

}
}
}
return

#IfWinActive, Update - ;###########################################################

`::PatternHotKey(".->GotoChart","..->SwapTextView")
\::PatternHotKey(".->EndUpdate", "..->EndUpdateToClinicalAssistant")
F1::PatternHotKey(".->OrderSearch")
F2::PatternHotKey(".->MedSearch", "..->UpdateMeds")
F3::PatternHotKey(".->ProblemSearch", "..->UpdateProblems")
F5::PatternHotKey(".->HPI")
F6::PatternHotKey(".->Preventive", "..->CommittoFlowsheet")
F7::PatternHotKey(".->PMH-SH-CCC", "..->InserttoNote")
F8::PatternHotKey(".->ROS", "..->ROS2")
F9::PatternHotKey(".->PE")
F10::PatternHotKey(".->CPOE", "..->AssessmentsDue")
F11::PatternHotKey(".->PatientInstructions", "..->PrintVisitSummary")
F12::PatternHotKey(".->Prescriptions", "..->SendPrescriptions")

[::
Send ^{PgUp}
return

]::
Send ^{PgDn}
return

#+s::
Gosub, EndUpdate
Gosub, SignUpdate
return

#IfWinActive, End Update ;###########################################################

\::PatternHotKey(".->HoldUpdate", "..->SendToClinicalAssistant")

#+s::
Gosub, SignUpdate
return

#IfWinActive, Chart - ;###########################################################

Space::PatternHotKey(".->FancyOpen","..->FancySign","_->FancyCPOEAppend")

`::
IfWinExist, Update
WinActivate, Update
IfWinNotExist, Update
gosub, GoChartDesktop
return

#+s::
Gosub, FancySign
return


Up::
FocusBlue()
Send {Up}
return

Down::
FocusBlue()
Send {Down}
return

#IfWinActive, Chart Desktop - ;###########################################################

Space::PatternHotKey(".->FancyOpen","..->FancySign","_->FancyCPOEAppend")
`::
IfWinExist, Update
WinActivate, Update
IfWinNotExist, Update
{
    WinGetPos,,,,winheight,A
    ypos := winheight - 182
    Click, 13, %ypos%
}
return

#+s::
Gosub, FancySign
return

#IfWinActive, Centricity Practice Solution Browser: ;###########################################################

Space::PatternHotKey(".->DownDocumentViewer", "..->CloseDocumentViewerandSave", "_->CloseDocumentViewerandAppend")

\::
gosub, CloseDocumentViewer
return

up::
WinGetPos,,,winwidth,,A
winwidth := winwidth - 10
Click %winwidth%, 67
return

down::
gosub, DownDocumentViewer
return

#+s::
Gosub, CloseDocumentViewerandSave
return

; End of Window Specific Hotkeys.  #########################################
#IfWinActive

; Miscellaneous Functions ##############################

CitrixSleep(){
Sleep, 150
}
return

; In Chart a selected item doesn't respond to arrows or hotkeys. Clicks to set focus
FocusBlue(){
WinGetPos,,,winwidth,winheight,A
ImageSearch, FoundX, FoundY, 200, 50, %winwidth%, %winheight%, %A_ScriptDir%/files/blue.png
    if (ErrorLevel = 0) {
        Click, %FoundX%, %FoundY%
    }
return
}
return

; Update Functions ##############################

GotoChart:
WinActivate, Chart
return

SwapTextView:
Send +{F8}
return

HoldUpdate:
Send !o
return

SignUpdate:
Send !s
WinWaitNotActive, End Update
gosub, GoChartDesktop
return

EndUpdate:
Send ^{PgDn}
citrixsleep()
citrixsleep()
citrixsleep()
Send ^e
WinWaitActive, End Update
return

EndUpdateToClinicalAssistant:
Gosub, EndUpdate
Gosub, SendToClinicalAssistant
return

SendToClinicalAssistant:
; Assumes End Update
IfWinActive, End Update
{
	Click, 316, 351
	WinWaitActive, New Routing Information
    CitrixSleep()
	SendInput %ClinicalAssistantName%{Enter}
	CitrixSleep()
	Click, 240, 345
	WinWaitActive, End Update
	CitrixSleep()
	Send !o
    WinWaitNotActive
    gosub, GoChartDesktop
}
else
{
return	
}
return

; Fancy Spacebar Functions ##############################
FancyOpen:
WinGetPos,,,winwidth,winheight,A
ImageSearch, FoundX, FoundY, 200, 50, %winwidth%, %winheight%, *n10 %A_ScriptDir%/files/attach.png
if (ErrorLevel = 0) {
    ImageSearch, FoundX1, FoundY1, 200, 50, %winwidth%, %winheight%, *n10 %A_ScriptDir%/files/paperclip.png
    if (ErrorLevel = 0) {
    MouseClick, left, %FoundX%, %FoundY%
    } ; end Paperclip
}
ImageSearch, FoundX, FoundY, 200, 50, %winwidth%, %winheight%, *n10 %A_ScriptDir%/files/open.png
if (ErrorLevel = 0) {
    ; if over 200 pixels y, we're in Chart Summary, might need to open attachment.
    MouseClick, left, %FoundX%, %FoundY%
    Sleep, 1000
    IfWinActive, Care Alert Warning - 
    {
    Send !c
    }
    Sleep, 1000
    WinGetPos,,,winwidth,winheight,A
    ImageSearch, FoundX, FoundY, 200, 50, %winwidth%, %winheight%, *n10 %A_ScriptDir%/files/attach.png
    if (ErrorLevel = 0) {
        ImageSearch, FoundX1, FoundY1, 200, 50, %winwidth%, %winheight%, *n10 %A_ScriptDir%/files/paperclip.png
        if (ErrorLevel = 0) {
            MouseClick, left, %FoundX%, %FoundY%
        } ; end Paperclip
}
WinWaitActive, Update, , 5
if (ErrorLevel = 0)
{
    Send, {WheelDown}
    Send, {WheelDown}
    Send, {WheelDown}
    Send, {WheelDown}
    Send, {WheelDown}

}
}
return

FancyCPOEAppend:
searchtext = "CPOE"
if winactive "Chart Desktop" or winactive "Chart" 
{
Progress, P0 b1 fm18 zh15 W400 cbGray, , Opening CPOE... , , Tahoma Bold
    ifWinActive, Chart Desktop -
    {
    Send ^j
    }
    ifWinActive, Chart -
    {
    WinGetPos,,,winwidth,winheight,A
    ImageSearch, FoundX, FoundY, 200, 50, %winwidth%, %winheight%, *n10 %A_ScriptDir%/files/append.png
    if (ErrorLevel = 0) {
        Click %FoundX%, %FoundY%
    }
    }
    ; Sometimes fails, use 3 second timeout
    WinWaitActive, Append to, , 3
    if (ErrorLevel = 0) {
        CitrixSleep()
        Progress, 5
        Send !F
        WinWaitActive, Append Document ; no timeout needed
        if (ErrorLevel = 0) {
            CitrixSleep()
            Progress, 10
            Send %searchtext%
            Progress, 15
            Send {Enter}
            ; Takes a while to open Append. Many progress updates.....
            Progress, 20
            CitrixSleep()
            Progress, 25
            CitrixSleep()
            Progress, 30
            CitrixSleep()
            Progress, 35
            CitrixSleep()
            Progress, 45
            CitrixSleep()
            Progress, 50
            CitrixSleep()
            Progress, 55
            CitrixSleep()
            Progress, 60
            CitrixSleep()
            Progress, 65
            CitrixSleep()
            Progress, 70
            WinWaitActive, Update ; no timeout needed
            if (ErrorLevel = 0) {
            CitrixSleep()
            Progress, 75
            Send +{F8}
            CitrixSleep()
            Progress, 80
            }
        }
    }
    template = ""
    ; From Generic FindTemplate but with progess updates
    ImageSearch, FoundX, FoundY, 20, 170, 203, 536, *n10 %A_ScriptDir%/files/CPOE-A&P-CCC.png
    if (ErrorLevel = 0) {
	Mouseclick, Left, %FoundX%, %FoundY%, 2
	CitrixSleep()
	Progress, 85
    CitrixSleep()
    Progress, 90
	CitrixSleep()
    Progress, 95
    ; Assessments Due
    MouseMove, 561, 108
    Progress, 100
    Progress, Off
}
}
return

FancySign:
ifWinActive, Chart Desktop -
    {
    Send ^s
    }
ifWinActive, Chart -
    {
    FocusBlue()
    Send ^s
    }
return

; Centricity Update Hotkey Functions ##############################

OrderSearch:
Click, 254, 38
WinWaitActive, Update Orders, , 3 ; Timeout
if (ErrorLevel = 0) {
    CitrixSleep()
    Click, 253, 287
    CitrixSleep()
    Click 412, 337
}
return

MedSearch:
Click, 524, 38
WinWaitActive, New Medication, , 3 ; Timeout
if (ErrorLevel = 0) {
	CitrixSleep()
	Click, 718, 81
	WinWaitActive, Find Medication
}
return

UpdateMeds:
Click, 350, 38
return

ProblemSearch:
Click, 593, 37
return

UpdateProblems:
Click, 428, 38
return

GoChartDesktop:
WinWaitActive, Chart -, , 4
if (ErrorLevel = 0) 
{
    ImageSearch, FoundX, FoundY, 0, 0, 200, %A_ScreenHeight%, *n10 %A_ScriptDir%/files/chart-desktop.png
    if (ErrorLevel = 0) {
        Mouseclick, Left, %FoundX%, %FoundY%, 
    WinWaitActive, Chart Desktop -, 3
    }
}
return

HPI:
FindTemplate("HPI-CCC")
Click, 767, 559
return

Preventive:
FindTemplate("Preventive-Care-Screening")
return

CommittoFlowsheet:
FindTemplate("Preventive-Care-Screening")
Click, 599, 113
return

PMH-SH-CCC:
FindTemplate("PMH-SH-CCC")
return

InserttoNote:
FindTemplate("PMH-SH-CCC")
Click, 921, 112
return

ROS:
FindTemplate("ROS-CCC")
return

ROS2:
FindTemplate("ROS-CCC")
Click, 351, 80
return

PE:
FindTemplate("PE-CCC")
if (ErrorLevel = 1)
{
FindTemplate("Pediatric-PE-Age-Specific-CCC")
}
Click, 522, 203
return

CPOE:
FindTemplate("CPOE-A&P-CCC")
return

AssessmentsDue:
FindTemplate("CPOE-A&P-CCC")
Click, 561, 108
return

PatientInstructions:
FindTemplate("Patient-Instructions-CCC")
Click, 594, 341
return

PrintVisitSummary:
FindTemplate("Patient-Instructions-CCC")
Click, 891, 130 
return

Prescriptions:
FindTemplate("Prescriptions")
MouseMove, 946, 563
return

SendPrescriptions:
FindTemplate("Prescriptions")
Click, 946. 563
return

FindTemplate(template) {
ImageSearch, FoundX, FoundY, 20, 170, 203, 536, *n10 %A_ScriptDir%/files/%template%.png
if (ErrorLevel = 0) {
	MouseMove, %FoundX%, %FoundY%
	Click 2
	CitrixSleep()
	CitrixSleep()
	CitrixSleep()
    Click 2
    MouseMove, 500, 0, 0, R
}
; if template not found, is it already selected?
if (ErrorLevel = 1) {
	ImageSearch, FoundX, FoundY, 20, 170, 203, 536, *n10 %A_ScriptDir%/files/%template%-highlighted.png
	; If found, errorlevel is now 0
}
if (ErrorLevel >= 1) {
	; Template Not Found So skip rest of hotkey except on Peds PE which needs to run to find Age exam.
	if (template != "PE-CCC")
    {
    exit
    }
}
}
return

; Centricty Browswer Functions ##############################

DownDocumentViewer:
WinGetPos,,,winwidth, winheight,A
xclick := winwidth - 10
yclick := winheight - 24
CLick, %xclick%, %yclick%
return

CloseDocumentViewer:
Send !{F4}
return

CloseDocumentViewerandSave:
Send !{F4}
WinWaitNotActive
Citrixsleep()
Focusblue()
Send ^s
return

CloseDocumentViewerandAppend:
Send !{F4}
WinWaitNotActive
Citrixsleep()
Focusblue()
Gosub, FancyCPOEAppend
return

; http://www.autohotkey.com/board/topic/66855-patternhotkey-map-shortlong-keypress-patterns-to-anything/?hl=%2Bpatternhotkey
PatternHotKey(arguments*)
{
    period = 0.2
    length = 1
    for index, argument in arguments
    {
        if argument is float
            period := argument, continue
        if argument is integer
            length := argument, continue
        separator := InStr(argument, ":", 1) - 1
        if ( separator >= 0 )
        {
            pattern   := SubStr(argument, 1, separator)
            command    = Send
            parameter := SubStr(argument, separator + 2)
        }
        else
        {
            separator := InStr(argument, "->", 1) - 1
            if ( separator >= 0 )
            {
                pattern := SubStr(argument, 1, separator)

                call := Trim(SubStr(argument, separator + 3))
                parenthesis := InStr(call, "(", 1, separator) - 1
                if ( parenthesis >= 0 )
                {
                    command   := SubStr(call, 1, parenthesis)
                    parameter := Trim(SubStr(call, parenthesis + 1), "()"" `t")
                }
                else
                {
                    command    = GoSub
                    parameter := call
                }
            }
            else
                continue
        }

        if ( Asc(pattern) = Asc("~") )
            pattern := SubStr(pattern, 2)
        else
        {
            StringReplace, pattern, pattern, ., 0, All
            StringReplace, pattern, pattern, -, [1-9A-Z], All
            StringReplace, pattern, pattern, _, [1-9A-Z], All
            StringReplace, pattern, pattern, ?, [0-9A-Z], All
            pattern := "^" . pattern . "$"
            if ( length < separator )
                length := separator
        }

        patterns%index%   := pattern
        commands%index%   := command
        parameters%index% := parameter
    }
    keypress := KeyPressPattern(length, period)
    Loop %index%
    {
        pattern   := patterns%A_Index%
        command   := commands%A_Index%
        parameter := parameters%A_Index%

        if ( pattern && RegExMatch(keypress, pattern) )
        {
            if ( command = "Send" )
                Send % parameter
            else if ( command = "GoSub" and IsLabel(parameter) )
                gosub, %parameter%
            else if ( IsFunc(command) )
                %command%(parameter)
        }
    }
}

KeyPressPattern(length = 2, period = 0.2)
{
    key := RegExReplace(A_ThisHotKey, "[\*\~\$\#\+\!\^]")
    IfInString, key, %A_Space%
        StringTrimLeft, key, key, % InStr(key, A_Space, 1)
    if key in Alt,Ctrl,Shift,Win
        modifiers := "{L" key "}{R" key "}"
    current = 0
    loop
    {
        KeyWait %key%, T%period%
        if ( ! ErrorLevel )
        {
            pattern .= current < 10
                       ? current
                       : Chr(55 + ( current > 36 ? 36 : current ))
            current = 0
        }
        else
            current++
        if ( StrLen(pattern) >= length )
            return pattern
        if ( ! ErrorLevel )
        {
            if ( key in Capslock, LButton, MButton, RButton or Asc(A_ThisHotkey) = Asc("$") )
            {
                KeyWait, %key%, T%period% D
                if ( ErrorLevel )
                    return pattern
            }
            else
            {
                Input,, T%period% L1 V,{%key%}%modifiers%
                if ( ErrorLevel = "Timeout" )
                    return pattern
                else if ( ErrorLevel = "Max" )
                    return
                else if ( ErrorLevel = "NewInput" )
                    return
            }
        }
    }
}
