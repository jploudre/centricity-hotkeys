; Setup
ClinicalAssistantName = "Handy"
CoordMode, Mouse, Window
#Persistent
SetKeyDelay, 30
return

#IfWinActive, Update - ;###########################################################

`::PatternHotKey(".->GotoChart","..->SwapTextView")
\::PatternHotKey(".->EndUpdate", "..->EndUpdateToClinicalAssistant")
F1::PatternHotKey(".->OrderSearch", "..->SignOrders")
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

#r::
OpenAppendType("Web")
return

#c::
OpenAppendType("CPOE")
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

#IfWinActive, Update Problems - ;###########################################################

; Long Hold is Top/Bottom. Tap is Up/Down
Up::PatternHotKey(".->UpdateProblemsUp","_->UpdateProblemsTop")

Down::PatternHotKey(".->UpdateProblemsDown","_->UpdateProblemsBottom")

Left::
Click, 736, 146
return

Right::
Click, 786, 146
return

BackSpace::
Delete::PatternHotKey(".->UpdateProblemsRemove")

\::
Enter::
Click, 694, 599
return

#IfWinActive, Update Medications - ;###########################################################

Enter::
click 559, 566
return

BackSpace::
Delete::
Send !r
WinWaitActive, Remove Medication
Click 285, 311
return

#ifWinActive, Update Orders - ;###########################################################

#s::
CLick 561, 656
return

F1::PatternHotKey("..->SignOrders")
return

LButton::
MouseGetPos, xpos, ypos
if ( 638 < xpos AND xpos < 709 AND 647 < ypos AND ypos < 667)
    {
    ; Click Sign, first
    Click 561, 656
    Citrixsleep()
    Soundplay *64
    Click %xpos%, %ypos%
    }
else
    {
    Click    
    }
return

#IfWinActive, Assessments Due ;###########################################################

Esc::
Enter::
WinGetPos,,,winwidth,winheight,A
xpos := winwidth - 20
ypos := winheight - 20
Click, %xpos%, %ypos%
return

#ifWinActive, Customize Letter ;###########################################################

#Space::
Send !p
WinWaitNotActive, Customize Letter
WinWaitActive, Customize Letter
Citrixsleep()
Send !s
WinWaitActive, Route Document
Citrixsleep()
Send !s
WinWaitActive, Print
Citrixsleep()
Click 568, 355
Soundplay *64
return

#ifWinActive, Care Alert Warning ;###########################################################

Space::
Enter::
\::
Send !c
return

#ifWinActive, New Medication ;###########################################################

LButton::
MouseGetPos, xpos, ypos
if ( 649 < xpos AND xpos < 712 AND 643 < ypos AND ypos < 669)
    {
    Soundplay *64
    Click %xpos%, %ypos%
    WinWaitActive, Update -
    CitrixSleep()
    Gosub, Prescriptions
    }
else 
{
    Click
}
return

; End of Window Specific Hotkeys.  #########################################
#IfWinActive
; Miscellaneous Functions ##############################

CitrixSleep(){
Sleep, 150
}
return

OpenAppendType(searchtext){
    ifWinActive, Chart Desktop -
    {
    Send ^j
    }
    ifWinActive, Chart -
    {
    WinGetPos,,,winwidth,winheight,A
    ImageSearch, FoundX, FoundY, 200, 50, %winwidth%, %winheight%, *n10 %A_ScriptDir%/files/append.png
    if (ErrorLevel = 0) {
        MouseMove, %FoundX%, %FoundY%
        Click
    }
    }
    ; Sometimes fails, use 3 second timeout
    WinWaitActive, Append to, , 3
    if (ErrorLevel = 0) {
        CitrixSleep()
        Send !F
        WinWaitActive, Append Document ; no timeout needed
        if (ErrorLevel = 0) {
            CitrixSleep()
            Send %searchtext%
            Send {Enter}
            WinWaitActive, Update ; no timeout needed
            if (ErrorLevel = 0) {
            CitrixSleep()
            Send +{F8}
            CitrixSleep()
            }
        }
    }
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
    MouseMove, %FoundX%, %FoundY%
    Click
    WinWaitNotActive, Chart,, 3
    if (ErrorLevel= 0) {
    Sleep, 1500
    IfWinActive, Centricity Practice Solution
        {
        Click, 508, 10, 2   ; Minimizes
        Sleep, 500
        WinGetPos, xpos, ypos, winwidth, winheight, A
        CoordMode, mouse, screen
        MouseClickDrag, Left, xpos + 200, ypos + 1, xpos + 200, 20
        CoordMode, mouse, relative
        
        WinGetPos, xpos, ypos, winwidth, winheight, A
        ychange := A_ScreenHeight - (winheight + 50)
        MouseClickDrag, Left, 200, ypos + winheight  - 20, 200, ypos + winheight - 20 + ychange
        }
    return
    }
    } ; end Paperclip
}
ImageSearch, FoundX, FoundY, 200, 50, %winwidth%, %winheight%, *n10 %A_ScriptDir%/files/open.png
if (ErrorLevel = 0) {
    ; if over 200 pixels y, we're in Chart Summary, might need to open attachment.
    MouseMove, %FoundX%, %FoundY%
    Click
    Sleep, 1000
    IfWinActive, Care Alert Warning - 
    {
    Send !c
    }
    Sleep, 500
    WinGetPos,,,winwidth,winheight,A
    ImageSearch, FoundX, FoundY, 200, 50, %winwidth%, %winheight%, *n10 %A_ScriptDir%/files/attach.png
    if (ErrorLevel = 0) {
        ImageSearch, FoundX1, FoundY1, 200, 50, %winwidth%, %winheight%, *n10 %A_ScriptDir%/files/paperclip.png
        if (ErrorLevel = 0) {
            MouseMove, %FoundX%, %FoundY%
            Click
        WinWaitNotActive, Chart,, 3
        if (ErrorLevel= 0) {
        Sleep, 1500
        IfWinActive, Centricity Practice Solution
            {
            Click, 508, 10, 2   ; Minimizes
            Sleep, 500
            WinGetPos, xpos, ypos, winwidth, winheight, A
            CoordMode, mouse, screen
            MouseClickDrag, Left, xpos + 200, ypos + 1, xpos + 200, 20
            CoordMode, mouse, relative
        
            WinGetPos, xpos, ypos, winwidth, winheight, A
            ychange := A_ScreenHeight - (winheight + 50)
            MouseClickDrag, Left, 200, ypos + winheight  - 20, 200, ypos + winheight - 20 + ychange
            }
        return
        }
        } ; end Paperclip
}
}
return
FancyCPOEAppend:
OpenAppendType("CPOE")
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

SignOrders:
Click, 254, 38
WinWaitActive, Update Orders, , 3 ; Timeout
if (ErrorLevel = 0) {
	CitrixSleep()
	Click, 561, 653
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
CitrixSleep()
WinGetPos,,,,winheight,A
ypos := winheight - 217
Click, 13, %ypos%
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
Click, 
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
	; Template Not Found So skip rest of hotkey that called this. 
	Exit
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

; Update Problems Functions ##############################

UpdateProblemsTop:
Click, 762, 100
return
UpdateProblemsBottom:
Click, 762, 189
return
UpdateProblemsUp:
Click, 762, 122
return
UpdateProblemsDown:
Click, 762, 170
return
UpdateProblemsRemove:
Click, 508, 572
WinWaitNotActive
Send {Enter}
return

; http://www.autohotkey.com/board/topic/66855-patternhotkey-map-shortlong-keypress-patterns-to-anything/?hl=%2Bpatternhotkey
; Usage : hotkey::PatternHotKey("command1", ["command2", "command3", length(integer), period(float)])
;     where commands match one of the following formats:
;         "pattern:keys"                  ; Maps pattern to send keys
;         "pattern->label"                ; Maps pattern to label (GoSub)
;         "pattern->function()"           ; Maps pattern to function myfunction with
;                                           no parameter
;         "pattern->function(value)"      ; Maps pattern to function myfunction with
;                                           the first parameter equal to 'value'
;         and patterns match the following formats:
;             '.' or '0' represents a short press
;             '-' or '_' represents a long press of any length
;             '?' represents any press
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
