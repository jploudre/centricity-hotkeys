; Setup
{
ClinicalAssistantName = "Handy"
CoordMode, Mouse, Window
#Persistent
SetKeyDelay, 30
SetTimer, CloseOutlook, 5000 
return
 
<#Esc::run taskmgr.exe
<#Up::Send {PgUp}
<#Down::Send {PgDn}
<#Backspace::Send {Delete}
RWin::return
LWin::return
#L::return
#D::return
}


#IfWinActive, Update - ;###########################################################
`::PatternHotKey(".->GotoChart","..->SwapTextView")
return
[::
Send ^{PgUp}
return
]::
Send ^{PgDn}
return
!Space::PatternHotKey(".->EndUpdate", "..->EndUpdateToClinicalAssistant")
#Space::PatternHotKey(".->EndUpdate", "..->EndUpdateToClinicalAssistant")
\::PatternHotKey(".->EndUpdate", "..->EndUpdateToClinicalAssistant")
F1::PatternHotKey(".->OrderSearch", "..->SignOrders")
F2::PatternHotKey(".->MedSearch", "..->UpdateMeds")
F3::PatternHotKey(".->ProblemSearch", "..->UpdateProblems")
F5::PatternHotKey(".->HPI")
F6::PatternHotKey(".->Preventive", "..->CommittoFlowsheet")
F7::PatternHotKey(".->PMH-SH-CCC", "..->InserttoNote")
F8::PatternHotKey(".->ROS", "..->ROS2")
F9::PatternHotKey(".->PE", "..->PE-XC", "...->PE-XU", "_->PE-XP")
F10::PatternHotKey(".->CPOE", "..->AssessmentsDue")
F11::PatternHotKey(".->PatientInstructions", "..->PrintVisitSummary")
F12::PatternHotKey(".->Prescriptions", "..->SendPrescriptions")

; Ends and signs an update. 
#+s::
GoSub, SendPortal
Gosub, EndUpdate
Gosub, SignUpdate
return

#+p::
gosub, CommittoFlowsheetandSign
return

#/::
GoSub, GotoChart
citrixsleep()
citrixsleep()
citrixsleep()
WinGetPos,,,winwidth,winheight,A
ImageSearch, FoundX, FoundY, 0, 112, %winwidth%, %winheight%, *n50 %A_ScriptDir%/files/documents.png
if (ErrorLevel = 0) {
    MouseMove, %FoundX%, %FoundY%
    Click
    citrixsleep()
    Click, 255, 212
}
return


#IfWinActive, End Update ;###########################################################
!Space::PatternHotKey(".->HoldUpdate", "..->SendToClinicalAssistant")
#Space::PatternHotKey(".->HoldUpdate", "..->SendToClinicalAssistant")
\::PatternHotKey(".->HoldUpdate", "..->SendToClinicalAssistant")
return
#s::
Gosub, SignUpdate
return


#IfWinActive, Chart - ;###########################################################
`::
IfWinExist, Update
WinActivate, Update
IfWinNotExist, Update
gosub, GoChartDesktop
return

; Sign a chart document
#s::
FocusBlue()
Send ^s
return

#j::
WinGetPos,,,winwidth,winheight,A
ImageSearch, FoundX, FoundY, 200, 50, %winwidth%, %winheight%, *n10 %A_ScriptDir%/files/append.png
if (ErrorLevel = 0) {
    MouseMove, %FoundX%, %FoundY%
    Click
}
return

#+p::
OpenAppendType("Clinical List Pr")
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
; Append
#j::
Send ^j
return

; Sign
#s::
Send ^s
return
; Preventive Append. Assumes in Documents.

#+p::
OpenAppendType("Clinical List Pr")
return

; Replies with Web Message. Assumes in Documents.
#r::
Send ^j
OpenAppendType("Web")
return

; eRx Append. Assumes in Documents.
#e::
Send ^j
OpenAppendType("* eSM")
return

; CPOE Append. Assumes in Documents.
#c::
Send ^j
OpenAppendType("CPOE")
return


#IfWinActive, Centricity Practice Solution Browser: ;###########################################################
#Space::
\::
Send !{F4}
return
down::
WinGetPos,,,winwidth,winheight,A
winwidth := winwidth - 10
winheight := winheight - 25
Click %winwidth%, %winheight%
return
up::
WinGetPos,,,winwidth,,A
winwidth := winwidth - 10
Click %winwidth%, 67
return
; Close and Sign
#s::
Send !{F4}
WinWaitNotActive
Citrixsleep()
Focusblue()
Send ^s
return

#+p::
Send Send !{F4}
Sleep, 1000
OpenAppendType("Clinical List Pr")
return


#IfWinActive, Blackbird Content ;###########################################################
#Space::
WinGetPos,,,winwidth,winheight,A
ImageSearch, FoundX, FoundY, 200, 50, %winwidth%, %winheight%, *n10 %A_ScriptDir%/files/Blackbird-OK.png
if (ErrorLevel = 0) {
    MouseMove, %FoundX%, %FoundY%
    Click
    Sleep, 500
    Send !{F4}
}
return
Enter::
Send \
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
; OK is 'Done'
\::
Click, 694, 599
return
; Enter should always do OK.
#Space::
!Space::
Enter::
Click, 694, 599
return


#IfWinActive, Update Medications - ;###########################################################
#Space::
!Space::
Enter::
click 559, 566
return
; Sends (intead of saves)
#s::
Send !p
citrixsleep()
Click, 559, 566
return
BackSpace::
Delete::
Send !r
WinWaitActive, Remove Medication
Click 285, 311
return


#ifWinActive, Update Orders - ;###########################################################

#Space::
!Space::
Click 561, 656
Citrixsleep()
Click 679, 656
return
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

#IfWinActive, Append to Document ;###########################################################
#s::
^s::
Send !s
return

#IfWinActive, Assessments Due ;###########################################################
#s::
^s::
#Space::
!Space::
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
Soundplay, %A_ScriptDir%/files/done.wav, Wait
return

#ifWinActive, Care Alert Warning ;###########################################################
#Space::
!Space::
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


; Open Attachment
#o::
if WinActive("Chart - \\Remote") or WinActive("Chart Desktop - \\Remote")
{
WinGetPos,,,winwidth,winheight,A
ImageSearch, FoundX, FoundY, 200, 50, %winwidth%, %winheight%, *n10 %A_ScriptDir%/files/attach.png
if (ErrorLevel = 0) {
    MouseMove, %FoundX%, %FoundY%
    Click
    WinWaitNotActive, Chart,, 1.5
    ; If Open Attachment fails, try open chart
    if (ErrorLevel= 0) {
    Sleep, 500
    IfWinActive, Centricity Practice Solution
        {
        Click, 508, 10, 2   ; Minimizes
        Sleep, 1000
        ImageSearch, FoundX, FoundY, 0, 0, %A_ScreenHeight%, %A_ScreenWidth%, *n10 %A_ScriptDir%/files/CPS-Browser-Top-Edge.png
         if (ErrorLevel = 0) {
            MouseMove, %FoundX%, %FoundY%
            CoordMode, Mouse, Screen
            MouseGetPos,, MouseY
            MouseY := (MouseY - 20) * -1
            CoordMode, Mouse, Relative
            MouseClickDrag, Left, %FoundX%, %FoundY%, %FoundX%, %MouseY%,
        }
        }
    return
    }
    if (ErrorLevel = 1) {
        ImageSearch, FoundX, FoundY, 200, 50, %winwidth%, %winheight%, *n10 %A_ScriptDir%/files/open.png
        if (ErrorLevel = 0) {
            MouseMove, %FoundX%, %FoundY%
            Click
        }
    }
}
}
return

; Open Patient Chart to the item
+#o::
if WinActive("Chart - \\Remote") or WinActive("Chart Desktop - \\Remote")
{
WinGetPos,,,winwidth,winheight,A
ImageSearch, FoundX, FoundY, 200, 50, %winwidth%, %winheight%, *n10 %A_ScriptDir%/files/open.png
if (ErrorLevel = 0) {
    MouseMove, %FoundX%, %FoundY%
    Click
}
}
return



; Replies with Web Message. Assumes in Documents.
#r::
OpenAppendType("Web")
return

; eRx Append. Assumes in Documents.
#e::
OpenAppendType("* eSM")
return


; CPOE Append. Assumes in Documents.
#c::
OpenAppendType("CPOE")
return

; Reply to a patient with a blank letter
+#R::
{
Send ^p
CitrixSleep()
Send l
CitrixSleep()
Send {Down 2}
CitrixSleep()
Send {Right 2}
CitrixSleep()
Send l
CitrixSleep()
Send {Down 2}
Click, 241, 59
Send B
CitrixSleep()
Click, 392, 351
}
return

; Quit All Windows
#+q::
WinGet, id, list,,, Program Manager
Loop, %id%
{
 this_id := id%A_Index%
 WinClose, ahk_id %this_id%
}
Sleep, 1000
DllCall("LockWorkStation")
Sleep, 200
SendMessage,0x112,0xF170,2,,Program Manager
return


; Functions ####################################################################

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

; In Chart a selected item doesn't respond to arrows or hotkeys.
; Click to set focus on the blue

FocusBlue(){
WinGetPos,,,winwidth,winheight,A
ImageSearch, FoundX, FoundY, 200, 50, %winwidth%, %winheight%, %A_ScriptDir%/files/blue.png
    if (ErrorLevel = 0) {
        Click, %FoundX%, %FoundY%
    }
return
}

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
WinWaitActive, Chart Desktop -,,5, ; Up to 5 seconds to complete
if (ErrorLevel = 0) {
    Soundplay, %A_ScriptDir%/files/done.wav, WAIT
}
return

EndUpdate:
Send ^{PgDn}
citrixsleep()
citrixsleep()
citrixsleep()
Send ^e
WinWaitActive, End Update
return

SendPortal:
WinGetPos,,,winwidth,winheight,A
ImageSearch, FoundX, FoundY, 0, 112, %winwidth%, %winheight%, *n50 %A_ScriptDir%/files/portal.png
if (ErrorLevel = 0) {
    
   Click, 406, 690
   Sleep, 1000
   Click, 406, 720
   Sleep, 1000
}
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

; Centricity Update Hotkey Functions
;#############################################################################

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
FindTemplate("Blackbird")
Click, 771, 87
WinWaitActive, Blackbird
Sendraw =
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

CommittoFlowsheetandSign:
Gosub, CommittoFlowsheet
Gosub, EndUpdate
Gosub, SignUpdate
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

; ###################### Broken PE second search
PE:
FindTemplate("PE-CCC")
if (ErrorLevel = 1)
{
FindTemplate("Pediatric-PE-Age-Specific-CCC")
}
Click, 522, 203
return

PE-XU:
FindTemplate("PE-CCC")
if (ErrorLevel = 1)
{
FindTemplate("Pediatric-PE-Age-Specific-CCC")
}
Click, 522, 203
Send xu{return}
return

PE-XC:
FindTemplate("PE-CCC")
if (ErrorLevel = 1)
{
FindTemplate("Pediatric-PE-Age-Specific-CCC")
}
Click, 522, 203
Send xc{return}
return

PE-XP:
FindTemplate("PE-CCC")
if (ErrorLevel = 1)
{
FindTemplate("Pediatric-PE-Age-Specific-CCC")
}
Click, 522, 203
Send xp{return}
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


; Update Problems Functions

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


; Downloaded Functions ----------------------------------------------------------------------------------

Clip(Text="", Reselect="") ; http://www.autohotkey.com/forum/viewtopic.php?p=467710 , modified August 2012
{
	Static BackUpClip, Stored, LastClip
	If (A_ThisLabel = A_ThisFunc) {
		If (Clipboard == LastClip)
			Clipboard := BackUpClip
		BackUpClip := LastClip := Stored := ""
	} Else {
		If !Stored {
			Stored := True
			BackUpClip := ClipboardAll
		} Else
			SetTimer, %A_ThisFunc%, Off
		LongCopy := A_TickCount, Clipboard := "", LongCopy -= A_TickCount
		If (Text = "") {
			Send, ^c
			ClipWait, LongCopy ? 0.5 : 0.25
		} Else {
			Clipboard := LastClip := Text
			ClipWait, 10
			Send, ^v
		}
		SetTimer, %A_ThisFunc%, -700
		If (Text = "")
			Return LastClip := Clipboard
		Else If (ReSelect = True) or (Reselect and (StrLen(Text) < 3000)) {
			Sleep 30
			StringReplace, Text, Text, `r, , All
			SendInput, % "{Shift Down}{Left " StrLen(Text) "}{Shift Up}"
		}
	}
	Return
	Clip:
	Return Clip()
}

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

::ujkp::
text := "Upcoming Appointment. ............................ Jonathan Ploudre, MD. " . A_MMM . " " . A_DD . ", " A_YYYY
clip(text)
sleep 100
Send !s
return
::sljkp::
text := "Send Letter with Results. ............................ Jonathan Ploudre, MD. " . A_MMM . " " . A_DD . ", " A_YYYY
clip(text)
sleep 100
Send !s
return
::cdn::
Send Call Doctor Note:{Enter 2}SITUATION:{Enter 3}BACKGROUND:{Enter 3}ASSESSMENT:{Enter 3}RECOMENDATION:{Enter 2}{Up 10}
return
:r:sbar::
Send SITUATION:{Enter 3}BACKGROUND:{Enter 3}ASSESSMENT:{Enter 3}RECOMENDATION:{Enter 2}{Up 10}
return
::sdjkp::
text := "............................ Jonathan Ploudre, MD. " . A_MMM . " " . A_DD . ", " A_YYYY
clip(text)
return
; Changes ";;" into "-->" to quickly type an arrow
::`;`;::-->

CloseOutlook: 
WinClose, Inbox - jkploudre 
Return
