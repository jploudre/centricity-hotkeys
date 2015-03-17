; Setup
{
CoordMode, Mouse, Window
SetKeyDelay, 30
return

<#Esc::run taskmgr.exe
<#Up::Send {PgUp}
<#Down::Send {PgDn}
<#Backspace::Send {Delete}
RWin::return
LWin::return
#L::return
}

#IfWinActive, Update
`::PatternHotKey(".->GotoChart","..->SwapTextView")
return
[::
Send ^{PgUp}
return
]::
Send ^{PgDn}
return
\::
Send ^e
return
Rwin::
#Space::
Run unfocus.exe
return
F1::PatternHotKey(".->OrderSearch", "..->SignOrders")
F2::PatternHotKey(".->MedSearch", "..->UpdateMeds")
F3::PatternHotKey(".->ProblemSearch", "..->UpdateProblems")
F5::PatternHotKey(".->HPI")
F6::PatternHotKey(".->Preventive", "..->CommittoFlowsheet")
F7::PatternHotKey(".->PMH-SH-CCC", "..->InserttoNote")
F8::PatternHotKey(".->ROS", "..->ROS2")
F9::PatternHotKey(".->PE", "..->PE-URI")
F10::PatternHotKey(".->CPOE", "..->AssessmentsDue")
F11::PatternHotKey(".->PatientInstructions", "..->PrintVisitSummary")
F12::PatternHotKey(".->Prescriptions", "..->SendPrescriptions")

#IfWinActive, End Update
\::PatternHotKey(".->HoldUpdate", "..->SendToBrandie")
return

#IfWinActive, Chart
`::
IfWinExist, Update
WinActivate, Update
IfWinNotExist, Update
{
	WinGetPos,,,,winheight,A
	ypos := winheight - 161
	Click, 13, %ypos%
}
return
; Preventive Update, Assumes in Documents.
#p::
Gosub, OpenAppendType
Send Clin{Down 4}{Enter}
WinWaitActive, Update
Send +{F8}
CitrixSleep()
Send ^{PgDn}
return
; Replies with Web Message. Assumes in Documents.
#r::
Gosub, OpenAppendType
Send Web{Enter}
WinWaitActive, Update
Send +{F8}
return
; CPOE Append. Assumes in Documents.
#c::
Gosub, OpenAppendType
Send CPOE{Enter}
WinWaitActive, Update
Send +{F8}
return

#IfWinActive, Centricity Practice Solution Browser:
\::
Send !{F4}
return

; Add Arrow keys to make organizing problem lists quicker. 
#IfWinActive, Update Problems
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
Enter::
Click, 694, 599
return

#IfWinActive, Append to Document
^s::
Send !s
return

#ifWinActive, Update Orders
F1::PatternHotKey("..->SignOrders")

; End of Window Specific Hotkeys. 
#IfWinActive

/*
Function Key Functions.

F2::
; Assumes in Patient Chart
IfWinActive, Chart
{
Click, 62, 522
CitrixSleep()
MouseClickDrag, Left, 677, 374, 472, 374
Send {Delete}
CitrixSleep()
MouseClickDrag, Left, 677, 491, 472, 491
Send {Delete}
CitrixSleep()
Click, 635 678
CitrixSleep()
WinWaitActive, Centricity Practice,, 1
if (ErrorLevel != 1) 
{
 Send, n   
 WinWaitNotActive, Centricity
}
CitrixSleep()
Send ^f
}
return
*/


+#L::
; Send a patient a blank letter
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


; Shift Ctrl-a: Open Append
^+a::
; Doesn't check location precisely yet, Assumes in Desktop:Documents
; Used with Gesture: Three Finger TipTap Middle

IfWinActive, Chart
	{
	; X position of control is defined from right boarder.
	WinGetPos,,,winwidth,,A
	xpos := winwidth - 205
	Click, %xpos%, 129
	}
return


; Shift Ctrl-O: Open Attachment
^+o::
; Doesn't check location precisely yet, Assumes in Desktop:Documents
; Used with Gesture: Two Finger 'Right TipTap

IfWinActive, Chart
	{
	; X position of control is defined from right boarder.
	WinGetPos,,,winwidth,,A
	xpos := winwidth - 70
	Click, %xpos%, 385
	}
return

; Functions ####################################################################

CitrixSleep(){
Sleep, 150
}
return

OpenAppendType:
Send ^j
WinWaitActive, Append to
Send !F
WinWaitActive, Append Document
return

GotoChart:
WinActivate, Chart
return

SwapTextView:
Send +{F8}
return

SaveUpdate:
Send !s
return

HoldUpdate:
Send !o
return

SendToBrandie:
; Assumes End Update
IfWinActive, End Update
{
	Click, 316, 351
	WinWaitNotActive
	CitrixSleep()
	SendInput Gaylor{Enter}
	CitrixSleep()
	Click, 240, 345
	WinWaitNotActive
	CitrixSleep()
	Send !o
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
WinWaitActive, Update Orders
CitrixSleep()
Click, 253, 287
CitrixSleep()
Click 400, 330
return

SignOrders:
Click, 254, 38
WinWaitActive, Update Orders
CitrixSleep()
Click, 550, 592
return


MedSearch:
Click, 320, 40
WinWaitActive, New Medication
Click, 706, 53
WinWaitActive, Find Medication
return

UpdateMeds:
FindTemplate("HPI-CCC.png")
Click, 931, 585
return

ProblemSearch:
Click, 407, 36
WinWaitActive, New Problem
Send !R
return

UpdateProblems:
FindTemplate("HPI-CCC.png")
Click, 841, 584
return


HPI:
FindTemplate("HPI-CCC.png")
Click, 767, 559
return

Preventive:
FindTemplate("Preventive-Care-Screening.png")
return

CommittoFlowsheet:
FindTemplate("Preventive-Care-Screening.png")
Click, 599, 113
return

PMH-SH-CCC:
FindTemplate("PMH-SH-CCC.png")
return

InserttoNote:
FindTemplate("PMH-SH-CCC.png")
Click, 921, 112
return

ROS:
FindTemplate("ROS-CCC.png")
return

ROS2:
FindTemplate("ROS-CCC.png")
Click, 351, 80
return

PE:
FindTemplate("PE-CCC.png")
if (ErrorLevel = 1)
{
FindTemplate("Pediatric-PE-Age-Specific-CCC.png")
}
Click, 522, 203
return

PE-URI:
FindTemplate("PE-CCC.png")
if (ErrorLevel = 1)
{
FindTemplate("Pediatric-PE-Age-Specific-CCC.png")
}
Click, 522, 203
Send xu{return}
return

CPOE:
FindTemplate("CPOE-A&P-CCC.png")
return

AssessmentsDue:
FindTemplate("CPOE-A&P-CCC.png")
Click, 561, 108
return

PatientInstructions:
FindTemplate("Patient-Instructions-CCC.png")
Click, 594, 341
return

PrintVisitSummary:
FindTemplate("Patient-Instructions-CCC.png")
Click, 891, 130
Click, 
return

Prescriptions:
FindTemplate("Prescriptions.png")
return

SendPrescriptions:
FindTemplate("Prescriptions.png")
Click, 946. 563
return

FindTemplate(imagefile) {
ImageSearch, FoundX, FoundY, 20, 170, 203, 536, *n10 %imagefile%
if (ErrorLevel = 0) {
MouseMove, %FoundX%, %FoundY%
Click 2
CitrixSleep()
CitrixSleep()
CitrixSleep()
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

; PatternHotKey / KeyPressPattern

; PatternHotKey registers one or multiple patterns of a hotkey to either sending keys, going to a
;               label or calling a function with zero or one parameter.
;
; Usage : hotkey::PatternHotKey("command1", ["command2", "command3", length(integer), period(float)])
;
;     where commands match one of the following formats:
;         "pattern:keys"                  ; Maps pattern to send keys
;         "pattern->label"                ; Maps pattern to label (GoSub)
;         "pattern->function()"           ; Maps pattern to function myfunction with
;                                           no parameter
;         "pattern->function(value)"      ; Maps pattern to function myfunction with
;                                           the first parameter equal to 'value'
;
;         and patterns match the following formats:
;             '.' or '0' represents a short press
;             '1' to '9' and 'A' to 'Z' represents a long press of
;                                       the specified length (base 36)
;             '-' or '_' represents a long press of any length
;             '?' represents any press
;             '~' as prefix marks the following string as a
;                 regular expression for the pattern
;
;     length : Maximum length of returned pattern. Automatically detected unless
;              using custom regular expression patterns. Keeping this value to the
;              minimum will speed up keypress detection.
;     period : Amount of time in seconds to wait for additional keypresses.
;
;     e.g. "01->mylabel" maps a short press followed by a 0.2 to 0.4 seconds press to
;                        the 'mylabel' label.
;          "_:{Esc}" maps a long press to sending the Esc key.
;          ".?-_0->myfunction(1)" maps a short press followed by any press followed by
;                                 2 long press to calling 'myfunction(1)'.
;          "~^[6-9A-Z]$->myfunction()" maps the regular expression '^[6-9A-Z]$' (exact
;                                      length match a long press of length '6' to 'Z')
;                                      to calling 'myfunction()'.
PatternHotKey(arguments*)
{
    period = 0.2
    length = 1

    ; Parse input
    for index, argument in arguments
    {
        ; Use any float as period
        if argument is float
            period := argument, continue

        ; Use any integer as length. Automatically calculated
        ; unless using custom patterns ('~' prefix).
        if argument is integer
            length := argument, continue

        ; Check for Send command (':')
        separator := InStr(argument, ":", 1) - 1
        if ( separator >= 0 )
        {
            pattern   := SubStr(argument, 1, separator)
            command    = Send
            parameter := SubStr(argument, separator + 2)
        }
        else
        {
            ; Check for Function or GoSub command ('->')
            separator := InStr(argument, "->", 1) - 1
            if ( separator >= 0 )
            {
                pattern := SubStr(argument, 1, separator)

                call := Trim(SubStr(argument, separator + 3))
                parenthesis := InStr(call, "(", 1, separator) - 1
                if ( parenthesis >= 0 )
                {
                    ; Parse function name and single parameter
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
            ; Short press
            StringReplace, pattern, pattern, ., 0, All

            ; Long press
            StringReplace, pattern, pattern, -, [1-9A-Z], All
            StringReplace, pattern, pattern, _, [1-9A-Z], All

            ; Any press
            StringReplace, pattern, pattern, ?, [0-9A-Z], All

            ; Exact length match
            pattern := "^" . pattern . "$"

            ; Record max pattern length
            if ( length < separator )
                length := separator
        }

        patterns%index%   := pattern
        commands%index%   := command
        parameters%index% := parameter
    }

    ; Record key press pattern
    keypress := KeyPressPattern(length, period)

    ; Try to find matching pattern
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

; KeyPressPattern returns a base-36 string pattern representing the recorded
;                 keypresses for the current hotkey. Each digit represents a
;                 keypress, indicating how much "pressed down time" (in amount
;                 of period) a key was down, with 0 representing a short press.
;
;     length       : Maximum length of returned pattern. Keeping this value to
;                    the minimum will speed up keypress detection.
;     period       : Amount of time in seconds to wait for additional keypresses.
;
;     e.g. (Using the default period of 0.2 seconds)
;          "01" is a short press followed by a 0.2 to 0.4 seconds press
;          "30" is a 0.6 to 0.8 seconds press followed by a short press
;          "000" is triple-click (3 short press)
KeyPressPattern(length = 2, period = 0.2)
{
    ; Find pressed key
    key := RegExReplace(A_ThisHotKey, "[\*\~\$\#\+\!\^]")
    IfInString, key, %A_Space%
        StringTrimLeft, key, key, % InStr(key, A_Space, 1)

    ; Find modifiers
    if key in Alt,Ctrl,Shift,Win
        modifiers := "{L" key "}{R" key "}"

    current = 0
    loop
    {
        ; Wait for key up
        KeyWait %key%, T%period%
        ; If key up was received...
        if ( ! ErrorLevel )
        {
            ; Append code in base 36 to pattern and reset current code
            pattern .= current < 10
                       ? current
                       : Chr(55 + ( current > 36 ? 36 : current ))
            current = 0
        }
        else
            current++

        ; Return pattern if it is of desired length
        if ( StrLen(pattern) >= length )
            return pattern

        ; If key up was received...
        if ( ! ErrorLevel )
        {
            ; Wait for next key down of the same key
            ;
            ; Capslock, mouse buttons and hotkeys using the no reentry ('$' prefix) cannot use
            ; Input. KeyWait is used instead, but it cannot detect cancelled patterns (a different
            ; key is pressed before timeout) or modifier keys.
            if ( key in Capslock, LButton, MButton, RButton or Asc(A_ThisHotkey) = Asc("$") )
            {
                KeyWait, %key%, T%period% D

                ; If key down timed out, return pattern
                if ( ErrorLevel )
                    return pattern
            }
            else
            {
                Input,, T%period% L1 V,{%key%}%modifiers%

                ; If key down timed out, return pattern
                if ( ErrorLevel = "Timeout" )
                    return pattern
                ; If a different key is pressed, cancel pattern
                else if ( ErrorLevel = "Max" )
                    return
                ; If input is cancelled, cancel pattern
                else if ( ErrorLevel = "NewInput" )
                    return
            }
        }
    }
}

; ############################ Abbreviations

::asscy::asymptomatically
::rxvitd::Recommend Vitamin D supplementation (50K IU Once weekly for 12 weeks, recheck serum level in 3 months) Please send to pts pharmacy.
::dnoabx::discussed reasoning behind no antibiotics for this.
::dusptf::reviewed all appropriate USPTF recommendations in checklist format
::dnytanxy::Recommended the patient read New York Times article 'Understanding the Anxious Mind'.
::dmedad::Discussed  on medication adherence issues.
::dpcl::Reviewed a personalized checklist of preventive recommendations based on USPTF.
::cgm::continuous glucose monitor
::nomedad::No significant medication adherence issues noted.
::dcgm::Discussed continuous glucose monitoring.
::pmprv::PMP reviewed, as expected.
::lcmp::We checked your CMP.
::+lcmp::We also checked your CMP.
::3xbp:: Average of 3 automated readings with 60 second pause between readings.
::3bp:: 3 blood pressures were measured 1 minute apart and averaged.
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
