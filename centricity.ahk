; Setup
{
CoordMode, Mouse, Window
#Persistent
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

; Preventive Append. Assumes in Documents.
#p::
OpenAppendType("Clinical List Pr")
return

; Replies with Web Message. Assumes in Documents.
#r::
OpenAppendType("Web")
return

; CPOE Append. Assumes in Documents.
#c::
OpenAppendType("CPOE")
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

OpenAppendType(searchtext){
    Send ^j
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
WinWaitActive, Update Orders, , 3 ; Timeout
if (ErrorLevel = 0) {
    CitrixSleep()
    Click, 253, 287
    CitrixSleep()
    Click 400, 330
}
return

SignOrders:
Click, 254, 38
WinWaitActive, Update Orders, , 3 ; Timeout
if (ErrorLevel = 0) {
	CitrixSleep()
	Click, 550, 592
}
return


MedSearch:
Click, 320, 40
WinWaitActive, New Medication, , 3 ; Timeout
if (ErrorLevel = 0) {
	CitrixSleep()
	Click, 706, 53
	WinWaitActive, Find Medication
}
return

UpdateMeds:
FindTemplate("HPI-CCC")
Click, 931, 585
return

ProblemSearch:
Click, 407, 36
WinWaitActive, New Problem
Send !R
return

UpdateProblems:
FindTemplate("HPI-CCC")
Click, 841, 584
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

; ###################### Broken PE second search
PE:
FindTemplate("PE-CCC")
if (ErrorLevel = 1)
{
FindTemplate("Pediatric-PE-Age-Specific-CCC")
}
Click, 522, 203
return

PE-URI:
FindTemplate("PE-CCC.png")
if (ErrorLevel = 1)
{
FindTemplate("Pediatric-PE-Age-Specific-CCC")
}
Click, 522, 203
Send xu{return}
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
ImageSearch, FoundX, FoundY, 20, 170, 203, 536, *n10 %template%.png
if (ErrorLevel = 0) {
	MouseMove, %FoundX%, %FoundY%
	Click 2
	CitrixSleep()
	CitrixSleep()
	CitrixSleep()
}
; if template not found, is it already selected?
if (ErrorLevel = 1) {
	ImageSearch, FoundX, FoundY, 20, 170, 203, 536, *n10 %template%-highlighted.png
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