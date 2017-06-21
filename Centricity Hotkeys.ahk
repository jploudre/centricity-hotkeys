; Setup
{
ClinicalAssistantName = "Handy"
CoordMode, Mouse, Window
#Persistent
SetKeyDelay, 30
SetTimer, CloseOutlook, 5000 
SetTimer, Focus, 100
SetTimer, AdjustMouse, 480000
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

#o::PatternHotKey(".->OrderSearch", "..->SignOrders")
F1::PatternHotKey(".->OrderSearch", "..->SignOrders")

#m::PatternHotKey(".->MedSearch", "..->UpdateMeds")
F2::PatternHotKey(".->MedSearch", "..->UpdateMeds")

#p::PatternHotKey(".->ProblemSearch", "..->UpdateProblems")
F3::PatternHotKey(".->ProblemSearch", "..->UpdateProblems")

#h::PatternHotKey(".->HPI")
F5::PatternHotKey(".->HPI")

#q::PatternHotKey(".->Preventive", "..->CommittoFlowsheet")
F6::PatternHotKey(".->Preventive", "..->CommittoFlowsheet", "...->CommittoFlowsheetandSign")

#z::PatternHotKey(".->PMH-SH-CCC")
F7::PatternHotKey(".->PMH-SH-CCC")

F8::PatternHotKey(".->ROS", "..->ROS2")

#x::PatternHotKey(".->PE", "..->PE-XC", "...->PE-XU", "_->PE-XP")
F9::PatternHotKey(".->PE", "..->PE-XC", "...->PE-XU", "_->PE-XP")

#c::PatternHotKey(".->CPOE", "..->AssessmentsDue")
F10::PatternHotKey(".->CPOE", "..->AssessmentsDue")

#v::PatternHotKey(".->PatientInstructions", "..->PrintVisitSummary")
F11::PatternHotKey(".->PatientInstructions", "..->PrintVisitSummary")

#r::PatternHotKey(".->Prescriptions", "..->SendPrescriptions")
F12::PatternHotKey(".->Prescriptions", "..->SendPrescriptions")

; Ends and signs an update. 
#+s::
SetTimer, Focus, Off ; prevent strobing.
Gosub, EndUpdate
Gosub, SignUpdate
SetTimer, Focus, On
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

;Process Clipboard, Paste back
#+x::
clipboard =
Send ^x
ClipWait  
Clipboard := RegExReplace(Clipboard, "; ", "`r`n") 
Clipboard := RegExReplace(Clipboard, ";", "`r`n") 

; Any whitespace at beginning of lines?
Clipboard := RegexReplace(Clipboard, "m)^\s(.*)$","$1")
Clipboard := RegexReplace(Clipboard, "m)^\s(.*)$","$1")
Clipboard := RegexReplace(Clipboard, "m)^\s(.*)$","$1")

Clipboard := RegexReplace(Clipboard, "m)^Cancer History:(.*)$","$1")
Clipboard := RegexReplace(Clipboard, "m)^Dermatology:(.*)$","$1")
Clipboard := RegexReplace(Clipboard, "m)^Endocrine:(.*)$","$1")
Clipboard := RegexReplace(Clipboard, "m)^Hematology:(.*)$","$1")
Clipboard := RegexReplace(Clipboard, "m)^Cardiovascular:(.*)$","$1")
Clipboard := RegexReplace(Clipboard, "m)^Pulmonary:(.*)$","$1")
Clipboard := RegexReplace(Clipboard, "m)^Other Problems:(.*)$","$1")
Clipboard := RegexReplace(Clipboard, "m)^Musculoskeletal:(.*)$","$1")
Clipboard := RegexReplace(Clipboard, "m)^Neurology/Genetic Dz.:(.*)$","$1")
Clipboard := RegexReplace(Clipboard, "m)^Mental Health History:(.*)$","$1")
Clipboard := RegexReplace(Clipboard, "m)Mood DisordersDepression","Depression")
Clipboard := RegexReplace(Clipboard, "m)^Renal/Genital/Urinary:(.*)$","$1")
Clipboard := RegexReplace(Clipboard, "m)^Infectious Disease History:(.*)$","$1")
Clipboard := RegexReplace(Clipboard, "m)^Substance Abuse History:(.*)$","$1")
Clipboard := RegexReplace(Clipboard, "m)^GI:(.*)$","$1")
Clipboard := RegexReplace(Clipboard, "m)^Gastrointestinal:(.*)$","$1")
Clipboard := RegexReplace(Clipboard, "m)^Allergy/Immunology:(.*)$","$1")
Clipboard := RegexReplace(Clipboard, "m)^Procedures:(.*)$","$1")

; Remove Current Medical Providers -- I have not been maintaining for years. 
Clipboard := RegexReplace(Clipboard, "m)^CURRENT MEDICAL PROVIDERS:(.*)$","")

Clipboard := RegexReplace(Clipboard, "m)^\s(.*)$","$1")
Clipboard := RegexReplace(Clipboard, "m)^\s(.*)$","$1")

; Remove Blank Lines
Loop
{
    StringReplace, Clipboard, Clipboard, `r`n`r`n, `r`n, UseErrorLevel
    if ErrorLevel = 0  ; No more replacements needed.
        break
}
Send ^v
return


#IfWinActive, End Update ;###########################################################

RButton::
MouseGetPos, xpos, ypos
; remove routing name
if ( 28 < xpos AND xpos < 515 AND 250 < ypos AND ypos < 331) { ; Routing Names area, right click
    Mouseclick, Left, %xpos%, %ypos%
    Citrixsleep()
    Send !m
}
; I'm Done
if ( 354 < xpos AND xpos < 444 AND 499 < ypos AND ypos < 520) { ; 'Route' button, right click
    Mouseclick, Left, %xpos%, %ypos%
    WinWaitNotActive
    Gosub, GoChartDesktop
}
else {
    Click right
}
return

!Space::PatternHotKey(".->HoldUpdate", "..->SendToClinicalAssistant")
#Space::PatternHotKey(".->HoldUpdate", "..->SendToClinicalAssistant")
\::PatternHotKey(".->HoldUpdate", "..->SendToClinicalAssistant")
return
#+s::
#s::
Gosub, SignUpdate
return

#n::
Send !n
return

#IfWinActive, Chart - ;###########################################################
Space::PatternHotKey(".->FancyOpen", "..->ChartDocumentSign")

#o::
Gosub, FancyOpen
return

`::
IfWinExist, Update
WinActivate, Update
IfWinNotExist, Update
gosub, GoChartDesktop
return

; Sign a chart document
#s::
Gosub, ChartDocumentSign
return

ChartDocumentSign:
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

#/::
WinGetPos,,,winwidth,winheight,A
ImageSearch, FoundX, FoundY, 0, 112, %winwidth%, %winheight%, *n50 %A_ScriptDir%/files/documents.png
if (ErrorLevel = 0) {
    MouseMove, %FoundX%, %FoundY%
    Click
    citrixsleep()
    Click, 255, 212
}
return


#IfWinActive, Chart Desktop - ;###########################################################

Space::PatternHotKey(".->FancyOpen")

+Space::
WinGetPos,,,winwidth,winheight,A
ImageSearch, FoundX, FoundY, 0, 0, %winwidth%, %winheight%, *n50 %A_ScriptDir%/files/open.png
if (ErrorLevel = 0) {
    MouseMove, %FoundX%, %FoundY%
    Click
    WinWaitNotActive, Chart Desktop
}
return

#o::
Gosub, FancyOpen
return

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

#/::
WinGetPos,,,winwidth,winheight,A
ImageSearch, FoundX, FoundY, 0, 112, %winwidth%, %winheight%, *n50 %A_ScriptDir%/files/documents.png
if (ErrorLevel = 0) {
    MouseMove, %FoundX%, %FoundY%
    Click
}
return

#IfWinActive, Centricity Practice Solution Browser: ;###########################################################
Space::PatternHotKey(".->DownDocumentViewer", "..->CloseDocumentViewerandSave")

#Space::
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

; Close and Sign
#s::
Gosub, CloseDocumentViewerandSave
return

#+p::
Send !{F4}
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

#IfWinActive, New Problem - ;###########################################################
#Enter::
#Space::
Click, 434, 532
return
F3::
#3::
Click, 135, 403
CitrixSleep()
Send 30
Click, 434, 532
return

#IfWinActive, Update Problems - ;###########################################################
Up::PatternHotKey(".->UpdateProblemsUp","_->UpdateProblemsTop","..->UpdateProblemsTop")
Down::PatternHotKey(".->UpdateProblemsDown","_->UpdateProblemsBottom","..->UpdateProblemsBottom")
Left::PatternHotKey(".->UpdateProblemsLeft")
Right::PatternHotKey(".->UpdateProblemsRight")

BackSpace::
Delete::PatternHotKey(".->UpdateProblemsRemove")
\::
#Space::
!Space::
Enter::
Gosub, UpdateProblemsOK
return

; Control Locations are relative for width/height
RButton::
MouseGetPos, xpos, ypos
; Problems
if ( 19 < xpos AND xpos < 995 AND 102 < ypos AND ypos < 251)
    {
    ; Click then Edit
    Click %xpos%, %ypos%
    Citrixsleep()
    GoSub, UpdateProblemsRemove
    MouseMove, %xpos%, %ypos%
    }
; Effect of Update
else if ( 19 < xpos AND xpos < 998 AND 412 < ypos AND ypos < 546)
    {
    Click %xpos%, %ypos%
    Citrixsleep()
    Send !k
    }    
else
    {
    Click    
    }
return


#IfWinActive, Update Medications - ;###########################################################
#n::
F2::
Send !n
return
Up::PatternHotKey(".->UpdateMedsUp","_->UpdateMedsTop","..->UpdateMedsTop")
Down::PatternHotKey(".->UpdateMedsDown","_->UpdateMedsBottom","..->UpdateMedsBottom")
UpdateMedsUp:
Click, 827, 84
return
UpdateMedsTop:
Click, 827, 202
return
UpdateMedsDown:
Click, 827, 114
return
UpdateMedsBottom:
Click, 827, 232
return
Left::
Click, 827, 147
return
Right::
Click, 827, 171
return

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
Gosub, UpdateMedicationsRemoveMedication
return

RButton::
MouseGetPos, xpos, ypos
if ( 19 < xpos AND xpos < 790 AND 94 < ypos AND ypos < 285)
    {
    Click %xpos%, %ypos%
    Citrixsleep()
    GoSub, UpdateMedicationsRemoveMedication
    }
else if ( 19 < xpos AND xpos < 859 AND 375 < ypos AND ypos < 516)
    {
    Click %xpos%, %ypos%
    Citrixsleep()
    Send !k
    }
else    
    {
    Click    
    }
return

UpdateMedicationsRemoveMedication:
Send !r
WinWaitActive, Remove Medication
CitrixSleep()
Send {Enter}
WinWaitNotActive
return

#ifWinActive, Update Orders - ;###########################################################

#Space::
!Space::
Click 561, 656
Citrixsleep()
Click 679, 656
gosub, OrdersFixBug
return
#s::
CLick 561, 656
gosub, OrdersFixBug
return

/::
Click, 253, 287
CitrixSleep()
Click 412, 337
return

; Order Details
#d::
Click, 341, 290
return

F1::PatternHotKey("..->SignOrders")
F3::PatternHotKey(".->OrdersNewProblem","..->OrdersEditProblem")

SignOrders:
Click, 254, 38
WinWaitActive, Update Orders, , 3 ; Timeout
if (ErrorLevel = 0) {
	CitrixSleep()
	Click, 561, 653
}
return

; After 'Ok' CPS goes to 'Chart' not back to 'Update'
OrdersFixBug:
WinWaitActive, Chart, , 4
if (ErrorLevel = 0) {
IfWinExist, Update
WinActivate, Update
}
return

OrdersNewProblem:
Click, 641, 261
return

OrdersEditProblem:
Click, 726, 263
return

OrdersDeleteOrder:
Click 51, 261
return

LButton::
MouseGetPos, xpos, ypos
if ( 638 < xpos AND xpos < 709 AND 647 < ypos AND ypos < 667)
    {
    ; Click Sign first
    Click 561, 656
    Citrixsleep()
    Click %xpos%, %ypos%
    gosub, OrdersFixBug
    }
else
    {
    Click    
    }
return

RButton::
MouseGetPos, xpos, ypos
if ( 619 < xpos AND xpos < 771 AND 81 < ypos AND ypos < 238)
    {
    ; Click then Edit
    Click %xpos%, %ypos%
    Citrixsleep()
    GoSub, OrdersEditProblem
    }
else if ( 16 < xpos AND xpos < 587 AND 117 < ypos AND ypos < 239)
    {
    Click %xpos%, %ypos%
    Citrixsleep()
    gosub, OrdersDeleteOrder
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
Space::
#Space::
!Space::
Enter::
\::
Send !c
return

#ifWinActive, New Medication ;###########################################################

F3::
Click, 153, 573
return

LButton::
MouseGetPos, xpos, ypos
; If Done, did diagnosis get associated?
if ( 649 < xpos AND xpos < 712 AND 643 < ypos AND ypos < 669)
    {
    WinGetPos,WinX,WinY,,,A
    DiagnosisControlX := WinX + 18
    DiagnosisControlY := WinY + 351
    ; Check if Diagnosis is Highlighted. If Not, Error.
    ImageSearch,,, 18, 351, 188, 531, *n10 %A_ScriptDir%/files/blue-little.png
    if (ErrorLevel = 0) {
        Click %xpos%, %ypos%
        WinWaitActive, Update -
        CitrixSleep()
        Gosub, Prescriptions
        }
    if (ErrorLevel > 0) { 
        Gui,5: +LastFound -Caption +ToolWindow +E0x20 +AlwaysOnTop
        Gui,5: Color,008080 
        Gui,5: Show, noactivate x%DiagnosisControlX% y%DiagnosisControlY% w188 h180, AssociateMeds
        WinGet,windowID1,ID
        WinSet, Transparent, 150, ahk_id %windowID1%
        CitrixSleep()
        Gui,5: Hide
        CitrixSleep()
        Gui,5: Show, noactivate x%DiagnosisControlX% y%DiagnosisControlY% w188 h180, AssociateMeds
        CitrixSleep()
        Gui,5: Destroy
        }
    }
else 
{
    Click
}
return

#IfWinActive, Route Document - ;###########################################################

\::
Send !R
return

#n::
Send !n
return


RButton::
MouseGetPos, xpos, ypos
; remove routing name
if ( 28 < xpos AND xpos < 515 AND 171 < ypos AND ypos < 255) { ; Routing Names area, right click
    Mouseclick, Left, %xpos%, %ypos%
    Citrixsleep()
    Send !m
}
; I'm Done
if ( 373 < xpos AND xpos < 445 AND 310 < ypos AND ypos < 331) { ; 'Route' button, right click
    Mouseclick, Left, %xpos%, %ypos%
    WinWaitNotActive
    GoSub, GoChartDesktop
}
else {
    Click right
}
return

#IfWinActive, New Routing Information - ;###########################################################

#Space::
!Space::
Click, 239, 352
WinWaitNotActive
CitrixSleep()
If WinActive("End Update")
    {
    Send !o
    gosub, GoChartDesktop
    return
    }


return

; End of Window Specific Hotkeys.  #########################################
#IfWinActive

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
+#R Up::
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
    SetTimer, Focus, Off
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
            clip(searchtext)
            Send {Enter}
            WinWaitActive, Update ; no timeout needed
            if (ErrorLevel = 0) {
            CitrixSleep()
            Send +{F8}
            CitrixSleep()
            }
        }
    }
    SetTimer, Focus, On
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
Send ^e
WinWaitActive, End Update, , 2
; Sometimes Fails
if (ErrorLevel = 1) {
    Send ^e
}
CitrixSleep()
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
        } ; End open.png
}
}
ImageSearch, , , 0, 0, %winwidth%, %winheight%, *n10 %A_ScriptDir%/files/pencil.png
if (ErrorLevel = 0) {
    xpos := winwidth -30
    ypos := winheight -76
    MouseMove, %xpos%, %ypos%
    Click
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

CommittoFlowsheetandSign:
SetTimer, Focus, Off ; prevent strobing.
Gosub, CommittoFlowsheet
CitrixSleep()
Gosub, EndUpdate
CitrixSleep()
Gosub, SignUpdate
CitrixSleep()
SetTimer, Focus, On
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


; Update Problems Functions

UpdateProblemsTop:
ClickFromRightTop(55, 102)
return
UpdateProblemsBottom:
ClickFromRightTop(55, 189)
return
UpdateProblemsUp:
ClickFromRightTop(55, 122)
return
UpdateProblemsDown:
ClickFromRightTop(55, 166)
return
UpdateProblemsLeft:
ClickFromRightTop(79, 143)
return
UpdateProblemsRight:
ClickFromRightTop(28, 143)
return
UpdateProblemsOK:
ClickFromRightBottom(212, 27)
return
UpdateProblemsRemove:
Click, 508, 572
WinWaitNotActive
CitrixSleep()
Send {Enter}
return

; For Buttons Positioned on Window Size
ClickFromRightTop(RelativeX, RelativeY) {
WinGetPos, , , winwidth, winheight
xpos := winwidth - RelativeX
Click, %xpos%, %RelativeY%
}
return

ClickFromRightBottom(RelativeX, RelativeY) {
WinGetPos, , , winwidth, winheight
xpos := winwidth - RelativeX
ypos := winheight - RelativeY
Click, %xpos%, %ypos%
}
return

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

AdjustMouse:
if (A_TimeIdlePhysical <= 1800000)
{
MouseClick, WU
}
return

::ujkp::
text := "Upcoming Appointment. ............................ Jonathan Ploudre, MD. " . A_MMM . " " . A_DD . ", " A_YYYY
clip(text)
CitrixSleep()
CitrixSleep()
Send !s
return
::sljkp::
text := "Send Letter with Results. ............................ Jonathan Ploudre, MD. " . A_MMM . " " . A_DD . ", " A_YYYY
clip(text)
CitrixSleep()
CitrixSleep()
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

; Excel name switcher
^Insert::
Send ^+{Right}
Send ^x
Send ^{Right 3}
Send %A_Space%^v{Enter}
return

CloseOutlook: 
WinClose, Inbox - jkploudre 
Return

Focus:
if WinActive("Append to Document") or WinActive("Assessments Due") or WinActive("Customize Letter") or WinActive("End Update") or WinActive("Care Alert Warning -") or WinActive("Find Medication") or WinActive("Change Medication") or WinActive("New Problem") or WinActive("Find Problem")  or WinActive("Edit Problem")   or WinActive("Edit Routing")   or WinActive("New Routing")  or WinActive("Centricity Practice Solution Browser") or WinActive("Route Document")
{
active_window := WinExist("A")
WinGetPos, Xpos, Ypos, winwidth , winheight , A

; Citrix treats some windows differently
if WinActive("Append to Document")  or WinActive("Customize Letter") or WinActive("End Update") or WinActive("Care Alert Warning -") or WinActive("Find Medication")  or WinActive("Change Medication")  or WinActive("Edit Problem") or WinActive("Edit Routing") or WinActive("New Routing") or WinActive("Route Document") {
Ypos := Ypos + 23
winheight := winheight - 23
}
Gui,1: +LastFound -Caption +ToolWindow +E0x20 +AlwaysOnTop
Gui,1: Color,008080
gui1h := Ypos + winheight 
Gui,1: Show, noactivate w%Xpos% h%gui1h% x0 y0,FocusWindow1
WinGet,windowID1,ID
WinSet, Transparent, 150, ahk_id %windowID1%

Gui,2: +LastFound -Caption +ToolWindow +E0x20 +AlwaysOnTop
Gui,2: Color,008080
gui2w := A_Screenwidth - Xpos
Gui,2: Show, noactivate w%gui2w% h%Ypos% x%Xpos% y0,FocusWindow2
WinGet,windowID2,ID
WinSet, Transparent, 150, ahk_id %windowID2%

Gui,3: +LastFound -Caption +ToolWindow +E0x20 +AlwaysOnTop
Gui,3: Color,008080
gui3w := A_Screenwidth - Xpos - winwidth
gui3x := Xpos + winwidth
gui3h := A_Screenheight - Ypos - 30
Gui,3: Show, noactivate w%gui3w% h%gui3h% x%gui3x% y%Ypos%,FocusWindow3
WinGet,windowID3,ID
WinSet, Transparent, 150, ahk_id %windowID3%

Gui,4: +LastFound -Caption +ToolWindow +E0x20 +AlwaysOnTop
Gui,4: Color,008080
gui4y := Ypos + winheight
gui4h := A_Screenheight - Ypos - winheight -30
gui4w := Xpos + winwidth
Gui,4: Show, noactivate w%gui4w% h%gui4h% x0 y%gui4y%,FocusWindow4
WinGet,windowID4,ID
WinSet, Transparent, 150, ahk_id %windowID4%

WinActivate,ahk_id %active_window%
}
else {
Gui,1: Destroy
Gui,2: Destroy
Gui,3: Destroy
Gui,4: Destroy	
}
return
