; Setup
{
FirstRun()
IniRead, Buddy, Settings.ini, Preferences, Buddy
ClinicalAssistantName = %Buddy%
IniRead, VisitSummaryType, Settings.ini, Preferences, CPOE-Summaries
VisitSummaryType = %VisitSummaryType%
CoordMode, Mouse, Window
#Persistent
SetKeyDelay, 30
SetTimer, Focus, 100
SetTimer, AdjustMouse, 480000
SendMode Input
Menu, Tray, NoStandard
Menu, Tray, Add, Problems?, SendEmailToJKP
Menu, Tray, Add, Reload, ReloadScript
Menu, Tray, Add, Exit, ExitScript
return

ExitScript:
ExitApp
return

ReloadScript:
Reload
Return

SendEmailToJKP:
OfficeVersion := 7
Loop, 10
{
	IfExist, C:\Program Files\Microsoft Office\Office%OfficeVersion%\outlook.exe
	{
	run, C:\Program Files\Microsoft Office\Office%OfficeVersion%\outlook.exe /c ipm.note /m jploudre@gmail.com&subject=Centricity`%20Hotkeys&body=Heres`%20an`%20issue:
	Exit
	}
	IfExist, C:\Program Files (x86)\Microsoft Office\Office%OfficeVersion%\outlook.exe
	{
	run, C:\Program Files (x86)\Microsoft Office\Office%OfficeVersion%\outlook.exe /c ipm.note /m jploudre@gmail.com&subject=Centricity`%20Hotkeys&body=Heres`%20an`%20issue:
	Exit
	}
	OfficeVersion++
}
MsgBox, 32, Issues?,
(
Would you mind sending me an e-mail?
    

   jploudre@gmail.com
   (Not my work e-mail, thanks!)
   
                -- Jonathan


)
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
MsgBox, 36, Visit Summaries,
(
Do you prefer Visit Summaries with CPOE:
    

   - Yes will print Handout with CPOE
   - No will do plain visit summaries


)
IfMsgBox, Yes
    IniWrite, "CPOE", Settings.ini, Preferences, CPOE-Summaries
IfMsgBox, No
    IniWrite, "No-CPOE", Settings.ini, Preferences, CPOE-Summaries
MsgBox, 64, Thanks,
(
If this is your first time:
    

   - Get some stickers for your keyboard
   - Maybe print the Help file?


)

}
}
}
return

#IfWinActive, Update - ;###########################################################
;; ## Update Hotkeys
;; 

;; * **`:** go back to chart (single), Swap between Text and Form views.
`::PatternHotKey(".->GotoChart","..->SwapTextView"  )
return

;; * **[:** go to previous form
[::
Send ^{PgUp}
return

;; * **]:** go to next form
]::
Send ^{PgDn}
return

;; * **'End' \:** End Update (single), *End to your CA* (double) 
;; 
;; ---
;; 
\::PatternHotKey(".->EndUpdate", "..->EndUpdateToClinicalAssistant")

;; * **F1:** Order Search (single), Sign Orders (double)
F1::PatternHotKey(".->OrderSearch", "..->SignOrders")

;; * **F2:** New Medication Search (single), Update Med List (double) 
F2::PatternHotKey(".->MedSearch", "..->UpdateMeds")

;; * **F3:** New Problem Search (single), Update Problem List (double) 
F3::PatternHotKey(".->ProblemSearch", "..->UpdateProblems")

;; * **F5:** Go to HPI form 
F5::PatternHotKey(".->HPI")

;; * **F6:** Go to Preventive form (single), Commit Preventive Flowsheet (double) 
F6::PatternHotKey(".->Preventive", "..->CommittoFlowsheet", "...->CommittoFlowsheetandSign")

;; * **F7:** Go to Medical Hx form (single), Insert Medical Hx into Note (double) 
F7::PatternHotKey(".->PMH-SH-CCC", "..->InserttoNote")

;; * **F8:** Go to ROS form (single), Go t ROS-2 form (double) 
F8::PatternHotKey(".->ROS", "..->ROS2")

;; * **F9:** Go to PE form (single), Basic CV Exam, Adults Only (Double), URI Exam, Adults Only (Triple), Psych Exam, Adults Only (Long Hold)
F9::PatternHotKey(".->PE", "..->PE-XC", "...->PE-XU", "_->PE-XP")

;; * **F10:** Go to CPEO form (single), Go CPOE, 'Assessments Due' (double) 
F10::PatternHotKey(".->CPOE", "..->AssessmentsDue")

;; * **F11:** Go to Patient Instructions form (single), Print the Visit Summary(double) 
F11::PatternHotKey(".->PatientInstructions", "..->PrintVisitSummary")

;; * **F12:** Go to Prescriptions form (single), Send Prescriptions(double) 
;; 
;; ---
;; 
F12::PatternHotKey(".->Prescriptions", "..->SendPrescriptions")

;; * **Window-Shift-S:** Ends update, Signs, (No Routing to anyone) and back to Chart Desktop.
#+s::
SetTimer, Focus, Off ; prevent strobing.
Gosub, EndUpdate
Gosub, SignUpdate
SetTimer, Focus, On
return

;; * **Window-/:** Go to Chart, Documents Section (to, say, review a scan, test result, consultation.)
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
;; ## End Update Hotkeys
;; 

;; * Right click in routing to remove. Right click "OK" to go back to Chart Desktop
RButton::
MouseGetPos, xpos, ypos
if ( 28 < xpos AND xpos < 515 AND 250 < ypos AND ypos < 331) { ; Routing Names area, right click
    Mouseclick, Left, %xpos%, %ypos%
    Citrixsleep()
    Send !m
}
if ( 354 < xpos AND xpos < 444 AND 499 < ypos AND ypos < 520) { ; 'Route' button, right click
    Mouseclick, Left, %xpos%, %ypos%
    WinWaitActive, Chart - , , 10
    If (ErrorLevel = 0) {
        Gosub, GoChartDesktop
    }
}
else {
    Click right
}
return

;; * **'End' \:** Hold Update (single), *End to your CA* (double) 
!Space::PatternHotKey(".->HoldUpdate", "..->SendToClinicalAssistant")
#Space::PatternHotKey(".->HoldUpdate", "..->SendToClinicalAssistant")
\::PatternHotKey(".->HoldUpdate", "..->SendToClinicalAssistant")
return

;; * **Window-Shift-S:** Signs, (No Routing to anyone) and back to Chart Desktop.
#+s::
#s::
Gosub, SignUpdate
return

#n::
Send !n
return

#IfWinActive, Chart - ;###########################################################
;; ## Chart Hotkeys
;; 

;; * **Space:** Opens item (Single), Signs Document (Double).
Space::PatternHotKey(".->FancyOpen", "..->ChartDocumentSign")

;; * **Window-R:** Reply to a patient with a Web Append.
#r::
OpenAppendType("Web")
return

;; * **Window-Shift-R:** Reply to a patient blank letter
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

;; * **Window-C:** CPOE Append.
#c::
OpenAppendType("CPOE")
return

;; * **`:** Swap between Chart/Chart Desktop/Update
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

;; * **Window-J:** Append document (makes this consistent between Chart, Chart Desktop.)
#j::
If (ImageMouseMove("append"))
    Click
return

;; * **Window-Shift-P:** Preventative Append
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

;; * **Window-/:** Go to Documents Section 
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
;; ## Chart Desktop Hotkeys
;; 

;; * **Space:** Open Item
Space::PatternHotKey(".->FancyOpen")

;; * **Shift-Space:** Open Patient Chart (not the item)
+Space::
If (ImageMouseMove("open"))
    Click
return

;; * **`:** Swap between Chart/Chart Desktop/Update
`::
IfWinExist, Update
WinActivate, Update
IfWinNotExist, Update 
{
    If (ImageMouseMove("chart"))
        Click
}
return

;; * **Window-J:** Append document (makes this consistent between Chart, Chart Desktop.)
#j::
Send ^j
return

;; * **Window-S/Window-Shift-S:** Signs Document
#+s::
#s::
Send ^s
return

;; * **Window-Shift-P:** Preventative Append
#+p::
OpenAppendType("Clinical List Pr")
return

;; * **Window-R:** Reply to a patient with a Web Append.
#r::
Send ^j
OpenAppendType("Web")
return

;; * **Window-E:** E-Rx Append.
#e::
Send ^j
OpenAppendType("* eSM")
return

;; * **Window-C:** CPOE Append.
#c::
Send ^j
OpenAppendType("CPOE")
return

;; * **Window-/:** Go to Documents Section 
#/::
WinGetPos,,,winwidth,winheight,A
ImageSearch, FoundX, FoundY, 0, 112, %winwidth%, %winheight%, *n50 %A_ScriptDir%/files/documents.png
if (ErrorLevel = 0) {
    MouseMove, %FoundX%, %FoundY%
    Click
}
return

#IfWinActive, Centricity Practice Solution Browser: ;###########################################################
;; ## Centricity Practice Solution Browser (Scans)
;; 

;; * **Space:** Page Down (Single), Close Document, Sign (Double)
Space::PatternHotKey(".->DownDocumentViewer", "..->CloseDocumentViewerandSave")

;; * **End \:** Close Document
#Space::
\::
gosub, CloseDocumentViewer
return

;; * **Up/Down Arrows:** Page Up and Down
up::
WinGetPos,,,winwidth,,A
winwidth := winwidth - 10
Click %winwidth%, 67
return

down::
gosub, DownDocumentViewer
return

;; * **Window-S/Window-Shift-S:** Close Document, Sign
#+s::
#s::
Gosub, CloseDocumentViewerandSave
return

;; * **Window-Shift-P:** Close Document, Preventative Append
#+p::
Send !{F4}
WinWaitActive, Chart, , 10
If (ErrorLevel = 0) {
    CitrixSleep
    focusblue()
    OpenAppendType("Clinical List Pr")
}
return


#IfWinActive, Blackbird Content ;###########################################################
;; ## Blackbird Hotkeys
;; 

;; * **Window-Space:** Done
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

;; * **Enter:** Selects

Enter::
Send \
return

#IfWinActive, New Problem - ;###########################################################
;; ## New Problem Hotkeys
;; 

;; * **Window-Space:** Done
#Space::
Click, 434, 532
return

;; * **F3:** Make 30 day Problem, Done.
F3::
Click, 135, 403
CitrixSleep()
Send 30
Click, 434, 532
return

#IfWinActive, Update Problems - ;###########################################################
;; ## Update Problems Hotkeys
;; 

;; * **Arrows:** Moves items up/down/right/left (single). Top/Bottom (double)
Up::PatternHotKey(".->UpdateProblemsUp","_->UpdateProblemsTop","..->UpdateProblemsTop")
Down::PatternHotKey(".->UpdateProblemsDown","_->UpdateProblemsBottom","..->UpdateProblemsBottom")
Left::PatternHotKey(".->UpdateProblemsLeft")
Right::PatternHotKey(".->UpdateProblemsRight")

;; * **Delete/Backspace:** Removes Problem
BackSpace::
Delete::PatternHotKey(".->UpdateProblemsRemove")

;; * Done with 'Enter' or 'End' or Window-Space
\::
#Space::
!Space::
Enter::
Gosub, UpdateProblemsOK
return

;; * Right-click problem to remove. Right click removed problem to change back.
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
;; ## Update Medications Hotkeys
;; 

;; * **F2:** New Medication
F2::
Send !n
return

;; * **Arrows:** Moves items up/down (single). Top/Bottom (double)
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

;; * Done with 'Enter' or Window-Space
#Space::
!Space::
Enter::
click 559, 566
return

;; * **Delete/Backspace:** Removes Medication
BackSpace::
Delete::
Gosub, UpdateMedicationsRemoveMedication
return

;; * Right-click medication to remove. Right click removed meds to change back.
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
WinWaitActive, Remove Medication, , 3
if (ErrorLevel = 0) {
    CitrixSleep()
    Send {Enter}
}
return

#ifWinActive, Update Orders - ;###########################################################
;; ## Update Medications Hotkeys
;; 

;; * **Window-space:** Signs Orders, Done. 
#Space::
!Space::
Click 561, 656
Citrixsleep()
Click 679, 656
gosub, OrdersFixBug
return

;; * **/:** Search Orders 
/::
Click, 253, 287
CitrixSleep()
Click 412, 337
return

;; * **Window-D:** Order Details 
#d::
Click, 341, 290
return

;; * **F1:** Signs Orders (Double). 
F1::PatternHotKey("..->SignOrders")
;; * **F3:** New Problem (Single) Edit Problem (Double).
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
    CitrixSleep()
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

;; * Left Click OK always signs first. (Prevents unsigned orders for x-rays, labs)
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

;; * Right click Order to delete
RButton::
MouseGetPos, xpos, ypos
if ( 16 < xpos AND xpos < 587 AND 117 < ypos AND ypos < 239)
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
;; ## Assessments Due Hotkeys
;; 

;; * You can use many keys to do 'OK' like Enter, Esc, Window-Space
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
;; ## Customize Letter Hotkeys
;; 


;; * **Window-Space:** Print and Save Letter
#Space::
Send !p
WinWaitNotActive, Customize Letter, , 10
if (ErrorLevel = 0) {
    CitrixSleep()
    WinWaitActive, Customize Letter, , 30
    if (ErrorLevel = 0) {
        Citrixsleep()
        Send !s
        WinWaitActive, Route Document, , 30
        if (ErrorLevel = 0) {
            Citrixsleep()
            Send !s
            WinWaitActive, Print, , 30
            if (ErrorLevel = 0) {
                Citrixsleep()
                Click 568, 355
            }
        }
    }
}
return

#ifWinActive, Care Alert Warning ;###########################################################
;; ## Customize Letter Hotkeys
;; 
;; * Space, Window-Space, Enter, 'End': All close the warning

Space::
#Space::
!Space::
Enter::
\::
Send !c
return

#ifWinActive, New Medication ;###########################################################
;; ## New Medication Hotkeys
;; 

;; * **F3:** New Problem 
F3::
Click, 153, 573
return

;; **Associating Diagnosis:** Will Force Association. Can use 'Enter' key to proceed without association.
LButton::
MouseGetPos, xpos, ypos
; If Done, did diagnosis get associated?
if ( 649 < xpos AND xpos < 712 AND 643 < ypos AND ypos < 669)
    {
    WinGetPos,WinX,WinY,,,A
    DiagnosisControlX := WinX + 18
    DiagnosisControlY := WinY + 351
    ; Check if Diagnosis is Highlighted. If Not, Error.
    If (ImageMouseMove("blue-little", 18, 351, 188, 531)) {
        Click
        WinWaitActive, Update -, , 5
        if (ErrorLevel = 0) {
            CitrixSleep()
            Gosub, Prescriptions
        }
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
;; ## Route Document Hotkeys
;; 

;; * **'End' \:** Route document 
\::
Send !R
return

#n::
Send !n
return

;; * **Right-Click:** in recipients to remove routing, Right Click 'Route' to route and go back to Chart Desktop.
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
    WinWaitActive, Chart - , , 5
    If (ErrorLevel = 0) {
        GoSub, GoChartDesktop
    }
}
else {
    Click right
}
return

#IfWinActive, New Routing Information - ;###########################################################
;; ## New Routing Hotkeys
;; 

;; * **Window-Space:** OK, Hold Document, Go back to Chart Desktop
#Space::
!Space::
Click, 239, 352
WinWaitNotActive, , , 10
CitrixSleep()
CitrixSleep()
if (ErrorLevel = 0) {
    If WinActive("End Update") {
        Send !o
        WinWaitActive, Chart - , , 10
        if (ErrorLevel = 0) {
            gosub, GoChartDesktop
            return
        }
    }
}
return

; End of Window Specific Hotkeys.  #########################################
#IfWinActive

;; ## Generic Hotkeys
;; 

;; * **Window-Shift-Q:** Quit All Windows, Log Out -- (End of day) 
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

; Returns Boolean & Moves Mouse if Found
ImageMouseMove(imagename, x1:=-2000, y1:=-2000, x2:=0, y2:=0){
    CoordMode, Pixel, Screen
    CoordMode, Mouse, Screen
    ImagePathandName := A_ScriptDir . "\files\" . imagename . ".PNG"
    if (x1 = -2000 AND y1 = -2000 AND x2 = 0 AND y2 = 0) {
    ImageSearch, FoundX, FoundY, x1, y1, %A_ScreenWidth%, %A_ScreenHeight%, *n10 %ImagePathandName%
    } else {
    ImageSearch, FoundX, FoundY, x1, y1, x2, y2, *n10 %ImagePathandName%
    }
    if (ErrorLevel = 0) {
        MouseMove, %FoundX%, %FoundY%
        CoordMode, Pixel, Window
        CoordMode, Mouse, Window
        return 1
    }
    CoordMode, Pixel, Window
    CoordMode, Mouse, Window
    ; If image is not found, do not continue Hotkey that called. 
    if (ErrorLevel = 1) {
    return 0
    }
}


OpenAppendType(searchtext){
    SetTimer, Focus, Off
    ifWinActive, Chart Desktop -
    {
        Send ^j
    }
    ifWinActive, Chart -
    {
        if (ImageMouseMove("append"))
            Click
    }
    WinWaitActive, Append to, , 3
    if (ErrorLevel = 0) {
        CitrixSleep()
        Send !F
        WinWaitActive, Append Document, , 5
        if (ErrorLevel = 0) {
            CitrixSleep()
            clip(searchtext)
            Send {Enter}
            WinWaitActive, Update, , 5
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
If (ImageMouseMove("blue", 200, 50, %winwidth%, %winheight%)) {
        Click
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
WinWaitActive, Chart -, , 10
if (ErrorLevel = 0) {
    gosub, GoChartDesktop
    WinWaitActive, Chart Desktop -, , 10,
}
return

EndUpdate:
Send ^e
WinWaitActive, End Update, , 2
; Sometimes Fails, Try a few times?
if (ErrorLevel = 1) {
    Send ^e
    WinWaitActive, End Update, , 1
    if (ErrorLevel = 1) {
    Send ^e
    }
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
    CitrixSleep()
	ClicktoNewWindow(316, 351,New Routing Information)
    CitrixSleep()
	Clip(ClinicalAssistantName)
    Send {Enter}
	CitrixSleep()
	ClicktoNewWindow(240, 345,End Update)
	CitrixSleep()
	Send !o
    WinWaitActive, Chart -, , 15
    if (ErrorLevel = 0) {
        CitrixSleep()
        gosub, GoChartDesktop
    }
}
else
{
return	
}
return

ClicktoNewWindow(x,y,WinTitle){
    Loop, 4
    {
        Click, %x%, %y%
        WinWaitActive, %WinTitle%, , 2
        if (ErrorLevel= 0) {
        CitrixSleep()
        return  
        }
        if (ErrorLevel=1){
        continue
        }
    }
    if (%A_Index%= 4) {
        Exit
    }
    return
}

FancyOpen:
WinGetPos,,,winwidth,winheight,A
ImageSearch, FoundX, FoundY, 200, 50, %winwidth%, %winheight%, *n10 %A_ScriptDir%/files/attach.png
if (ErrorLevel = 0) {
    ImageSearch, FoundX1, FoundY1, 200, 50, %winwidth%, %winheight%, *n10 %A_ScriptDir%/files/paperclip.png
    if (ErrorLevel = 0) {
    MouseMove, %FoundX%, %FoundY%
    Click
    WinWaitNotActive, Chart, , 3
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
        WinWaitNotActive, Chart, , 3
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
WinWaitActive, Update Orders, , 3
if (ErrorLevel = 0) {
    CitrixSleep()
    Click, 253, 287
    CitrixSleep()
    Click 412, 337
}
return




MedSearch:
Click, 524, 38
WinWaitActive, New Medication, , 3
if (ErrorLevel = 0) {
	CitrixSleep()
	Click, 718, 81
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
If (ImageMouseMove("chart-desktop"))
    Click
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
CitrixSleep()
Click, 921, 112
return

ROS:
FindTemplate("ROS-CCC")
return

ROS2:
FindTemplate("ROS-CCC")
CitrixSleep()
Click, 351, 80
return

; Special Case, Needs to Search for Peds if it doesn't work. 
PE:
FindPETemplate()

return

PE-XU:
FindPETemplate()
CitrixSleep()
CitrixSleep()
CitrixSleep()
Click, 401, 136
CitrixSleep()
CitrixSleep()
CitrixSleep()
Click, 401, 321
CitrixSleep()
CitrixSleep()
CitrixSleep()
Click, 401, 408
CitrixSleep()
CitrixSleep()
CitrixSleep()
Click, 401, 500
CitrixSleep()
CitrixSleep()
CitrixSleep()
Click, 401, 586
CitrixSleep()
CitrixSleep()
CitrixSleep()
    Click, 397, 80
    CitrixSleep()
    CitrixSleep()
    CitrixSleep()
    CitrixSleep()
    CitrixSleep()
    CitrixSleep()
    CitrixSleep()
Click, 401, 110
CitrixSleep()
CitrixSleep()
CitrixSleep()
Click, 401, 337
CitrixSleep()
CitrixSleep()
CitrixSleep()
Click, 401, 509
CitrixSleep()
CitrixSleep()
CitrixSleep()
    Click, 656, 80
    CitrixSleep()
    CitrixSleep()
    CitrixSleep()
    CitrixSleep()
    CitrixSleep()
    CitrixSleep()
    CitrixSleep()
Click, 401, 298
CitrixSleep()
CitrixSleep()
CitrixSleep()
Click, 401, 531
return

PE-XC:
FindPETemplate()
CitrixSleep()
CitrixSleep()
CitrixSleep()
Click, 401, 136
CitrixSleep()
CitrixSleep()
CitrixSleep()
    Click, 397, 80
    CitrixSleep()
    CitrixSleep()
    CitrixSleep()
    CitrixSleep()
    CitrixSleep()
    CitrixSleep()
    CitrixSleep()
Click, 401, 110
CitrixSleep()
CitrixSleep()
CitrixSleep()
Click, 401, 337
CitrixSleep()
CitrixSleep()
CitrixSleep()
Click, 401, 509
CitrixSleep()
CitrixSleep()
CitrixSleep()
    Click, 656, 80
    CitrixSleep()
    CitrixSleep()
    CitrixSleep()
    CitrixSleep()
    CitrixSleep()
    CitrixSleep()
    CitrixSleep()
Click, 401, 531
return

PE-XP:
FindPETemplate()
CitrixSleep()
CitrixSleep()
CitrixSleep()
Click, 401, 136
CitrixSleep()
CitrixSleep()
CitrixSleep()
    Click, 656, 80
    CitrixSleep()
    CitrixSleep()
    CitrixSleep()
    CitrixSleep()
    CitrixSleep()
    CitrixSleep()
    CitrixSleep()
Click, 401, 531
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
If (VisitSummaryType = "No-CPOE") {
    FindTemplate("Patient-Instructions-CCC")
    Click, 891, 130
}
if (VisitSummaryType = "CPOE") {
    FindTemplate("CPOE-A&P-CCC")
    ClicktoNewWindow(933, 346, Print Patient)
    Sleep, 1500
    Clip("Clinical Visit Summary with CPOE Assessments")
    Send {Enter}
    Sleep, 3000
    Send !p
    Sleep, 1000
    Send !{F4}
}
return


Prescriptions:
FindTemplate("Prescriptions")
MouseMove, 946, 563
return

SendPrescriptions:
FindTemplate("Prescriptions")
Click, 946, 563
return

FindTemplate(template) {
ImageSearch, FoundX, FoundY, 20, 170, 203, 536, *n10 %A_ScriptDir%/files/%template%.png
if (ErrorLevel = 0) {
	MouseMove, %FoundX%, %FoundY%
	Click 2
	CitrixSleep()
    Click 2
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

FindPETemplate() {
ImageSearch, FoundX, FoundY, 20, 170, 203, 536, *n10 %A_ScriptDir%/files/PE-CCC.png
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
	ImageSearch, FoundX, FoundY, 20, 170, 203, 536, *n10 %A_ScriptDir%/files/Pediatric-PE-Age-Specific-CCC.png
    if (ErrorLevel = 0) {
        MouseMove, %FoundX%, %FoundY%
        Click 2
        CitrixSleep()
        CitrixSleep()
        CitrixSleep()
        Click 2
    MouseMove, 500, 0, 0, R
    Exit
}
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
Citrixsleep()
Citrixsleep()
IfWinExist, Chart Desktop -
    WinActivate, Chart Desktop -
IfWinExist, Chart -
    WinActivate, Chart -
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
WinWaitActive, Remove Problem -, , 5
if (ErrorLevel = 0) {
    CitrixSleep()
    Send {Enter}
}
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

;; ## Typing Aids (Quicktexts that work everywhere -- not just where CPS allows)
;; 

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
;; * **SBAR** a template for call notes

:r:sbar::
Send SITUATION:{Enter 3}BACKGROUND:{Enter 3}ASSESSMENT:{Enter 3}RECOMENDATION:{Enter 2}{Up 10}
return

::sdjkp::
text := "............................ Jonathan Ploudre, MD. " . A_MMM . " " . A_DD . ", " A_YYYY
clip(text)
return
;; * Changes ";;" into "-->" to quickly type an arrow
::`;`;::-->

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
