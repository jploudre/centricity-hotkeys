#Persistent
SetTimer, Focus, 100
return

Focus:
if WinActive("focus.ahk") or WinActive("Update Problems -") or WinActive("Update Medications -") or WinActive("Update Orders -") or WinActive("Update Orders -") or WinActive("Append to Document") or WinActive("Assessments Due") or WinActive("Customize Letter") or WinActive("End Update") or WinActive("Care Alert Warning -") or WinActive("New Medication") or WinActive("New Problem") or WinActive("Find Problem") or WinActive("Find Medication")
{
active_window := WinExist("A")
WinGetPos, Xpos, Ypos, winwidth , winheight , A

Gui,1: +LastFound -Caption +ToolWindow +E0x20 +AlwaysOnTop
Gui,1: Color,b8b8b8
gui1h := Ypos + winheight 
Gui,1: Show, noactivate w%Xpos% h%gui1h% x0 y0,FocusWindow1
WinGet,windowID1,ID
WinSet, Transparent, 150, ahk_id %windowID1%

Gui,2: +LastFound -Caption +ToolWindow +E0x20 +AlwaysOnTop
Gui,2: Color,b8b8b8
gui2w := A_Screenwidth - Xpos
Gui,2: Show, noactivate w%gui2w% h%Ypos% x%Xpos% y0,FocusWindow2
WinGet,windowID2,ID
WinSet, Transparent, 150, ahk_id %windowID2%

Gui,3: +LastFound -Caption +ToolWindow +E0x20 +AlwaysOnTop
Gui,3: Color,b8b8b8
gui3w := A_Screenwidth - Xpos - winwidth
gui3x := Xpos + winwidth
gui3h := A_Screenheight - Ypos - 30
Gui,3: Show, noactivate w%gui3w% h%gui3h% x%gui3x% y%Ypos%,FocusWindow3
WinGet,windowID3,ID
WinSet, Transparent, 150, ahk_id %windowID3%

Gui,4: +LastFound -Caption +ToolWindow +E0x20 +AlwaysOnTop
Gui,4: Color,b8b8b8
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

