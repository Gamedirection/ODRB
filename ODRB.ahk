; Creator: GameDirection - Alex Sierputowski
; Date Created: 5/1/23
; Script Name: Outlooks Dynamic Reading Window
; Script Description: 
;This script is designed to dynamically update the reading pane in Outlook based on the window width. Before running the script, please make sure that Outlook.exe is already open. Additionally, please ensure that your third "Quick Access Toolbar" is set to "Reading Pane." If not, you can find the Reading Pane in the "View" tab at the top. Right-click and select "Add to Quick Access Toolbar" to add it.

;Once you have confirmed these settings, you can run the script. The script will automatically adjust the reading pane based on the window width.

;Please note that there is a known bug where the Quick Access Toolbar may disable if you are not in an emails tab. If this occurs, the script may attempt to run and will result in a Windows ding. Unfortunately, there is no current solution to this bug that I have implimented.

;Thank you for using this Auto Hot Key Script!
; Version: V1.6
; Website: https://github.com/Gamedirection
; Contact: Join the Discord @ Join.GameDirection.Net
; Copyright 2023 GameDirection LLC
; All rights reserved. This script may be reproduced and distributed in any form with the a credit to the author "GameDirection".

#Persistent
prev_width := 0
prev_height := 0
inbox_names := [""]
SetTimer, DetectWindowChange, 1000
return

DetectWindowChange:
if (WinActive("ahk_exe outlook.exe"))
{
    ; get the window titles from Outlook
    WinGet, WinList, List, Outlook
    WinCount := 0
    WinText := ""
    Loop % WinList
    {
        WinTitle := WinGetTitle(WinList%A_Index%)
        If (WinTitle != "")
        {
            WinText .= WinTitle "`n"
            WinCount += 1
        }
    }
    ; update the inbox_names array
    if (WinCount > 0)
    {
        inbox_names := StrSplit(WinText, "`n")
        Loop % inbox_names.MaxIndex()
        {
            if (inbox_names[A_Index] = "")
                inbox_names.Remove(A_Index)
        }
    }
    ; check for changes in the active window's size
    for each, inbox_name in inbox_names
    {
        if (WinActive(inbox_name))
        {
            WinGetPos, , , width, height, A
            if (width <> prev_width || height <> prev_height)
            {
                prev_width := width
                prev_height := height
                if (width >= 700 && height >= 500)  ; only resize if width is >= 700 pixels
                {
                    if (width > height)
                    {
                        SendInput, !3
                        Sleep, 1
                        SendInput, r
                    }
                    else
                    {
                        SendInput, !3
                        Sleep, 1
                        SendInput, b
                    }
                }
            }
            ;break  exit the loop if a matching window is found
        }
    }
}
return

WinGetTitle(hwnd)
{
    VarSetCapacity(title, 255, 0)
    DllCall("GetWindowText", "Ptr", hwnd, "Str", title, "Int", 255)
    Return title
}
