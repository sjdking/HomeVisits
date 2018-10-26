#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

Gui, Add, ListView, r20 w500 gMyListView, Name|Size (KB)

LV_Add("",1,2)
LV_Add("",3,4)

Gui, Show
Return

MyListView:
{
	If (A_GuiEvent = "DoubleClick")
	{
		LV_GetText(RowText, A_EventInfo, 1)
		LV_GetText(Row2Text, A_EventInfo, 2)
		ToolTip You double clicked row number %A_EventInfo%. Text: %RowText%`, %Row2Text%
	}
	Return
}

GuiClose:
ExitApp