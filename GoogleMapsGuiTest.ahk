#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

Prev := FixIE()

Gui, Add, Text, x10 y10, Address
Gui, Add, Edit, x100 y10 vstrAddress
Gui, Add, Button, x30 y50 ggetAddress, Get Address

Gui, Show
Return

getAddress:
{
	Gui, Submit, NoHide
	Gui, Destroy
	
	addressArray := StrSplit(strAddress, " ")
	Loop % addressArray.MaxIndex()
	{
		loopVal := addressArray[A_Index]
		addressLink .= loopVal "+"
	}
	;MsgBox, %addressLink%
	Gui, Add, ActiveX, xm w980 h640 vWB, http://www.google.com.au/maps/place/%addressLink%
	Gui, Show
	Return
}

;Gui, Add, ActiveX, xm w980 h640 vWB, http://maps.google.com.au
;Gui, Add, ActiveX, xm w980 h640 vWB, http://www.google.com.au/maps/place/8+Jingana+Rd,+Banksia+Grove+WA+6031
;WB.Navigate("http://www.news.com.au")
;Gui, Show
;Return

FixIE(Version=0, ExeName="")
{
	static Key := "Software\Microsoft\Internet Explorer"
	. "\MAIN\FeatureControl\FEATURE_BROWSER_EMULATION"
	, Versions := {7:7000, 8:8888, 9:9999, 10:10001, 11:11001}
	
	if Versions.HasKey(Version)
		Version := Versions[Version]
	
	if !ExeName
	{
		if A_IsCompiled
			ExeName := A_ScriptName
		else
			SplitPath, A_AhkPath, ExeName
	}
	
	RegRead, PreviousValue, HKCU, %Key%, %ExeName%
	if (Version = "")
		RegDelete, HKCU, %Key%, %ExeName%
	else
		RegWrite, REG_DWORD, HKCU, %Key%, %ExeName%, %Version%
	return PreviousValue
}


GuiClose:
ExitApp

^Esc::ExitApp