#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#SingleInstance Force
#INCLUDE ADOSQL.AHK

Global ADOSQL_LastError, ADOSQL_LastQuery

OasisConnect := "Driver={SQL Server};Server=WSQE004SQL\OASIS9_2;Database=OASIS;Uid=OasisUserWSQE004SQL;Pwd="
TherapyConnect := "Driver={SQL Server};Server=WSQE004SQL;Database=TherapyClinic;Uid=WASDRITC;Pwd=Therapy"

Gui, Add, Text, x20 y20 w77 h21, Patient URN:
Gui, Add, Text, x20 y50 w77 , Name:
Gui, Add, Text, x20 y80 w77 , Date of Birth:
;Gui, Add, Text, x20 y110 w77, Address:
Gui, Font, s08 w600

Gui, Add, Edit, x97 y18 w80 h21 gURNCheckSUM vpatURN, 
Gui, Add, Text, x98 y50 w130 h42 vpatName
Gui, Add, Text, x98 y80 w120 h21 vpatDOB
Gui, Add, Text, x98 yp30 w180 vStrAddress1
Gui, Add, Text, yp20 w180 vStrSuburb
Gui, Font, s08 w400

;Gui, Add, Button, x18 y110 gshowAddress, Address:`nClick to show
Gui, Add, Text, x20 y110, Address:
Gui, Add, Link, x98 y110 w150 h100 vstrAddressLink
;Gui, Add, Checkbox, x200 y20 vpatActive, Active
;GuiControl, Disable, patActive

Gui, Add, GroupBox, x270 y20 w170 h190, Details
Gui, Add, Text, x290 yp20 w77 vstrPatActive
Gui, Add, Text, x290 yp20 w77 h21 vstrFundingSource
;Gui, Add, DDL, yp15 w100 h23 r4 vfundingSource, DSC|MND|Public Fund|Self Funded
Gui, Add, Text, yp20 w100 vstrConsultant
;Gui, Add, DDL, yp15 w100 h23 r12 vpatConsultant, Dr B. Singh|Dr C. Kosky|Dr N. McArdle|Dr A. James|Dr I. Ling|Dr R. Warren|Dr J. Leong|Dr S. Phung|Registrar 1|Registrar 2|Registrar 3

Gui, Add, Text, yp20 w100 vstrDiagnosis
;Gui, Add, Edit, xp77 yp-2 vpatDiagnosis

Gui, Add, Button, x300 y180 w105 gbtnDetails, Edit Details

Gui, Add, Text, x460 y20, Current Appointment Date:
Gui, Add, MonthCal, x460 y40 vMyDateTime

;GuiControl,,MyDateTime, 20040805

Gui, Add, GroupBox, x10 y210 w675 h370, Equipment
Gui, Add, Text, x30 yp30, Machine:
Gui, Add, Edit, xp90 yp0 vpatMachine
Gui, Add, Text, x30 yp30, Mask:
Gui, Add, Edit, xp90 yp0 vpatMask
Gui, Add, Text, x30 yp30, Battery Backup:
Gui, Add, Edit, xp90 yp0 vpatBatteryBackup
Gui, Add, Text, x30 yp30, Backup Machine:
Gui, Add, Edit, xp90 yp0 vpatBackupMachine
;Gui, Add, Checkbox, x30 yp30 vPatWheelchair, Wheelchair mounted
;Gui, Add, Checkbox, x30 yp30 vhoursUsage, 18+ hours use
;Gui, Add, Checkbox, x30 yp30 vlifeSupport, Life Support Register
Gui, Add, Text, x30 yp30, Comments:
Gui, Add, Edit, xp70 yp0 w220 h110 vpatComments

Gui, Add, GroupBox, x340 y230 w325 h330, Settings
Gui, Add, Text, x360 yp30, Mode:
Gui, Add, DDL,  yp15 w50 r3 vtimedMode, ST|S|T
Gui, Add, Text, yp32, IPAP:
Gui, Add, Edit, yp15 w50 vintIPAP
Gui, Add, Text, xp60 yp3, cmH2O
Gui, Add, Text, xp-60 yp32, Rate:
Gui, Add, Edit, yp15 w50 vintBPM
Gui, Add, Text, xp60 yp3, breaths/min
Gui, Add, Text, xp-60 yp32, Ti Min:
Gui, Add, Edit, yp15 w50 vintTiMin
Gui, Add, Text, yp32, Trigger:
Gui, Add, Edit, yp15 w50 vtrigger
Gui, Add, Text, yp32, Ramp:
Gui, Add, Edit, yp15 w50 vintRamp
Gui, Add, Text, xp60 yp3, minutes

Gui, Add, Text, x520 y260, Start EPAP:
Gui, Add, Edit, yp15 w50 vstartEPAP
Gui, Add, Text, xp60 yp3, cmH2O
Gui, Add, Text, xp-60 yp32, EPAP:
Gui, Add, Edit, yp15 w50 vintEPAP
Gui, Add, Text, xp60 yp3, cmH2O
Gui, Add, Text, xp-60 yp32, Rise Time:
Gui, Add, Edit, yp15 w50 vriseTime
Gui, Add, Text, yp32, Ti Max:
Gui, Add, Edit, yp15 w50 vintTiMax
Gui, Add, Text, yp32, Cycle:
Gui, Add, Edit, yp15 w50 vcycle

Gui, Font, s08 w600
GuiControl, Font, timedMode
GuiControl, Font, intIPAP
GuiControl, Font, intBPM
GuiControl, Font, intTiMin
GuiControl, Font, trigger
GuiControl, Font, intRamp
GuiControl, Font, startEPAP
GuiControl, Font, intEPAP
GuiControl, Font, riseTime
GuiControl, Font, intTiMax
GuiControl, Font, cycle
Gui, Font, s08 w400

Gui, Add, Button, x550 y590 gbtnSaveAppointment, Save
Gui, Add, Button, x600 y590 gbtnCancel, Cancel

Gui, Show, w700 h630, Home Visits

Return

btnDetails:
{
	Gui, Submit, NoHide
	Gui, 2:Add, CheckBox, x10 y10 vpatActive, Patient Active
	Gui, 2:Add, Text, yp30, Funding
	Gui, 2:Add, DDL, yp15 w100 h23 r5 vfundingSource, DSC|MND|Public Fund|Self Funded|Not yet on NIV
	Gui, 2:Add, Text, yp30, Consultant
	Gui, 2:Add, DDL, yp15 w100 h23 r12 vpatConsultant, Dr B. Singh|Dr C. Kosky|Dr N. McArdle|Dr A. James|Dr I. Ling|Dr R. Warren|Dr J. Leong|Dr S. Phung|Registrar 1|Registrar 2|Registrar 3
	Gui, 2:Add, Text, yp30, Diagnosis
	Gui, 2:Add, DDL, yp15 w100 h23 r10 vpatDiagnosis, 1|2|3
	Gui, 2:Add, Checkbox, yp40 vLifeSupport, Life Support Register
	Gui, 2:Add, Checkbox, yp30 vWheelchairMount, Wheelchair Mounted
	Gui, 2:Add, Checkbox, yp30 vHoursUse, 18+ hours use
	Gui, 2:Add, Checkbox, yp30 vCountryPatient, 16+ hours use AND country patient
	Gui, 2:Add, Checkbox, yp30 vTrach, 8+ hours via tracheostomy
	Gui, 2:Add, Button, x180 y360 w50 gbtnDetailsSave, Save
	Gui, 2:Add, Button, x240 yp0 w50 gbtnDetailsCancel, Cancel
	GuiControl,2:,patActive,%patDBActive%
	
	Gui, 2:Show, w300 h400
	Return
}

btnDetailsSave:
{
	Gui, 2:Submit, NoHide
	If (patActive = 0)
		strDBActive := "Not active"
	Else
		strDBActive := "Patient active"

	If (newPatient = True)
	{
		insertDetails := % "INSERT INTO HomeVisits VALUES ('" . patURN . "', '" . patDiagnosis . "', '" . fundingSource . "', '" . patConsultant . "', '" . patActive . "')"
		objReturn := ADOSQL(TherapyConnect, insertDetails)
		newPatient := False
	}
	Else
	{
		updateDetails := % "UPDATE HomeVisits SET Diagnosis = '" . patDiagnosis . "', Funding = '" . fundingSource . "', Consultant = '" . patConsultant . "', Active = '" . patActive . "' WHERE patURN = '" . patURN . "'"
		objReturn := ADOSQL(TherapyConnect, updateDetails)
	}
	MsgBox, Details saved
	Gui, 2: Destroy
	GuiControl,1:, strPatActive, %strDBActive%
	GuiControl,1:, strFundingSource, %fundingSource%
	GuiControl,1:, strConsultant, %patConsultant%
	GuiControl,1:, strDiagnosis, %patDiagnosis%
	Return
}

btnDetailsCancel:
{
	Gui, 2:Destroy
	Return
}

btnSaveAppointment:
{
	Return
}

btnCancel:
{
	Return
}

/*
The setting that I’m thinking are 
ST/S/T 
Ipap
Epap
BPM
Ti min
Ti Max
Rise time
Trigger
Cycle
Start Epap
Ramp

Also maybe a tick box to say yes they have registered on the life support register.
*/


;Subroutine to check URN is valid, load details from Oasis and and therapy clinic database
URNCheckSUM:
{
	GuiControlGet, patURN
	lengthURN := StrLen(patURN)
	If (lengthURN = 8)	;Check for valid URN length of 8 characters
	{
		FoundPos := RegExMatch(patURN, "\d+(.*)", numURN)	;Remove alpha character(s) from start of URN, save in numURN variable
		URNLength := StrLen(patURN)	;Get length of URN
		numURNLength := StrLen(numURN)	;Get length of numURN variable
		StringLeft, letterURN, patURN, 1	;Get the alpha character from start of URN, save in variable letterURN
		;URN has checksum, first letter determined by dividing numeric characters by 11
		remURN := Mod(numURN, 11)
		If (remURN = 0)
			alpha := "A"
		If (remURN = 1)
			alpha := "B"
		If (remURN = 2)
			alpha := "C"
		If (remURN = 3)
			alpha := "D"
		If (remURN = 4)	
			alpha := "E"
		If (remURN = 5)
			alpha := "F"
		If (remURN = 6)
			alpha := "G"
		If (remURN = 7)
			alpha := "H"
		If (remURN = 8)
			alpha := "J"
		If (remURN = 9)
			alpha := "K"
		If (remURN = 10)
			alpha := "L"
		
		If (letterURN <> alpha) || (URNLength <> 8) || (numURNLength <>7)	;Check that URN has valid format
		{
			MsgBox, URN error, please check details
			ierr := 1
			return
		}
		Else
		{
			;retrieve details from Oasis
			sqlURN = '%patURN%'	;remove any whitespace characters and add quotes around variable for use in SQL query
			query_Statement := % "SELECT * FROM PBPATMAS WHERE FileNo = "sqlURN
			objReturn := ADOSQL(OasisConnect, query_Statement)
			dateOfBirth := % objReturn[2,11]
			patFirstName := % objReturn[2,6]
			patLastName := % objReturn[2,4]
			patTitle := % objReturn[2,5]
			patSex := % objReturn[2,12]
			If (StrLen(dateOfBirth)<10)	;correct date if single digit at DD location
				dateOfBirth = 0%dateOfBirth%
			patDOB_YYYY = % SubStr(dateOfBirth,7,4)
			patDOB_MM = % SubStr(dateOfBirth,4,2)
			patDOB_DD = % SubStr(dateOfBirth,1,2)
			patAge := % A_MM-patDOB_MM<0 ? A_YYYY-patDOB_YYYY-1 : A_YYYY-patDOB_YYYY	;calculate patient age today
			patAge = %patAge%0	;required in case age is 0
			StringTrimRight, patientAge, patAge, 1	;remove '0' added in previous line
			patDVANumber := % objReturn[2,153]
			patPensionNumber := % objReturn[2,26]
			
			patFirstName = %patFirstName% ;trim whitespaces
			patLastName = %patLastName%
			StringUpper, patLastName, patLastName
			
			patAddress1 := % objReturn[2,13]
			patAddress2 := % objReturn[2,14]
			patSuburb := % objReturn[2,15]
			patPostcode := % objReturn[2,16]
			
			patAddress1 = %patAddress1%
			patAddress2 = %patAddress2%
			patSuburb = %patSuburb%
			
			arrAddress1 := StrSplit(patAddress1, " ")
			arrAddress2 := StrSplit(patAddress2, " ")
			arrSuburb := StrSplit(patSuburb, " ")
			patPostcode = %patPostcode%
			
			
			Loop % arrAddress1.MaxIndex()
			{
				loopVal := arrAddress1[A_Index]
				If (A_Index < arrAddress1.MaxIndex())
					addressLink .= loopVal "+"
				Else
					addressLink .= loopVal ","
			}
			Loop % arrAddress2.MaxIndex()
			{
				loopVal := arrAddress2[A_Index]
				If (A_Index < arrAddress2.MaxIndex())
					addressLink .= loopVal "+"
				Else
					addressLink .= loopVal 
			}
			Loop % arrSuburb.MaxIndex()
			{
				loopVal := arrSuburb[A_Index]
				addressLink .= loopVal "+"
			}
			addressLink .= patPostcode

;			patAddress := % patAddress1 " " patAddress2 "," patSuburb " " patPostcode
			
			;GuiControl,, patName, % lastName . ", " . firstName
			;GuiControl,, patDOB, %dateOfBirth%
			;GoSub, PatientInformationEnable
			GoSub, LoadDetails
			;GoSub, LoadDetails
		}
	}
	Else If (lengthURN <> 8)	;disables GUI components if URN is not 8 characters
	{
;		GoSub, PatientDetailsDisable
	}
	Return
}

/*
showAddress:
{
	Prev := FixIE()
	Gui, Submit, NoHide
	
;	MsgBox, %addressLink%
;	addressArray := StrSplit(patAddress, " ")
;	Loop % addressArray.MaxIndex()
;	{
;		loopVal := addressArray[A_Index]
;		MsgBox, %A_Index% = /%loopVal%/
;		addressLink .= loopVal "+"
;	}
;	MsgBox, %addressLink%
	Gui, 2:Add, ActiveX, xm w980 h640 vWB, http://www.google.com.au/maps/place/%addressLink%
	Gui, 2:Show
	
	Return
}
*/

LoadDetails:
{
	GuiControl, Disable, patURN	;disable patURN field to avoid conflicts once a patient has been loaded
	
	GuiControl,,patName, % patLastName . ", " . patFirstName
	GuiControl,,patDOB, %dateOfBirth%
	;GuiControl,,StrAddress1, % patAddress1 . ", " . patAddress2
	;GuiControl,,StrSuburb, % patSuburb . ", " . patPostcode
	GuiControl,,strAddressLink, % "<a href=""http://www.google.com.au/maps/place/" addressLink """>" patAddress1 . ", " . patAddress2 . "`n" . patSuburb . ", " . patPostcode "</a>"
	
	tcURN = '%patURN%'
	patientQuery := "SELECT * FROM HomeVisits WHERE patURN = "tcURN
	patientObjReturn := ADOSQL(TherapyConnect, patientQuery)
	
	If (patientObjReturn[2,1] <> "")	;Check if patient already exists in the database
	{
		newPatient := False
		patDBDiagnosis := % patientObjReturn[2,2]
		patDBFunding := % patientObjReturn[2,3]
		patDBConsultant := % patientObjReturn[2,4]
		patDBActive := % patientObjReturn[2,5]
		
		If (patDBActive = 1)
			strDBActive := "Patient Active"
		Else
			strDBActive := "Not active"
		
		GuiControl,, strFundingSource, %patDBFunding%
		GuiControl,, strConsultant, %patDBConsultant%
		GuiControl,, strDiagnosis, %patDBDiagnosis%
		GuiControl,, strPatActive, %strDBActive%
	}
	Else
		newPatient := True
	
	Return
}

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