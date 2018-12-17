#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

#SingleInstance Force
#INCLUDE ADOSQL.AHK
#Include Gdip.ahk

Global ADOSQL_LastError, ADOSQL_LastQuery

OasisConnect := "Driver={SQL Server};Server=WSQE004SQL\OASIS9_2;Database=OASIS;Uid=OasisUserWSQE004SQL;Pwd="
TherapyConnect := "Driver={SQL Server};Server=WSQE004SQL;Database=TherapyClinic;Uid=WASDRITC;Pwd=Therapy"

GoSub, CreateGrid

Gui, Add, Text, x20 y20 w77 h21, Patient URN:
Gui, Add, Text, x20 y50 w77 , Name:
Gui, Add, Text, x20 y80 w77 , Date of Birth:
Gui, Font, s08 w600

Gui, Add, Edit, x97 y18 w80 h21 gURNCheckSUM vpatURN, 
Gui, Add, Text, x98 y50 w130 h42 vpatName
Gui, Add, Text, x98 y80 w120 h21 vpatDOB
Gui, Add, Text, x98 yp30 w180 vStrAddress1
Gui, Add, Text, yp20 w180 vStrSuburb
Gui, Font, s08 w400

Gui, Add, Text, x20 y110, Address:
Gui, Add, Link, x98 y110 w150 h60 vstrAddressLink gaddressClick

Gui, Font, s08 w600
Gui, Add, GroupBox, x20 y170 w280 h280, Details
Gui, Font, s08 w400
Gui, Add, Picture, xp20 yp20 w250 h222, gridDetails.png
Gui, Add, Text, xp10 yp5 +BackgroundTrans, Active
Gui, Add, Text, yp22 +BackgroundTrans, Funding
Gui, Add, Text, yp22 +BackgroundTrans, Consultant
Gui, Add, Text, yp22 +BackgroundTrans, Diagnosis
Gui, Add, Text, xp10 yp22 +BackgroundTrans, - Sub Diagnosis
Gui, Add, Text, xp-10 yp22 +BackgroundTrans, Life Support Register
Gui, Add, Text, yp22 +BackgroundTrans, Wheelchair Mounted
Gui, Add, Text, yp22 +BackgroundTrans, 18+ Hours Use
Gui, Add, Text, yp22 +BackgroundTrans, 16+ Hours Country
Gui, Add, Text, yp22 +BackgroundTrans, Tracheostomy
Gui, Add, Checkbox, xp120 y195 vCBPatActive +BackgroundTrans,
Gui, Add, ComboBox, xp-10 yp17 w130 vCBOFundingSource +BackgroundTrans, DSC|MND|Public Fund|Self Funded|Not yet on NIV
Gui, Add, ComboBox, yp22 w130 vCBOConsultant +BackgroundTrans, Dr B. Singh|Dr C. Kosky|Dr N. McArdle|Dr A. James|Dr I. Ling|Dr R. Warren|Dr J. Leong|Dr S. Phung|Registrar 1|Registrar 2|Registrar 3
Gui, Add, Combobox, yp22 w130 vCBODiagnosis gcboDiagnosis +AltSubmit +BackgroundTrans, Neuromuscular Disease|Obesity Hypoventilation|Pulmonary Disease|Chest Wall Disorder
Gui, Add, Combobox, yp22 w130 vCBOSubDiagnosis +BackgroundTrans,
Gui, Add, Checkbox, xp10 yp26 vCBLifeSupport +BackgroundTrans
Gui, Add, DateTime, xp30 yp-3 w90 vdateLifeSupport +BackgroundTrans, dd/MM/yyyy
Gui, Add, Checkbox, xp-30 yp25 vCBWheelchair +BackgroundTrans
Gui, Add, Checkbox, yp22 vCBHoursUse +BackgroundTrans
Gui, Add, Checkbox, yp22 vCBCountry +BackgroundTrans
Gui, Add, Checkbox, yp22 vCBTrach +BackgroundTrans

;MyDateTime

Gui, Add, Button, x50 yp30 w105 vbtnEditDetails gbtnEditDetails, Edit Details
Gui, Add, Button, xp120 w105 vbtnSaveDetails gbtnSaveDetails, Save Details

Gui, Add, Text, x330 y30, Appointment Date:
Gui, Add, DDL, xp100 yp-2 w100 vApptList gApptList, %arrApptList%
Gui, Add, Button, xp120 yp-2 vbtnNewAppt gbtnNewAppt, New Appointment

;GuiControl,,MyDateTime, 20040805

Gui, Font, s08 w600
Gui, Add, GroupBox, xp-240 yp47 w435 h420, Equipment
Gui, Font, s08 w400
Gui, Add, Text, xp10 yp30, Machine:
Gui, Add, Edit, xp90 yp0 w100 vpatMachine
Gui, Add, Text, xp-90 yp30, Mask:
Gui, Add, Edit, xp90 yp0 w100 vpatMask
Gui, Add, Text, xp-90 yp30, Battery Backup:
Gui, Add, Edit, xp90 yp0 w100 vpatBatteryBackup
Gui, Add, Text, xp-90 yp30, Backup Machine:
Gui, Add, Edit, xp90 yp0 w100 vpatBackupMachine

;Gui, Font, s08 w600
;Gui, Add, GroupBox, x330 y260 w325 h315, Settings
;Gui, Font, s08 w400
Gui, Add, Picture, xp120 yp-90 w200 h244, grid.png
Gui, Add, Text, xp5 yp5 vtextMode +BackgroundTrans, Mode
Gui, Add, Text, yp22 vtextIPAP +BackgroundTrans, IPAP
Gui, Add, Text, yp22 vtextEPAP +BackgroundTrans, EPAP
Gui, Add, Text, yp22 vtextStartEPAP +BackgroundTrans, Start EPAP
Gui, Add, Text, yp22 vtextTiMin +BackgroundTrans, Ti Min
Gui, Add, Text, yp22 vtextTiMax +BackgroundTrans, Ti Max
Gui, Add, Text, yp22 vtextRate +BackgroundTrans, Rate
Gui, Add, Text, yp22 vtextTrigger +BackgroundTrans, Trigger
Gui, Add, Text, yp22 vtextRamp +BackgroundTrans, Ramp
Gui, Add, Text, yp22 vtextRiseTime +BackgroundTrans, Rise Time
Gui, Add, Text, yp22 vtextCycle +BackgroundTrans, Cycle

Gui, Add, ComboBox, xp75 yp-225 w120 h19 r3 -E0x200 veditMode gDDL1, ST|S|T

Gui, Add, Edit, xp4 yp26 w116 h18 -E0x200 veditIPAP gudIPAP Number
;Gui, Add, UpDown, vIPAP gudIPAP Range0-20
Gui, Add, Text, xp50 +BackgroundTrans, cmH2O

Gui, Add, Edit, xp-50 yp22 w45 h18 -E0x200 veditEPAP gudEPAP Number
;Gui, Add, UpDown, vEPAP gudEPAP Range0-20
Gui, Add, Text, xp50 +BackgroundTrans, cmH2O

Gui, Add, Edit, xp-50 yp22 w45 h18 -E0x200 veditStartEPAP Number
Gui, Add, Text, xp50 +BackgroundTrans, cmH2O

Gui, Add, Edit, xp-50 yp22 w115 h18 -E0x200 veditTiMin
Gui, Add, Edit, yp22 w115 h18 -E0x200 veditTiMax
Gui, Add, Edit, yp22 w115 h18 -E0x200 veditBreathRate
Gui, Add, Edit, yp22 w115 h18 -E0x200 veditTrigger
Gui, Add, Edit, yp22 w115 h18 -E0x200 veditRampTime
Gui, Add, Edit, yp22 w115 h18 -E0x200 veditRiseTime
Gui, Add, Edit, yp22 w115 h18 -E0x200 veditCycle

Gui, Add, Text, xp-285 yp30, Comments:
Gui, Add, Edit, yp20 w400 h100 vapptComments

Gui, Add, Button, xp100 yp125 gbtnSaveAppointment, Save Appointment
Gui, Add, Button, xp100 yp0 gbtnClose, Close

Gui, Show, w760 h630, Home Visits

Return

btnNewAppt:
{
	Gui,2:-SysMenu +AlwaysOnTop
	Gui,2:Add, MonthCal, h160 vnewAppt
	Gui,2:Add, Button, yp200 gbtnCreateAppt, Create Appointment
	Gui,2:Add, Button, xp100 gbtnCancelAppt, Cancel
	Gui,2:Show, h300, Select new appointment date
	GuiControl, Disable, btnNewAppt
	Return
}

btnCreateAppt:
{
	Gui, Submit, NoHide
	Gui,2:Destroy
	FormatTime, newAppt, %newAppt%, ddMMyyyy
	GuiControl,1:Enable, btnNewAppt
	arrApptList = %arrApptList%
	arrApptList := SubStr(arrApptList,2,StrLen(arrApptList)-1)
	arrApptList .= newAppt . "||"
	GuiControl,1:,ApptList, %arrApptList%
	Return
}

btnCancelAppt:
{
	Gui,2:Destroy
	GuiControl, Enable, btnNewAppt
	Return
}

apptList:
{
	GuiControlGet, ApptList
	apptKey := % "'" . patURN . "_" . ApptList . "'"
	selApptQuery := "SELECT * FROM HomeVisitAppointments WHERE apptKey = "apptKey
	objSelAppt := % ADOSQL(TherapyConnect, selApptQuery)
	
	strMode := % objSelAppt[2,4]
	strIPAP := % objSelAppt[2,5]
	strEPAP := % objSelAppt[2,6]
	strStartEPAP := % objSelAppt[2,7]
	strTiMin := % objSelAppt[2,8]
	strTiMax := % objSelAppt[2,9]
	strRate := % objSelAppt[2,10]
	strTrigger := % objSelAppt[2,11]
	strRamp := % objSelAppt[2,12]
	strRiseTime := % objSelAppt[2,13]
	strCycle := % objSelAppt[2,14]
	
	strMode = %strMode%
	
	GuiControl, Choose, editMode, %strMode%
	GuiControl,, editIPAP, %strIPAP%
	GuiControl,, editEPAP, %strEPAP%
	GuiControl,, editStartEPAP, %strStartEPAP%
	GuiControl,, editTiMin, %strTiMin%
	GuiControl,, editTiMax, %strTiMax%
	GuiControl,, editBreathRate, %strRate%
	GuiControl,, editTrigger, %strTrigger%
	GuiControl,, editRampTime, %strRamp%
	GuiControl,, editRiseTime, %strRiseTime%
	GuiControl,, editCycle, %strCycle%
	
	;MsgBox, %strMode%`n%strIPAP%`n%strEPAP%`n%strStartEPAP%`n%strTiMin%`n%strTiMax%`n%strRate%`n%strTrigger%`n%strRamp%`n%strRiseTime%`n%strCycle%
	Return
}

cboDiagnosis:
{
	GuiControlGet, CBODiagnosis
	GuiControl,Enable, CBOSubDiagnosis
	If CBODiagnosis = 1
		GuiControl,,CBOSubDiagnosis, |Motor Neurone Disease|Duchenne Muscular Dystrophy|Other Muscular Dystrophy|Diaphragm Weakness|Central Hypoventilation|Spinal Cord Damage|Post-Polio|Other
	Else If CBODiagnosis = 2
	{
		GuiControl,, CBOSubDiagnosis, |
		GuiControl,Disable,CBOSubDiagnosis
	}
	Else If CBODiagnosis = 3
		GuiControl,, CBOSubDiagnosis, |COPD|CF|Non-CF Bronchiectasis|Other
	Else If CBODiagnosis = 4
		GuiControl,, CBOSubDiagnosis, |Kyphoscoliosis/Scoliosis|Other
	Return
}

addressClick:
{
	WinActivate, ahk_exe chrome.exe
	IfWinNotExist, ahk_exe chrome.exe
	{
		Run, chrome.exe
		Sleep, 100
		while (A_Cursor = "AppStarting" or A_Cursor = "Wait")
			Continue
		Sleep, 100
		while (A_Cursor = "AppStarting" or A_Cursor = "Wait")
			Continue
		Sleep, 3000
	}

	WinActivate, ahk_exe chrome.exe
	Sleep, 100
	ControlSend, ,{CTRL down}t{CTRL up}, ahk_exe chrome.exe
	Sleep, 100
	SendRaw, % "http://www.google.com.au/maps/place/" addressLink
	Send, {ENTER}
	Return
}

DDL1:
{
	GuiControlGet, editMode
	If (editMode = "ST") OR (editMode = "S") OR (editMode = "T")
		GuiControl,Choose, editMode, editMode
	Else
	{
		MsgBox, Select a valid mode: ST, S or T
		GuiControl,Choose, editMode, ST
	}
	Return
	
}

udIPAP:
{
	Gui, Submit, NoHide
	If (editIPAP < editEPAP)
	{
		MsgBox, IPAP should not be lower than EPAP
		editIPAP := % editIPAP + 1
	}
	GuiControl,, editIPAP, % editIPAP 
	Return
}

udEPAP:
{
	Gui, Submit, NoHide

	If (editIPAP < editEPAP)
	{
		MsgBox, IPAP should not be lower than EPAP
		editEPAP := % editEPAP - 1
	}
	GuiControl,, editEPAP, %editEPAP%
	Return
}

udStartEPAP:
{
	Gui, Submit, NoHide
	GuiControl,, editStartEPAP, % StartEPAP
	If (EPAP < StartEPAP)
		MsgBox, Start EPAP should not be higher than EPAP
	Return
}

btnSaveAppointment:
{
	Gui, Submit, NoHide
	
	;FormatTime, MyDateTime, %MyDateTime%, ddMMyyyy
	apptKey := % patURN "_" ApptList
	checkAppt := % "SELECT * FROM HomeVisitAppointments WHERE apptKey = '" . apptKey . "'"
	objCheckAppt := ADOSQL(TherapyConnect, checkAppt)
	apptExists := % objCheckAppt[2,1]
	If (apptExists = "")
	{
		insertAppointment := % "INSERT INTO HomeVisitAppointments (apptKey, patURN, Date, strMode, intIPAP, intEPAP, intStartEPAP, TiMin, TiMax, intRate, strTrigger, strRamp, intRiseTime, strCycle) VALUES ('" . apptKey . "', '" . patURN . "', '" . newAppt . "', '" . editMode . "', '" . editIPAP . "', '" . editEPAP . "', '" . editStartEPAP . "', '" . editTiMin . "', '" . editTiMax . "', '" . editBreathRate . "', '" . editTrigger . "', '" . editRampTime . "', '" . editRiseTime . "', '" . editCycle . "')"
		objApptReturn := ADOSQL(TherapyConnect, insertAppointment)
;		MsgBox, %ADOSQL_LastQuery%`n`n%ADOSQL_LastError%
	}
	Else
	{
		MsgBox, 4,, Do you want to update the details for the appointment on %ApptList%?
		IfMsgBox Yes
		{
			updateAppointment := % "UPDATE HomeVisitAppointments SET Date = '" . ApptList . "', strMode = '" . editMode . "', intIPAP = '" . editIPAP . "', intEPAP = '" . editEPAP . "', intStartEPAP = '" . editStartEPAP . "', TiMin = '" . editTiMin . "', TiMax = '" . editTiMax . "', intRate = '" . editBreathRate . "', strTrigger = '" . editTrigger . "', strRamp = '" . editRampTime . "', intRiseTime = '" . editRiseTime . "', strCycle = '" . editCycle . "' WHERE apptKey = '" . apptKey . "'"
			objApptReturn := ADOSQL(TherapyConnect, updateAppointment)
		}
		Else
			Return
	}
	MsgBox, Appointment Saved
	Return
}

btnEditDetails:
{
/*	Gui, Submit, NoHide
	Gui, 2:Add, CheckBox, x10 y10 vpatActive, Patient active
	Gui, 2:Add, Text, yp30, Funding
	Gui, 2:Add, DDL, yp15 w100 h23 r5 vfundingSource, DSC|MND|Public Fund|Self Funded|Not yet on NIV
	Gui, 2:Add, Text, yp30, Consultant
	Gui, 2:Add, DDL, yp15 w100 h23 r12 vpatConsultant, Dr B. Singh|Dr C. Kosky|Dr N. McArdle|Dr A. James|Dr I. Ling|Dr R. Warren|Dr J. Leong|Dr S. Phung|Registrar 1|Registrar 2|Registrar 3
	Gui, 2:Add, Text, yp30, Diagnosis
	Gui, 2:Add, DDL, yp15 w100 h23 r10 vpatDiagnosis, MND|Duchennes
	Gui, 2:Add, Checkbox, yp40 vLifeSupport, Life Support Register
	Gui, 2:Add, Checkbox, yp30 vWheelchairMount, Wheelchair Mounted
	Gui, 2:Add, Checkbox, yp30 vHoursUse, 18+ hours use
	Gui, 2:Add, Checkbox, yp30 vCountryPatient, 16+ hours use AND country patient
	Gui, 2:Add, Checkbox, yp30 vTrach, 8+ hours via tracheostomy
	Gui, 2:Add, Button, x180 y360 w50 gbtnDetailsSave, Save
	Gui, 2:Add, Button, x240 yp0 w50 gbtnDetailsCancel, Cancel
	
	If (newPatient = False)
	{
		GuiControl,2:,patActive,%patDBActive%
		GuiControl,2:Choose,fundingSource,%patDBFunding%
		GuiControl,2:Choose,patDiagnosis,%patDBDiagnosis%
		GuiControl,2:Choose,patConsultant,%patDBConsultant%
		GuiControl,2:,LifeSupport,%patDBLife%
		GuiControl,2:,WheelchairMount,%patDBWheelchair%
		GuiControl,2:,HoursUse,%patDBHoursUse%
		GuiControl,2:,CountryPatient,%patDBCountry%
		GuiControl,2:,Trach,%patDBTrach%
	}
	
	Gui, 2:Show, w300 h400
*/
	GoSub, EnableDetails

	
	Return
}

btnSaveDetails:
{
	Gui, Submit, NoHide
	If (newPatient = True)
	{
		insertDetails := % "INSERT INTO HomeVisits VALUES ('" . patURN . "', '" . CBODiagnosis . "', '" . CBOFundingSource . "', '" . CBOConsultant . "', '" . CBPatActive . "', '" . CBLifeSupport . "', '" . CBWheelchair . "', '" . CBHoursUse . "', '" . CBCountry . "', '" . CBTrach . "')"
		objReturn := ADOSQL(TherapyConnect, insertDetails)
		newPatient := False
	}
	Else
	{
		updateDetails := % "UPDATE HomeVisits SET Diagnosis = '" . CBODiagnosis . "', Funding = '" . CBOFundingSource . "', Consultant = '" . CBOConsultant . "', Active = '" . CBPatActive . "', LifeSupport = '" . CBLifeSupport . "', Wheelchair = '" . CBWheelchair . "', HoursUse = '" . CBHoursUse . "', CountryPatient = '" . CBCountry . "', Tracheostomy = '" . CBTrach . "' WHERE patURN = '" . patURN . "'"
		objReturn := ADOSQL(TherapyConnect, updateDetails)
	}
	MsgBox, Details saved
	
;	MsgBox, %ADOSQL_LastQuery%`n`n%ADOSQL_LastError%
	GuiControl,, CBPatActive, %strDBActive%
	GuiControl, Choose, CBOFundingSource, %fundingSource%
	GuiControl, Choose, CBOConsultant, %patConsultant%
	GuiControl, Choose, CBODiagnosis, %patDiagnosis%
	GuiControl, Choose, CBOSubDiagnosis, %patSubDiagnosis%
	GuiControl,, CBLifeSupport, %CBLifeSupport%
	GuiControl,, CBHoursUse, %HoursUse%
	GuiControl,, CBCountry, %CountryPatient%
	GuiControl, ,CBTrach, %Trach%
	
	Return
}

btnDetailsCancel:
{
	Gui, 2:Destroy
	Return
}

btnClose:
{
	ExitApp
}

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
			StringUpper, patURN, patURN
			GuiControl,, patURN, %patURN%
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

LoadDetails:
{
	GuiControl, Disable, patURN	;disable patURN field to avoid conflicts once a patient has been loaded
	
	GuiControl,,patName, % patLastName . ", " . patFirstName
	GuiControl,,patDOB, %dateOfBirth%
	;GuiControl,,strAddressLink, % "<a href=""http://www.google.com.au/maps/place/" addressLink """>" patAddress1 . ", " . patAddress2 . "`n" . patSuburb . ", " . patPostcode "</a>"
	GuiControl,,strAddressLink, % "<a id=""http://www.google.com.au/maps/place/" addressLink """>" patAddress1 . ", " . patAddress2 . "`n" . patSuburb . ", " . patPostcode "</a>"
	tcURN = '%patURN%'
	patientQuery := "SELECT * FROM HomeVisits WHERE patURN = "tcURN
	patientObjReturn := ADOSQL(TherapyConnect, patientQuery)
	
	If (patientObjReturn[2,1] <> "")	;Check if patient already exists in the database
	{
		newPatient := False
		patDBDiagnosis := % patientObjReturn[2,2]
		patDBFunding := % patientObjReturn[2,3]
		patDBFunding = %patDBFunding%
		patDBConsultant := % patientObjReturn[2,4]
		patDBActive := % patientObjReturn[2,5]
		patDBLife := % patientObjReturn[2,6]
		patDBWheelchair := % patientObjReturn[2,7]
		patDBHoursUse := % patientObjReturn[2,8]
		patDBCountry := % patientObjReturn[2,9]
		patDBTrach := % patientObjReturn[2,10]
		
/*		If (patDBActive = 1)
			strDBActive := "Patient Active"
		Else
			strDBActive := "Not active"
		If patDBLife = 1
			strDBLife := "Y"
		Else
			strDBLife := "N"
		If patDBWheelchair = 1
			strDBWheelchair := "Y"
		Else
			strDBWheelchair := "N"
		If patDBHoursUse = 1
			strDBHoursUse := "Y"
		Else
			strDBHoursUse := "N"
		If patDBCountry = 1
			strDBCountry := "Y"
		Else
			strDBCountry := "N"
		If patDBTrach = 1
			strDBTrach := "Y"
		Else
			strDBTrach := "N"
*/
		GuiControl, Choose, CBOFundingSource, %patDBFunding%
		GuiControl, Choose, CBOConsultant, %patDBConsultant%
		GuiControl, Choose, CBODiagnosis, %patDBDiagnosis%
		GuiControl, Choose, CBOSubDiagnosis, %patDBSubDiagnosis%
		GuiControl,, CBPatActive, %patDBActive%
		GuiControl,, CBLifeSupport, %patDBLife%
		GuiControl,, CBWheelchair, %patDBWheelchair%
		GuiControl,, CBHoursUse, %patDBHoursUse%
		GuiControl,, CBCountry, %patDBCountry%
		GuiControl,, CBTrach, %patDBTrach%
		GoSub, DisableDetails
		
		apptQuery := "SELECT * FROM HomeVisitAppointments WHERE patURN = "tcURN
		apptObjReturn := ADOSQL(TherapyConnect, apptQuery)
		
		arrApptList := " ||"
		
		Loop
		{
			i := % A_Index + 1
			dateVal := % apptObjReturn[i,3]
			If (dateVal = "")
				Break
			arrApptList .= dateVal . "|"
		}
		GuiControl,, ApptList, %arrApptList%
	}
	Else
		newPatient := True
	
	Return
}

EnableDetails:
{
	GuiControl, Enable, CBOFundingSource
	GuiControl, Enable, CBOConsultant
	GuiControl, Enable, CBODiagnosis
	GuiControl, Enable, CBOSubDiagnosis
	GuiControl, Enable, CBPatActive
	GuiControl, Enable, CBLifeSupport
	GuiControl, Enable, CBWheelchair
	GuiControl, Enable, CBHoursUse
	GuiControl, Enable, CBCountry
	GuiControl, Enable, CBTrach
	GuiControl, Enable, btnSaveDetails
	
	GuiControl, Disable, btnEditDetails
	Return
}

DisableDetails:
{
	GuiControl, Disable, CBOFundingSource
	GuiControl, Disable, CBOConsultant
	GuiControl, Disable, CBODiagnosis
	GuiControl, Disable, CBOSubDiagnosis
	GuiControl, Disable, CBPatActive
	GuiControl, Disable, CBLifeSupport
	GuiControl, Disable, CBWheelchair
	GuiControl, Disable, CBHoursUse
	GuiControl, Disable, CBCountry
	GuiControl, Disable, CBTrach
	GuiControl, Disable, btnSaveDetails
	
	GuiControl, Enable, btnEditDetails
	Return
}

CreateGrid:
{
	; Start gdi+
	If !pToken := Gdip_Startup()
	{
		MsgBox, 48, gdiplus error!, Gdiplus failed to start. Please ensure you have gdiplus on your system
		ExitApp
	}

	pBitmap := Gdip_CreateBitmap(200, 244)
	G := Gdip_GraphicsFromImage(pBitmap)
	dBitmap := Gdip_CreateBitmap(250, 222)
	D := Gdip_GraphicsFromImage(dBitmap)

	pBrush := Gdip_BrushCreateSolid(0xFFFFFFFF)
	Gdip_FillRectangle(G, pBrush, 0, 0, 200, 244)
	Gdip_FillRectangle(D, pBrush, 0, 0, 250, 222)
	Gdip_DeleteBrush(pBrush)

	pPen := Gdip_CreatePen("0xFFD3D3D3", 1)
	Gdip_DrawLine(G, pPen,80,0,80,244)
	Gdip_DrawLine(D, pPen,120,0,120,222)
	Loop, 10
	{
		i := % A_Index * 22
		j := % A_Index
		Gdip_DrawLine(G, pPen,0,i,200,i)
		Gdip_DrawLine(D, pPen,0,i,250,i)
		If (j = 9)
		{
			Gdip_SaveBitmapToFile(dBitmap, "gridDetails.png")
		}
	}
	Gdip_DeletePen(pPen)
	Gdip_SaveBitmapToFile(pBitmap, "grid.png")
	Return
}


GuiClose:
ExitApp

^Esc::ExitApp