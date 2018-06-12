#NoEnv  ; Recommended for performance and compatibility with future AutoHotkey releases.
; #Warn  ; Enable warnings to assist with detecting common errors.
SendMode Input  ; Recommended for new scripts due to its superior speed and reliability.
SetWorkingDir %A_ScriptDir%  ; Ensures a consistent starting directory.

;Version 0.1

;Application for managing WASDRI Therapy Clinic

#SingleInstance Force
#INCLUDE ADOSQL.AHK

;Show loading popup until main GUI is ready
Gui, -SysMenu
Gui, Add, Progress, w200
Gui, Add, Text, wp center, Loading...
Gui, Show, x100 y100, WASDRI Therapy Clinic... LOADING

SetFormat, float, 0.2	;format to 2 decimal places
Global ADOSQL_LastError, ADOSQL_LastQuery	;for troubleshooting SQL errors

;Setup SQL connection strings for Oasis and TherapyClinic databases
	OasisConnect := "Driver={SQL Server};Server=WSQE004SQL\OASIS9_2;Database=OASIS;Uid=OasisUserWSQE004SQL;Pwd="
	TherapyConnect := "Driver={SQL Server};Server=WSQE004SQL;Database=TherapyClinic;Uid=WASDRITC;Pwd=Therapy"

;Initialise list of Sleep Scientists for use in lists
	numSCNames := 0
	queryStatement := "SELECT * FROM Staff WHERE Active = '1' AND StaffID LIKE 'SC%'"
	objReturn := ADOSQL(TherapyConnect, queryStatement)
	SCNames := ""
	SCfirstName := "NULL"

	Loop
{
		i := A_Index + 1
		numSCNames++
		SCfirstName := % objReturn[i,2]
		SClastName := % objReturn[i,3]
		If (SCfirstName <> "")
			SCNames .= SClastName . ", " . SCfirstName . "|"
}
	Until SCfirstName = ""
	Sort, SCNames, D|

;Initialise list of Doctors for use in lists
	numDRNames := 0
	queryStatement := "SELECT * FROM Staff WHERE Active = '1' AND StaffID LIKE 'DR%'"
	objReturn := ADOSQL(TherapyConnect, queryStatement)
	DRNames := ""
	DRfirstName := "NULL"

	Loop
	{
		i := A_Index + 1
		numDRNames++
		DRfirstName := % objReturn[i,2]
		DRlastName := % objReturn[i,3]
		If (DRfirstName <> "")
			DRNames .= DRlastName . ", " . DRfirstName . "|"
	}
	Until DRfirstName = ""
	Sort, DRNames, D|
	
;Initialise list of CPAP Side Effects for use in lists
	numSENames := 0
	queryStatement := "SELECT * FROM SideEffects WHERE SideEffectID LIKE 'SE%'"
	objReturn := ADOSQL(TherapyConnect, queryStatement)
	SENames := ""

	Loop
	{
		i := A_Index + 1
		numSENames++
		sideEffect := % objReturn[i,2]
		If (sideEffect <> "")
			SENames .= sideEffect . "|"
	}
	Until sideEffect = ""

;Initialise number of tabs and tab names for TC visits
	tabs := 1
	tabList := "Equipment Issue"

Gui, Destroy

;Subroutine for initial view on loading application
InitialView:
{
	Gui, Add, Text, x14 y24 w77 h21, Patient URN
	Gui, Add, Text, x14 y49 w77 h20 +0x200, Name
	Gui, Add, Text, x14 y74 w77 h20 +0x200, Date of Birth
	Gui, Font, s08 w600
	Gui, Add, Edit, x107 y22 w120 h21 gURNCheckSUM vpatURN, 
	Gui, Add, Text, x108 y53 w120 h25 vpatName
	Gui, Add, Text, x108 y77 w120 h21 vpatDOB
	
	Gui, Font, s08 w400

	Gui, Show, w1366 h768, WASDRI Therapy Clinic
	
	;FOR TESTING ONLY, REMOVE THE COMMENT CHARACTER ";" ON THE LINE BELOW
	;GuiControl,,patURN, H8561945
	Return
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
			
			;GuiControl,, patName, % lastName . ", " . firstName
			;GuiControl,, patDOB, %dateOfBirth%
			;GoSub, PatientInformationEnable
			GoSub, StudyDetails
			GoSub, LoadWASDRITC
		}
	}
	Else If (lengthURN <> 8)	;disables GUI components if URN is not 8 characters
	{
		GoSub, PatientDetailsDisable
	}
	Return
}
	
LoadWASDRITC:
{
	;Load visit details from TherapyClinic database
	tcURN = '%patURN%`%'
	
	prescriptionQuery := "SELECT * FROM Prescriptions WHERE PrescriptionID LIKE "tcURN
	prescriptionObjReturn := ADOSQL(TherapyConnect, prescriptionQuery)
	If (prescriptionObjReturn <> "")	;if Prescription exists the fill prescription details from therapy SQL DB
		GoSub, PrescriptionFill
		
	If (prescriptionObjReturn[2,1] <> "")	;if Prescription exists check if Issue exists in therapy SQL DB
	{
		issueQuery := "SELECT * FROM Issues WHERE IssueID LIKE "tcURN
		issueObjReturn := ADOSQL(TherapyConnect, issueQuery)
		GoSub, IssueFill
		
		If (issueObjReturn[2,1] <> "")	;if Issue exists check if TC Visits exist in therapy SQL DB
		{
			visitQuery := "SELECT * FROM Visits WHERE VisitID LIKE "tcURN
			visitObjReturn := ADOSQL(TherapyConnect, visitQuery)
			GoSub, VisitFill
		}
	}
	issueObjReturn := ADOSQL(TherapyConnect, issueQuery)
	VisitDate := % objReturn[3,3]
	Return
}

StudyDetails:
{
	GuiControl, Disable, patURN	;disable patURN field to avoid conflicts once a patient has been loaded
	
	GuiControl,,patName, % patLastName . ", " . patFirstName
	GuiControl,,patDOB, %dateOfBirth%
	
	Gui, Add, Button, x108 y100 w120 h21, Change Patient	;use this button to change the patient
	
	Gui, Add, Checkbox, x270 y24 w70 h23 vVeteran, Veteran	;autoselects if DVA number pulled from Oasis
	Gui, Add, Checkbox, x270 y49 w70 h23 vPensioner, Pensioner
	Gui, Add, Checkbox, x270 y74 w70 h23 vDSC, DSC
	Gui, Add, Checkbox, x360 y24 w70 h23 vPrisoner, Prisoner
	Gui, Add, Checkbox, x360 y49 w200 h23 vPMImpairment, Physical/mental impairment
	Gui, Add, Checkbox, x360 y74 w120 h23 vCountry, Country patient

	patDVANumber = %patDVANumber%
	If (patDVANumber <> "")
		GuiControl,,Veteran,1	;autoselect Veteran checkbox
	
	Gui, Font, s10 w600	;change font size and weight to bold
	Gui, Add, GroupBox, x15 y130 w315 h200, Study Details
	Gui, Font, s08 Normal	;reset to normal size and weight font
	
	Gui, Add, Text, x30 y160 w70 h21, Study Date
	Gui, Add, DateTime, x120 y158 w120 h21 vStudyDate
	Gui, Add, Text, x30 y185 w70 h21, AHI
	Gui, Add, Edit, x120 y183 w40 h21 vAHI
	Gui, Add, Text, x165 y185 w120 h21, events/hour
	Gui, Add, Text, x30 y210 w70 h21, SpO2 Nadir
	Gui, Add, Edit, x120 y208 w40 h21 vSpO2Nadir
	Gui, Add, Text, x165 y210 w20 h21, `%
	Gui, Add, Text, x30 y235 w70 h21, ESS
	Gui, Add, Edit, x120 y233 w40 h21 vESS
	Gui, Add, Text, x30 y260 w90 h21, Main Symptoms:
	Gui, Add, Edit, x120 y260 w195 h55 vMainSymptomsText
	
	Gui, Font, s10 w600
	Gui, Add, GroupBox, x345 y130 w320 h200, Urgency
	Gui, Font, s08 Normal
	
	Gui, Add, Checkbox, x360 y160 w170 h21 vurgInpatient, Inpatient
	Gui, Add, Checkbox, x360 y185 w170 h21 vurgRespFailure, Respiratory Failure
	Gui, Add, Checkbox, x360 y210 w170 h21 vurgOccDriver, Occupational Driver
	Gui, Add, Checkbox, x360 y235 w170 h21 vurgSevereOSA, Severe OSA and/or ESS > 16
	Gui, Add, Checkbox, x360 y260 w50 h21 vurgOther, Other:
	Gui, Add, Edit, x455 y260 w195 h55 vurgOtherText
	
	Gui, Font, s10 w600
	Gui, Add, GroupBox, x15 y340 w650 h410, Prescription Details
	Gui, Font, s08 Normal
	
	Gui, Font, Italic w600
	Gui, Add, GroupBox, x30 y380 w300 h100, Prescription Type
	Gui, Font, Normal
	
	Gui, Add, Radio, x50 y400 w140 h21 vRadioPrescriptionType, 4 Week Physician Trial
	Gui, Add, Radio, x50 y425 w140 h21, Therapy Clinic
	Gui, Add, Radio, x50 y450 w140 h21, Purchase
	
	Gui, Font, Italic w600
	Gui, Add, GroupBox, x30 y502 w300 h210, Device Details
	Gui, Font, Normal
	
	Gui, Add, Radio, x50 y530 w140 h21 vRadioTreatmentTypeFixed gTreatmentCheck, Fixed Pressure CPAP Unit
	Gui, Add, Radio, x50 y605 w130 h21 vRadioTreatmentTypeAuto gTreatmentCheck, Autoset CPAP Unit	
	Gui, Add, Text, x70 y555 w130 h21 vstrCPAPRecommendedPressure, Recommended Pressure
	Gui, Add, Edit, x205 y553 w45 h21 vCPAPRecommendedPressure
	Gui, Add, Text, x260 y555 w50 h21 vstrCPAPRecommendedPressurecmH2O, cmH2O
	Gui, Add, Text, x70 y580 w130 h21 vstrRampTime, Ramp Time
	Gui, Add, Edit, x205 y578 w45 h21 vRampTime
	Gui, Add, Text, x260 y580 w50 h21 vstrRampTimeMinutes, minutes
	Gui, Add, Text, x70 y630 w130 h21 vstrMinCPAP, Minimum CPAP
	Gui, Add, Edit, x205 y628 w45 h21 vMinCPAP
	Gui, Add, Text, x260 y630 w50 h21 vstrMinCPAPcmH2O, cmH2O
	Gui, Add, Text, x70 y655 w130 h21 vstrMaxCPAP, Maximum CPAP
	Gui, Add, Edit, x205 y653 w45 h21 vMaxCPAP
	Gui, Add, Text, x260 y655 w50 h21 vstrMaxCPAPcmH2O, cmH2O
	Gui, Add, Text, x70 y680 w130 h21 vstrSettlingTime, Settling Time
	Gui, Add, Edit, x205 y678 w45 h21 vSettlingTime
	Gui, Add, Text, x260 y680 w50 h21 vstrSettlingTimeMinutes, minutes
	
	GuiControl, Disable, CPAPRecommendedPressure
	GuiControl, Disable, strCPAPRecommendedPressure
	GuiControl, Disable, strCPAPRecommendedPressurecmH2O
	GuiControl, Disable, RampTime
	GuiControl, Disable, strRampTime
	GuiControl, Disable, strRampTimeMinutes
	GuiControl, Disable, MaxCPAP
	GuiControl, Disable, strMaxCPAP
	GuiControl, Disable, strMaxCPAPcmH2O
	GuiControl, Disable, MinCPAP
	GuiControl, Disable, strMinCPAP
	GuiControl, Disable, strMinCPAPcmH2O
	GuiControl, Disable, SettlingTime
	GuiControl, Disable, strSettlingTime
	GuiControl, Disable, strSettlingTimeMinutes
	
	Gui, Font, Italic w600
	Gui, Add, GroupBox, x345 y380 w305 h100, Options
	Gui, Font, Normal
	
	Gui, Add, Checkbox, x360 y400 w120 h21 vOximetry, Oximetry in first week
	Gui, Add, Checkbox, x360 y425 w120 h21 vHeatedHumidifier, Heated Humidifier
	Gui, Add, Checkbox, x360 y450 w120 h21 vChinstrap, Chinstrap
	Gui, Add, Checkbox, x520 y400 w120 h21 vNoseMask, Nose Mask
	Gui, Add, Checkbox, x520 y425 w120 h21 vFullFaceMask, Full Face Mask
	Gui, Add, Checkbox, x520 y450 w120 h21 vPressureRelief, Pressure Relief
	
	Gui, Font, Italic w600
	Gui, Add, Text, x345 y490 w120 h21, Doctor's comments:
	Gui, Font, Normal
	
	Gui, Add, Edit, x345 y510 w305 h120 vDoctorsComments
	
	
	Gui, Add, DateTime, x500 y650 w120 h21 ChooseNone vPrescriptionDate
	Gui, Add, DropDownList, x500 y680 w120 h21 vPrescribingDoctor r%numDRNames%, %DRNames%

	Gui, Font, s10 w600
	Gui, Add, Text, x360 y652 w130 h21, Prescription Date:
	Gui, Add, Text, x360 y682 w100 h21, Doctor:
	Gui, Add, Button, x375 y710 w230 h30 vbtnGenPresc, Generate Prescription
	
	Gui, Font, s10 w600
	Gui, Add, Button, x690 y710 w140 h30 vbtnAddTC, Add TC Visit
	Gui, Font, s08 Normal

	Gui, Add, Tab3, x690 y30 w660 h660 AltSubmit vVisitTab -wrap, %tabList%
	Sleep, 10
	GoSub, EquipmentIssueGUI	;subroutine to load Issue GUI
	
	Return
}

;subroutine when Generate Prescription button is pressed
ButtonGeneratePrescription:
{
	Gui, Submit, NoHide
	If (PrescribingDoctor = "")
	{
		MsgBox, Please select the prescribing doctor
		Return
	}
	If (PrescriptionDate = "")
	{
		MsgBox, Please select the prescription date
		Return
	}
	If (RadioPrescriptionType = "0")
	{
		MsgBox, Please select the prescription Type
		Return
	}
	
	If (RadioTreatmentTypeFixed = "0" AND RadioTreatmentTypeAuto = "0")
	{
		MsgBox, Please select the treatment type (Trial, TC or Purchase)
		Return
	}
	
	If (RadioTreatmentTypeFixed = "1")
		treatmentType = Fixed
	Else If (RadioTreatmentTypeAuto = "1")
		treatmentType = Auto
		
	PrescriptionID := % patURN . "P01"	;set SQL ID code for Prescription table
;FUTURE UPDATE: Allow for more than Prescrition to be added ie: URNXXXXXP02, URNXXXXP03 etc
		
	;convert checkbox options into a CSV string for insertion into SQL
	strOption := % Oximetry . "," . HeatedHumidifier . "," . Chinstrap . "," . NoseMask . "," . FullFaceMask . "," . PressureRelief
	
	staff_Array := StrSplit(PrescribingDoctor, ",")
	staffFirstName := staff_Array[2]
	staffLastName := staff_Array[1]
	staffFirstName = %staffFirstName%
	staffLastName = %staffLastName%
	queryStatement := % "SELECT StaffID FROM Staff WHERE FirstName = '" . staffFirstName . "' AND LastName = '" . staffLastName . "'"
	objReturn := ADOSQL(TherapyConnect, queryStatement)
	PrescribingDoctor := % objReturn[2,1]	;returns DR code from Therapy SQL DB
	PrescYYYY := SubStr(PrescriptionDate, 1, 4)
	PrescMM := SubStr(PrescriptionDate, 5, 2)
	PrescDD := SubStr(PrescriptionDate, 7, 2)
	PrescriptionDate := % PrescDD . "/" . PrescMM . "/" . PrescYYYY	;formats date for insertion into Therapy SQL DB
	
	prescStatement1 := % "INSERT INTO Prescriptions VALUES ('" . PrescriptionID . "'
		, '" . patURN . "', '" . PrescriptionDate . "', '" . RadioPrescriptionType . "', '" . treatmentType . "', '" . strOption . "'
		, '" . CPAPRecommendedPressure . "', '" . RampTime . "', '" . MinCPAP . "', '" . MaxCPAP . "'
		, '" . SettlingTime . "', '" . DoctorsComments . "', '" . PrescribingDoctor . "')"

	objReturn := ADOSQL(TherapyConnect, prescStatement1)	;execute SQL query
	GuiControl, Disable, btnGenPresc	;disable this button so that no more edits can be made to prescription
	MsgBox, Prescription saved!
	Return
}

;Disable or Enable edit fields depending on state of 'Fixed' or 'Auto' treatment type	
TreatmentCheck:
{
	Gui, Submit, NoHide
	If (RadioTreatmentTypeFixed = 1)
	{
		GuiControl, Enable, CPAPRecommendedPressure
		GuiControl, Enable, RampTime
		GuiControl, Disable, MaxCPAP
		GuiControl, Disable, MinCPAP
		GuiControl, Disable, SettlingTime
		GuiControl, Enable, strCPAPRecommendedPressure
		GuiControl, Enable, strCPAPRecommendedPressurecmH2O
		GuiControl, Enable, strRampTime
		GuiControl, Enable, strRampTimeMinutes
		GuiControl, Disable, strMaxCPAP
		GuiControl, Disable, strMaxCPAPcmH2O
		GuiControl, Disable, strMinCPAP
		GuiControl, Disable, strMinCPAPcmH2O
		GuiControl, Disable, strSettlingTime
		GuiControl, Disable, strSettlingTimeMinutes
	}
	If (RadioTreatmentTypeAuto = 1)
	{
		GuiControl, Disable, CPAPRecommendedPressure
		GuiControl, Disable, RampTime
		GuiControl, Enable, MaxCPAP
		GuiControl, Enable, MinCPAP
		GuiControl, Enable, SettlingTime
		GuiControl, Disable, strCPAPRecommendedPressure
		GuiControl, Disable, strCPAPRecommendedPressurecmH2O
		GuiControl, Disable, strRampTime
		GuiControl, Disable, strRampTimeMinutes
		GuiControl, Enable, strMaxCPAP
		GuiControl, Enable, strMaxCPAPcmH2O
		GuiControl, Enable, strMinCPAP
		GuiControl, Enable, strMinCPAPcmH2O
		GuiControl, Enable, strSettlingTime
		GuiControl, Enable, strSettlingTimeMinutes
	}
	Return
}

;Subroutine to reset GUI when changing patient
ButtonChangePatient:
{
	Gui, Destroy
	GoSub, InitialView
	Return
}

;Subroutine to write Issue details to Therapy SQL DB and **FUTURE** generate Issue summary for processing by admin staff
ButtonFinaliseIssue:
{
	;MsgBox,4,, Have you completed education and maintenance instructions for this patient?
	;IfMsgBox No
	;	Return
	;Else
	{
		Gui, Submit, NoHide
		StringUpper, patURN, patURN
		IssueID := % PatURN . "S01"
		
		staff_Array := StrSplit(IssueScientist, ",")
		staffFirstName := staff_Array[2]
		staffLastName := staff_Array[1]
		staffFirstName = %staffFirstName%
		staffLastName = %staffLastName%
		queryStatement := % "SELECT StaffID FROM Staff WHERE FirstName = '" . staffFirstName . "' AND LastName = '" . staffLastName . "'"
		objReturn := ADOSQL(TherapyConnect, queryStatement)
		StaffID := % objReturn[2,1]
		IssueYYYY := SubStr(IssueDate, 1, 4)
		IssueMM := SubStr(IssueDate, 5, 2)
		IssueDD := SubStr(IssueDate, 7, 2)
		IssueDateFormatted := % IssueDD . "/" . IssueMM . "/" . IssueYYYY
		
		queryStatement1 := % "INSERT INTO Issues VALUES ('" . IssueID . "'
		, '" . patURN . "', '" . IssueDateFormatted . "', '" . StaffID . "', '" . RadioMonitored . "'
		, '" . RadioPatientSleep . "', '" . IssueDevice . "', '" . IssuePressure . "', '" . IssueEPR . "'
		, '" . IssueRamp . "', '" . IssueMask . "', '" . IssueMaskSize . "', '" . IssueChinstrap . "', '" . IssueChinstrapSize . "'
		, '" . IssueOther . "', '" . IssueComments . "')"
		objReturn := ADOSQL(TherapyConnect, queryStatement1)
		
		If RadioMonitored = 1
			RadioMonitored := "Y"
		Else
			RadioMonitored := "N"
		
		If RadioPatientSleep = 1
			RadioPatientSleep := "Y"
		Else
			RadioPatientSleep := "N"
	}
	GuiControl, Disable, btnFinIssue
	
	issueTemplatePath := "C:\temp\AutoHotKeyScripts\TherapyClinic\TherapyClinic\Templates\EquipmentIssue.dotx"
	issueDocName := % IssueID . "_" . IssueDate
	issueDocumentPath = C:\IO8Takeup\%issueDocName%.docx
	
	wdApp := ComObjCreate("Word.Application")
	wdApp.Visible := true
	issueDoc := wdApp.Documents.Add(issueTemplatePath)
	
	bmarkPatientName := issueDoc.bookmarks.item("patientName").Range
	bmarkPatientName.Select()
	wdApp.Selection.TypeText(patLastName ", " patFirstName)
	
	bmarkPatientURN := issueDoc.bookmarks.item("patientURN").Range
	bmarkPatientURN.Select()
	wdApp.Selection.TypeText(patURN)
	
	bmarkPatientDOB := issueDoc.bookmarks.item("patientDOB").Range
	bmarkPatientDOB.Select()
	wdApp.Selection.TypeText(dateOfBirth)
	
	bmarkIssueDate := issueDoc.bookmarks.item("issueDate").Range
	bmarkIssueDate.Select()
	wdApp.Selection.TypeText(IssueDateFormatted)
	
	bmarkIssueComments := issueDoc.bookmarks.item("issueComments").Range
	bmarkIssueComments.Select()
	wdApp.Selection.TypeText(IssueComments)
	
	bmarkIssueDoctor := issueDoc.bookmarks.item("issueDoctor").Range
	bmarkIssueDoctor.Select()
	wdApp.Selection.TypeText(PrescribingDoctor)
	
	bmarkIssueScientist := issueDoc.bookmarks.item("issueScientist").Range
	bmarkIssueScientist.Select()
	wdApp.Selection.TypeText(IssueScientist)
	
	bmarkCpapDevice := issueDoc.bookmarks.item("cpapDevice").Range
	bmarkCpapDevice.Select()
	wdApp.Selection.TypeText(IssueDevice)
	
	bmarkCpapPressure := issueDoc.bookmarks.item("cpapPressure").Range
	bmarkCpapPressure.Select()
	wdApp.Selection.TypeText(IssuePressure)
	
	bmarkCpapEPR := issueDoc.bookmarks.item("cpapEPR").Range
	bmarkCpapEPR.Select()
	wdApp.Selection.TypeText(IssueEPR)
	
	bmarkCpapRamp := issueDoc.bookmarks.item("cpapRamp").Range
	bmarkCpapRamp.Select()
	wdApp.Selection.TypeText(IssueRamp)
	
	bmarkIssueMask := issueDoc.bookmarks.item("maskMask").Range
	bmarkIssueMask.Select()
	wdApp.Selection.TypeText(IssueMask)
	
	bmarkIssueMaskSize := issueDoc.bookmarks.item("maskMaskSize").Range
	bmarkIssueMaskSize.Select()
	wdApp.Selection.TypeText(IssueMaskSize)
	
	bmarkIssueChinstrap := issueDoc.bookmarks.item("maskChinstrap").Range
	bmarkIssueChinstrap.Select()
	wdApp.Selection.TypeText(IssueChinstrap)
	
	bmarkIssueChinstrapSize := issueDoc.bookmarks.item("maskChinstrapSize").Range
	bmarkIssueChinstrapSize.Select()
	wdApp.Selection.TypeText(IssueChinstrapSize)
	
	bmarkIssueOther := issueDoc.bookmarks.item("maskOther").Range
	bmarkIssueOther.Select()
	wdApp.Selection.TypeText(IssueOther)
	
	bmarkIssueMonitored := issueDoc.bookmarks.item("issueMonitored").Range
	bmarkIssueMonitored.Select()
	wdApp.Selection.TypeText(RadioMonitored)
	
	bmarkIssuePatientSleep := issueDoc.bookmarks.item("issueSleep").Range
	bmarkIssuePatientSleep.Select()
	wdApp.Selection.TypeText(RadioPatientSleep)
		
	wdApp.ActiveDocument.SaveAs(issueDocumentPath)

	
	MsgBox, Issue saved to IO takeup folder!
	Return
}

;Subroutine to add Therapy Clinic Visit tab
ButtonAddTCVisit:
{
	MsgBox,4,, Add a new Therapy Clinic Visit?
	IfMsgBox Yes
	{
		tabList .= "|Visit " . (tabs++)	;update list of names for tabs
			
		Gui, Tab, %tabs%	;select the newest tab
		GoSub, VisitDetailsGui
		GuiControl,,VisitTab,|%tabList%	;add the new tab
		GuiControl,Choose, VisitTab, |%tabs%	;set focus to the new tab
	}
	Return
}

;Subroutine to populate GUI with equipment issue fields
EquipmentIssueGUI:
{
;	Gui, Font, s08 w600
;	Gui, Add, Groupbox, x690 y20 w660 h245, Equipment Issue
	Gui, Font, s08 Normal

	Gui, Add, Text, x720 y100, Issue Date
	Gui, Add, DateTime, x790 y98 w120 h21 vIssueDate
	Gui, Add, Text, x720 y130, Scientist
	Gui, Add, DropDownList, x790 y128 w120 h21 r%numSCNames% vIssueScientist, %SCNames%	;populated with active scientist names
	
	Gui, Add, Text, x1000 y100, Monitored:
	Gui, Add, Radio, x1110 y100 vRadioMonitored, Y
	Gui, Add, Radio, x1155 y100, N
	Gui, Add, Text, x1000 y130, Did the patient sleep?
	Gui, Add, Radio, x1110 y130 vRadioPatientSleep, Y
	Gui, Add, Radio, x1155 y130, N
	
	Gui, Font, s08 w600 Italic
	Gui, Add, Groupbox, x705 y160 w630 h195, Issued Equipment
	Gui, Font, Normal	
	
	Gui, Font, w600	
	Gui, Add, GroupBox, x730 y190 w250 h150, CPAP
	Gui, Font, Normal
	Gui, Add, Text, x745 y220, Device
	Gui, Add, Edit, x805 y220 w140 h21 vIssueDevice
	Gui, Add, Text, x745 y250, Pressure
	Gui, Add, Edit, x805 y250 w30 h21 vIssuePressure
	Gui, Add, Text, x845 y250, cmH2O
	Gui, Add, Text, x745 y280, C-Flex/EPR
	Gui, Add, Edit, x805 y280 w30 h21 vIssueEPR
	Gui, Add, Text, x875 y280, Ramp
	Gui, Add, Edit, x915 y280 w30 h21 vIssueRamp
	
	Gui, Font, w600
	Gui, Add, GroupBox, x1020 y190 w300 h150, Mask
	Gui, Font, Normal
	Gui, Add, Text, x1035 y220, Mask
	Gui, Add, Edit, x1085 y220 w110 h21 vIssueMask
	Gui, Font, Italic
	Gui, Add, Text, x1215 y220, Size
	Gui, Font, Normal
	Gui, Add, Edit, x1250 y220 w45 h21 vIssueMaskSize
	Gui, Add, Text, x1035 y250, Chinstrap
	Gui, Add, Edit, x1085 y250 w110 h21 vIssueChinstrap
	Gui, Font, Italic
	Gui, Add, Text, x1215 y250, Size
	Gui, Font, Normal
	Gui, Add, Edit, x1250 y250 w45 h21 vIssueChinstrapSize
	Gui, Add, Text, x1035 y280, Other
	Gui, Add, Edit, x1085 y280 w110 h21 vIssueOther
	
	Gui, Font, s08 w600 Italic
	Gui, Add, Text, x705 y385, Comments:
	Gui, Font, Normal
	Gui, Add, Edit, x705 y410 w630 h150 vIssueComments
	
	Gui, Font, s10 w600
	Gui, Add, Button, x1135 y600 w200 h30 vbtnFinIssue, Finalise Issue
	Gui, Font, s08 Normal
	
	;Gui, Font, s10 w600
	;Gui, Add, Button, x710 y710 w140 h30 vbtnAddTC, Add TC Visit
	;Gui, Font, s08 Normal
	
	Return
}

;Subroutine to populate GUI tab with therapy clinic fields
VisitDetailsGUI:
{
	Gui, Add, Text, x705 y80, Scientist
	Gui, Add, DropDownList, x765 y80 w120 h21 r%numSCNames% vScientist%tabs%, %SCNames%
	Gui, Add, Text, x705 y110, Doctor
	Gui, Add, DropDownList, x765 y110 w120 h21 r%numDRNames% vDoctor%tabs%, %DRNames%
	Gui, Add, Text, x705 y140, Visit Date
	Gui, Add, DateTime, x765 y140 w120 h21 vVisitDate%tabs%
	
	Gui, Add, Text, x945 y80, ESS:
	Gui, Add, Edit, x1010 y80 w45 h21 vESS%tabs%
	Gui, Add, Text, x945 y110, Compliance:
	Gui, Add, Edit, x1010 y110 w45 h21 vCompliance%tabs%
	Gui, Add, Text, x1065 y110, hours/day
	Gui, Add, Text, x1165 y80, AHI:
	Gui, Add, Edit, x1220 y80 w45 h21 vAHI%tabs%
	Gui, Add, Text, x1275 y80, events/hour
	Gui, Add, Text, x1165 y110, Leak:
	Gui, Add, Edit, x1220 y110 w45 h21 vLeak%tabs%
	Gui, Add, Text, x1275 y110, Litres/min
	
	Gui, Font, w600
	Gui, Add, GroupBox, x730 y170 w290 h235, Outcome
	Gui, Font, Normal
	Gui, Add, Text, x745 y195, Clinical Response:
	Gui, Add, Edit, x745 y220 w260 h70 vClinicalResponse%tabs%
	Gui, Add, Text, x745 y300, Side Effects:
	Gui, Add, ListBox, x745 y325 w260 h70 Multi vSideEffects%tabs%, %SENames%
	
	Gui, Font, w600
	Gui, Add, GroupBox, x1040 y170 w290 h235, Plan
	Gui, Font, Normal
	Gui, Add, Text, x1055 y195, Alterations
	Gui, Add, Edit, x1055 y220 w260 h70 vAlterations%tabs%
	Gui, Add, Text, x1055 y300, Follow up:
	Gui, Add, Edit, x1055 y325 w260 h70 vFollowUp%tabs%

;Add Equipment below
	Gui, Font, s08 w600 Italic
	Gui, Add, Groupbox, x730 y410 w600 h160, Equipment
	Gui, Font, Normal	
	
	Gui, Font, w600	
	Gui, Add, GroupBox, x745 y435 w250 h120, CPAP
	Gui, Font, Normal
	Gui, Add, Text, x760 y465, Device
	Gui, Add, Edit, x820 y465 w140 h21 vDevice%tabs%
	Gui, Add, Text, x760 y495, Pressure
	Gui, Add, Edit, x820 y495 w30 h21 vPressure%tabs%
	Gui, Add, Text, x860 y495, cmH2O
	Gui, Add, Text, x760 y525, C-Flex/EPR
	Gui, Add, Edit, x820 y525 w30 h21 vEPR%tabs%
	Gui, Add, Text, x890 y525, Ramp
	Gui, Add, Edit, x930 y525 w30 h21 vRamp%tabs%
	
	;**FUTURE** Add ability to have up to 3 masks per visit
	Gui, Font, w600
	Gui, Add, GroupBox, x1020 y435 w300 h120, Mask
	Gui, Font, Normal
	Gui, Add, Text, x1035 y465, Mask
	Gui, Add, Edit, x1085 y465 w110 h21 vMask%tabs%
	Gui, Font, Italic
	Gui, Add, Text, x1215 y465, Size
	Gui, Font, Normal
	Gui, Add, Edit, x1250 y465 w45 h21 vMaskSize%tabs%
	Gui, Add, Text, x1035 y495, Chinstrap
	Gui, Add, Edit, x1085 y495 w110 h21 vChinstrap%tabs%
	Gui, Font, Italic
	Gui, Add, Text, x1215 y495, Size
	Gui, Font, Normal
	Gui, Add, Edit, x1250 y495 w45 h21 vChinstrapSize%tabs%
	
	Gui, Font, s08 w600 Italic
	Gui, Add, Text, x730 y580, Comments:
	Gui, Font, Normal
	Gui, Add, Edit, x800 y580 w300 h70 vComments%tabs%
	
	Gui, Font, s10 w600
	Gui, Add, Button, x1130 y580 w140 h30 vVisitSummary%tabs%, Visit Summary
	Gui, Add, Button, x1130 y620 w140 h30 vTherapySummary%tabs%, Therapy Summary
	Gui, Font, s08 Normal
	
	Return
}

;Subroutine to complete individual visit and **FUTURE** create summary document for admin staff
;**UNDER DEVELOPMENT STILL**
ButtonVisitSummary:
{
	Gui, Submit, NoHide
	GuiControlGet,CurrentTab,,VisitTab
	Gui 2: Add, Text, x30 y30 w40 h21, Doctor:
	Gui 2: Add, Text, x100 y30 w120 h21 vIsSmDr
	Gui 2: Add, Text, x30 y55 w120 h21 vIsSmPatURN
	Gui 2: Add, Text, x30 y80 w120 h21 vIsSmPatName
	Gui 2: Add, Text, x30 y105 w120 h21 vIsSmPatDOB

	CurrentDoctor := Doctor%CurrentTab%
	CurrentScientist := Scientist%CurrentTab%
	CurrentVisitDate := VisitDate%CurrentTab%
	GuiControl 2:, IsSmDr, %CurrentDoctor%
	GuiControl 2:, IsSmPatURN, %patURN%
	GuiControl 2:, IsSmPatName, % patLastName . ", " . patFirstName
	GuiControl 2:, IsSmPatDOB, %dateOfBirth%
	Gui 2: Add, Button, gbtnIsSmOK, OK
	Gui 2: Show, w300 h300

	Return
}

;TESTING BELOW FOR ButtonVisitSummary subroutine (directly above)
btnIsSmOK:
{
	Gui 2: Destroy
	Return
}
	
;subroutine to display GUI for selecting correct options for the therapy summary document, then create the document in the InfoOrg takeup folder
;ButtonTherapySummary:
;		{
;		Gui 2: Add, Text, , Hi there
;		Gui 2: Show
;		Return
;		}

;Subroutine to fill prescription details from Therapy SQL DB	
PrescriptionFill:
{
	NumPresc = 0	;counter for number of prescriptions, for future use when adding support for multiple prescriptions
	prescDates := ""
	eqDevices := ""
	eqRecommendedPressures := ""
	
	prescDate := % prescriptionObjReturn[2,3]
	
	If (prescDate = "")	;if prescription date is empty then no valid prescription exists
	{
		Return
	}
	Else
	{
		prescYYYY := SubStr(prescDate, 7, 4)
		prescMM := SubStr(prescDate, 4, 2)
		prescDD := SubStr(prescDate, 1, 2)
		prescDate := % prescYYYY . prescMM . prescDD
		MsgBox, Presc date is %prescDate%	;TESTING
	}
	
	prescType := % prescriptionObjReturn[2,4]
	prescOptions := % prescriptionObjReturn[2,6]
	prescRecPress := % prescriptionObjReturn[2,7]
	prescRamp := % prescriptionObjReturn[2,8]
	prescMinCPAP := % prescriptionObjReturn[2,9]
	prescMaxCPAP := % prescriptionObjReturn[2,10]
	prescSettlingTime := prescriptionObjReturn[2,11]
	prescComments := % prescriptionObjReturn[2,12]
	prescDoc := % prescriptionObjReturn[2,13]
	
	prescDoc = '%prescDoc%'
	staffQuery := "SELECT * FROM Staff WHERE StaffID = "prescDoc
	staffObjReturn := ADOSQL(TherapyConnect, staffQuery)
	
	DRfirstName := % staffObjReturn[2,2]
	DRlastName := % staffObjReturn[2,3]
	DRName .= DRlastName . ", " . DRfirstName
	
	prescOptions := StrSplit(prescOptions, ",")
	
	If (prescType = "Auto")
	{		
		prescTypeAuto := 1
		prescTypeFixed := 0
		GuiControl, Disable, CPAPRecommendedPressure
		GuiControl, Disable, RampTime
		GuiControl, Enable, MaxCPAP
		GuiControl, Enable, MinCPAP
		GuiControl, Enable, SettlingTime
	}
	Else
	{
		prescTypeAuto := 0
		prescTypeFixed := 1
		GuiControl, Enable, CPAPRecommendedPressure
		GuiControl, Enable, RampTime
		GuiControl, Disable, MaxCPAP
		GuiControl, Disable, MinCPAP
		GuiControl, Disable, SettlingTime
	}
	
	GuiControl,, PrescriptionDate, %prescDate%
	GuiControl, , RadioTreatmentTypeAuto, %prescTypeAuto%
	GuiControl, , RadioTreatmentTypeFixed, %prescTypeFixed%
	GuiControl,, CPAPRecommendedPressure, %prescRecPress%
	GuiControl,, RampTime, %prescRamp%
	GuiControl,, MaxCPAP, %prescMaxCPAP%
	GuiControl,, MinCPAP, %prescMinCPAP%
	GuiControl,, SettlingTime, %prescSettlingTime%
	GuiControl,, DoctorsComments, %prescComments%
	GuiControl, Choose, PrescribingDoctor, %DRName%
	GuiControl, Disable, btnGenPresc
	
;	Loop
;	{
;		
;		i := A_Index + 1
;		numPresc++
;		prescDate := % prescriptionObjReturn[i,3]
;		prescDevice := % prescriptionObjReturn[i,4]
;		prescRecommendedPressure := % PrescriptionObjReturn[i,6]
;		If (prescDate <> "")
;			prescDates .= prescDate . "|"
;			prescDevices .= prescDevice . "|"
;			prescRecommendedPressures .= prescRecommendedPressure . "|"
;			MsgBox, %prescDate%
;			GuiControl,, PrescriptionDate, %prescDate%
;	}
;	Until prescDate = ""


	Return
}

;Subroutine to fill Equipment Issue details from Therapy SQL DB	
IssueFill:
{
	
	Return
}
	
;Subroutine to fill Therapy Clinic visit details from Therapy SQL DB	
VisitFill:
{
	
	Return
}

;Subroutine to disable patient details
PatientDetailsDisable:
{
	patName := ""
	patDOB := ""
	Veteran := ""
	GuiControl,, patName,
	GuiControl,, patDOB,
	Return
}
	
^Esc::ExitApp

GuiClose:
{
IfWinExist, Word
	wdApp.Application.Quit()
ExitApp
}