'GATHERING STATS----------------------------------------------------------------------------------------------------
'name_of_script = "ACTIONS - INTAKE.vbs"
'start_time = timer

'LOADING ROUTINE FUNCTIONS FROM GITHUB REPOSITORY---------------------------------------------------------------------------
url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, FALSE									'Attempts to open the URL
req.send													'Sends request
IF req.Status = 200 THEN									'200 means great success
	Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
	Execute req.responseText								'Executes the script code
ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
	MsgBox 	"Something has gone wrong. The code stored on GitHub was not able to be reached." & vbCr &_ 
			vbCr & _
			"Before contacting Veronica Cary, please check to make sure you can load the main page at www.GitHub.com." & vbCr &_
			vbCr & _
			"If you can reach GitHub.com, but this script still does not work, ask an alpha user to contact Veronica Cary and provide the following information:" & vbCr &_
			vbTab & "- The name of the script you are running." & vbCr &_
			vbTab & "- Whether or not the script is ""erroring out"" for any other users." & vbCr &_
			vbTab & "- The name and email for an employee from your IT department," & vbCr & _
			vbTab & vbTab & "responsible for network issues." & vbCr &_
			vbTab & "- The URL indicated below (a screenshot should suffice)." & vbCr &_
			vbCr & _
			"Veronica will work with your IT department to try and solve this issue, if needed." & vbCr &_ 
			vbCr &_
			"URL: " & url
			StopScript
END IF

'VARIABLES THAT NEED DECLARING----------------------------------------------------------------------------------------------------
checked = 1
unchecked = 0

'DIALOGS----------------------------------------------------------------------------------------------------
BeginDialog PRISM_case_number_dialog, 0, 0, 186, 50, "PRISM case number dialog"
  EditBox 100, 10, 80, 15, PRISM_case_number
  ButtonGroup ButtonPressed
    OkButton 35, 30, 50, 15
    CancelButton 95, 30, 50, 15
  Text 5, 10, 90, 20, "PRISM case number (XXXXXXXXXX-XX format):"
EndDialog

BeginDialog CS_intake_dialog, 0, 0, 371, 315, "CS intake dialog"
  CheckBox 15, 30, 145, 10, "Case Opening - Welcome Letter", NCP_welcome_ltr_check
  CheckBox 15, 45, 140, 10, "Court Order Summary", ncp_court_order_summary_check
  CheckBox 15, 60, 50, 10, "DORD F0999 - PIN Notice", NCP_PIN_Notice_Check
  CheckBox 15, 75, 150, 10, "DORD F0924 - Health Insurance Verification", NCP_health_ins_verif_check
  CheckBox 15, 90, 120, 10, "Notice of Arrears Reported", arrears_reported_check
  CheckBox 50, 115, 65, 10, "DORD F0100", dord_F0100_check
  CheckBox 50, 145, 65, 10, "DORD F0109", dord_F0109_check
  CheckBox 50, 175, 60, 10, "DORD F0107", dord_F0107_check
  CheckBox 15, 220, 125, 10, "Set File Location to QC 30", qc_30_file_loc_check
  CheckBox 15, 235, 115, 10, "Set File Location to SAFETY", safety_file_loc_check
  CheckBox 195, 30, 130, 10, "Case Opening - Welcome Letter", CP_welcome_ltr_check
  CheckBox 195, 45, 115, 10, "CP New Order Summary", CP_new_order_summary_check
  CheckBox 195, 60, 50, 10, "DORD F0999 - PIN Notice", CP_PIN_Notice_check
  CheckBox 195, 75, 155, 10, "DORD F0924 - Health Insurance Verification", NCP_Health_Ins_check
  CheckBox 195, 90, 130, 10, "Child Care Verification", child_care_verif_check
  CheckBox 195, 105, 125, 10, "CP Statement of Arrears Letter", CP_Stmt_of_Arrears_check
  CheckBox 200, 135, 105, 10, "10 day tickler to call NCP", 10_day_ticker_check
  CheckBox 200, 150, 110, 10, "30 day tickler to load arrears", 30_day_to_load_arrears_check
  CheckBox 200, 165, 105, 10, "30 day case review", 30_day_case_review
  EditBox 210, 175, 140, 15, 30_day_cawd_txt
  CheckBox 200, 195, 105, 10, "60 day case review", 60_day_case_review
  EditBox 210, 205, 140, 15, 60_day_cawd_txt
  EditBox 240, 235, 110, 15, worker_name
  EditBox 240, 255, 110, 15, worker_phone
  EditBox 260, 275, 90, 15, worker_signature
  ButtonGroup ButtonPressed
    OkButton 225, 290, 50, 20
    CancelButton 275, 290, 50, 20
  Text 185, 275, 70, 10, "Sign your CAAD note:"
  Text 35, 105, 30, 10, "NPA"
  Text 35, 135, 95, 10, "MFIP, DWP, CCA"
  Text 35, 165, 90, 10, "MA only"
  Text 185, 255, 50, 10, "Worker phone:"
  Text 185, 235, 50, 10, "Worker name:"
  GroupBox 5, 205, 170, 50, "File Location on CAST"
  GroupBox 5, 15, 170, 180, "Letters to NCP"
  GroupBox 185, 15, 170, 105, "Letters to CP"
  Text 5, 0, 325, 15, "Enforcement Intake Script"
  GroupBox 185, 125, 170, 105, "CAWD notes to add"
EndDialog

'CUSTOM FUNCTION***************************************************************************************************************


FUNCTION send_dord_doc(recipient, dord_doc)
	call navigate_to_PRISM_screen("DORD")
	EMWriteScreen "C", 3, 29
	transmit
	EMWriteScreen "A", 3, 29
	EMWriteScreen dord_doc, 6, 36
	EMWriteScreen recipient, 11, 51
	transmit
END FUNCTION
	

'THE SCRIPT----------------------------------------------------------------------------------------------------

'Connects to BlueZone
EMConnect ""

'Finds the PRISM case number using a custom function
call PRISM_case_number_finder(PRISM_case_number)

'Shows case number dialog
Do
	Do
		Dialog PRISM_case_number_dialog
		If buttonpressed = 0 then stopscript
		call PRISM_case_number_validation(PRISM_case_number, case_number_valid)
		If case_number_valid = False then MsgBox "Your case number is not valid. Please make sure it uses the following format: ''XXXXXXXXXX-XX''"
	Loop until case_number_valid = True
	transmit
	EMReadScreen PRISM_check, 5, 1, 36
	If PRISM_check <> "PRISM" then MsgBox "You seem to be locked out of PRISM. Are you in PRISM right now? Are you passworded out? Check your BlueZone screen and try again."
Loop until PRISM_check = "PRISM"

'Clearing case info from PRISM
call navigate_to_PRISM_screen("REGL")
transmit

'Navigating to CAPS
call navigate_to_PRISM_screen("CAPS")

'Entering case number and transmitting
EMSetCursor 4, 8
EMSendKey replace(PRISM_case_number, "-", "")									'Entering the specific case indicated
EMWriteScreen "d", 3, 29												'Setting the screen as a display action
transmit															'Transmitting into it

'Getting worker info for case note
EMSetCursor 5, 53
PF1
EMReadScreen worker_name, 27, 6, 50
EMReadScreen worker_phone, 12, 8, 35
PF3

'Cleaning up worker info
worker_name = trim(worker_name)
call fix_case(worker_name, 1)
worker_name = change_client_name_to_FML(worker_name)


'Shows intake dialog, checks to make sure we're still in PRISM (not passworded out)
Do
	Dialog CS_intake_dialog
	If buttonpressed = 0 then stopscript
	transmit
	EMReadScreen PRISM_check, 5, 1, 36
	If PRISM_check <> "PRISM" then MsgBox "You seem to be locked out of PRISM. Are you in PRISM right now? Are you passworded out? Check your BlueZone screen and try again."
Loop until PRISM_check = "PRISM"


'Creating the Word application object (if any of the Word options are selected), and making it visible 
If _
	NCP_welcome_ltr_check = checked or _
	ncp_court_order_summary_check = checked or _
	arrears_reported_check = checked or _
	CP_welcome_ltr_check = checked or _
	CP_Stmt_of_Arrears_check = checked or _
	child_care_verif_check = checked or _
	CP_new_order_summary_check = checked then
		Set objWord = CreateObject("Word.Application")
		objWord.Visible = True
End if


'NCP Welcome Letter
If NCP_welcome_ltr_check = checked then
'	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\New Folder\CP Paternity Request Sheet.dotx")
	With objDoc
'		.FormFields("field_childs_name").Result = childs_name
'		.FormFields("field_CP_name").Result = CP_name
'		.FormFields("field_AF_name").Result = NCP_name
'		.FormFields("field_case_number").Result = PRISM_case_number
	End With
End if

'NCP Court Order Summary
If ncp_court_order_summary_check = checked then
'	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\New Folder\Financial Affidavit OCS.dotx")
	With objDoc
'		.FormFields("field_case_number").Result = PRISM_case_number
'		.FormFields("field_all_children").Result = CAPS_kids
'		.FormFields("field_CP_name").Result = CP_name
	End With
End if

'Arrears Reported
If arrears_reported_check = checked then
'	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\New Folder\Paternity Cover letter to CP - Normal.dotx")
	With objDoc
'		.FormFields("field_name").Result = CP_name
'		.FormFields("field_street_address").Result = street_address
'		.FormFields("field_city_state_zip").Result = city_state_zip
'		.FormFields("field_NCP_gender").Result = NCP_gender
'		.FormFields("field_NCP_gender_02").Result = NCP_gender
'		.FormFields("field_NCP_gender_03").Result = NCP_gender
'		.FormFields("field_NCP_gender_04").Result = NCP_gender
'		.FormFields("field_case_number").Result = PRISM_case_number
'		.FormFields("field_date_plus_five").Result = dateadd("d", date, 5)
'		.FormFields("field_phone").Result = worker_phone
	End With
End if

'CP Welcome Letter
If CP_welcome_ltr_check = checked then
'	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\New Folder\Paternity Cover letter to CP - Relative Caretaker.dotx")
	With objDoc
'		.FormFields("field_name").Result = CP_name
'		.FormFields("field_street_address").Result = street_address
'		.FormFields("field_city_state_zip").Result = city_state_zip
'		.FormFields("field_case_number").Result = PRISM_case_number
'		.FormFields("field_date_plus_five").Result = dateadd("d", date, 5)
'		.FormFields("field_phone").Result = worker_phone
	End With
End if

'CP Statment of Arrears
If CP_Stmt_of_Arrears_check = checked then
'	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\New Folder\Paternity Cover letter to CP - Minor with GAL attachment.dotx")
	With objDoc
'		.FormFields("field_name").Result = CP_name
'		.FormFields("field_street_address").Result = street_address
'		.FormFields("field_city_state_zip").Result = city_state_zip
'		.FormFields("field_case_number").Result = PRISM_case_number
'		.FormFields("field_date_plus_five").Result = dateadd("d", date, 5)
'		.FormFields("field_phone").Result = worker_phone
'		.FormFields("field_name_02").Result = CP_name
'		.FormFields("field_case_number_02").Result = PRISM_case_number
	End With
End if

'Child Care Verification
If child_care_verif_check = checked then
'	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\New Folder\Establishment Intake Letter.dotx")
	With objDoc
'		.FormFields("CPName").Result = CP_name
'		.FormFields("CP_address").Result = street_address
'		.FormFields("CP_CSZ").Result = city_state_zip
'		.FormFields("PRISM_No").Result = PRISM_case_number
'		.FormFields("CPName_2").Result = CP_name
'		.FormFields("Due_Date").Result = dateadd("d", date, 5)
'		.FormFields("worker").Result = worker_name
	End With
End if
'CP New Order Summary
If CP_new_order_summary_check = checked then
'	set objDoc = objWord.Documents.Add("Q:\Blue Zone Scripts\Word documents for script use\New Folder\Establishment Intake Letter.dotx")
	With objDoc
'		.FormFields("CPName").Result = CP_name
'		.FormFields("CP_address").Result = street_address
'		.FormFields("CP_CSZ").Result = city_state_zip
'		.FormFields("PRISM_No").Result = PRISM_case_number
'		.FormFields("CPName_2").Result = CP_name
'		.FormFields("Due_Date").Result = dateadd("d", date, 5)
'		.FormFields("worker").Result = worker_name
	End With
End if

   CheckBox 15, 30, 145, 10, "Case Opening - Welcome Letter", NCP_welcome_ltr_check
  CheckBox 15, 45, 140, 10, "Court Order Summary", ncp_court_order_summary_check
  CheckBox 15, 60, 50, 10, "PIN Notice", NCP_PIN_Notice_Check
  CheckBox 15, 75, 150, 10, "DORD F0924 - Health Insurance Verification", NCP_health_ins_verif_check
  CheckBox 15, 90, 120, 10, "Notice of Arrears Reported", arrears_reported_check
  CheckBox 50, 115, 65, 10, "DORD F0100", dord_F0100_check
  CheckBox 50, 145, 65, 10, "DORD F0109", dord_F0109_check
  CheckBox 50, 175, 60, 10, "DORD F0107", dord_F0107_check
  CheckBox 15, 220, 125, 10, "Set File Location to QC 30", qc_30_file_loc_check
  CheckBox 15, 235, 115, 10, "Set File Location to SAFETY", safety_file_loc_check
  CheckBox 195, 30, 130, 10, "Case Opening - Welcome Letter", CP_welcome_ltr_check
  CheckBox 195, 45, 115, 10, "CP New Order Summary", CP_new_order_summary_check
  CheckBox 195, 60, 50, 10, "PIN Notice", CP_PIN_Notice_check
  CheckBox 195, 75, 155, 10, "DORD F0924 - Health Insurance Verification", NCP_Health_Ins_check
  CheckBox 195, 90, 130, 10, "Child Care Verification", child_care_verif_check
  CheckBox 195, 105, 125, 10, "CP Statement of Arrears Letter", CP_Stmt_of_Arrears_check


'If F0018 is indicated on the dialog then it navigates to DORD to send it.
If F0018_check = checked then 'send Your Privacy Rights to NCP
	call navigate_to_PRISM_screen("DORD")
	EMWriteScreen "C", 3, 29
	transmit
	EMWriteScreen "A", 3, 29
	EMWriteScreen "F0018", 6, 36
	transmit
End if

'If F0100 is indicated on the dialog then it navigates to DORD to send it.
If F0100_check = checked then  'send Authorization for Support to NCP
	call navigate_to_PRISM_screen("DORD")
	EMWriteScreen "C", 3, 29
	transmit
	EMWriteScreen "A", 3, 29
	EMWriteScreen "F0100", 6, 36
	transmit
End if

'If F0022 is indicated on the dialog then it navigates to DORD to send it.
If F0022_check = checked then
	call navigate_to_PRISM_screen("DORD")
	send_msg = MsgBox("Do you want to send the F0022 Important Statement of Rights to both parties? Click Yes for both, or click No to send it to CP only.", vbYesNo)
	If send_msg = vbYes Then
		call send_dord_doc("CPP", "F0022")
		call send_dord_doc("NCP", "F0022")
	Else
		call send_dord_doc("CPP", "F0022")
	End If
End if

'If F5000 is indicated on the dialog then it navigates to DORD to send it.
If F5000_checkbox = checked then 'Send waiver of personal service to CP
	if caretaker_checkbox = checked then
		call navigate_to_PRISM_screen("DORD")
		EMWriteScreen "C", 3, 29
		transmit
		EMWriteScreen "A", 3, 29
		EMWriteScreen "F5000", 6, 36
		transmit
		Pf14
		EMWriteScreen "U", 20, 14
		transmit
		EMWriteScreen "S", 12, 5
		transmit
		EMWriteScreen "12", 16, 15
		transmit
		PF3
		EMWriteScreen "M", 3, 29
		transmit 
		PF3
	End if
End if
'If F0109 is indicated on the dialog then it navigates to DORD to send it.
If F0109_checkbox = checked then
	if caretaker_checkbox = checked then
		call send_dord_doc("NCP", "F0109")
	else
		call send_dord_doc("CPP", "F0109")
		call send_dord_doc("NCP", "F0109")
	end if
End if
'If F0021 is indicated on the dialog then it navigates to DORD to send it.
If F0021_checkbox = checked then
	if caretaker_checkbox = checked then
		call send_dord_doc("NCP", "F0021")
	else
		call send_dord_doc("CPP", "F0021")
		call send_dord_doc("NCP", "F0021")
	end if
End if


If CAAD_note_check = checked then

	'Going to CAAD, adding a new note
	call navigate_to_PRISM_screen("CAAD")
	EMWriteScreen "A", 8, 5
	transmit
	EMReadScreen case_activity_detail, 20, 2, 29
	If case_activity_detail <> "Case Activity Detail" then script_end_procedure("The script could not navigate to a case note. There may have been a script error. Add case note manually, and report the error to a script writer.")


	'Setting the type
	EMWriteScreen "M2123", 4, 54

	'Setting cursor in write area and writing note details
	EMSetCursor 16, 4
	call write_new_line_in_PRISM_case_note("* Paternity packet sent to CP with the following docs:")
	If child_only_MA_check = checked then call write_new_line_in_PRISM_case_note("    * Child Only MA Choice of Service letter")
	If child_only_MA_relative_caretaker_check = checked then call write_new_line_in_PRISM_case_note("    * Child Only MA - relative caretaker letter")
	If CP_and_child_MA_check = checked then call write_new_line_in_PRISM_case_note("    * CP and Child MA choice of service letter")
	If CP_paternity_request_sheet_check = checked then call write_new_line_in_PRISM_case_note("    * CP Paternity Request sheet")
	If financial_affidavit_OCS_check = checked then call write_new_line_in_PRISM_case_note("    * Financial Affidavit OCS")
	If issues_paternity_to_be_decided_check = checked then call write_new_line_in_PRISM_case_note("    * Issues-Paternity-to be Decided")
	If parenting_time_schedules_check = checked then call write_new_line_in_PRISM_case_note("    * Parenting Time Schedules")
	If paternity_cover_letter_normal_check = checked then call write_new_line_in_PRISM_case_note("    * Normal Paternity Cover Letter to CP")
	If paternity_cover_letter_relative_caretaker_check = checked then call write_new_line_in_PRISM_case_note("    * Relative Caretaker Paternity Cover Letter to CP")
	If paternity_cover_letter_minor_check = checked then call write_new_line_in_PRISM_case_note("    * Minor with GAL Attachment")
	If paternity_information_form_memo_check = checked then call write_new_line_in_PRISM_case_note("    * Paternity Information Form Memo")
	If paternity_information_form_check = checked then call write_new_line_in_PRISM_case_note("    * Paternity Information Form")
	If supplemental_paternity_information_form_check = checked then call write_new_line_in_PRISM_case_note("    * Supplemental Paternity Information Form")
	If Est_Ltr_checkbox = checked then call write_new_line_in_PRISM_case_note("    * Establishment Intake Letter")
	If F0018_checkbox = checked then call write_new_line_in_PRISM_case_note("    * DORD F0018")
	If F0021_checkbox = checked then call write_new_line_in_PRISM_case_note("    * DORD F0021")
	If F0022_check = checked then call write_new_line_in_PRISM_case_note("    * DORD F0022")
	If F0100_check = checked then call write_new_line_in_PRISM_case_note("    * DORD F0100")
	If F0109_checkbox = checked then call write_new_line_in_PRISM_case_note("    * DORD F0109")
	If F5000_checkbox = checked then call write_new_line_in_PRISM_case_note("    * DORD F5000")


	call write_new_line_in_PRISM_case_note("---")
	call write_new_line_in_PRISM_case_note("* CP to return by " & dateadd("d", date, 5) & ".")
	call write_new_line_in_PRISM_case_note("---")
	call write_new_line_in_PRISM_case_note(worker_signature)

	transmit
End if

If CAWD_check = checked then
	'Going to CAWD to write worklist
	call navigate_to_PRISM_screen("CAWD")
	EMWriteScreen "A", 8, 4
	transmit

	'Setting type as "free" and writing note	
	EMWriteScreen "FREE", 4, 37
	EMWriteScreen "*** Intake Docs due from CP", 10, 4
	EMWriteScreen dateadd("d", date, 7), 17, 21
	transmit
End if

script_end_procedure("")
