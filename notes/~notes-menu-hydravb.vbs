'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "~notes-menu-hydravb.vbs"
start_time = timer

'LOADING FUNCTIONS LIBRARY FROM GITHUB REPOSITORY===========================================================================
IF IsEmpty(FuncLib_URL) = TRUE THEN	'Shouldn't load FuncLib if it already loaded once
	IF run_locally = FALSE or run_locally = "" THEN	   'If the scripts are set to run locally, it skips this and uses an FSO below.
		IF use_master_branch = TRUE THEN			   'If the default_directory is C:\DHS-MAXIS-Scripts\Script Files, you're probably a scriptwriter and should use the master branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/master/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		Else											'Everyone else should use the release branch.
			FuncLib_URL = "https://raw.githubusercontent.com/MN-Script-Team/BZS-FuncLib/RELEASE/MASTER%20FUNCTIONS%20LIBRARY.vbs"
		End if
		SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a FuncLib_URL
		req.open "GET", FuncLib_URL, FALSE							'Attempts to open the FuncLib_URL
		req.send													'Sends request
		IF req.Status = 200 THEN									'200 means great success
			Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
			Execute req.responseText								'Executes the script code
		ELSE														'Error message
			critical_error_msgbox = MsgBox ("Something has gone wrong. The Functions Library code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
                                            "FuncLib URL: " & FuncLib_URL & vbNewLine & vbNewLine &_
                                            "The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
                                            vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
            StopScript
		END IF
	ELSE
		FuncLib_URL = "C:\BZS-FuncLib\MASTER FUNCTIONS LIBRARY.vbs"
		Set run_another_script_fso = CreateObject("Scripting.FileSystemObject")
		Set fso_command = run_another_script_fso.OpenTextFile(FuncLib_URL)
		text_from_the_other_script = fso_command.ReadAll
		fso_command.Close
		Execute text_from_the_other_script
	END IF
END IF
'END FUNCTIONS LIBRARY BLOCK================================================================================================

'CHANGELOG BLOCK ===========================================================================================================
'Starts by defining a changelog array
changelog = array()

'INSERT ACTUAL CHANGES HERE, WITH PARAMETERS DATE, DESCRIPTION, AND SCRIPTWRITER. **ENSURE THE MOST RECENT CHANGE GOES ON TOP!!**
'Example: call changelog_update("01/01/2000", "The script has been updated to fix a typo on the initial dialog.", "Jane Public, Oak County")
call changelog_update("01/24/2018", "Initial version.", "Veronica Cary, DHS")

'Actually displays the changelog. This function uses a text file located in the My Documents folder. It stores the name of the script file and a description of the most recent viewed change.
changelog_display
'END CHANGELOG BLOCK =======================================================================================================





' Predeclaring a number which will match what Hydra provides to ButtonPressed, does not actually connect with Hydra
button_incrementer = 1


BeginDialog menu_dialog, 0, 0, 506, 390, "Notes menu dialog"
  ButtonGroup ButtonPressed
    CancelButton 450, 370, 50, 15

    PushButton 5, 5, 120, 10, "Adjustments", btn_adjustments
    Text 130, 5, 370, 10, "Creates CAAD note for documenting adjustments made to the case."
    btn_adjustments = button_incrementer
    button_incrementer = button_incrementer + 1

	PushButton 5, 20, 120, 10, "Arrears Management Review", btn_arrears_management_review
	Text 130, 20, 370, 10, "Creates CAAD note for documenting an arrears management review."
	btn_arrears_management_review = button_incrementer
	button_incrementer = button_incrementer + 1

	PushButton 5, 35, 120, 10, "Case Initiation Docs Received", btn_case_initiation_docs_received
	Text 130, 35, 370, 10, "Creates CAAD note for recording receipt of intake/case initiation docs."
	btn_case_initiation_docs_received = button_incrementer
	button_incrementer = button_incrementer + 1

    PushButton 5, 50, 120, 10, "Client Contact", btn_client_contact
    Text 130, 50, 370, 10, "Creates a uniform CAAD note for when you have contact with or about client."
    btn_client_contact = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 65, 120, 10, "Contempt Hearing", btn_contempt_hearing
    Text 130, 65, 370, 10, "Creates a hearing date CAAD note for a contempt hearing."
    btn_contempt_hearing = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 80, 120, 10, "Court Order Request", btn_court_order_request
    Text 130, 80, 370, 10, "Creates B0170 CAAD note for requesting a court order, which also creates worklist to remind worker of order request."
    btn_court_order_request = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 95, 120, 10, "CSENET Info", btn_csenet_info
    Text 130, 95, 370, 10, "Creates T0111 CAAD note with text copied from INTD screen."
    btn_csenet_info = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 110, 120, 10, "E-Filing", btn_e_filing
    Text 130, 110, 370, 10, "Template for adding CAAD note about e-filing."
    btn_e_filing = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 125, 120, 10, "Fraud Referral", btn_fraud_referral
    Text 130, 125, 370, 10, "Template for adding CAAD note about a fraud referral."
    btn_fraud_referral = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 140, 120, 10, "Hearing Notes", btn_hearing_notes
    Text 130, 140, 370, 10, "CAAD note template for sending details about hearing notes."
    btn_hearing_notes = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 155, 120, 10, "Invoices", btn_invoices
    Text 130, 155, 370, 10, "Creates CAAD note for recording invoices."
    btn_invoices = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 170, 120, 10, "IW CAAD CAWT", btn_iw_caad_cawt
    Text 130, 170, 370, 10, "Creates CAAD and CAWT about IW."
    btn_iw_caad_cawt = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 185, 120, 10, "Maintaining County", btn_maintaining_county
    Text 130, 185, 370, 10, "Creates CAAD note for requesting maintaining county."
    btn_maintaining_county = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 200, 120, 10, "MES Financial Docs Sent", btn_mes_financial_docs_sent
    Text 130, 200, 370, 10, "Creates CAAD note for recording documents sent to parties."
    btn_mes_financial_docs_sent = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 215, 120, 10, "Mod CAAD Note - Contact Checklist", btn_mod_caad_note_contact_checklist
    Text 130, 215, 370, 10, "Creates CAAD note for recording contact with Client regarding possible Mod."
    btn_mod_caad_note_contact_checklist = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 230, 120, 10, "Pay or Report", btn_pay_or_report
    Text 130, 230, 370, 10, "CAAD note for contempt/''pay or report'' instances."
    btn_pay_or_report = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 245, 120, 10, "Quarterly Reviews", btn_quarterly_reviews
    Text 130, 245, 370, 10, "CAAD note for quarterly review processes."
    btn_quarterly_reviews = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 260, 120, 10, "Waiver of Personal Service", btn_waiver_of_personal_service
    Text 130, 260, 370, 10, "Creates CAAD note of the date a CP signed the waiver of personal service document."
    btn_waiver_of_personal_service = button_incrementer
    button_incrementer = button_incrementer + 1



    ' These scripts don't appear to have worked in Hydra
    if engine <> "cscript.exe" then

    end if


EndDialog







Dialog menu_dialog
IF ButtonPressed = 0 THEN script_end_procedure("")

if ButtonPressed = btn_adjustments then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/notes/adjustments.vbs"
elseif ButtonPressed = btn_arrears_management_review then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/notes/arrears-management-review.vbs"
elseif ButtonPressed = btn_case_initiation_docs_received then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/notes/case-initiation-docs-received.vbs"
elseif ButtonPressed = btn_client_contact then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/notes/client-contact.vbs"
elseif ButtonPressed = btn_contempt_hearing then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/notes/contempt-hearing.vbs"
elseif ButtonPressed = btn_court_order_request then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/notes/court-order-request.vbs"
elseif ButtonPressed = btn_csenet_info then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/notes/csenet-info.vbs"
elseif ButtonPressed = btn_e_filing then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/notes/e-filing.vbs"
elseif ButtonPressed = btn_fraud_referral then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/notes/fraud-referral.vbs"
elseif ButtonPressed = btn_hearing_notes then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/notes/hearing-notes.vbs"
elseif ButtonPressed = btn_invoices then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/notes/invoices.vbs"
elseif ButtonPressed = btn_iw_caad_cawt then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/notes/iw-caad-cawt.vbs"
elseif ButtonPressed = btn_maintaining_county then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/notes/maintaining-county.vbs"
elseif ButtonPressed = btn_mes_financial_docs_sent then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/notes/mes-financial-docs-sent.vbs"
elseif ButtonPressed = btn_mod_caad_note_contact_checklist then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/notes/mod-caad-note---contact-checklist.vbs"
elseif ButtonPressed = btn_pay_or_report then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/notes/pay-or-report.vbs"
elseif ButtonPressed = btn_quarterly_reviews then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/notes/quarterly-reviews.vbs"
elseif ButtonPressed = btn_waiver_of_personal_service then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/notes/waiver-of-personal-service.vbs"
end if


'Determining the script selected from the value of ButtonPressed
'Since we start at 100 and then go up, we will simply subtract 100 when determining the position in the array
call parse_and_execute_bzs(script_to_run)
