'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "~actions-menu-hydravb.vbs"
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


BeginDialog actions_menu_dialog, 0, 0, 506, 390, "Actions menu dialog"
  ButtonGroup ButtonPressed
    CancelButton 450, 370, 50, 15

    PushButton 5, 5, 120, 10, "Admin Redirect", btn_admin_redirect
    Text 130, 5, 370, 10, "Creates redirection docs and redirection worklist items."
    btn_admin_redirect = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 50, 120, 10, "COLA", btn_cola
    Text 130, 50, 370, 10, "Leads you through performing a COLA. Adds CAAD note when completed."
    btn_cola = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 65, 120, 10, "Emancipation DORD docs", btn_emancipation_dord_docs
    Text 130, 65, 370, 10, "Sends emancipation DORD docs."
    btn_emancipation_dord_docs = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 80, 120, 10, "Employment Verification", btn_employment_verification
    Text 130, 80, 370, 10, "Complete an Employment Verification in NCID or CPID, includes info on CAAD note."
    btn_employment_verification = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 95, 120, 10, "Enforcement Intake", btn_enforcement_intake
    Text 130, 95, 370, 10, "Intake workflow on enforcement cases."
    btn_enforcement_intake = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 110, 120, 10, "Establishment DORD docs - NPA", btn_establishment_dord_docs_npa
    Text 130, 110, 370, 10, "Generates establishment DORD docs for NPA case."
    btn_establishment_dord_docs_npa = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 125, 120, 10, "Establishment DORD docs - PA", btn_establishment_dord_docs_pa
    Text 130, 125, 370, 10, "Generates establishment DORD docs for PA case."
    btn_establishment_dord_docs_pa = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 140, 120, 10, "Fee Suppression Override", btn_fee_suppression_override
    Text 130, 140, 370, 10, "Overrides a fee suppression."
    btn_fee_suppression_override = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 155, 120, 10, "Financial Statement Follow-up", btn_financial_statement_follow_up
    Text 130, 155, 370, 10, "Sends follow-up memo to parties regarding financial statements."
    btn_financial_statement_follow_up = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 170, 120, 10, "Find Name on CALI", btn_find_name_on_cali
    Text 130, 170, 370, 10, "Searches CALI for a specific CP or NCP."
    btn_find_name_on_cali = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 185, 120, 10, "Generic Enforcement Intake", btn_generic_enforcement_intake
    Text 130, 185, 370, 10, "Creates various docs related to CS intake as well as DORD docs and enters CAAD."
    btn_generic_enforcement_intake = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 200, 120, 10, "Income Verification", btn_income_verification
    Text 130, 200, 370, 10, "Generates Word document regarding payments CP has received on their case."
    btn_income_verification = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 215, 120, 10, "Interview Information Sheet", btn_interview_information_sheet
    Text 130, 215, 370, 10, "Creates a Word document with general and case-specific information to be used as a reference when meeting with clients."
    btn_interview_information_sheet = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 230, 120, 10, "NCP Locate", btn_ncp_locate
    Text 130, 230, 370, 10, "Walks you through processing an NCP locate."
    btn_ncp_locate = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 245, 120, 10, "Non Pay", btn_non_pay
    Text 130, 245, 370, 10, "Sends DORD doc and creates CAAD related to Non-Pay."
    btn_non_pay = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 260, 120, 10, "Record IW Info", btn_record_IW_info
    Text 130, 260, 370, 10, "Record IW withholding info in a CAAD note, worklist, or view in a message box."
    btn_record_IW_info = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 275, 120, 10, "Refer to Mod", btn_refer_to_mod
    Text 130, 275, 370, 10, "Starts REAM and sends docs to include employer verifs."
    btn_refer_to_mod = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 305, 120, 10, "Sanction", btn_sanction
    Text 130, 305, 370, 10, "Takes actions on the case to apply or remove public assistance sanction for non-cooperation with child support."
    btn_sanction = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 350, 120, 10, "Unreimbursed Uninsured Sending Docs", btn_unreimbursed_uninsured_sending_docs
    Text 130, 350, 370, 10, "Prints DORD docs for collecting unreimbursed and uninsured expenses."
    btn_unreimbursed_uninsured_sending_docs = button_incrementer
    button_incrementer = button_incrementer + 1

    ' These scripts don't appear to have worked in Hydra
    if engine <> "cscript.exe" then
        PushButton 5, 20, 120, 10, "Affidavit of Service by Mail Docs", btn_affadavit_of_service_by_mail_docs
        Text 130, 20, 370, 10, "Sends Affidavits of Service to multiple participants on the case."
        btn_affadavit_of_service_by_mail_docs = button_incrementer
        button_incrementer = button_incrementer + 1

        PushButton 5, 35, 120, 10, "Case Transfer", btn_case_transfer
        Text 130, 35, 370, 10, "Transfers single case and creates CAAD about why."
        btn_case_transfer = button_incrementer
        button_incrementer = button_incrementer + 1

        PushButton 5, 290, 120, 10, "Returned Mail", btn_returned_mail
        Text 130, 290, 370, 10, "Updates address to new or unknown, and creates CAAD note."
        btn_returned_mail = button_incrementer
        button_incrementer = button_incrementer + 1

        PushButton 5, 320, 120, 10, "Send F0104 DORD memo", btn_send_f0104_dord_memo
        Text 130, 320, 370, 10, "Sends F0104 DORD Memo Docs, with options to send a memo to both parties and preview memo text."
        btn_send_f0104_dord_memo = button_incrementer
        button_incrementer = button_incrementer + 1

        PushButton 5, 335, 120, 10, "Unreimbursed Uninsured Returned Docs", btn_unreimbursed_uninsured_returned_docs
        Text 130, 335, 370, 10, "Sends DORD docs when unreimbursed and uninsured docs are returned."
        btn_unreimbursed_uninsured_returned_docs = button_incrementer
        button_incrementer = button_incrementer + 1
    end if


EndDialog







Dialog actions_menu_dialog
IF ButtonPressed = 0 THEN script_end_procedure("")

if ButtonPressed = btn_admin_redirect then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/actions/admin-redirect.vbs"
elseif ButtonPressed = btn_affadavit_of_service_by_mail_docs then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/actions/affidavit-of-service-by-mail-docs.vbs"
elseif ButtonPressed = btn_case_transfer then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/actions/case-transfer.vbs"
elseif ButtonPressed = btn_cola then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/actions/cola.vbs"
elseif ButtonPressed = btn_emancipation_dord_docs then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/actions/emancipation-dord-docs.vbs"
elseif ButtonPressed = btn_employment_verification then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/actions/employment-verification.vbs"
elseif ButtonPressed = btn_enforcement_intake then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/actions/enforcement-intake.vbs"
elseif ButtonPressed = btn_establishment_dord_docs_npa then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/actions/establishment-dord-docs---npa.vbs"
elseif ButtonPressed = btn_establishment_dord_docs_pa then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/actions/establishment-dord-docs---pa.vbs"
elseif ButtonPressed = btn_fee_suppression_override then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/actions/fee-suppression-override.vbs"
elseif ButtonPressed = btn_financial_statement_follow_up then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/actions/financial-statement-follow-up.vbs"
elseif ButtonPressed = btn_find_name_on_cali then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/actions/find-name-on-cali.vbs"
elseif ButtonPressed = btn_generic_enforcement_intake then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/actions/generic-enforcement-intake.vbs"
elseif ButtonPressed = btn_income_verification then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/actions/income-verification.vbs"
elseif ButtonPressed = btn_interview_information_sheet then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/actions/interview-information-sheet.vbs"
elseif ButtonPressed = btn_ncp_locate then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/actions/ncp-locate.vbs"
elseif ButtonPressed = btn_non_pay then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/actions/non-pay.vbs"
elseif ButtonPressed = btn_record_IW_info then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/actions/record-iw-info.vbs"
elseif ButtonPressed = btn_refer_to_mod then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/actions/refer-to-mod.vbs"
elseif ButtonPressed = btn_returned_mail then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/actions/returned-mail.vbs"
elseif ButtonPressed = btn_sanction then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/actions/sanction.vbs"
elseif ButtonPressed = btn_send_f0104_dord_memo then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/actions/send-f0104-dord-memo.vbs"
elseif ButtonPressed = btn_unreimbursed_uninsured_returned_docs then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/actions/unreimbursed-uninsured-returned-docs.vbs"
elseif ButtonPressed = btn_unreimbursed_uninsured_sending_docs then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/actions/unreimbursed-uninsured-sending-docs.vbs"
end if


'Determining the script selected from the value of ButtonPressed
'Since we start at 100 and then go up, we will simply subtract 100 when determining the position in the array
if engine = "cscript.exe" then
	call parse_and_execute_bzs(script_to_run)
else
	CALL run_from_GitHub(script_to_run)
end if
