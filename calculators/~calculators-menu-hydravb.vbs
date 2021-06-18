'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "~calculators-menu-hydravb.vbs"
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

' ERROR HANDLING
 on error resume next				' So this way I can catch errors
 ButtonGroup ""						' Sending a test ButtonGroup "" function... if it errors then we need to load functions
 if err.number = 13 then LoadFuncs 	' Declared below
 on error goto 0 					' Further errors should behave as expected

' Predeclaring a number which will match what Hydra provides to ButtonPressed, does not actually connect with Hydra
button_incrementer = 1


BeginDialog menu_dialog, 0, 0, 506, 100, "Calculators menu dialog"

  ButtonGroup ButtonPressed
    CancelButton 450, 75, 50, 15

    PushButton 5, 5, 120, 10, "DDPL", btn_ddpl
    Text 130, 5, 370, 10, "Calculates direct deposits made over user-provided date range."
    btn_ddpl = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 20, 120, 10, "IW", btn_iw
    Text 130, 20, 370, 10, "Calculator for determining the amount of IW over a given period."
    btn_iw = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 35, 120, 10, "PALC", btn_palc
    Text 130, 35, 370, 10, "Calculates voluntary and involuntary amounts from the PALC screen."
    btn_palc = button_incrementer
    button_incrementer = button_incrementer + 1

    PushButton 5, 50, 120, 10, "Prorate Support", btn_prorate_support
    Text 130, 50, 370, 10, "Calculator for determining pro-rated support for patrial months."
    btn_prorate_support = button_incrementer
    button_incrementer = button_incrementer + 1

    ' These scripts don't appear to have worked in Hydra (commented out one for use as a sample)
    if engine <> "cscript.exe" then
'        PushButton 5, 20, 120, 10, "Affidavit of Service by Mail Docs", btn_affadavit_of_service_by_mail_docs
'        Text 130, 20, 370, 10, "Sends Affidavits of Service to multiple participants on the case."
'        btn_affadavit_of_service_by_mail_docs = button_incrementer
'        button_incrementer = button_incrementer + 1
    end if


EndDialog







Dialog menu_dialog
IF ButtonPressed = 0 THEN script_end_procedure("")

if ButtonPressed = btn_ddpl then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/calculators/ddpl.vbs"
elseif ButtonPressed = btn_iw then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/calculators/iw.vbs"
elseif ButtonPressed = btn_palc then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/calculators/palc.vbs"
elseif ButtonPressed = btn_prorate_support then
    script_to_run = "https://raw.githubusercontent.com/MN-Script-Team/DHS-PRISM-Scripts/master/calculators/prorate-support.vbs"
end if


'Determining the script selected from the value of ButtonPressed
'Since we start at 100 and then go up, we will simply subtract 100 when determining the position in the array
call parse_and_execute_bzs(script_to_run)

function LoadFuncs
	script_URL = "https://raw.githubusercontent.com/MN-Script-Team/hydra/master/vbs-libs/bzio-helper-functions.vbs"
	SET req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
	req.open "GET", script_URL, FALSE									'Attempts to open the URL
	req.send													'Sends request
	IF req.Status = 200 THEN									'200 means great success
		Set fso = CreateObject("Scripting.FileSystemObject")	'Creates an FSO
		ExecuteGlobal req.responseText								'Executes the script code
	ELSE														'Error message, tells user to try to reach github.com, otherwise instructs to contact Veronica with details (and stops script).
		critical_error_msgbox = MsgBox ("Something has gone wrong. The code stored on GitHub was not able to be reached." & vbNewLine & vbNewLine &_
										"Script URL: " & script_URL & vbNewLine & vbNewLine &_
										"The script has stopped. Please check your Internet connection. Consult a scripts administrator with any questions.", _
										vbOKonly + vbCritical, "BlueZone Scripts Critical Error")
		StopScript
	END IF
end function
