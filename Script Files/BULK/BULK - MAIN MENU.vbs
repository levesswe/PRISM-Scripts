'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "BULK - MAIN MENU.vbs"
start_time = timer

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

'DIALOGS---------------------------------------------------------------------------
BeginDialog BULK_main_menu_dialog, 0, 0, 351, 180, "BULK main menu dialog"
  ButtonGroup ButtonPressed
    PushButton 10, 15, 60, 10, "CALI to Excel", BULK_cali_to_excel_button
    PushButton 10, 40, 60, 10, "Case Transfer", BULK_case_transfer_button
    PushButton 10, 75, 100, 10, "Companion Case Finder - CP", BULK_cp_companion_case_finder_button
    PushButton 10, 100, 105, 10, "Companion Case Finder - NCP", BULK_ncp_companion_case_finder_button
    PushButton 10, 125, 60, 10, "Evaluate NOCS", BULK_evaluate_nocs_button
    CancelButton 295, 160, 50, 15
  Text 80, 15, 265, 20, "This script builds a list in Microsoft Excel of case numbers, function types, program codes, interstate codes, and parent names."
  Text 80, 40, 265, 25, "-- NEW!!! 08/2015  This script allows users to transfer up to 15 cases to as many workers as they need OR to transfer an entire caseload to as many workers as needed."
  Text 115, 75, 225, 20, "--- NEW!!! 08/2015 -- This script builds a list of companion cases for your CPs on a given CALI."
  Text 120, 100, 220, 20, "--- NEW!!! 08/2015 -- This script builds a list of companion cases for your NCPs on a given CALI."
  Text 80, 125, 260, 10, "This script evaluates D0800 worklist items for continued services."
EndDialog

'THE SCRIPT-----------------------------------------------------------------------------------------------
'Shows the dialog
Dialog BULK_main_menu_dialog
If buttonpressed = cancel then StopScript
IF ButtonPressed = BULK_cali_to_excel_button 					THEN CALL run_from_GitHub(script_repository & "BULK/BULK - CALI TO EXCEL.vbs")
IF ButtonPressed = BULK_case_transfer_button 					THEN CALL run_from_GitHub(script_repository & "BULK/BULK - CASE TRANSFER.vbs")
IF ButtonPressed = BULK_cp_companion_case_finder_button 			THEN CALL run_from_GitHub(script_repository & "BULK/BULK - CP COMPANION CASE FINDER.vbs")
IF ButtonPressed = BULK_ncp_companion_case_finder_button			THEN CALL run_from_GitHub(script_repository & "BULK/BULK - NCP COMPANION CASE FINDER.vbs")
IF ButtonPressed = BULK_evaluate_nocs_button 					THEN CALL run_from_GitHub(script_repository & "BULK/BULK - EVALUATE NOCS.vbs")
