'GATHERING STATS----------------------------------------------------------------------------------------------------
name_of_script = "ACTIONS - MAIN MENU.vbs"
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
BeginDialog ACTIONS_main_menu_dialog, 0, 0, 306, 145, "ACTIONS main menu dialog"
  ButtonGroup ButtonPressed
    PushButton 5, 10, 70, 10, "Find Name on CALI", ACTIONS_find_name_on_cali_button
    PushButton 5, 30, 30, 10, "Intake", ACTIONS_intake_button
    PushButton 5, 55, 60, 10, "PALC calculator", ACTIONS_PALC_calculator_button
    PushButton 5, 75, 60, 10, "Prorate Support", ACTIONS_prorate_support_button
    CancelButton 250, 125, 50, 15
  Text 80, 10, 215, 10, "-- NEW 06/2015!!! Searches CALI for a specific CP or NCP."
  Text 70, 55, 230, 10, "-- Calculates voluntary and involuntary amounts from the PALC screen."
  Text 40, 30, 260, 15, "-- Creates various documents related to Child Support intake, as well as DORD documents, and enters a note on CAAD."
  Text 70, 75, 225, 10, "- Calculator for deteremining pro-rated support for partial months."
EndDialog


'THE SCRIPT-----------------------------------------------------------------------------------------------

'Shows the dialog
Dialog ACTIONS_main_menu_dialog
If buttonpressed = cancel then stopscript
IF BUttonPressed = ACTIONS_find_name_on_cali_button THEN CALL run_from_GitHub(script_repository & "ACTIONS/ACTIONS - FIND NAME ON CALI.vbs")
IF ButtonPressed = ACTIONS_prorate_support_button THEN call run_from_GitHub(script_repository & "ACTIONS/ACTIONS - PRORATE SUPPORT.vbs")
IF ButtonPressed = ACTIONS_intake_button then call run_from_GitHub(script_repository & "ACTIONS/ACTIONS - INTAKE.vbs")
IF ButtonPressed = ACTIONS_PALC_calculator_button then call run_from_GitHub(script_repository & "ACTIONS/ACTIONS - PALC CALCULATOR.vbs")
