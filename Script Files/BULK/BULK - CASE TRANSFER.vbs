'Gathering stats-------------------------------------------------------------------------------------
name_of_script = "BULK - CASE TRANSFER.vbs"
start_time = timer

'LOADING ROUTINE FUNCTIONS (FOR PRISM)---------------------------------------------------------------
Dim URL, REQ, FSO					'Declares variables to be good to option explicit users
If beta_agency = "" then 			'For scriptwriters only
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/master/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
ElseIf beta_agency = True then		'For beta agencies and testers
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/beta/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
Else								'For most users
	url = "https://raw.githubusercontent.com/MN-CS-Script-Team/PRISM-Scripts/release/Shared%20Functions%20Library/PRISM%20Functions%20Library.vbs"
End if
Set req = CreateObject("Msxml2.XMLHttp.6.0")				'Creates an object to get a URL
req.open "GET", url, False									'Attempts to open the URL
req.send													'Sends request
If req.Status = 200 Then									'200 means great success
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

BeginDialog worker_numbers_dlg, 0, 0, 231, 190, "Enter Worker Numbers"
  Text 10, 10, 210, 10, "Please enter a list of CSO Worker Numbers to transfer cases to. "
  Text 10, 30, 210, 30, "NOTE: You can enter either the 8-digit Worker ID or the 11-digit code (County, Office Team, Position). The script can decipher between the different numbers."
  Text 10, 70, 210, 20, "The script will give you a list of workers that are not found in PRISM."
  Text 10, 100, 125, 10, "Separate each worker with a comma."
  EditBox 10, 120, 210, 15, worker_list
  CheckBox 10, 145, 210, 10, "Check HERE to transfer a caseload from one user to another.", transfer_all_cases_check
  ButtonGroup ButtonPressed
    OkButton 125, 170, 50, 15
    CancelButton 175, 170, 50, 15
EndDialog

BeginDialog transfer_all_cases_dlg, 0, 0, 191, 130, "Enter Worker Numbers"
  EditBox 90, 55, 75, 15, transfer_from
  EditBox 90, 75, 75, 15, transfer_to
  ButtonGroup ButtonPressed
    OkButton 80, 110, 50, 15
    CancelButton 130, 110, 50, 15
  Text 10, 15, 175, 25, "You can enter either the 8-digit worker number or the 11-digit position number. The script will sort it out."
  Text 10, 60, 75, 10, "Transfer Cases From:"
  Text 10, 80, 65, 10, "Transfer Cases To:"
EndDialog


'===== CUSTOM FUNCTION FOR DIALOGS FOR EACH WORKER
FUNCTION create_case_numbers_dlg(i, worker_array)

	BeginDialog case_numbers_dlg, 0, 0, 276, 195, "Enter Case Numbers"
	Text 10, 15, 60, 10, "Worker Number"
	Text 10, 35, 60, 10, "Worker Name"
	Text 75, 15, 85, 10, worker_array(i, 0)
	Text 75, 35, 85, 10, worker_array(i, 1)
	EditBox 10, 55, 80, 15, worker_array(i, 2)
	EditBox 10, 75, 80, 15, worker_array(i, 3)
	EditBox 10, 95, 80, 15, worker_array(i, 4)
	EditBox 10, 115, 80, 15, worker_array(i, 5)
	EditBox 10, 135, 80, 15, worker_array(i, 6)
	EditBox 100, 55, 80, 15, worker_array(i, 7)
	EditBox 100, 75, 80, 15, worker_array(i, 8)
	EditBox 100, 95, 80, 15, worker_array(i, 9)
	EditBox 100, 115, 80, 15, worker_array(i, 10)
	EditBox 100, 135, 80, 15, worker_array(i, 11)
	EditBox 190, 55, 80, 15, worker_array(i, 12)
	EditBox 190, 75, 80, 15, worker_array(i, 13)
	EditBox 190, 95, 80, 15, worker_array(i, 14)
	EditBox 190, 115, 80, 15, worker_array(i, 15)
	EditBox 190, 135, 80, 15, worker_array(i, 16)
	ButtonGroup ButtonPressed
		OkButton 170, 175, 50, 15
		PushButton 220, 175, 50, 15, "STOP SCRIPT", stop_script_button
	EndDialog
	
	DIALOG case_numbers_dlg
		IF ButtonPressed = stop_script_button THEN stopscript
END FUNCTION 


'===== THE SCRIPT =====
EMConnect ""
CALL check_for_PRISM(False)

DIALOG worker_numbers_dlg
	IF ButtonPressed = stop_script_button THEN stopscript
	IF InStr(worker_list, "UUDDLRLRBA") <> 0 THEN 
		developer_mode = True
		MsgBox "Developer mode enabled."
	END IF

IF transfer_all_cases_check = 0 THEN 
	worker_list = replace(worker_list, " ", "")
	worker_list = split(worker_list, ",")
	
	number_of_workers = UBound(worker_list)
	ReDim worker_array(number_of_workers, 16)
	
	i = 0
	FOR EACH cso_worker IN worker_list
		IF cso_worker <> "" THEN 
			worker_array(i, 0) = cso_worker
			i = i + 1
		END IF
	NEXT
	
	FOR i = 0 TO number_of_workers
		IF len(worker_array(i, 0)) = 8 THEN 
			'If the length of the worker number is 8 then the script goes to CALI to gather the 11-digit worker position number.
			CALL navigate_to_PRISM_screen("CALI")
			EMWriteScreen left(worker_array(i, 0), 3), 20, 18
			EMWriteScreen "001", 20, 30
			transmit
			
			EMSetCursor 20, 49
			EMSendKey "X"
	
			PF1
			
			CALI_row = 13
			DO
				EMReadScreen worker_id, 8, CALI_row, 39
				EMReadScreen end_of_data, 11, CALI_row, 39
				IF end_of_data = "End of Data" THEN 
					worker_array(i, 1) = "WORKER NOT FOUND"
					EXIT DO
				END IF
				IF UCASE(worker_id) = UCASE(worker_array(i, 0)) THEN 
					EMReadScreen worker_array(i, 1), 30, CALI_row, 49
					worker_array(i, 1) = trim(worker_array(i, 1))
					EMReadScreen CALI_position, 20, CALI_row, 18
					CALI_position = replace(CALI_position, " ", "")
					worker_array(i, 0) = CALI_position
					PF3
					CALL create_case_numbers_dlg(i, worker_array)
					EXIT DO
				ELSE
					CALI_row = CALI_row + 1
					IF CALI_row = 19 THEN 
						PF8
						CALI_row = 13
					END IF
				END IF
			LOOP		
		ELSEIF len(worker_array(i, 0)) = 11 THEN
			CALL navigate_to_PRISM_screen("CALI")
			EMSetCursor 20, 18
			EMSendKey worker_array(i, 0)
			transmit
			
			EMReadScreen error_message_on_bottom_of_screen, 20, 24, 2
			error_message_on_bottom_of_screen = trim(error_message_on_bottom_of_screen)
			IF error_message_on_bottom_of_screen = "" THEN 
				CALL find_variable("Name: ", worker_array(i, 1), 30)
				CALL create_case_numbers_dlg(i, worker_array)
			ELSEIF error_message_on_bottom_of_screen <> "" THEN 
				worker_array(i, 1) = "WORKER NOT FOUND"
			END IF
		ELSE
			worker_array(i, 1) = "WORKER NOT FOUND"
		END IF
	NEXT		
	
	'Navigating to CAAS to let the case transferring begin!!
	CALL navigate_to_PRISM_screen("CAAS")
	
	err_workers = ""
	FOR i = 0 TO number_of_workers
		msgbox worker_array(i, 0) & vbCr & worker_array(i, 1)
		IF worker_array(i, 1) = "WORKER NOT FOUND" THEN 
			err_workers = err_workers & vbCr & "     " & worker_array(i, 0) 
		ELSEIF worker_array(i, 1) <> "WORKER NOT FOUND" THEN 
			FOR j = 2 TO 16
				IF worker_array(i, j) <> "" THEN 
					CAAS_county = left(worker_array(i, 0), 3)
					CAAS_office = right(left(worker_array(i, 0), 6), 3)
					CAAS_team = left(right(worker_array(i, 0), 5), 3)
					CAAS_position = right(worker_array(i, 0), 2)
					
					EMWriteScreen "D", 3, 29
					EMwriteScreen left(worker_array(i, j), 10), 4, 8
					EMWriteScreen right(worker_array(i, j), 2), 4, 19
					transmit
					
					EMWriteScreen "M", 3, 29
					EMWriteScreen CAAS_county, 9, 20
					EMWriteScreen CAAS_office, 10, 20
					EMWriteScreen CAAS_team, 11, 20
					EMWriteScreen CAAS_position, 12, 20
					
					IF developer_mode = True THEN 
						MsgBox "*** Developer Mode Enabled ***" & vbCr & vbCr & _
							"Transferring Case " & worker_array(i, j) & " to " & worker_array(i, 1)				
					ELSE
						DO
							transmit
							EMReadScreen confirmation_message, 70, 24, 2
						LOOP UNTIL InStr(confirmation_message, "modified successfully") <> 0
					END IF
				END IF
			NEXT
		END IF
	NEXT
	
	'Displaying the list of workers that were skipped because they could not be found.
	IF err_workers <> "" THEN MsgBox ("*** NOTICE!!! ***" & vbCr & vbCr & "The script could not transfer cases to the following worker ID/code(s): " & vbCr & err_workers & vbCr & vbCr & "The script has determined that ID/code is not a valid ID/code assigned to a worker. You may need to reconsider the worker ID/code you selected and try again." & vbCr & vbCr & "If the script erred in its determination of valid worker ID/codes, please report this to your scripts administrator." & vbCr & vbCr & "Thank you.")

ELSEIF transfer_all_cases_check = 1 THEN 
	DO
		DO
			DO
				err_msg = ""
				DIALOG transfer_all_cases_dlg
					IF ButtonPressed = 0 THEN stopscript
					transfer_from = trim(transfer_from)
					transfer_to = trim(transfer_to)
					IF transfer_from = "" THEN err_msg = err_msg & vbCr & "* Please enter a valid worker/position number to transfer cases FROM."
					IF transfer_to = "" THEN err_msg = err_msg & vbCr & "* Please enter a valid worker/position number to transfer cases TO."
					IF err_msg <> "" THEN MsgBox "*** NOTICE!!! ***" & vbCr & err_msg & vbCr & vbCr & "Please resolve for the script to continue."
			LOOP UNTIL err_msg = "" 
			
			'Getting the worker names for the confirmation message
			IF len(transfer_from) = 8 THEN 
				'If the length of the worker number is 8 then the script goes to CALI to gather the 11-digit worker position number.
				CALL navigate_to_PRISM_screen("CALI")
				EMWriteScreen left(transfer_from, 3), 20, 18
				EMWriteScreen "001", 20, 30
				transmit
				
				EMSetCursor 20, 49
				EMSendKey "X"
		
				PF1
				
				CALI_row = 13
				DO
					EMReadScreen worker_id, 8, CALI_row, 39
					EMReadScreen end_of_data, 11, CALI_row, 39
					IF end_of_data = "End of Data" THEN 
						transfer_from_name = "WORKER NOT FOUND"
						EXIT DO
					END IF
					IF UCASE(worker_id) = UCASE(transfer_from) THEN 
						EMReadScreen transfer_from_name, 30, CALI_row, 49
						transfer_from_name = trim(transfer_from_name)
						EMReadScreen CALI_position, 20, CALI_row, 18
						CALI_position = replace(CALI_position, " ", "")
						transfer_from_position_number = CALI_position
						PF3
						EXIT DO
					ELSE
						CALI_row = CALI_row + 1
						IF CALI_row = 19 THEN 
							PF8
							CALI_row = 13
						END IF
					END IF
				LOOP		
			ELSEIF len(transfer_from) = 11 THEN
				CALL navigate_to_PRISM_screen("CALI")
				EMSetCursor 20, 18
				EMSendKey transfer_from
				transmit
				
				EMReadScreen error_message_on_bottom_of_screen, 20, 24, 2
				error_message_on_bottom_of_screen = trim(error_message_on_bottom_of_screen)
				IF error_message_on_bottom_of_screen = "" THEN 
					CALL find_variable("Name: ", transfer_from_name, 30)
					transfer_from_name = trim(transfer_from_name)
					transfer_from_position_number = transfer_from
				ELSEIF error_message_on_bottom_of_screen <> "" THEN 
					transfer_from_name = "WORKER NOT FOUND"
				END IF
			ELSE
				transfer_from_name = "WORKER NOT FOUND"
			END IF
			
			'Getting the worker names for the confirmation message
			IF len(transfer_to) = 8 THEN 
				'If the length of the worker number is 8 then the script goes to CALI to gather the 11-digit worker position number.
				CALL navigate_to_PRISM_screen("CALI")
				EMWriteScreen left(transfer_to, 3), 20, 18
				EMWriteScreen "001", 20, 30
				transmit
				
				EMSetCursor 20, 49
				EMSendKey "X"
		
				PF1
				
				CALI_row = 13
				DO
					EMReadScreen worker_id, 8, CALI_row, 39
					EMReadScreen end_of_data, 11, CALI_row, 39
					IF end_of_data = "End of Data" THEN 
						transfer_from_name = "WORKER NOT FOUND"
						EXIT DO
					END IF
					IF UCASE(worker_id) = UCASE(transfer_to) THEN 
						EMReadScreen transfer_to_name, 30, CALI_row, 49
						transfer_to_name = trim(transfer_to_name)
						EMReadScreen CALI_position, 20, CALI_row, 18
						CALI_position = replace(CALI_position, " ", "")
						transfer_to_position_number = CALI_position
						PF3
						EXIT DO
					ELSE
						CALI_row = CALI_row + 1
						IF CALI_row = 19 THEN 
							PF8
							CALI_row = 13
						END IF
					END IF
				LOOP		
			ELSEIF len(transfer_to) = 11 THEN
				CALL navigate_to_PRISM_screen("CALI")
				EMSetCursor 20, 18
				EMSendKey transfer_to
				transmit
				
				EMReadScreen error_message_on_bottom_of_screen, 20, 24, 2
				error_message_on_bottom_of_screen = trim(error_message_on_bottom_of_screen)
				IF error_message_on_bottom_of_screen = "" THEN 
					CALL find_variable("Name: ", transfer_to_name, 30)
					transfer_to_name = trim(transfer_to_name)
					transfer_to_position_number = transfer_to
				ELSEIF error_message_on_bottom_of_screen <> "" THEN 
					transfer_to_name = "WORKER NOT FOUND"
				END IF
			ELSE
				transfer_to_name = "WORKER NOT FOUND"
			END IF
			
			IF transfer_to_name = "WORKER NOT FOUND" THEN MsgBox "Worker " & transfer_to & " not found. Please try again."
			IF transfer_from_name = "WORKER NOT FOUND" THEN MsgBox "Worker " & transfer_from & " not found. Please try again."
			
		LOOP UNTIL transfer_from_name <> "WORKER NOT FOUND" AND transfer_to_name <> "WORKER NOT FOUND"
		
		confirmation_message = MsgBox("*** PLEASE CONFIRM ***" & vbCr & "Transfer FROM Caseload: " & transfer_from_name & " (" & transfer_from_position_number & ")" & vbCr & "Transfer TO Caseload: " & transfer_to_name & " (" & transfer_to_position_number & ")." & vbCr & vbCr & "Is this correct? Press YES to continue, press NO to retry.", vbYesNo)		

	LOOP UNTIL confirmation_message = vbYes

	'>>>> Now going to TRANSFER FROM to gather all cases.
	CALL navigate_to_PRISM_screen("CALI")
	EMSetCursor 20, 18
	EMSendKey transfer_from_position_number
	transmit
	
	IF developer_mode = True THEN 
		Set objExcel = CreateObject("Excel.Application")
		objExcel.Visible = True
		Set objWorkbook = objExcel.Workbooks.Add()
		objExcel.DisplayAlerts = True
		objExcel.Cells(1, 1).Value = "CASES TO TRANSFER"
		objExcel.Cells(1, 1).Font.Bold = True
		objExcel.Columns(1).AutoFit()
		objExcel.Cells(1, 2).Value = "CASES TRANSFERRED"
		objExcel.Cells(1, 2).Font.Bold = True
		objExcel.Columns(2).AutoFit()
		excel_row = 2
	END IF
	
	CALI_row = 8
	DO
		EMReadScreen end_of_data, 11, CALI_row, 32
		EMReadScreen case_number, 14, CALI_row, 7
		case_number = replace(case_number, " ", "")
		IF end_of_data <> "End of Data" THEN 
			all_cases_array = all_cases_array & case_number & "~~~"
			IF developer_mode = True THEN 
				objExcel.Cells(excel_row, 1).Value = case_number
				excel_row = excel_row + 1
			END IF
		ELSEIF end_of_data = "End of Data" THEN 
			EXIT DO
		END IF
		CALI_row = CALI_row + 1
		IF CALI_row = 19 THEN 
			CALI_row = 8
			PF8
		END IF		
	LOOP UNTIL end_of_data = "End of Data"
	
	all_cases_array = trim(all_cases_array)
	all_cases_array = split(all_cases_array, "~~~")
	
	IF developer_mode = True THEN excel_row = 2
	
	CALL navigate_to_PRISM_screen("CAAS")
	FOR EACH PRISM_case_number IN all_cases_array
		IF PRISM_case_number <> "" THEN 
			EMSetCursor 3, 29
			EMSendKey "D"
			EMSendKey PRISM_case_number
			transmit
			
			EMWriteScreen "M", 3, 29
			EMSetCursor 9, 20
			EMSendKey transfer_to_position_number
			IF developer_mode = False THEN 
				DO
					transmit
					EMReadScreen confirmation_message, 70, 24, 2
				LOOP UNTIL InStr(confirmation_message, "modified successfully") <> 0
			ELSEIF developer_mode = True THEN 
				objExcel.Cells(excel_row, 2).Value = PRISM_case_number
				excel_row = excel_row + 1
			END IF
		END IF
	NEXT
	
END IF	

script_end_procedure("Success!!")
