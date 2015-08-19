'GATHERING STATS---------------------------------------------------------------------------------------------------- 
name_of_script = "BULK - FAILURE POF - RSDI DFAS.vbs" 
start_time = timer 

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

' >>>>> THE SCRIPT <<<<<
EMConnect ""

'>>>>> GOING TO USWT <<<<<
Call navigate_to_Prism_screen("USWT")

' >>>>> SELECTING THE SPECIFIC WORKLIST TYPE <<<<<
EMWriteScreen "E0014", 20, 30
transmit

USWT_row = 7
' >>>>> STARTING THE DO LOOP. THE SCRIPT NEEDS TO HANDLE THESE CASES ONE AT A TIME <<<<<
DO
	EMReadScreen USWT_type, 5, USWT_row, 45
	IF USWT_type = "E0014" THEN
		EMReadScreen USWT_case_number, 13, USWT_row, 8
		EMWriteScreen "s", USWT_row, 4
		transmit

		' >>>>> MAKING SURE THAT THERE IS INFORMATION ON PAPL <<<<
		EMReadScreen end_of_data, 11, USWT_row, 32
		IF end_of_data <> "End of Data" THEN

			' >>>>> READING THE MOST RECENT PAY DATE AND CONVERTING IT TO A USABLE DATE <<<<<
			EMReadScreen PAPL_most_recent_pay_date, 6, 7, 7
			Call PRISM_date_conversion(PAPL_most_recent_pay_date)

			' >>>> CHECKING THAT THE DATE IN THE PAYMENT ID IS FROM THE CURRENT MONTH MINUS 1 <<<<<
			current_month_minus1 = DateAdd("m", -1, date)
			IF current_month_minus1 = DatePart ("m", PAPL_most_recent_pay_date) THEN
				' >>>>> IF THE PAYMENT IS FROM LAST MONTH, THE SCRIPT GRABS THE EMPLOYER/SOURCE ID <<<<<
				PF11
				EMReadScreen PAPL_name, 30, 7, 38
				' >>>>> LISTING OUT THE CONDITIONS THAT CAN BE PURGED AUTOMATICALLY <<<<<
				IF InStr(PAPL_name, "DFAS") <> 0 OR _
				   InStr(PAPL_name, "U S SOCIAL") <> 0 OR _ 
				   InStr(PAPL_name, "U S DEPT OF TREASURY") <> 0 THEN purge = True
			END IF
		END IF	

		' >>>>> GOING BACK TO USWT <<<<<
		Call navigate_to_PRISM_screen ("USWT")
		EMWriteScreen "E0014", 20, 30
		transmit

		' >>>>> FINDING THE CASE NUMBER THAT WE WERE WORKING ON <<<<<
		USWT_row = 7
		DO
			EMReadScreen case_number, 13, USWT_row, 8
			IF case_number <> USWT_case_number THEN 
				USWT_row = USWT_row + 1
				IF USWT_row = 19 THEN 
					PF8
					USWT_row = 7
				END IF
			END IF
		LOOP UNTIL USWT_case_number = case_number

		' >>>>> IF THE WORKLIST ITEM IS ELIGIBLE TO BE PURGED, THE SCRIPT PURGES...
		IF purge = True THEN 
			EMWriteScreen "P", USWT_row, 4
			transmit
			transmit
			PF3
		ELSE
		'  ...  IF THE WORKLIST ITME IS NOT ELIGIBLE TO BE PURGED, THE SCRIPT INCREASES USWT_ROW + 1 <<<<<
			USWT_row = USWT_row + 1
			IF USWT_row = 19 THEN 
				PF8
				USWT_row = 7
			END IF
		END IF
	END IF
LOOP UNTIL USWT_type <> "E0014"

script_end_procedure("")
