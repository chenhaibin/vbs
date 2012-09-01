
'On Error Resume Next

' ** Server to log to **
strLocal = "NTW324100"
boolLogtoSQL = TRUE
'strLog = "SYSTEM"
strLog = "APPLICATION"

' ** If the number of arguments isn't three then error out
if wscript.arguments.count <> 3 then
	wscript.echo " * Error: Incorrect No of Arguments"

	wscript.echo " * USAGE: cscript afeventlog.vbs [start time] [finish time] [Server]" & vbcrlf
	wscript.quit
else

	' ** Assign arguments to strings	
	strStartArg = wscript.arguments (0)
	strFinishArg = wscript.arguments (1)
	strServerArg = wscript.arguments (2)
	
	'wscript.echo isdate (wscript.arguments(0))
	
	' ** Check 1st two arguments are time's in the correct format
	subCheckTime strStartArg
	subCheckTime strFinishArg
	
	' ** Check to make sure first time is less than 2nd time
	if strFinishArg < strStartArg then
		wscript.echo " * Error: Finish time is before start time"
		wscript.quit
	end if
	
	' ** Call main subroutine which does most of the work
	sublogreader strServerArg, strStartArg, strFinishArg
	
end if


'wscript.echo " * End Script"
' ** finish the script
wscript.quit

' Subroutine to check times are valid
Sub SubCheckTime (strTime)
			
	' if string does not contain a : (eg. 1450) then error
	if Instr (strTime, ":") = 0 then
		wscript.echo " * Error: Argument is not a time "
		wscript.quit
	else
		' if string is longer than 5 characters (eg 14:50:43) then error
		if len (trim (strTime)) <> 5 then 
			wscript.echo " * Error: Time is longer than 5 characters long" & _
					 " * Error: Please specify time as HH:mm"
			wscript.quit
		else
			
		end if
		
		' split the time into an hour string and a minute string
		arrTime = Split (strTime, ":")
		
		' check the hour string isn't below 0 or above 23
		if arrTime(0) < 0 or arrTime(0) > 23 then
			wscript.echo " * Error: No of Hours is incorrect"
			wscript.quit
				
		end if
		
		' check the minute string isn't below 0 or above 59
		if arrTime(1) < 0 or arrTime(1) > 59 then
			wscript.echo " * Error: No of Minutes is incorrect"
			wscript.quit
		end if

	end if

end Sub




' subroutine which does most of the work

Sub SubLogReader (strServer, StrStartDate, strFinishDate)
	
	' On error allows us to go past errors, which have no error text
	' otherwise script bombs out.
	On Error Resume Next	
	
	' start free third party object to connect to eventlogs
	set oELogReader = createobject("Softwing.EventLogReader")
	
	' start ado object to allow use of ADO
	set oConnection = createobject("ADODB.Connection")

	' Open connection to local sql server using oledb
	oConnection.Open "Provider=sqloledb;" & _
           	"Data Source=(local);" & _
           	"Initial Catalog=EventLogs;" & _
          	"User ID=test;" & _
           	"Password=test"	
	
	' clear the logs table
	' set oCommand = createObject ("ADODB.Command")
	'
	' set the command text (this is our sql query)
	' oCommand.CommandText =   "DELETE FROM " & strLocal & ".EventLogs.dbo.Logs "
		
	' Use the connection already set up by ADO
	' oCommand.ActiveConnection = oConnection

	' Execute the sql query, return any text into the results string
	' oCommand.Execute strResults

	''wscript.echo "Results " & Results & " added to DB."

	' close the command object
	' set oCommand = nothing
	
	'strDate = "17/11/2002"
	strDate = Date
	
	' add current date to start and end dates
	 strStartDate = strDate & " " & strStartDate
	 strFinishDate = strDate & " " & strFinishDate
        
        'wscript.echo " * Opening " & strServer
	
	' open application log
	oELogReader.OpenLog strServer, strLog
	
	' get the highest record number (which is the first record)
	nRecord = oElogReader.Getfirst
	
	''wscript.echo "First record is " & nRecord	
	
	' setup strings for maximum and minimum
	nRecordMax = nRecord
	nRecordMin = 0
	
	'wscript.echo " * Opening records " & nRecordMax & " to " & nRecordmin
	wscript.echo " * Between times " & strStartDate & " and " & strFinishDate & vbcrlf
	
	
	' start for next loop to enumerate back through the records from most recent to zero
	for x = nRecordMax to nRecordMin step -1
		
		'clear the err object for error trap
		err.clear
	
		' retrieve record
		oElogReader.GetRecord x, EvID, EvType, strSource, EvDate, strComputer, strAccount, strDesc
		
		select case lcase(strSource)
		
			case "perflib"
			
			case "iisinfoctrs"
			
			case else
			
				' error trap in case error text of record can not be interpreted
				if err.number = -2147418113 then strDesc = "Can not obtain error text"

				' remove cr, lf, backticks and -- for input checking (if this isn't done
				' an incorrectly formed error record could start a sql query)
				strDesc = Replace (strDesc, vbcrlf, "")
				strDesc = Replace (strDesc, vbcr, "")
				strDesc = Replace (strDesc, vblf, "")
				strDesc = Replace (strDesc, "'", "")
				strDesc = Replace (strDesc, "--", "")

				if Len (strDesc) > 500 then
					strDesc = "Error text too Long: " & Left (strDesc, 500)
				end if

				'wscript.echo x& " " & FormatDateTime (Evdate) & " " & FormatDateTime (strFinishDate)

				' checks if the error date is less than the finish time
				if FormatDateTime (EvDate) < FormatDateTime (strFinishDate) then

					'checks if the error date is greater than the start time
					if formatDateTime (EvDate) > FormatDateTime(strStartDate) then
						wscript.echo " " & x & " # " & EvType & " # " & FormatDateTime (EvDate) & " # " & EvID & " # " & strSource & " # " & strComputer & " # " & strDesc

							' Change the date format to be a SQL compatible date/time
							strDate = FormatDateTime (EvDate)

							arrSplit = Split (strDate)
							arrDate = Split (arrSplit(0), "/")

							strDD = arrDate (0)
							strMM = arrDate (1)
							strYYYY = arrDate (2)

							strDateReconstruct = strMM & "-" & strDD & "-" & strYYYY & " " & arrSplit(1)

							if boolLogtoSQL = TRUE then

								' Here's the sql I want to perform


								' start a command object
								set oCommand = createObject ("ADODB.Command")

								' set the command text (this is our sql query)
								oCommand.CommandText =   "INSERT INTO " & strLocal & ".EventLogs.dbo.Logs " & _ 
								"(EntryID, EventClass, oDateTime, EventID, Source, Server, oDescription) " & _ 
								"VALUES ('" & x & "', '" & EvType & "', '" & strDateReconstruct & "', '" & EvID & "', '" & strSource &"', '" & strComputer & "', '" & Left (strDesc,500) & "')"

								 ' Use the connection already set up by ADO
								oCommand.ActiveConnection = oConnection

								err.clear

								' Execute the sql query, return any text into the results string
								oCommand.Execute strResults

								if err.number <> 0 then 
									wscript.echo "**** I'VE ERRORED *****"
								end if

								''wscript.echo "Results " & Results & " added to DB."

								' close the command object
								'set oCommand = nothing
							else
							end if

					else
						' In here error records are not valid and outside of our specified range
						' so we exit the for loop to speed up script
						'wscript.echo vbcrlf & " * Exit for loop"
						exit for

					end if

				end if
		end select

	next
	
	' close the event log we've been reading
	oElogReader.closelog()
	
	' close the log reader object
	set oElogreader = nothing
	
	' close the ADO object we've been using to connect.
	set oConnection = nothing

end sub

