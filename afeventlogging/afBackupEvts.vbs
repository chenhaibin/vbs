' ***********************************************************
' Name    : afBackupEvts.vbs
' Author  : Adrian Farnell
' Date    : 19.11.2002
' Version : 1.0
' Notes   : Script to backup and clear the eventlogs
'
' Subroutines included in this script:
'
'  subLogToFile ( strMsg, strLogDir ) - Logs Message into file in directory specified, filename will be based on date.
'  subMsg ( strMsg )                  - Logs Message to console/messagebox
'  subDebugMsg ( strMsg )             - If boolDebug = TRUE, displays all debug messages
'  subErrMsg ( strMsg )               - Encounters an error, displays message in console/messagebox and eventlog. Quits script
'  subDisplayUsage		      - Displays usage (this'll have to be doctored for each script)
'  subProcessSwitches ( strNum )      - Process switches (again will have to be doctored for each script)
'  subCreateDir ( strDir )            - Creates Directory
'
' ************************************************************

Option Explicit

Dim boolDebug

'debugging off by default
boolDebug = FALSE

'process the arguments
subProcessSwitches (2)

' -----------------------------------( insert code here )-----------------------------------------

'Dim strings
Dim strComputer, strLogDir, strDate, strDD, strMM, strYYYY, errbackuplog

'Dim objects
Dim oWshNetwork, oFSO, objWMIService, objLogFile

'Dim collections
Dim colLogFiles

'sets the script to look at localhost
strComputer = "."

'sets the log directory
strLogDir = "c:\temp\logs\evts"

'get date string ready for processing
strDate = Date

'what is the day
strDD = Left ( Date, 2)

'what is the month
strMM = Mid (Date, 4,2)

'what is the year
strYYYY = Mid (Date, 7,4)

'checks and creates the log directory (if neccesary)
subCreateDir strLogDir

'hook into the WshNetwork object (allows us to get computername)
set owshnetwork = createobject ("wscript.network")

'hook into filesystem object (allows us to check if file exists
set ofso = createobject ("scripting.filesystemobject")

'start the WMI object 
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate,(Backup, Security)}!\\" & strComputer & "\root\cimv2")

'get all logfiles called application
Set colLogFiles = objWMIService.ExecQuery ("Select * from Win32_NTEventLogFile where LogFileName='Application'")

'for each logfile in the collection....(should only be one)
For Each objLogfile in colLogFiles
	
	'check if a backup already exists for today
	if ofso.fileexists (strLogDir & "/" & oWshnetwork.computername & "_" & strYYYY & strMM & strDD & "_app.evt") = true then
		
		subDebugMsg "Backup Exists of Application Log"
	
	else
		'if no backup is exists then create the backup
		subDebugMsg "Creating " &strLogDir & "/" & oWshnetwork.computername & "_" & strYYYY & strMM & strDD & "_app.evt"
		errBackupLog = objLogFile.BackupEventLog(strLogDir & "/" & oWshnetwork.computername & "_" & strYYYY & strMM & strDD & "_app.evt")
		
		'if there were problems with backing up, then error
		If errBackupLog <> 0 Then        

			subErrMsg "The Application event log could not be backed up."

		Else
			'clear the event log file
			subdebugmsg "I would have deleted your logs here (App)"
			objLogFile.ClearEventLog()

		End If
		
	end if
	
Next

'get logfiles called system
Set colLogFiles = objWMIService.ExecQuery ("Select * from Win32_NTEventLogFile where LogFileName='System'")

'for each logfile in the collection.....(should only be one)
For Each objLogfile in colLogFiles
	
	'check to see whether backup exists
	if ofso.fileexists (strLogDir & "/" & oWshnetwork.computername & "_" & strYYYY & strMM & strDD & "_sys.evt") = true then
		subDebugMsg "Backup Exists of System Log"
	else
		
		'backup doesn't exist it, create it
		subDebugMsg "Creating " &strLogDir & "/" & oWshnetwork.computername & "_" & strYYYY & strMM & strDD & "_Sys.evt"
		errBackupLog = objLogFile.BackupEventLog(strLogDir & "/" & oWshnetwork.computername & "_" & strYYYY & strMM & strDD & "_sys.evt")
		
		'if problems we're experienced then error
		If errBackupLog <> 0 Then        

			subErrMsg "The System event log could not be backed up."

		Else

			subdebugmsg "I would have deleted your logs here (Sys)"
			objLogFile.ClearEventLog()

		End If
	end if
Next

'get all log files called security
Set colLogFiles = objWMIService.ExecQuery ("Select * from Win32_NTEventLogFile where LogFileName='Security'")

'for each logfile in the collection (there should only be one)
For Each objLogfile in colLogFiles

	'check to see if backup already exists
	if ofso.fileexists (strLogDir & "/" & oWshnetwork.computername & "_" & strYYYY & strMM & strDD & "_sec.evt") = true then
		subDebugMsg "Backup Exists of Security Log"
	else
		'backup doesn't exist, so create it
		subDebugMsg "Creating " &strLogDir & "/" & oWshnetwork.computername & "_" & strYYYY & strMM & strDD & "_Sec.evt"
		errBackupLog = objLogFile.BackupEventLog(strLogDir & "/" & oWshnetwork.computername & "_" & strYYYY & strMM & strDD & "_Sec.evt")
		
		'any probs with backup then error
		If errBackupLog <> 0 Then        

			subErrMsg "The Security event log could not be backed up."

		Else
			
			'clear the logfiles.
			subdebugmsg "I would have deleted your logs here (Sec)"
			objLogFile.ClearEventLog()


		End If

	end if

Next

'clean up objects
set ofso = nothing
set objWMIService = nothing
set owshnetwork = nothing

' ------------------------------------------( cut ) ----------------------------------------------

sub subTestRun

	call subLogtoFile ("Hello", "c:\temp\logs\test")
	call subCreateDir ("c:\temp\CreateDir\test")
	call subMsg ("Test Message")
	call subDebugMsg ("Test Debug Message")
	call SubErrMsg ("I've Errored")

end sub

sub subCreateDir (strDir)

	' Dim objects
	Dim oFSO, oFile
	
	' Dim strings
	Dim strDate, strDD, strMM, strYYYY, strProcessDir
	
	' Dim Arrays
	Dim arrLogDir
	
	' Dim counters
	Dim i

	' create hook into filesystem
	set oFSO = createobject ("scripting.filesystemobject")
	
	'get date string ready for processing
	strDate = Date
	
	'what is the day
	strDD = Left ( Date, 2)
	
	'what is the month
	strMM = Mid (Date, 4,2)
	
	'what is the year
	strYYYY = Mid (Date, 7,4)
	
	'debug for values of date
	subDebugMsg strDD & " " & strMM & " " & strYYYY

	'split each of the subdirectories up so that we can check if they exist
	arrLogDir = split (strdir, "\")
	
	'start of with no text in drive string
	strProcessDir = ""
	
	'enumerate each of the directories one by one
	for i = 0 to Ubound (arrLogDir)
		
		'if this is the first array value e.g c:\ just add it to strProcessDir
		if i = 0 then
			strProcessDir = arrLogDir(i)	
		else
			'if this is any other array value, append it on to the end of the string	
			strProcessDir = strProcessDir & "\" & arrLogDir(i) 
		end if
		
		' check to see whether the folder specified in strProcessDir exists.  if not create it
		if ofso.folderexists ( strProcessDir ) = true then
			subDebugMsg strProcessDir & " exists"
		else
			subDebugMsg strProcessDir & " doesn't exist, creating..."
			
			'create folder
			ofso.createfolder ( strProcessDir ) 
			
			'check it was created. Else error.
			if ofso.folderexists (strProcessDir) = true then
				subDebugMsg strProcessDir & " created"
			else
				subErrMsg strProcessDir & " could not be created"
			end if
			
		end if
		
	next 
	

end sub

sub subLogtoFile (strMsg, strLogDir)

	' Dim objects
	Dim oFSO, oFile
	
	' Dim strings
	Dim strDate, strDD, strMM, strYYYY, strProcessDir
	
	' Dim Arrays
	Dim arrLogDir
	
	' Dim counters
	Dim i

	' create hook into filesystem
	set oFSO = createobject ("scripting.filesystemobject")
	
	'get date string ready for processing
	strDate = Date
	
	'what is the day
	strDD = Left ( Date, 2)
	
	'what is the month
	strMM = Mid (Date, 4,2)
	
	'what is the year
	strYYYY = Mid (Date, 7,4)
	
	'debug for values of date
	subDebugMsg strDD & " " & strMM & " " & strYYYY

	'create directories if needed
	subCreateDir (strlogdir)

	' check whether logfile exists in the directory. if not create a new file using the date as basis
	if ofso.fileexists (strLogDir & "\" & strYYYY & strMM & strDD & ".log") = true then
		subDebugMsg "file does exist"
		
		'open file for appending
		set ofile = ofso.opentextfile (strLogDir & "\" & strYYYY & strMM & strDD & ".log",8)
	else
		subDebugMsg "file doesn't exist"
		
		'create file
		set ofile = ofso.createtextfile (strLogDir & "\" & strYYYY & strMM & strDD & ".log")
	end if

	'write the message into the file
	ofile.write strMsg
	
	'and save
	ofile.close
	
	'clean up objects
	set ofile = nothing
	set ofso = nothing 

end sub

sub subMsg (strMsg)
	
	'display message to console/messagebox
	wscript.echo Now & " Msg: " & strMsg

end sub

sub subDebugMsg (strMsg)
	
	'only if debugging is on display debugging messages to console/messagebox
	if BoolDebug = TRUE then	
		'Dim strMsg
		wscript.echo Now & " Debug: " & strMsg
	else
	end if	
		
	
end sub

sub subErrMsg (strMsg)
	
	'Dim objects
	Dim oWshShell
	
	'start oWshShell so we can log events
	set oWshShell = createobject ("wscript.shell")
	
	'display error to console/messagebox
	wscript.echo Now & " Err: " & strMsg
	
	'write error into eventlog
	oWshShell.logevent 1,Now & " " & wscript.scriptname & " Err: " & strMsg
	
	'cleanup objects
	set oWshShell = nothing
	
	'end script, set errorlevel to 1 for failure
	wscript.quit(1)
	
end sub

sub SubDisplayUsage
	
	'display usage
	submsg	vbcrlf & vbcrlf & " Usage: afBackupEvts.vbs [/?] [/debug]" & vbcrlf & _
		vbcrlf & "    /?        - Displays this help" & _
		vbcrlf & "    /debug    - Shows debug info"
		
end sub

sub subProcessSwitches (strNum)

	'Dim counters
	Dim nArgs

	'get arguments from running the script, if there are too many then error
	if wscript.arguments.count > strNum then 
		subMsg " Too many Arguments"
		subDisplayUsage	
		wscript.quit
	else 
		'enumerate through arguments and choose what to do based on those arguments
		for nArgs = 0 to wscript.arguments.count - 1
			select case wscript.arguments(nArgs)
				case "/?"
					subDisplayUsage
					wscript.quit
				case "/debug"
					'display debug info where appropriate
					boolDebug = TRUE
				case else
					'do not display debug info
					boolDebug = FALSE
			end select
		next
	end if 

end sub