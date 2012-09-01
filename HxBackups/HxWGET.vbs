' ***********************************************************
' Name    : HxWGET.vbs - Launches a wget to each of the websites on a server
' Author  : Adrian Farnell
' Date    : 6.12.2002
' Version : 1.0
' Notes   : Designed to get round the problem if IIS log file locking. And to generate a 
'           a valid logfile for each website.  
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
Dim strLogDir, strPrefix

'debugging off by default
boolDebug = FALSE

'set the string below for log directory (needs to be set)
'strLogDir = "c:\temp"

'set the string below for prefix on logs (doesn't need to be set)
'strPrefix = "TMP"

'process the arguments
subProcessSwitches (2)

' -----------------------------------( insert code here )-----------------------------------------

'Dim strings
Dim strIpMap, strComputerName, strLine

'Dim Objects
Dim oFile, oFSO, oWshNetwork

'Dim counters
Dim Counter

'Dim arrays
Dim arrLine

'ipmap to interrogate
strIpMap = "c:\winnt\system32\drivers\etc\ebusipmap.txt"

'find out my computers name

set oWshNetwork = createObject ("wscript.network")
strComputerName = oWshNetwork.ComputerName

'hook into filesystem
set oFSO = createobject ("scripting.filesystemobject")

'read through file, obtain ip addresses for each of the websites.

if oFSO.fileExists ( strIpmap ) = true then
	'file does exist
	set ofile = ofso.opentextfile(strIPmap) 
else
	'file does not exist
	subErrMsg "File, " & stripMap & " does not exist"
end if

'Loop through file
Do while ofile.atendofstream <> True
	
	'read one line
	strLine = ofile.readline
	
	' if there's nothing on the line, or theres a hash ignore it
	if Len(strLine) = 0 or Left(strLine,1) ="#" then
		'do nothing
	else
		' if the computer name matches the computername specified on the line process it.
		If Lcase (strcomputername) = Lcase (Left (strLine, Instr(strLine, ",")-1 )) then
			
			'debug message about line being processed
			subDebugMsg strLine
			
			'split the line by comma
			arrLine = split (strLine, ",")
			
			'send a http request to the third entry on the line (15 times)
			for counter = 1 to 15			
				subWget arrLine(2)			
			next
						
		else 

		end if
	end if

Loop

'wget subroutine
sub SubWget (strIP)

	'Dim objects
	Dim oWshShell

	'start up shell object
	set oWshShell = createobject ("wscript.shell")
	
	'debug message describing what its hitting
	subDebugMsg "Running wget against http://" & strIP
	
	'run wget on the command line
	oWshShell.Run "wget --delete-after http://" & strIP
	
	'clean up object
	set oWshShell = nothing

end sub 

' ------------------------------------------( cut ) ----------------------------------------------

sub subTestRun

	call subLogtoFile ("Hello", strLogDir)
	call subCreateDir (strLogDir)
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

	if strLogDir = "" then subErrMsg "strLogDir not set"

	' Dim objects
	Dim oFSO, oFile, oWshNetwork
	
	' Dim strings
	Dim strDate, strDD, strMM, strYYYY, strProcessDir
	
	' Dim Arrays
	Dim arrLogDir
	
	' Dim counters
	Dim i

	' create hook into filesystem
	set oFSO = createobject ("scripting.filesystemobject")
	
	' create hook into network
	set oWshNetwork = createobject ("wscript.network")
	
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
	if ofso.fileexists (strLogDir & "\" & strPrefix & "_" &oWshNetwork.ComputerName & "_" & strYYYY & strMM & strDD & ".log") = true then
		subDebugMsg "file does exist"
		
		'open file for appending
		set ofile = ofso.opentextfile (strLogDir & "\" & strPrefix & "_"& oWshNetwork.ComputerName & "_" & strYYYY & strMM & strDD & ".log",8)
	else
		subDebugMsg "file doesn't exist"
		
		'create file
		set ofile = ofso.createtextfile (strLogDir & "\" & strPrefix & "_" &oWshNetwork.ComputerName & "_" & strYYYY & strMM & strDD & ".log")
	end if

	'write the message into the file
	ofile.write strMsg
	
	'and save
	ofile.close
	
	'clean up objects
	set ofile = nothing
	set ofso = nothing 
	set owshnetwork = nothing

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
	subDebugMsg Now & " Err: " & strMsg
	
	'write error into eventlog (first method is for Scripting host 5.1 or higher)
	'oWshShell.logevent 1,Now & " " & wscript.scriptname & " Err: " & strMsg
	owshShell.run "logevent -s E " & chr(34) & strMsg & chr(34) ,0, true	

	'cleanup objects
	set oWshShell = nothing
	
	'end script, set errorlevel to 1 for failure
	wscript.quit(1)
	
end sub

sub SubDisplayUsage
	
	'display usage
	submsg	vbcrlf & vbcrlf & " Usage: template.vbs [/?] [/debug]" & vbcrlf & _
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