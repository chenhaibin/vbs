' ***********************************************************
' Name    : HxCheckMem.vbs
' Author  : Adrian Farnell
' Date    : 01.04.2003
' Version : 1.0
' Notes   :
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
Dim strLogDir, strPrefix, strServer,strScore, strMaster, strSlave
	
'debugging off by default
boolDebug = FALSE

'set the string below for log directory (needs to be set)
'strLogDir = "c:\temp"

'set the string below for prefix on logs (doesn't need to be set)
'strPrefix = "TMP"

strServer = ""
strScore = 0

'process the arguments
subProcessSwitches (3)

' -----------------------------------( insert code here )-----------------------------------------

if strServer = "" or Lcase (Left (strServer, 2)) = "/s" then 
	subDisplayUsage
else
	subQueryMem (strServer)
end if



' ------------------------------------------( cut ) ----------------------------------------------


sub subQueryMem (strServer)
	
	'Dim Objects
	Dim oConnection, RSsubs
	
	'Dim strings
	Dim strConnectString, strSQL
	
	'Hook into ADO
	set oConnection = createobject ("ADODB.Connection")
	
	'set Connection string to attach to SQL
	strConnectString = "driver={SQL Server}; server=" & strServer & ";uid=hxecomadmin;pwd=2difficult4u2c;database=SSMembership_Master"
	
	'start connection to sql server
	oConnection.Open strConnectString
	
	'set sql statement to be run on SQL server
	strSQL = "select * from sysobjects where name = 'sysSubscriptions'"
	
	'Execute statement
	set RSsubs  = oConnection.Execute (strSQL)
	
	'test for records in subscriptions table
	if RSSubs.BOF or RSSubs.EOF then
		subMsg "Server does not have any subscriptions"
		
		'here we'll have to create a fall back position to query the web servers on their MASTER/SLAVE
	else
		subDebugMsg "Subscriptions are present"
		
		'here we'll use another sub to query each of the databases and verify MASTER/SLAVE
		subQuerySubs "SSMembership_Master", strServer
		subQuerySubs "SSMembership_P1", strServer
		subQuerySubs "SSMembership_P2", strServer
		subQuerySubs "SSMembership_P3", strServer
		subQuerySubs "SSMembership_P4", strServer
		
	end if 
	
	select case strScore
	
		case 0
			'here we'll have to create a fall back position to query the web servers on their MASTER/SLAVE
		case 44
			subMsg "Master is: " & strServer & ", (100% certainty)"
		case else
			subMsg "Master is: " & strServer & ", (" & FormatNumber (strScore * 100 /44, 1 ) & "% certainty)"
			subMsg "Checking web servers...."
			'here we'll have to create a fall back position to query the web servers on their MASTER/SLAVE
		
	end select
	
	oConnection.close
	
	set oConnection = nothing
	
end sub 

sub subQuerySubs (strDB, strServer)

	Dim oConnection, RSSubscriptions
	
	Dim strConnectString, strQuery

	set oConnection = createObject ("ADODB.Connection")
	
	strConnectString = "driver={SQL Server}; server=" & strServer & ";uid=hxecomadmin;pwd=2difficult4u2c;database=" & strDB
	
	oConnection.open strConnectString
	
	strQuery = "Select * from " & strDB & "..sysSubscriptions"

	set RSsubscriptions = oConnection.execute (strQuery)
	
	do while not RSsubscriptions.eof
	
		Select Case RSSubscriptions ("subscription_type")
		
			case 1
				subDebugMsg strDB & " subscription Type is Pull"
			case 0
				subDebugMsg strDB & " subscription Type is Push"
				strScore = strScore + 1
			case else
				subDebugMsg strDB & " could not identify subscription type"
		end select
		
		Select Case RSSubscriptions ("status")
		
			case 0
				subDebugMsg strDB & " subscription is Inactive"
			case 1 
				subDebugMsg strDB & " subscription is Subscribed"			
			case 2
				subDebugMsg strDB & " subscription is Active"			
				strScore = strScore + 1
			case else
				subDebugMsg strDB & " Could not identify subscription status"		
		
		end select
		
		RSSubscriptions.movenext
		
	Loop

	oConnection.close
	
	set oConnection = nothing

end sub

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
	submsg	vbcrlf & vbcrlf & " Usage: hxCheckMem.vbs [/?] [/debug] [/server:<servername>]" & vbcrlf & _
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
			
			select case Left (wscript.arguments(nArgs),2)
				case "/?"
					subDisplayUsage
					wscript.quit
				case "/d"
					'display debug info where appropriate
					boolDebug = TRUE
					subDebugMsg "Debugging on"
					
				case "/s"	
					strServer = Replace (wscript.arguments(nArgs), "/server:", "")
					subDebugMsg "strServer set to: " & strServer
				case else
					'do not display debug info
					boolDebug = FALSE
			end select
			subDebugMsg "Argument (" & nargs & "):" & Left (wscript.arguments(nArgs),2)
		next
	end if 

end sub