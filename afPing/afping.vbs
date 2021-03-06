' ***********************************************************
' Name    : 
' Author  :
' Date    :
' Version :
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
Dim strLogDir, strPrefix, strServertoPing, iCount, strTotalPings, strPacketsLost

'debugging off by default
boolDebug = FALSE

'set the string below for log directory (needs to be set)
strLogDir = "c:\temp"

'set the string below for prefix on logs (doesn't need to be set)
strPrefix = "Ping_"

strServertoPing = ""

strTotalPings = 0
strPacketsLost = 0

'process the arguments
subProcessSwitches (3)

' -----------------------------------( insert code here )-----------------------------------------

if Len (strServertoPing) = 0 then
	subErrMsg "No server to ping"
	wscript.quit
else

end if

Do
	subPing strServerToPing
	subping "ntw313755"
	subping "ntw309277"
	wscript.sleep 1000
Loop

' ------------------------------------------( cut ) ----------------------------------------------

sub subPing (strServer)

	Dim strTarget, strHost, colPing, oResult, strError
	
	strTarget = strServer
	strHost = "."
	strError = 0
	strTotalPings = strTotalPings + 1
	
	set colPing = GetObject("winmgmts:{impersonationLevel=impersonate}//" & _
		 strHost & "/root/cimv2"). ExecQuery("SELECT * FROM Win32_PingStatus " & _
		"WHERE Address = '" + strTarget + "'")

	For each oResult in colPing
	
		select case oResult.Statuscode
		
			case 0 
				subdebugmsg strServer & " is responding"
				strError = 0
			case 11001
				subLogtoFile "Ping to " & strServer & "  Failed: Buffer to small", strLogDir
				subdebugmsg "Ping to " & strServer & " Failed: Buffer to small"
				strError = 1
			case 11002
				subLogtoFile "Ping to " & strServer & " Failed: Destination Net Unreachable", strLogDir
				subdebugmsg "Ping to " & strServer & " Failed: Destination Net Unreachable"
				strError = 1			
			case 11003
				strPacketsLost = strPacketsLost + 1
				subLogtoFile "Ping to " & strServer & " Failed: Destination Host Unreachable (Packet Loss: " & formatnumber ((1 -(strTotalPings - strPacketsLost)/strTotalPings) * 100, 1) & "% Loss)" , strLogDir
				subdebugmsg "Ping to " & strServer & " Failed: Destination Host Unreachable"
				strError = 1			
			case 11004
				subLogtoFile "Ping to " & strServer & " Failed: Destination Protocol Unreachable", strLogDir
				subdebugmsg "Ping to " & strServer & " Failed: Destination Protocol Unreachable"	
				strError = 1		
			case 11005
				subLogtoFile "Ping to " & strServer & " Failed: Destination Port Unreachable", strLogDir
				subdebugmsg "Ping to " & strServer & " Failed: Destination Port Unreachable"
				strError = 1			
			case 11006
				subLogtoFile "Ping to " & strServer & " Failed: No Resources ", strLogDir
				subdebugmsg "Ping to " & strServer & " Failed: No Resources "	
				strError = 1		
			case 11007
				subLogtoFile "Ping to " & strServer & " Failed: Bad Option", strLogDir
				subdebugmsg "Ping to " & strServer & " Failed: Bad Option"
				strError = 1		
			case 11008
				subLogtoFile "Ping to " & strServer & " Failed: Hardware Error", strLogDir
				subdebugmsg "Ping Failed: Hardware Error"
				strError = 1			
			case 11009
				subLogtoFile "Ping to " & strServer & " Failed: Packet Too Big", strLogDir
				subdebugmsg "Ping Failed:to " & strServer & "  Packet Too Big"	
				strError = 1		
			case 11010
				strPacketsLost = strPacketsLost + 1
				subLogtoFile "Ping to " & strServer & " Failed: Request Timed Out (Packet Loss: " & formatnumber ((1 -(strTotalPings - strPacketsLost)/strTotalPings) * 100, 1) & "% Loss)" , strLogDir
				subdebugmsg "Ping to " & strServer & " Failed: Request Timed Out"	
				strError = 1			
			case 11011
				subLogtoFile "Ping to " & strServer & " Failed: Bad Request", strLogDir
				subdebugmsg "Ping to " & strServer & " Failed: Bad Request"
				strError = 1		
			case 11012
				subLogtoFile "Ping to " & strServer & " Failed: Bad Route", strLogDir
				subdebugmsg "Ping to " & strServer & " Failed: Bad Route"	
				strError = 1		
			case 11013
				subLogtoFile "Ping to " & strServer & " Failed: TimeToLive Expired in Transit", strLogDir
				subdebugmsg "Ping to " & strServer & " Failed: TimeToLive Expired in Transit"	
				strError = 1
			case 11014
				subLogtoFile "Ping to " & strServer & " Failed: TimeToLive Expired Reassembly", strLogDir
				subdebugmsg "Ping to " & strServer & " Failed: TimeToLive Expired Reassembly"
				strError = 1		
			case 11015
				subLogtoFile "Ping to " & strServer & " Failed: Parameter Problem", strLogDir
				subdebugmsg "Ping to " & strServer & " Failed: Parameter Problem"	
				strError = 1				
			case 11016
				subLogtoFile "Ping to " & strServer & " Failed: Source Quench", strLogDir
				subdebugmsg "Ping to " & strServer & " Failed: Source Quench"	
				strError = 1	
			case 11017
				subLogtoFile "Ping to " & strServer & " Failed: Option Too Big", strLogDir
				subdebugmsg "Ping to " & strServer & " Failed: Option Too Big"	
				strError = 1		
			case 11018
				subLogtoFile "Ping to " & strServer & " Failed: Bad Destination", strLogDir
				subdebugmsg "Ping to " & strServer & " Failed: Bad Destination"	
				strError = 1	
			case 11032
				subLogtoFile "Ping to " & strServer & " Failed: Negotiating IPSEC", strLogDir
				subdebugmsg "Ping to " & strServer & " Failed: Negotiating IPSEC"	
				strError = 1			
			case 11050
				subLogtoFile "Ping to " & strServer & " Failed: General Failure", strLogDir
				subdebugmsg "Ping to " & strServer & " Failed: General Failure"
				strError = 1			
			case else
				subLogtoFile "Ping to " & strServer & " Failed: Unknown Failure, Error code: " & oResult.Statuscode, strLogDir
				subdebugmsg "Ping to " & strServer & " Failed: Unknown Failure, Error code: " & oResult.Statuscode
				strError = 1				
		end select 
	
		if strError = 0 then
			subDebugMsg "Displaying other Properties....."
			subDebugMsg "Address: " & oResult.Address
			subDebugMsg "BufferSize: " & oResult.BufferSize
			subDebugMsg "NoFragmentation: " & oResult.NoFragmentation
			subDebugMsg "PrimaryAddressResolutionStatus: " & oResult.PrimaryAddressResolutionStatus
			subDebugMsg "ProtocolAddress: " & oResult.ProtocolAddress
			subDebugMsg "ProtocolAddressResolved: " & oResult.ProtocolAddressResolved
			subDebugMsg "RecordRoute: " & oResult.RecordRoute
			subDebugMsg "ReplyInconsistency: " & oResult.ReplyInconsistency
			subDebugMsg "ReplySize: " & oResult.ReplySize
			subDebugMsg "ResolveAddressNames: " & oResult.ResolveAddressNames
			subDebugMsg "ResponseTime: " & oResult.ResponseTime
			subDebugMsg "ResponseTimeToLive: " & oResult.ResponseTimeToLive
			subDebugMsg "RouteRecord: " & oResult.RouteRecord
			subDebugMsg "RouteRecordResolved: " & oResult.RouteRecordResolved
			subDebugMsg "SourceRoute: " & oResult.SourceRoute
			subDebugMsg "SourceRouteType: " & oResult.SourceRouteType
			subDebugMsg "StatusCode: " & oResult.StatusCode
			subDebugMsg "Timeout: " & oResult.Timeout
			subDebugMsg "TimeStampRecord: " & oResult.TimeStampRecord
			subDebugMsg "TimeStampRecordAddress: " & oResult.TimeStampRecordAddress
			subDebugMsg "TimeStampRecordAddressResolved: " & oResult.TimeStampRecordAddressResolved
			subDebugMsg "TimestampRoute: " & oResult.TimestampRoute
			subDebugMsg "TimeToLive: " & oResult.TimeToLive
			subDebugMsg "TypeofService: " & oResult.TypeofService
		else
		
		
		
		end if

	
	next

	set colPing = nothing 
	
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

	' add date to the message
	strMsg = vbcrlf & Now & " " &  strMsg 

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
	submsg	vbcrlf & vbcrlf & " Usage: afPing.vbs [/?] [/debug] [/server:<servername>]" & vbcrlf & _
		vbcrlf & "    /?                     - Displays this help" & _
		vbcrlf & "    /debug                 - Shows debug info" & _
		vbcrlf & "    /server:<servername>   - Pings <servername>, logs failures to a textfile"
		
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
			select case lcase (Left (wscript.arguments(nArgs),2))
				case "/?"
					subDisplayUsage
					wscript.quit
				case "/d"
					'display debug info where appropriate
					boolDebug = TRUE
				case "/s"
					strServertoPing = Mid (wscript.arguments(nArgs), Instr (wscript.arguments(nArgs), ":")+1 )
					subDebugMsg "strServertoPing is set to: " & strServertoPing
				case else
					'do not display debug info
					boolDebug = FALSE

			end select
		next
	end if 

end sub