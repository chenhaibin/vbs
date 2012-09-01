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
Dim strLogDir, strPrefix, strSize, strFile, oFSO, ofile, oNewFile

strFile = "new.tmp"

'debugging off by default
boolDebug = FALSE

'set the string below for log directory (needs to be set)
'strLogDir = "c:\temp"

'set the string below for prefix on logs (doesn't need to be set)
'strPrefix = "TMP"

'process the arguments
subProcessSwitches (3)

' -----------------------------------( insert code here )-----------------------------------------

if IsNumeric (strSize) = False then
	subErrMsg "Error! Size specified is not a number"
else
	subDebugMsg strSize
end if

set oFSO = createobject ("scripting.filesystemobject")

if oFSO.fileexists (strfile) = true then
	set oFile = oFSO.getfile (strFile)
else
	set onewFile = oFSO.createTextFile (strFile)
	
	onewFile.writeline Now & " Created"
	
	onewFile.close
	
	set onewFile = nothing 
	
	set oFile = oFSO.getfile (strFile)
	
end if

subDebugMsg oFile.size & " is the size of the file"

if oFile.size < strSize then
	set oFile = nothing
	subGrowFile
end if

set oFSO = nothing

' ------------------------------------------( cut ) ----------------------------------------------

sub subGrowFile

	Dim iCount, strMax, strMin, strRnd, strLine, strEighty, strPercent, i, strchr

	subDebugMsg "Growing File...."

	set ofile = oFSO.opentextfile ( strFile, 8, true )
	
	subMsg "1 dot for each line added"
	
	For iCount = 1 to int (strSize/80)
	
		'strPercent = ( strSize - Len(strLine) )/100
		'
		'Select case strPercent
		'
		'	case 10,20,30,40,50,60,70,80,90
		'		subDebugMsg (100 - strPercent) & "% complete"
		'	case else
		'	
		'end select
	

		
		'strLine = strLine & "...................................................................................................." & vbcrlf 'Chr(strRnd)
		
		
	wscript.stdout.write "."	
			
		
		
		strLine = ""
		
		For i = 1 to 80
							
			Randomize 

			strMax = 60
			strMin = 146

			strRnd = Int ((strMax - strMin + 1) * Rnd() + strMin )

			strLine = strLine & Chr (strRnd)
		next 
		
		strLine = strLine & vbcrlf
		
		ofile.write strLine
		'strEighty = iCount / 80
		'
		'if int(strEighty) = strEighty then
		'	strLine = strLine & vbcrlf
		'end if
	
	next

	
	
	ofile.close
	
	set ofile = nothing

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
	submsg	vbcrlf & vbcrlf & " Usage: afCreateFileSize.vbs [/?] [/debug] [/sz:number_of_Bytes]" & vbcrlf & _
		vbcrlf & "    /?                   - Displays this help" & _
		vbcrlf & "    /debug               - Shows debug info" & _
		vbcrlf & "    /sz:number_of_Bytes  - Creates File of Size specified"
		
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
			select case Left (wscript.arguments(nArgs), 4)
				case "/?"
					subDisplayUsage
					wscript.quit
				case "/deb"
					'display debug info where appropriate
					boolDebug = TRUE
				case "/sz:"
					subDebugMsg "found sz"
					
					strSize = Mid (wscript.arguments(nArgs), Instr (wscript.arguments(nArgs), ":")+ 1)
					
					if IsNumeric (strSize) = False then 
						subErrMsg "Error: Input was not a number"
					end if 
					
					
				case else
					'do not display debug info
					boolDebug = FALSE
					

			end select
		next
	end if 

end sub