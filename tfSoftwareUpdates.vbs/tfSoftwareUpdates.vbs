
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

' Dim 
Dim strVersionFileLocation

strVersionFileLocation = "\\Jupiter\buildshare\version.txt"

if funcCheckVersion = true then
	subMsg "System is up to date"
	wscript.quit
else
	subMsg "System needs updating"
	subApplyNextPackage
end if

' what OS am i?
wscript.quit

' ------------------------------------------( cut ) ----------------------------------------------

function funcCheckVersion

	Dim oShell, oFSO, oFile
	
	Dim strRegVer, strLine, strFileVersion
	
	' try to read the version control key

	set oShell = createObject("wscript.shell")
	
	set oFSO = createObject ("scripting.filesystemobject")
	
	err.clear
	
	On error resume next
	
	strRegVer = oshell.regread ("HKLM\Software\TechFrontier\VersionNo")
	
	If err.number <> 0 then
		
		err.clear
		
		'reg key doesn't exist create it
		oshell.regwrite "HKLM\Software\TechFrontier\VersionNo", "0000"
		
		if err.number <> 0 then
	
			subErrMsg "Error writing registry value"
			wscript.quit (1)
		else
			subDebugMsg "Registry entry written"
			strRegVer = "0000"
		end if
		
	else
		'do nothing
	end if
	
	On error goto 0
	
		
	' try to read the version control text file on server
	if ofso.fileexists ( strVersionFileLocation ) = false then
		subErrMsg "Version file does not exist in " & strVersionFileLocation
		wscript.quit (1)
	else
	
	end if
	
	set oFile = oFSO.Opentextfile ( strVersionFileLocation,1 )
	
	Do while oFile.atendofstream <> true
	
		strLine = oFile.readline
	
		if Len (strLine) = 0 or Left (strLine,1) = "#" then
			'ignore line
		else
			
			
			If lcase (Left (strLine, 10)) = "versionno:" then
				'found version number
				strFileVersion =  Replace (strLine, "VersionNo: ", "")
			else
				subErrMsg "Could not find version data within file"
			end if
			
		end if
	
	Loop
	
	subDebugMsg "Version in registry: " & strRegVer
	subDebugMsg "Version in file: " & strFileVersion
	
	if strRegVer < strFileVersion then
	
		'need more updates
		funcCheckVersion = FALSE
		
	else
		
		'machine is up to date
		funccheckVersion = TRUE
		'subMsg "System is up to date"
		
	end if

	set oShell = nothing
	set oFSO = nothing

	
	
end function

sub subApplyNextPackage

	Dim oShell, oFSO, oFile
	
	Dim strRegVer, strLine, strFileVersion, strUpdate
	
	Dim BoolSuccess
	
	' try to read the version control key

	set oShell = createObject("wscript.shell")
	
	set oFSO = createObject ("scripting.filesystemobject")
	
	err.clear
	
	On error resume next
	
	strRegVer = oshell.regread ("HKLM\Software\TechFrontier\VersionNo")
	
	If err.number <> 0 then
		
		err.clear
		
		'reg key doesn't exist create it
		oshell.regwrite "HKLM\Software\TechFrontier\VersionNo", "0000"
		
		if err.number <> 0 then
	
			subErrMsg "Error writing registry value"
			wscript.quit (1)
		else
			subDebugMsg "Registry entry written"
			strRegVer = "0000"
		end if
		
	else
		'do nothing
	end if
	
	On error goto 0

	strUpdate = strRegVer + 1
	
	subDebugMsg "strUpdate: " & strUpdate
	
	select case strUpdate
	
		case 1
			subDebugMsg "Installing Package 1"
			
			oshell.logevent 4, "Installation of Package 1 started"
			
			boolSuccess = true
			
			if boolSuccess = true then
				
				oshell.logevent 4, "Installation of Package 1 Finished"
			
				oshell.regwrite "HKLM\Software\TechFrontier\VersionNo", "0001"
				
				subMsg "Installation of Package 1 Finished"
				
			end if
		case 2
			subDebugMsg "Installing Package 2"
			
			oshell.logevent 4, "Installation of Package 2 started"
			
			boolSuccess = true
			
			if boolSuccess = true then
				
				oshell.logevent 4, "Installation of Package 2 Finished"
			
				oshell.regwrite "HKLM\Software\TechFrontier\VersionNo", "0002"
				
				subMsg "Installation of Package 2 Finished"
				
			end if
		case 3
			subDebugMsg "Installing Package 3"
			
			oshell.logevent 4, "Installation of Package 3 started"
			
			boolSuccess = true
			
			if boolSuccess = true then
				
				oshell.logevent 4, "Installation of Package 3 Finished"
			
				oshell.regwrite "HKLM\Software\TechFrontier\VersionNo", "0003"
				
				subMsg "Installation of Package 3 Finished"
				
			end if
		case 4
			subDebugMsg "Installing Package 4"
			
			oshell.logevent 4, "Installation of Package 4 started"
			
			boolSuccess = true
			
			if boolSuccess = true then
				
				oshell.logevent 4, "Installation of Package 4 Finished"
			
				oshell.regwrite "HKLM\Software\TechFrontier\VersionNo", "0004"
				
				subMsg "Installation of Package 4 Finished"
				
			end if
		case 5
			subDebugMsg "Installing Package 5"
		case 6
			subDebugMsg "Installing Package 6"
		case 7
			subDebugMsg "Installing Package 7"
		case 8
			subDebugMsg "Installing Package 8"
		case 9
			subDebugMsg "Installing Package 9"
		case else
			subDebugMsg "Unable to determine next package update"
	end select

	if funcCheckVersion = true then
		subMsg "System is up to date"
		wscript.quit
	else
		subMsg "System still needs updating"
		subApplyNextPackage
	end if

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
	strMsg = vbcrlf & Now & strMsg 

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