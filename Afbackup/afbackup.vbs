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

Dim boolDebug, blnQuit
Dim strLogDir, strPrefix, strReport

'debugging off by default
boolDebug = FALSE

'quit on errors by default
blnQuit = TRUE

'set the string below for log directory (needs to be set)
'strLogDir = "c:\temp"

'set the string below for prefix on logs (doesn't need to be set)
'strPrefix = "TMP"

'process the arguments
subProcessSwitches (2)

' -----------------------------------( insert code here )-----------------------------------------

' Dim objects
Dim oFSO, ofile

' Dim strings
Dim strLine, strBackupFolder, strDir, strWStart, strWFinish

' Open backup.ini

'strReport = "Backup begun on: " & date & vbcrlf

strWstart = Timer

set oFSO = createObject ("scripting.filesystemobject")

' Does backup.ini exist
if oFSO.fileexists ("backup.ini") = true then
	subDebugMsg "Backup.ini File Exists"
else
	subErrMsg "Backup.ini does not exist", TRUE
end if

set oFile = ofso.opentextfile ("backup.ini") 

Do while ofile.atendofstream = false

	strline = ofile.readline
	
	if Len (strLine) = 0 or Left (strLine,1) = "#" or Left(strLine,1) = "'" then
	else
		subDebugMsg strLine
		
		if lcase (Left (strLine, 9)) = "backupdir" then
			
			strLine = Mid ( strLine, InstrRev (strLine, "=") + 1 )
			
			subCreateDir strLine
			
			strBackupFolder = strLine
		
			subDebugMsg "Backup Directory is: " & strBackupFolder
		else 
		end if

		if Left (strLine, 1 ) = "[" then
			
			strLine = Replace (strLine, "[", "")
			strLine = Replace (strLine, "]", "")
			
			strDir = strBackupFolder & "\" & strLine
			
			subCreateDir strDir

		else
		
		end if
		
		if instr (strLine, "file=") <> 0 then
			
			strline = replace (strLine, "file=", "")
		
			subCopyFile strDir, strLine
		
		end if
		
		if instr (strLine, "files=") <> 0 then 
		
			strLine = replace (strLine, "files=", "")
			
			subCopyWildCardFiles strDir, strLine
			
		end if
		
		if instr (strLine, "dir=") <> 0 then
		
			strLine = replace (strLine, "dir=", "")
			
			subCopyFolder strDir, strLine
			
		end if
		
		
	end if
Loop

strWFinish = Timer

subMsg  "Backup took " & FormatNumber (strWFinish - strWStart,2) & " seconds"

subReportToEvt strReport

' ------------------------------------------( cut ) ----------------------------------------------

sub subCopyFile (strDir, strfile)
	
	Dim strStart, strFinish, ofile1, ofile2
	
	subDebugMsg "Entered subCopyFile (" & strDir & "," & strfile & ")"
	
	strStart = Timer
	
	'Does source file exist
	if ofso.fileexists ( strFile ) = true then
		
		'Does file already exist in backup folder
		if oFso.fileExists ( strDir & "\" & Mid (strFile, InstrRev (strFile, "\")+1)) = true then
			subDebugMsg strDir & "\" & Mid (strFile, InstrRev (strFile, "\")+1) & " exists, overwriting"
		else
		end if

		'Overwrite file regardless
		on error resume next
		err.clear
		ofso.copyfile strFile, strDir & "\", True

		if err.number <> 0 then 
			subErrMsg error.description & "(" & err.number & ")", FALSE
		end if

		on error goto 0

		

		set ofile1 = ofso.getfile (strFile)

		set ofile2 = ofso.getfile (strDir & "\" & Mid (strFile, InstrRev (strFile, "\")+1))

		if ofile1.datelastmodified = ofile2.dateLastmodified and ofile1.size = ofile2.size then
			subMsg "Backup of " & strFile & " Success"
		else
			subErrMsg "Backup of " & strFile & " may have failed.", FALSE
		end if

		set oFile1 = nothing
		set ofile2 = nothing

	else
		subErrMsg strFile & " source file does not exist, skipping", FALSE
	end if
	
	strFinish = Timer
	
	subDebugMsg "Backup Files took " & strFinish - strStart & " seconds"
	
end sub

sub subCopyWildCardFiles (strDir, strLine)

	Dim strStart, strFinish, ofile1, ofile2, colFiles, oFolder,strFolder, file
	
	subDebugMsg "Entered subCopyWildCardFiles (" & strDir & "," & strLine & ")"

	strStart = Timer
	
	if ofso.folderexists (strDir & "\" & Mid (strFolder, InstrRev (strFolder, "\")+1)) = true then
		
		set oFolder = ofso.getfolder (Left (strLine, InstrRev (strLine, "\")-1))
		set colFiles = oFolder.files

		strLine = Mid (strLine, InstrRev (strLine, "\")+1)
		
		strLine = Replace (strLine, "*", "")
		
		subDebugMsg "Searching for files with '" & strLine & "'"
		
		For each file in colFiles
		
			'wscript.echo file.name
						
			if Instr (file.name, strLine) <> 0 then  
				on error resume next 
				err.clear

				subDebugMsg "Copying File " & file.name
				ofso.copyfile file.path,  strDir & "\" & Mid (strFolder, InstrRev (strFolder, "\")+1) & "\", TRUE

				set ofile1 = ofso.getfile (file.path)
				set ofile2 = ofso.getfile (strDir & "\" & Mid (strFolder, InstrRev (strFolder, "\")+1) & "\" & file.name)

				if ofile1.size = ofile2.size and ofile1.dateLastModified = ofile2.dateLastModified then
					subMsg "File " & file.name & " Backed Up Success"
				else
					subErrMsg "File " & file.name & " Failed", FALSE
				end if


				if err.number <> 0 then
					subErrMsg error.description & "(" & err.number & ")"
				end if
				on error goto 0
			else
				'file not relevant
				
		
			end if
		
		Next
		
	else
		subErrMsg strFolder & " source folder does not exist", FALSE
	end if
	
	strFinish = Timer

	subDebugMsg "Backup Files took " & strFinish - strStart & " seconds"
	
end sub

sub subCopyFolder (strDir, strFolder)

	Dim strStart, strFinish, ofolder1, ofolder2

	strStart = Timer 
	
	subDebugMsg "Entered subCopyFolder (" & strDir & ", " & strFolder & ")"
	
	if ofso.folderexists (strFolder) = true then
	
		'Does folder already exist
		if ofso.folderExists (strDir & "\" & Mid (strFolder, InstrRev (strFolder, "\")+1)) = true then
			subDebugMsg strDir & "\" & Mid (strFolder, InstrRev (strFolder, "\")+1) & " exists, overwriting"
			ofso.deletefolder strDir & "\" & Mid (strFolder, InstrRev (strFolder, "\")+1), true
		else
		end if
		
		'overwrite folder
		on error resume next
		err.clear
		ofso.copyfolder strFolder, strDir & "\" & Mid (strFolder, InstrRev (strFolder, "\")+1), true
		
		if err.number <> 0 then
			subErrMsg error.description & "(" & err.number & ")"
		end if
		
		on error goto 0
		
		set oFolder1 = ofso.getfolder (strFolder)
		set oFolder2 = ofso.getfolder (strDir & "\" & Mid (strFolder, InstrRev (strFolder, "\")+1))
		
		if ofolder1.size = ofolder2.size then 'and ofolder.
			subMsg "Backup of " & strFolder & " Success"
		else
			subErrMsg "Backup of " & strFolder & " may have failed", FALSE
		end if
		
		set ofolder1 = nothing
		set ofolder2 = nothing
	else 
		subErrMsg strFolder & " source folder does not exist", FALSE
	end if
	
	strFinish = Timer

	subDebugMsg "Backup Folder took " & strFinish - strStart & " seconds"
	
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
	
	strReport = strReport & Now & " Msg: " & strMsg & vbcrlf

end sub

sub subDebugMsg (strMsg)
	
	'only if debugging is on display debugging messages to console/messagebox
	if BoolDebug = TRUE then	
		'Dim strMsg
		wscript.echo Now & " Debug: " & strMsg
		strReport = strReport & Now & " Debug: " & strMsg & vbcrlf
	else
	end if	
		
	
end sub

sub subErrMsg (strMsg, blnQuit)
	
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
	
	if blnQuit = true then
	
		'end script, set errorlevel to 1 for failure
		wscript.quit(1)
		
	else
		'carry on
	end if
	
	strReport = strReport & Now & " Err: " & strMsg & vbcrlf
	
end sub

sub subReportToEvt (strMsg)

	'Dim objects
	Dim oWshShell
	
	'start oWshShell so we can log events
	set oWshShell = createobject ("wscript.shell")
	
	'write error into eventlog
	oWshShell.logevent 4,Now & " " & wscript.scriptname & " Report: " & vbcrlf & strMsg	

	'cleanup objects
	set oWshShell = nothing

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