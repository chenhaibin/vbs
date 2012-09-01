
Dim arrProcess ()

strLogDir = "c:\temp\logs\Process"
BoolDebug = FALSE

subProcessSwitches (2)

'On Error Resume Next
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
set oWshNetwork = createobject ("wscript.network")

Do

	Set colItems = objWMIService.ExecQuery("Select * from Win32_PerfRawData_PerfProc_Process",,48)
	Set colOperatingSystems = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")

	for each oOS in colOperatingSystems

		strOS = oOS.Version

	next

	count = 0

	For Each objItem in colItems

		count = count + 1

		Redim Preserve arrProcess(4, count)

		arrProcess (1, count) = objItem.Name
		arrProcess (2, count) = objItem.IDProcess
		arrProcess (3, count) = objItem.PercentProcessorTime

		select case strOS

			case "5.0.2195"
				arrProcess (4,count) = ""
			case "XP" 
				arrProcess (4, count) = objItem.Commandline
			case Else
				subErrMsg " Unidentified OS"
		end select

		if arrProcess (1,count) = "_Total" then
			strTotal = objItem.PercentProcessorTime
		else

		end if


	Next

	strMax = Ubound (arrProcess,2)

	strLog = Now & " " & owshnetwork.computername

	For i = 1 to strMax

		if arrProcess (1,i) = "_Total" then
		else
			strLog = strLog & " #" & arrProcess (1,i) & "(PID:" & arrProcess(2,i) & ", %:" & Round(arrProcess(3,i)/strTotal * 100,2) & ", CommandLine:" &arrProcess(4,1)& ")"
		end if

	next

	strLog = strLog & vbcrlf

	subDebugMsg strLog

	subLogTofile strLog, strLogDir

	wscript.sleep 60000

Loop

sub subLogtoFile (strMsg, strLogDir)

	set oFSO = createobject ("scripting.filesystemobject")

	strDate = Date

	strDD = Left ( Date, 2)
	strMM = Mid (Date, 4,2)
	strYYYY = Mid (Date, 7,4)

	subDebugMsg strDD & " " & strMM & " " & strYYYY

	arrLogDir = split (strlogdir, "\")
	strProcessDir = ""
	
	for i = 0 to Ubound (arrLogDir)
		
		if i = 0 then
			strProcessDir = arrLogDir(i)	
		else
			strProcessDir = strProcessDir & "\" & arrLogDir(i) 
		end if
		
		
		if ofso.folderexists ( strProcessDir ) = true then
			subDebugMsg strProcessDir & " exists"
		else
			subDebugMsg strProcessDir & " doesn't exist, creating..."
			ofso.createfolder ( strProcessDir ) 
			
			if ofso.folderexists (strProcessDir) = true then
				subDebugMsg strProcessDir & " created"
			else
				subErrMsg strProcessDir & " could not be created"
			end if
			
		end if
		
	next 

	if ofso.fileexists (strLogDir & "\Process_" & oWshNetwork.Computername &"_"& strYYYY & strMM & strDD & ".log") = true then
		subDebugMsg "file does exist"
		
		set ofile = ofso.opentextfile (strLogDir & "\Process_" & oWshNetwork.Computername &"_"& strYYYY & strMM & strDD & ".log",8)
	else
		subDebugMsg "file doesn't exist"
		set ofile = ofso.createtextfile (strLogDir & "\Process_" & oWshNetwork.Computername &"_"& strYYYY & strMM & strDD & ".log")
	end if

	ofile.write strMsg
	
	ofile.close
	
	set ofile = nothing
	set ofso = nothing 

end sub

sub subMsg (strMsg)
	
	wscript.echo Now & " Msg: " & strMsg

end sub

sub subDebugMsg (strMsg)
	if BoolDebug = TRUE then	
		'Dim strMsg
		wscript.echo Now & " Debug: " & strMsg
	else
	end if	
		
	
end sub

sub subErrMsg (strMsg)
	
	'Dim strMsg
	
	set oWshShell = createobject ("wscript.shell")
	
	wscript.echo Now & " Err: " & strMsg
	oWshShell.logevent 1,Now & " " & wscript.scriptname & " Err: " & strMsg
	
	set oWshShell = nothing
	
	wscript.quit(1)
	
end sub

sub SubDisplayUsage

	submsg	vbcrlf & vbcrlf & " Usage: afLogProcess.vbs [/?] [/debug]" & vbcrlf & _
		vbcrlf & "    /?        - Displays this help" & _
		vbcrlf & "    /debug    - Shows debug info"
end sub

sub subProcessSwitches (strNum)

	if wscript.arguments.count > strNum then 
		subMsg " Too many Arguments"
		subDisplayUsage	
		wscript.quit
	else 
		'carry on.....from arguments decide what to do.
		for nArgs = 0 to wscript.arguments.count - 1
			select case wscript.arguments(nArgs)
			
				case "/?"
					subDisplayUsage
					wscript.quit
				case "/debug"
					boolDebug = TRUE
				case else
					boolDebug = FALSE
			end select
		next
	end if 

end sub