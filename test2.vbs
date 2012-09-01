strComputer = "."

strLog = "Application"

' start connection to sql
set oConnection = createobject ("ADODB.connection")

oConnection.Open 	"Provider=sqloledb;" & _
			"Data Source=(local);" & _
			"Initial Catalog=Dev;" & _
			"Integrated Security=SSPI"

			'"User ID=test;" & _
			'"Password=test"

set objRS = createobject ("ADODB.RecordSet")


Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colLoggedEvents = objWMIService.ExecQuery _
    ("Select * from Win32_NTLogEvent Where Logfile = '" & strLog & "'")

For Each objEvent in colLoggedEvents
	strStart = Timer
    	on error resume next
	strCategory = objEvent.Category
	strComputer = objEvent.ComputerName
	strEvent = objEvent.EventCode
	
	if IsNull (objEvent.Message) = True then
		strDescription = "no error text"
	else 
		err.clear
		strDescription = Replace (objEvent.Message, vbcrlf, " ")
		strDescription = Replace (strDescription, vbcr, " ")
		strDescription = Replace (strDescription, vblf, " ")
		strDescription = Replace (strDescription, "'", "|")
	
		if err.number <> 0 then 
			wscript.echo " ** " & objEvent.Message & ", " & objEvent.SourceName & ", " & objEvent.timewritten & " ** " 
			wscript.echo "Err: " & Err.Number & ", " & Err.description
		end if
	
	end if
	'Wscript.Echo "Record Number: " & objEvent.RecordNumber
	strSource = objEvent.SourceName
	strType = objEvent.Type
	strUser = objEvent.User


	strYYYY = Left (objEvent.TimeWritten, 4)
	strMM = Mid (objEvent.TimeWritten, 5, 2)
	strDD = Mid (objEvent.TimeWritten, 7, 2)
	strHH = Mid (objEvent.TimeWritten, 9, 2)
	strMin = Mid (objEvent.TimeWritten, 11, 2)
	strSS = Mid (objEvent.TimeWritten, 13, 2)

	strDateTime = strYYYY & "-" & strMM & "-" & strDD & " " & strHH & ":" & strMin & ":" & strSS

	'wscript.echo strDateTime

	strSource = funcReturnSourceID(strSource)
	strUser = funcReturnUserID (strUser)
	strComputer = funcReturnComputerID (strComputer)
	strDescription = funcReturnDescriptionID (strDescription)

	subAddEventRecord  strType, strUser, strComputer, strSource, strCategory, strEvent, strDescription, strDateTime 
    	on error goto 0
    	strFinish = Timer
    	if strFinish - strStart > 1 then 
    		wscript.echo "Took " & formatnumber (strFinish - strStart,2) & "s to add"
    	end if
Next

' end connection to sql 
set oConnection = nothing


sub subAddEventRecord ( strType, strUser, strComputer, strSource, strCategory, strEvent, strDescription, strDateTime )

	strAdd = "INSERT INTO uEventlog (Type, oDate, SourceID, Category, Event, UserID, ComputerID, DescriptionID, oLog) values " & _
		"('" & strType & "', '" & strDateTime & "', '" & strSource & "','" & strCategory & "','" & strEvent & "','" & strUser & "','" & strComputer &"','" & strDescription & "','" & strLog & "')"
	err.clear	
	oConnection.execute (strAdd)
	
	if err.number <> 0 then 
		wscript.echo "Err in Sub: " & Err.number & ": " & Err.description
		'wscript.echo strType & " :: " & strUser & " :: " & strComputer & " :: " & strSource & " :: " & strCategory & " :: " & strEvent & " :: " & strDescription & " :: " & strDateTime
	end if

	wscript.echo strType & " :: " & strUser & " :: " & strComputer & " :: " & strSource & " :: " & strCategory & " :: " & strEvent & " :: " & strDescription & " :: " & strDateTime

end sub

function funcReturnDescriptionID (strDescription)

	strCheck = "SELECT DescriptionID FROM uDescription WHERE DescriptionText = '" & strDescription & "'"
	strAdd = "INSERT INTO uDescription(DescriptionText) VALUES ('" & strDescription & "')"
	
	set objRS = oConnection.execute (strcheck)
	
	If objRS.eof then 
		oConnection.execute (strAdd)
	else
		'do nothing record exists.
	end if

	set objRS = oConnection.execute (strcheck)
	
	Do while not objRS.eof 
	
		strResults = objRS("DescriptionID")
		objRS.movenext
	
	Loop
	
	funcReturnDescriptionID = strResults
	
	'objRS.close
	
	'set objRS = nothing 

end function

function funcReturnComputerID (strComputer)

	strCheck = "SELECT ComputerID FROM uComputer WHERE Computername = '" & strComputer & "'"
	strAdd = "INSERT INTO uComputer(Computername) VALUES ('" & strComputer & "')"
	
	set objRS = oConnection.execute (strcheck)
	
	If objRS.eof then 
		oConnection.execute (strAdd)
	else
		'do nothing record exists.
	end if

	set objRS = oConnection.execute (strcheck)
	
	Do while not objRS.eof 
	
		strResults = objRS("ComputerID")
		objRS.movenext
	
	Loop
	
	funcReturnComputerID = strResults
	
	'objRS.close
	
	'set objRS = nothing 

end function

function funcReturnUserID (strUser)

	'wscript.echo "strUser: " & strUser & " :"

	strCheck = "SELECT UserID FROM uUser WHERE UserName = '" & strUser & "'"
	strAdd = "INSERT INTO uUser(UserName) VALUES ('" & strUser & "')"
	
	set objRS = oConnection.execute (strcheck)
	
	If objRS.eof then 
		oConnection.execute (strAdd)
	else
		'do nothing record exists.
	end if

	set objRS = oConnection.execute (strcheck)
	
	Do while not objRS.eof 
	
		strResults = objRS("UserID")
		objRS.movenext
	
	Loop
	
	funcReturnUserID = strResults
	
	'objRS.close
	
	'set objRS = nothing 

end function

function funcReturnSourceID (strSource)

	strCheck = "SELECT SourceID FROM uSource WHERE sourceText = '" & strSource & "'"
	strAdd = "INSERT INTO uSource(SourceText) VALUES ('" & strSource & "')"
	
	set objRS = oConnection.execute (strcheck)
	
	If objRS.eof then 
		oConnection.execute (strAdd)
	else
		'do nothing record exists.
	end if

	set objRS = oConnection.execute (strcheck)
	
	Do while not objRS.eof 
	
		strResults = objRS("SourceID")
		objRS.movenext
	
	Loop
	
	funcReturnSourceID = strResults
	
	'objRS.close
	
	'set objRS = nothing 

end function 