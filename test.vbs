
' simulated event record

strDate = "20/03/2003"
strTime = "08:00:00"
strType = "Information"
strUser = "N/A"
strComputer = "NTW324100"
strSource = "VMnetx"
strCategory = "None"
strEvent = "34"
strDescription = "(\Device\Vmxctr) Connecting VMnet8."
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


strSource = funcReturnSourceID(strSource)
strUser = funcReturnUserID (strUser)
strComputer = funcReturnComputerID (strComputer)
strDescription = funcReturnDescriptionID (strDescription)

strDateTime = Now

subAddEventRecord  strType, strUser, strComputer, strSource, strCategory, strEvent, strDescription, strDateTime 

' end connection to sql 
set oConnection = nothing


sub subAddEventRecord ( strType, strUser, strComputer, strSource, strCategory, strEvent, strDescription, strDateTime )

	strAdd = "INSERT INTO uEventlog (Type, oDate, SourceID, Category, Event, UserID, ComputerID, DescriptionID, oLog) values " & _
		"('" & strType & "', '03-20-2003 08:00:00', '" & strSource & "','" & strCategory & "','" & strEvent & "','" & strUser & "','" & strComputer &"','" & strDescription & "','" & strLog & "')"
		
	oConnection.execute (strAdd)
	
	if err.number <> 0 then 
		wscript.echo Err.number & ": " & Err.description
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
	
	objRS.close
	
	set objRS = nothing 

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
	
	objRS.close
	
	set objRS = nothing 

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
	
	objRS.close
	
	set objRS = nothing 

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
	
	objRS.close
	
	set objRS = nothing 

end function 