' ************************************************************
' Script	: afnamecheck.vbs - Works out names of ecommerce
'			    servers.
'
' Author 	: Adrian Farnell  
' Date Amended	: 05.12.2000
' Version	: pre-alpha
'
' Prerequisites : None
'
' ************************************************************

on error resume next

' Start arguments for command line
set collargs = wscript.arguments

if collargs.count <> 1 then 

	' ********************************************
	' Usage
	' ********************************************

	wscript.echo ""
	wscript.echo 	" USAGE: 	afnamecheck.vbs <name of server>" & vbcrlf & _
			"		Works Out Names of ecommerce Servers" & vbcrlf & _
	wscript.quit(1)
else
end if

strname = collargs(0)






namecheck strname

function namecheck (strname)

	'Check strname to see if it has a value
	if strname = "" then
		wscript.echo " Error: Name of server has not been entered, Refer to Usage."
		wscript.quit(1)
	else
	end if
	
	'Display strname
	wscript.echo " Server Name to test is : " & strname
	wscript.echo vbcrlf & " Script will now work out what kind of server this is from it's name"

	'First Four Letters analysis

	strWhoseServer = Left (strname, 4)
	wscript.echo " First four letters : " & strwhoseserver

	'*****************************************************
	' Servers in this select case statement are put into types
	' The type of server is an indication of where the variables
	' for SQL servers etc are set. 
	' 1 - Means Platform contains one server for each role, and first four letters are same for each server
	' 2 - Means Platform contains one server for each role, first four letters are different
	' 3 - Means Platform contains one server for each role, Server uses specialist Minnows DB
	' 4 - Minirig environment (SQL roles shared), assumes Copley
	' 5 - IF Minirig environment (Only one SQL box), assumes Copley
	' 6 - Unknown
	'*****************************************************
	
	select case strwhoseserver
		case "ocom"
			strwho = "Open"
			strtype = 2
		case "ecom"
			strwho = "RFS Ecommerce"
			strtype = 1
		case "icom"
			strwho = "IF"
			strtype = 1
		case "ncom"
			strwho = "NED"
			strtype = 2
		case "ccom"
			strwho = "CMIG Ecommerce"
			strtype = 3
		case "cbus"
			strwho = "CMIG Ebusiness"
			strtype = 3
		case "tcom"
			strwho = "Treasury"
			strtype = 3
		case "bcom"
			strwho = "BMID"
			strtype = 3
		case "ebus"
			strwho = "RFS Ebusiness"
			strtype = 1
		case "mrig"
			strwho = "RFS Minirig Environment (probably)"
			strtype = 4
		case "mrga"
			strwho = "IF Minirig Environment"
			strtype = 5
		case "mrgb"
			strwho = "Unidentifiable Minirig Environment"
			strtype = 6
		case "mrgc"
			strwho = "Unidentifiable Minirig Environment"
			strtype = 6
		case "deva"
			strwho = "Development Rig"
			strtype = 6
		case "idev"
			strwho = "Development Rig"
			strtype = 6
		case else
			strwho = "Unknown, Contact ECP"	
			strtype = 6
	End Select
	
	wscript.echo " Server belongs to : " & strwho & " (" & strtype & ")"
	
	'Server Location
	strlocationandlevel = mid (strname, 5, 1)
	'wscript.echo " Location Digit : " & strlocationandlevel

	select case strtype
		case 4
			strlocation = "Copley"
		case 5
			strlocation = "Copley"
		case 1,2,3
			If strlocation < 5 then 
				strlocation = "Copley"
			else
				strlocation = "Pudsey"
			end if
		case else 
			strlocation = "Unknown"
	End Select		

	
	Wscript.echo " Location of server is : " & strlocation

	'Test, Live or Development?
	
	if strtype => 4 then 
		strlevel = "Development"
	else
		select case strlocationandlevel
			case 0,5
				strlevel = "LIVE"
			case 1,6
				if strwhoseserver = "icom" then
					strlevel = "PRE-PRODUCTION"
				else
					strlevel = "LEVEL1"
				end if
			case 2,7
				strlevel = "LEVEL2"
			case 3
				strlevel = "BAU"
			case else
				strlevel = "Unknown"
		end select 	
	end if
	Wscript.echo " Server is in " & strlevel & " environment."


	'Type of server?

end function

 



' ********************************************
' Error subroutine
' ********************************************

sub Showerr (sDescription)
	Wscript.echo "Error Subroutine: " & Sdescription
	'Wscript.echo "Error Subroutine: An error occured - code: " & Hex (Err.Number) & " - "& Err.description
End sub

'*********************************************
' log to file function
'*********************************************

Function funclogtofile (strmessage)

	strdate = Date
	strdate = replace (strdate, "/", "_")
	Set ofso = createobject ("scripting.filesystemobject")
	set ofsologfile = ofso.Opentextfile ( strdate & ".log",8,true)

	strlog = date & " " & Time & " - : " & strmessage & vbcrlf
	ofsologfile.write strlog
	ofsologfile.close
	
End Function