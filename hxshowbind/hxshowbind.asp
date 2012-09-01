'*************************************************************
' Script	: hxshowbind.vbs - try to make a ipmap, showbind
'				   userfriendly
' Author 	: Adrian Farnell  
' Date Amended	: 14.11.2000
' Version	: v1.0
'
'
' Prerequisites : Ipmap and shbind need to be present in same dir
'
' ************************************************************



' ********************************************
' Usage
' ********************************************

response.write ""
response.write 	" USAGE: cscript hxshowbind.vbs" & vbcrlf & _
		"        When run in the same directory as an ipmap" & vbcrlf & _
		"        and showbind, this script will display a" & Vbcrlf & _
		"        userfriendly representation of it." & vbcrlf

on error resume next

' ********************************************
' setup important constants
' ********************************************

constipmapfile = "ecomipmap.txt"
constshbindfile = "shbind.txt"

' ********************************************
' Start filesystem object
' ********************************************

set objfso = server.createobject ("scripting.filesystemobject")

' ********************************************
' Open ipmap file for read-only
' ********************************************
err.clear
set objipmap = objfso.opentextfile (constipmapfile,1)
set objipmapgot = objfso.getfile (constipmapfile)

' ********************************************
' error out if file doesn't exist
' ********************************************

if Err.number <> 0 then
	response.write " Failed to open " & constipmapfile & ". File does not exist in same directory."
	'set objFso = nothing
	server.quit(1)
end if

' ********************************************
' Check version info in ipmapfile.
' ********************************************

Do while objipmap.atendofstream <> true

	strveripmap = objipmap.readline
	if instr (strveripmap, "Version") <> 0 then
		response.write " " & constipmapfile & " " & strveripmap
	else
		'response.write " Version info on " & constipmapfile & " not available."
	end if 
Loop

stripmapmodified = objipmapgot.datelastModified
response.write " " & constipmapfile & " was last modified on " & stripmapmodified 


' ********************************************
' Open shbind file for read-only
' ********************************************
err.clear
set objshbind = objfso.opentextfile (constshbindfile,1)
set objshbindgot = objfso.getfile (constshbindfile)


' ********************************************
' error out if file doesn't exist
' ********************************************

if Err.number <> 0 then
	response.write " Failed to open " & constshbindfile & ". File does not exist in same directory." & vbcrlf
	'set objFso = nothing
	server.quit(1)
end if

strshbindmodified = objshbindgot.datelastModified
response.write " " & constshbindfile & " was last modified on " & strshbindmodified & vbcrlf

' ********************************************
' Read through showbind
' ********************************************


Do while objshbind.AtEndOfStream <> true

	'puts one line of file into string at a time	
	strline = trim (objshbind.readline)
	
	'filters out (OOS)
	if instr (strline, "(IS)") = 0 then
	else
	

		'filters out all lines which do not contain 443
		if instr (strline, "443") = 0 then
		else
			'tidies up line so that only IP address is left
			if instr (strline, ":") = 0 then
			Else
				
				'string that is left will only contain IP address of 443 if it's in service or out of service
				strline = Left (strline, instr(strline, ":")-1)
				'response.write strline	

				objipmap.close
				set objipmap = objfso.opentextfile (constipmapfile,1)
				Do while objipmap.atendofstream <> true
					
					
					stripmapline = objipmap.readline
					
					'Take out any preceding zeroes from ipmap
					if instr (stripmapline, ".00") <> 0 then
						'response.write stripmapline
						stripmapline = replace(stripmapline, ".00", ".")
						'response.write stripmapline
					end if
					if instr (stripmapline, ".00") <> 0 then
						'response.write stripmapline
						stripmapline = replace(stripmapline, ".0", ".")
						'response.write stripmapline
					end if
					
					'compare ipmap to shbind and display results					
					if  instr (stripmapline, strline) <> 0 then 
					'response.write " " & stripmapline & " : " & strline
					arrwebsitesup = split (stripmapline, ",")
					response.write " " & arrwebsitesup(0) & " " & arrwebsitesup(1)
					End if
					

				Loop


			End if
		end if
	end if
Loop


' ***************************************END OF SCRIPT, FUNCTIONS FOLLOW*********************************************************

' ********************************************
' Error subroutine
' ********************************************

sub Showerr (sDescription)
	response.write "Error Subroutine: " & Sdescription
	'response.write "Error Subroutine: An error occured - code: " & Hex (Err.Number) & " - "& Err.description
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