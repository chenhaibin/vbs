' ************************************************************
' Script	: hxipcheck.vbs - Uses ipmap files to check
'				  if all IP addresses are installed
' Author 	: Adrian Farnell  
' Date Amended	: 17.10.2000
' Version	: pre-alpha
'
' Prerequisites : None
'
' ************************************************************



' ********************************************
' Usage
' ********************************************

wscript.echo ""
wscript.echo 	" USAGE: 	cscript hxipcheck.vbs " & vbcrlf & _
		"		Uses ipmap files to check if all	" & vbcrlf & _
		"		ip addresses are installed correctly"

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