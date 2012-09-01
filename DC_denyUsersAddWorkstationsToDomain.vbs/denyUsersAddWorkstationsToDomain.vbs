
' ********************************************************
' ScriptName: denyUsersAddWorkstationsToDomain.vbs
'
' Notes: Amends the ms-DS-MachineAccountQuote attribute to  
'        Zero. This stops Authenticated users adding 
'        Computers to a domain.
'
' Author: Adrian Farnell
'
' Version: 1.0
'
' Date: 25/07/2009
'
' ********************************************************

' ********************************************************
'  Option Explicit
' ********************************************************
Option Explicit


' ********************************************************
' Dim Objects
' ********************************************************
Dim objLDAP
Dim strDomain


' ********************************************************
' Set Domain
' ********************************************************
strDomain = "LDAP://DC=Contoso,DC=com"

' ********************************************************
' Connect to LDAP (incl Err Checking)
' ********************************************************
err.clear

on error resume next

set objLDAP = GetObject (strDomain)

if err.number <> 0 then

	subError err.message, err.Number 

end if

on error goto 0

' ********************************************************
' Set Attribute
' ********************************************************
objLDAP.Put "ms-DS-MachineAccountQuota", "0"

' ********************************************************
' SetInfo() (incl Err Checking)
' ********************************************************
err.clear

on error resume next

objLDAP.SetInfo

if err.number <> 0 then

	subError err.message, err.Number 

end if

on error goto 0

' ********************************************************
' End Script with a zero Errorlevel
' ********************************************************
wscript.echo Now & " [SUCCESS] : " & err.number


' ********************************************************
' Error Subroutine
' ********************************************************
sub subError (strMsg, strErrCode)

	'Event log it

		'Dim object
		Dim oWshShell
	
		'start oWshShell so we can log events
		set oWshShell = createobject ("wscript.shell")
	
	
		'write error into eventlog
		oWshShell.logevent 1, Now & " " & wscript.scriptname & " Err: " & strMsg
		
		'cleanup objects
		set oWshShell = nothing
	
	'Display it
	wscript.echo Now() & " [ERR " & strErrCode & ", " & wscript.scriptname & "] : " & strMsg


	'And quit with the error code
	wscript.quit(strErrCode)

end sub