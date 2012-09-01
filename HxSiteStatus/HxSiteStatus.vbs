'/****************************************************************************/
'/*                                                                          */
'/* VBScript                                                                 */
'/*  SiteSatus.VBS    - Web Site Checker                                     */
'/*                                                                          */
'/* SYNOPSIS                                                                 */
'/* Returns current status of web sites on the server                        */
'/*                                                                          */
'/*  PROJECT:  E-COMMERCE PLATFORM (P0429)                                   */
'/*  PLATFORM: E-Commerce                                                    */
'/*                                                                          */
'/* AUTHOR                                                                   */
'/*  Andrew Sewell                                                           */
'/*                                                                          */
'/* DATE                                                                     */
'/*  7th June 2000                  	- (1.0) Creation 		     */
'/*  3rd July 2000			- (2.0) Added code to display        */
'/*                                             IPaddress bindings and       */
'/*  					        selection of server   
'/*  4th July 2000      		- (3.0) Added SMTP reporting, error handling
'*/  11th July 2000			- (4.0b) Extra comments,  */
'/****************************************************************************/

on error resume next
Dim objW3SVC, objSITE, strOUT, retval, status, retvalipbinding, objArgs, sserver, strSMTPout

' First check arguments to see of there is one

set objArgs = Wscript.arguments

if objArgs.count <> 1 then 
	Wscript.echo "Error, no arguments specified using default localhost"
	Wscript.echo ""
	sserver = "127.0.0.1"
Else
	sserver = objARgs(0)
 
End if 

'usage

	Wscript.echo "USAGE:	cscript hxsitestatus.vbs [server_ip]"
	Wscript.echo ""
	Wscript.echo "		Displays sites installed on a server,"
	Wscript.echo "		Whether the installed sites are running,"
	Wscript.echo "		and the ip address and port https is"
	Wscript.echo " 		is bound to."
	wscript.echo "		Script will exit after 20secs if there are problems"
	wscript.echo ""
	wscript.echo "		Note: User must be logged on the domain of the" 
	wscript.echo "		server in question or have a user account which"
	wscript.echo "		name and password both exist on the domain you are"
	wscript.echo "		currently logged in as and on the domain of "
	wscript.echo "		the server (passwords must match)"
	wscript.echo ""

'set up array 

status = array("Starting", "Started", "Stopping", "Stopped", "Pausing", "Paused", "Continuing")

Set objW3SVC = GetObject("IIS://" & sserver & "/W3SVC")      

' error handling

If (Err <> 0) Then
	strOUT = "Error " & Hex(Err.Number) & "("
	strOUT = strOUT & Err.Description & ") occurred."      
Else

'  Enumerate web sites & status, skipping the Default Web Site (no 1)

wscript.echo "Metabase         WebSiteName            status?  Bound to IP"
wscript.echo "---------------------------------------------------------------------"
	For Each objSITE In objW3SVC      
		If objSITE.class = "IIsWebServer" Then
			If objSITE.Name <> 1 Then

				strOUT = strOUT & "LM/W3SVC/" & objSITE.Name & Chr(9)
				strOUT = strOUT & Chr(34) & objSITE.ServerComment
				retval = objSITE.ServerState
				retvalipbinding = join (objsite.serverbindings, ", ")
				strOUT = strOUT & Chr(34)
				'if len (strout) < 225 then strout = strout & Chr(9) else strout = strout
				if len (objSITE.servercomment) < 14 then strtabs = Chr(9) else strtabs = ""
				strout = strout & Chr(9) & strtabs & status(retval-1) & " (" & retvalipbinding & ")" & Chr(13) & Chr(10)
			End If
			  
		End If      
	Next
End If      

' Display Output

WScript.Echo strOUT 

set objSMTPSVC = GetObject ("IIS://" & sserver & "/SMTPSVC/1")
	
	wscript.echo ""
	wscript.echo "SMTPSVCName                      Smarthost      status?  Bound to IP"
	wscript.echo "---------------------------------------------------------------------"

If Err.number <> 0 Then

	showerr "IIS Server " & sserver & " Does not have SMTP configured/Installed"
	wscript.exit err
	
Else



	strSMTPout = strSMTPout & objSMTPsvc.Servercomment & chr(9)
	strSMTPout = strSMTPout & Chr(34) & objSMTPsvc.smarthost & chr(34)
	retval = objSMTPsvc.Serverstate
	retvalipbinding = join (objsmtpsvc.serverbindings, ", ")
	if retvalipbinding = ":25:" then retvalipbinding = "All Unassigned:25"
	strSMTPout = strSMTPout & Chr(9) & status (retval-1) & " (" & retvalipbinding & ")" & chr(13) & chr(10)

End IF

wscript.echo strSMTPout 



' Destroy Object

Set objW3SVC = nothing



sub Showerr (sDescription)
	Wscript.echo "HxSiteStatus.vbs: " & Sdescription
	'Wscript.echo "HXsitestatus.vbs: An error occured - code: " & Hex (Err.Number) & " - "& Err.description
End Sub



