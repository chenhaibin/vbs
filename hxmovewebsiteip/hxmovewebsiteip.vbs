' ************************************************************
' Script	: hxmovewebsiteip.vbs - moves installed websites 
'					to ip address in ipmap.txt 
'				 
'
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
wscript.echo 	" USAGE: 	hxmovewebsitesip.vbs /restore" & vbcrlf & _
		"		Displays Websites current config" & vbcrlf & _
		"		then changes that config to match " & vbcrlf & _
		"		ipmap.txt" & vbcrlf & vbcrlf & _
		"		/restore - if this switch is included the " & vbcrlf & _
		"		backupfile will be used to restore" & vbcrlf

set wsharguments = wscript.arguments
if wsharguments.count <> 1 then
	strRestore = ""
else
	strRestore = wsharguments(0)
end if

' *********************************************
' Display current config
' *********************************************

Stripmapfile = "e:\winnt\system32\drivers\etc\ecomipmap.txt"
'Stripmapfile = "95l085673.txt"

'start ADSI, filesystem and network objects
Set objW3SVC = getObject("IIS://127.0.0.1/W3SVC")      
set objwshnetwork = createobject ("wscript.network")
set objfso = createobject ("scripting.filesystemobject")

Strcomputername = objwshnetwork.computername

'start backout string to be able to recover the old ip addresses
Strbackout = "# Backout file for Website ipaddress move " & vbcrlf & "# Date: " & Date & " " & time

' error handling
If (Err <> 0) Then
	strOUT = "Error " & Hex(Err.Number) & "("
	strOUT = strOUT & Err.Description & ") occurred."      
Else

	'enumerate through w3svc object
	For Each objSITE In objW3SVC      
		
		'find IISwebserver instances
		If objSITE.class = "IIsWebServer" Then
			
			'as long as the site isn't 1 (usually the default Website), then Display current settings
			If objSITE.Name <> 1 Then
				
				strserverbindings = (Join(objsite.ServerBindings, ", "))
				strlength = Len (strserverbindings)
				strdifference = Len (strserverbindings) - 4
				strserverbindings = Left (strserverbindings, strdifference)
				strbackout = objwshnetwork.computername & "," & objsite.servercomment & ","  & strserverbindings
				'wscript.echo strbackout
			
			End If
			  
		End If
		
		'prepare backouttext string
		strbackouttext = strbackouttext & strbackout & vbcrlf
		
	Next
End If      

'wscript.echo "strbackouttext : " & strbackouttext

'write Backout text file to disk and close it

If not objfso.fileexists (strcomputername & ".txt") then
	Set objbackoutfile = objfso.opentextfile ( objwshnetwork.computername & ".txt",2,true)
	objbackoutfile.write strbackouttext & vbcrlf
	objbackoutfile.close
	'error handling
	If err.number <> 0 then 
		wscript.echo " Failed to open file, the Error number is (" & err.number & ") " & Err.description 
		funclogtofile "Error writing backout file"
		wscript.quit
	Else
		
		funclogtofile "writing to backout file (.txt)"
					
	end If

Else

	Set objbackoutfile = objfso.opentextfile ( objwshnetwork.computername & ".old",2,true)
	objbackoutfile.write strbackouttext & vbcrlf
	objbackoutfile.close
	'error handling
	If err.number <> 0 then 
		wscript.echo " Failed to open file, the Error number is (" & err.number & ") " & Err.description 
		funclogtofile "Failed to write backout file"
		wscript.quit
	Else
		
		funclogtofile "Written Backout file (.old)"
					
	end If

end if

'read in values from ipmap.txt into array.

if strRestore = "/restore" then 
	stripmapfile = strcomputername & ".txt"
	funclogtofile "Started restore of old data"
end if

strcomputername = Ucase(Strcomputername)
set objfile = objfso.opentextfile (stripmapfile, 1)

'error handling
	If err.number <> 0 then 
		wscript.echo " Failed to open file, the Error number is (" & err.number & ") " & Err.description 
		funclogtofile "Failed to read ipmap file"
		wscript.quit
	Else
		
		funclogtofile "Opened ipmap file successfully"

	End if	
	
'read through text file
Do while objfile.AtEndOfstream <> true
	
	'read one line at a time
	strline = objfile.readline
	
	'If computername matches line then process
	If Left(strline,8) = Left (Strcomputername, 8) then
		
		strfirstcomma = Instr (strline, ",")
		strline = Ucase (Mid (strline, strfirstcomma+1))
		'wscript.echo strline
		
		'array which contains the two values for Name of Website and its corresponding ip address
		Arrwebsiteandip = split (strline, ",")
		
		'wscript.echo Arrwebsiteandip(0) & " : " & Arrwebsiteandip(1)
		
		'enumerate websites using ADSI, and insert new values into the correct places
		For Each objSITE In objW3SVC      
			If objSITE.class = "IIsWebServer" then
				If objSITE.Name <> 1 Then
				
					if Ucase (objSITE.servercomment) = Arrwebsiteandip (0) then
						
						'wscript.echo objSITE.servercomment & ":" & Arrwebsiteandip(0)
						funclogtofile "Before Put operation: " & objSITE.servercomment & " : " & Join(objsite.ServerBindings, ", ")
						objSITE.put "serverbindings", Arrwebsiteandip(1) & ":80:"
						objSITE.setinfo
						objSITE.getinfo
						funclogtofile "After Put operation: " & objSITE.servercomment & " : " & Join(objsite.ServerBindings, ", ")
						
					End if
			
				End If
			  
			End If
		
  
		Next
		
	End If

Loop
objfile.close

if strRestore = "/restore" then
	objfso.deletefile stripmapfile, true
	funclogtofile "Backout restored, Original backout file deleted."
End if

Function funclogtofile (strmessage)
	
	'basis of Name of logfile
	strdate = Date
	strdate = replace (strdate, "/", "_")
	
	'start an instance of the filesystemobject
	Set ofso = createobject ("scripting.filesystemobject")
	
	'open a text file to append to
	set ofsologfile = ofso.Opentextfile ( strdate & ".log",8,true)
	
	'write string (will be only one line)
	strlog = date & " " & Time & " - : " & strmessage & vbcrlf
	
	'write string to text file
	ofsologfile.write strlog
	
	'close text file
	ofsologfile.close
	
End Function




