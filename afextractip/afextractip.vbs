
' afextractip.vbs - Extracts the IP address from $winnt$.inf and dumps it to screen. 
' Author : Adrian Farnell
' Date Amended : 30.06.00
' Version no: Pre-Alpha

' create filesystemobject
set objfso = createobject ("scripting.filesystemobject")
set objwshnetwork = wscript.createobject ("wscript.network")

' Assign servername to a string
strServerName = objwshnetwork.computername

wscript.echo strservername

' Open $Winnt$.inf for reading
Wscript.echo "Opening $winnt$.inf for reading"
set objfile = objfso.opentextfile ("$winnt$.inf",1)

' Open %servername%.htm for writing 
wscript.echo "Opening " & strservername & ".htm for writing"
set objservernamehtm = objfso.opentextfile ("" & strservername & ".htm", 2)

' Read text file

' Read file line by line, if it finds the magic string dump that line to a string

Do While objfile.AtEndOfstream <> True
	strLine = objfile.readline
	strfindrightline = Left (strline, 9)
	if strfindrightline = "IPAddress" then strIPaddress = strLine
	if strfindrightline = "Subnet = " then strsubnetMask = strLine
	'wscript.echo strfindrightline
Loop


' Trim ipaddress and subnet mask to just the numbers

strIPaddress = Mid (strIPaddress, 14)
strsubnetMask = Mid (strsubnetMask, 11)

strIpaddress = replace (strIpaddress, chr(34), chr(0))
strsubnetmask = replace (strsubnetmask, chr(34), chr(0))

wscript.echo strsubnetmask

