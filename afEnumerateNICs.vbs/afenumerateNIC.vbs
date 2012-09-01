
' *********************************************************
' Name:		afenumeratnic.vbs
' Notes:	Enumerates through NICs using WMI
' Version:	1.0
' Date:		26/11/2001
' Author:	Adrian Farnell
' *********************************************************

Option Explicit

' ** Dim Objects, variables, constants **

Dim objWMI, objNICs,  objNICproperty, objAdapter, objWMIserver, objwshNetwork, objfso
Dim objConfigFile
Dim NIC, property, property2,ip
Dim strAdapterName, strNICindex, strServer, strline, strAdapterIP, contrim
Dim strNewIPAddress, strNewSubnet, intNumAddr, intRes,wsharguments
Dim arrIP, arrMask

' ** Set wshnetwork **
set objwshnetwork = createobject ("wscript.network")


' ** Get Arguments from commandline **
set wsharguments = wscript.arguments
if wsharguments.count <> 0 then
if wsharguments(0) = "/?" then subDisplayUsage
if wsharguments(0) = "-?" then subDisplayUsage
if wsharguments(0) = "/h" then subDisplayUsage
if wsharguments(0) = "/help" then subDisplayUsage
end if

' ** Set variables **
'strAdapterName = "MADGE"
'strNewIpaddress = "10.208.23.54"
'strNewSubnet = "255.255.240.0"
'strServer = objwshnetwork.computername
'wscript.echo " Searching for adapter type: " & strAdapterName & " ...."

' ** start filesystem object **
set objfso = createobject ("scripting.filesystemobject")

' ** Get file to process **
if objfso.fileexists ("ips.dat") = true then
	wscript.echo "Found ips.dat. processing...."
	set objConfigFile = objfso.opentextfile ("ips.dat", 1)
	'Collect Details
	do while objConfigFile.atendofstream = false
		strline = objConfigfile.readline
		if Left (strline,1 ) = "#" or Len (strline) = 0 then
			'ignoring line
		else
			wscript.echo strline
			if Left (strline,1) = "[" then
				
				strAdapterName = Mid (strLine, 2, Instr (strline, ":")-2)
				conTrim = Len (strline) - Instr(strline, ":") -1 
				strAdapterIP = Mid (strline, Instr (strline, ":")+1, conTrim)
				wscript.echo strAdapterName
				wscript.echo strAdapterIP
				if 
			else
			
			end if
			
		end if
	loop
	wscript.quit
else
	wscript.echo " Error! - File ips.dat does not exist"
	wscript.quit(1)
end if


' ** Start WMI Object **
set objWMI = getobject ("winmgmts:")

' ** Enumerate thru instances of win32_NetworkNic
set objNICs = objWMI.InstancesOf ("win32_NetworkAdapterConfiguration")

For each NIC in objNICs
	
	set objNICproperty = NIC.properties_
	
	For each property in ObjNICproperty
	
		if lcase (property.name) = "description" then
			if Instr (property.value,stradaptername) <> 0 then
			
					wscript.echo vbcrlf & " MADGE Card Found, Properties are: "
					wscript.echo " ---------------------------------"
					'set objNICproperty2 = NIC.properties_
					For each property2 in objNICproperty
						

						
						if Vartype (property2.value) = 8204 then
							wscript.echo "  " & property2.name & ": ** VarType is Array **"
						else
							wscript.echo "  " & property2.Name & ": " & property2.value
						end if			
						
						if lcase (property2.name) = "index" then
							strNICindex = property2.value
						end if
						
			
					Next

			else
				wscript.echo " Message, (" & property.value & ") adapter found."
			end if
		else
		end if

	Next
Next

set objwmi = nothing

Wscript.echo vbcrlf & " Found " & strAdaptername & " card on Index " & strNICindex

set ObjWMIserver = getobject ("WINMGMTS:{impersonationLevel=impersonate,(Security)}!\\" & strServer & "\ROOT\CIMV2")
set ObjAdapter = objWMIserver.get ("win32_NetworkAdapterConfiguration=" & strNICindex)

arrIP = objAdapter.IPaddress
arrMask = objAdapter.IPSubnet

'For each IP in arrip
For IP = 0 to ubound (arrip)

	if ip = strNewIPaddress then
		wscript.echo " Error! This address, " & strNewIPaddress & ", already exists on this server"
	else
		wscript.echo " Address not found on machine, installing...."
	end if
	
	intNumAddr = Ubound (arrIP) + 1
	Redim Preserve arrIP (intNumAddr)
	arrIP (intNumAddr) = strNewIpAddress
	arrMask = objAdapter.IPSubnet
	Redim Preserve arrMask (intNumAddr)
	arrMask (intNumAddr) = strNewSubnet
	intRes = objAdapter.EnableStatic (arrIP, arrMask)
	
	If intRes = 0 then 
		wscript.echo " IP address, " & arrIP(intNumAddr) & ". successfully added with subnet mask " & arrmask (intNumaddr) 
	else
		wscript.echo " Failed to add IP address. " & chr(13) & "Error: " & err.number
		wscript.echo " Description: " & err.description
		wscript.quit(1)
	End if
next

wscript.quit(0)

' ************************ funcs and subs ***************************

sub subDisplayUsage
	wscript.echo "Usage:" 
	wscript.quit(1)
end Sub