dim arrips, arrdns

strIPAddress = "169.254.1.1"
strServer= "95l085673"

strAddSubnet = "255.255.255.0"
strAddIP = "169.254.1.2"
strAddGateway = "169.254.1.254"
strAddDNS = "169.254.1.252"

set objWMI = getobject ("winmgmts:")
set objNICs = objWMI.InstancesOf ("win32_NetworkAdapterConfiguration")

For each NIC in objNICs

	set objNICproperty = NIC.properties_
	For each property in objNICproperty
	
		'get ips
		
		if lcase (property.name) = "index" then strIndex = property.value
		
		if lcase (property.name) = "ipaddress" then
			
			if vartype (property.value) = 8204  then
				'ipaddresses are in an array here
				
				arrips = property.value
				for ip = 0 to ubound (arrips)
					if trim (arrips (ip)) = strIPaddress then
						'wscript.echo " Found IP address " & arrips(ip) & " on NIC " & strIndex
						
						'call subroutine to add/replace ip addresses
						subAddIp strAddIP, strAddSubnet, strIndex
						
						'do we need a dns search order
						if Instr (strAddDNS, ".") > 1 then
							subAddDNS strAddDNS, strIndex
						end if
					
						'do we need a gateway
						if Instr (strAddGateway, ".") > 1 then
							subAddGateway strAddGateway, strIndex
						end if
					end if
				next 
			
			else
			'NIC has no assigned ip addresses
			end if
		end if
		
					 
			
		
	next

next

set objWMI = nothing
set objNICs = nothing

sub subAddDNS (strDNS, strNIC)

	set ObjWMIserver = getobject ("WINMGMTS:{impersonationLevel=impersonate,(Security)}!\\" & strServer & "\ROOT\CIMV2")
	set ObjAdapter = objWMIserver.get ("win32_NetworkAdapterConfiguration=" & strNIC)

	arrDNS = objAdapter.DNSserverSearchOrder
	
	if vartype (arrDNS) = 8204 then
		'there is DNS, so enumerate thru entries and then set
		
		for DNS = 0 to Ubound (arrDNS)
			if arrDNS (DNS) = strAddDNS then
				strDNSfound = "y"
			else
				if strDNSfound = "y" then
				else
					strDNSfound = "n"
				end if
			end if
		next
		
		if strDNSfound = "n" then
			if ubound (arrDNS) = 1 then
				wscript.echo " Error! Too many DNS server entries!"
			else
				intNumAddr = Ubound (arrDNS) + 1
				Redim Preserve arrDNS (intNumAddr)
				arrDNS (intNumAddr) = strAddDNS	
				intRes = objAdapter.setDNSServerSearchOrder (arrDNS)
				wscript.echo " Set DNSserverSearchOrder to: " & strAddDNS
			end if
		else
			wscript.echo " Error! Found DNS server already exists"
		end if
	else
		'no DNS set here
		IntNumAddr = 0
		Dim arrDNS2(0)
		arrDNS2(0) = strAddDNS
		
		intRes = objAdapter.setDNSServerSearchOrder (arrDNS2)
		wscript.echo " Set DNSServerSearchOrder to: " & arrDNS2(0)
	end if
	
	set objwmiserver = nothing
	set objadapter = nothing

end sub

sub subADDgateway (strGW, strNIC)

	set ObjWMIserver = getobject ("WINMGMTS:{impersonationLevel=impersonate,(Security)}!\\" & strServer & "\ROOT\CIMV2")
	set ObjAdapter = objWMIserver.get ("win32_NetworkAdapterConfiguration=" & strNIC)
	
	arrGateway = objadapter.DefaultIPGateway
	
	
 	if vartype (arrGateway) = 8204 then
 		'is a gateway
 		wscript.echo "There is already a gateway on this NIC."
 	else
 		'no gateway
 		IntNumAddr = 0
 		Dim arrGateway2(0)
 		arrGateway2(0) = strAddGateway
 		
 		intres = objAdapter.setGateways (arrGateway2)
 		wscript.echo " Set Gateway to: " & strAddGateway
 	end if 
 
 
end sub

sub subAddIP (strAddIP, strAddSubnet, strNIC)

	set ObjWMIserver = getobject ("WINMGMTS:{impersonationLevel=impersonate,(Security)}!\\" & strServer & "\ROOT\CIMV2")
	set ObjAdapter = objWMIserver.get ("win32_NetworkAdapterConfiguration=" & strNIC)

	arrIP = objAdapter.IPaddress
	arrMask = objAdapter.IPSubnet

	' ** IP Address, subnet mask section **
	For IP = 0 to ubound (arrip)	
		
		if arrip(ip) = strAddIp then
			strIPfound = "y"
		else
			if strIPfound = "y" then 
			else
				strIPfound = "n"
			end if

		end if
		'wscript.echo arrip(ip) & " " & arrAddIP & strIPfound
		
			
	next
	
	if strIPfound = "n" then
		wscript.echo " Applying IP address change"
		intNumAddr = Ubound (arrIP) + 1
		Redim Preserve arrIP (intNumAddr)
		arrIP (intNumAddr) = strAddIp	
		arrMask = objAdapter.IPSubnet
		Redim Preserve arrMask (intNumAddr)
		arrMask (intNumAddr) = strAddSubnet	
		intRes = objAdapter.EnableStatic (arrIP, arrMask)

		If intRes = 0 then 
			wscript.echo " IP address, " & strAddIp & ". successfully added with subnet mask " & arrmask (intNumaddr) 
		else
			wscript.echo " Error! Failed to add IP address. " & chr(13) & " Error No: " & err.number
			wscript.echo " Error Description: " & err.description
			wscript.quit(1)
		End if	
	else 
		wscript.echo " Error! IP address Exists on NIC"
	end if
	
	set objWMIserver = nothing
	set objAdapter = nothing

end sub
