


' dim all arrays, variables, constants
dim arrips, arrdns, arrgateway, arrGateway2, arrIP, arrMask

dim objwshnetwork, objwshshell, objwshenvironment, objwmi, objNICs, objNICproperty
dim strIPAddress, strServer, strAddSubnet, strAddGateway, strAddDNS1, strAddDNS2, strAddIP
dim strIndex
dim NIC, property, ip

dim objwmiserver, objadapter
dim strDNS, strNIC, strDNSfound
dim IntNumAddr, IntRes



' start wshnetwork
set objwshnetwork = createobject ("wscript.network")

' ip address used to find which card to change ip address on
strIPAddress = "10.0.0.1"

' get server name
strServer= objwshnetwork.computername

'get environment variables
set objwshshell = createobject ("wscript.shell")
set objwshenvironment = objwshshell.environment

strAddSubnet = objwshenvironment.item ("svrSubMask")
strAddGateway = objwshenvironment.item ("svrGateway")
strAddDNS1 = objwshenvironment.item ("DNSPri")
strAddDNS2 = objwshenvironment.item ("DNSSec")
strAddIP = objwshenvironment.item ("svrIP")


' ** For/next loop to find 10.0.0.1 card **
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
					'Having found the card, call various subroutines to replace ips, gateways and dns
					if trim (arrips (ip)) = strIPaddress then
						wscript.echo " Found IP address " & arrips(ip) & " on NIC " & strIndex
						
						'call subroutine to add/replace ip addresses
						subAddIp strAddIP, strAddSubnet, strIndex
						
						'do we need a dns search order
						if Instr (strAddDNS1, ".") > 1 or Instr (strAddDNS2, ".") > 1 then
							if len (strAddDNS1) <> 0 then
								subAddDNS strAddDNS1, strIndex
							end if
							if len (strAddDNS2) <> 0 then
								subAddDNS strAddDNS2, strIndex
							end if
						end if
					
						'do we need a gateway
						if Instr (strAddGateway, ".") > 1 then
							if len (straddGateway) <> 0 then
								subAddGateway strAddGateway, strIndex
							end if
						end if
					else
						
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

' **************************** subs and funcs ************************************

sub subAddDNS (strDNS, strNIC)

	'hook into WMI and NIC
	set ObjWMIserver = getobject ("WINMGMTS:{impersonationLevel=impersonate,(Security)}!\\" & strServer & "\ROOT\CIMV2")
	set ObjAdapter = objWMIserver.get ("win32_NetworkAdapterConfiguration=" & strNIC)
	
	'setup array of dns servers
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
		arrDNS2(0) = strDNS
		
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
	
	intNumAddr = Ubound (arrIP)
	redim preserve arrIP (intNumaddr)
	arrIP (intNumAddr) = strAddIP
	
	redim preserve arrMask (intNumaddr)
	arrMask (intNumAddr) = strAddSubnet
	
	intRes = objAdapter.EnableStatic (arrIP, arrMask)
	
	if IntRes = 0 then
		wscript.echo " IP address " & arrIP (intNumAddr) & " successfully added with Subnet Mask " & arrMask (intNumAddr)
	else
		wscript.echo " Error! Failed to add IP address. " & chr(13) & " Error No: " & err.number
		wscript.echo " Error Description: " & err.description
		wscript.quit(1)
	
	end if
	set objWMIserver = nothing
	set objAdapter = nothing

end sub
