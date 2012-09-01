
' ******************************************************
'  Name: HxIpAddress.vbs - Adds Ip addresses to 2k/XP servers
'  			(would work on NT with SP4 + WMI)
'
'  Author: Adrian Farnell
'  Date:    25 Oct 2001
'  Version: Alpha
'
'
'  Notes: Rescripted from another script found on the 'net
'
' ******************************************************

Option Explicit

' ** Dim Objects, variables, etc **
Dim ObjWshNetwork, ObjServer, objNICadaptor
Dim constServer
Dim strIP, strTypeofIp
Dim objWshArguments
Dim strStartIP, strEndIP, StrSubnetMaskIP, StrDefGwayIP
Dim arrIPs
Dim bloop

' ** Start wshnetwork **
set objWshNetwork = createobject ("wscript.network")

' ** Start wsharguments, to take command line options **
set objWshArguments = wscript.arguments
if objWshArguments.count <> 4 then
	wscript.echo " Error! Number of Arguments specified is incorrect" & vbcrlf & _
			vbcrlf & " USAGE: cscript HxIPAddress.vbs <Starting IP> <End IP> <SubnetMask> <DefGway>" & _
			"" & vbcrlf & _
			"  <starting ip> - IP Address of the start of the range to add" & vbcrlf & _
			"  <end ip> -  IP address of the end of the range to add" & vbcrlf & _ 
			"  <SubnetMask> - Subnet Mask for all of the addresses" & vbcrlf & _
			"  <DefGway> - Gateway (if no gateway specify NONE)" & vbcrlf
			wscript.quit(1)
			
else

end if

' ** Start Constants **
constServer = objwshnetwork.computername

' ** Starting IP Address? **
strStartIP = objwsharguments(0)
call CheckIPs ( strStartIP, "IP" )

' ** End IP Address? **
strEndIP = objwsharguments(1)
call CheckIPs (strEndIP, "IP")

' ** Subnet Mask? **
strSubnetMaskIP = objwsharguments(2)
call CheckIPs (strSubnetMaskIP, "mask")

' ** Default Gateway? **
if lcase (objwsharguments(3)) = "none" then
	wscript.echo " No default gateway specified, continuing"
	strDefGwayIP = "none"
else
	strDefGwayIP = objwsharguments(3)
	call CheckIPs (strDefGwayIp, "IP")
end if

wscript.echo "strStartIP: " & strStartIp
wscript.echo "strEndIP: " & strEndIp
wscript.echo "strSubnetMaskIP: " & strSubnetMaskIP
wscript.echo "strGatewayIP: " & strDefGwayIp


' ** Start WMI Object **
set objServer = GetObject ("WINMGMTS:{ImpersonationLevel=Impersonate,(Security)}!\\" & constServer & "\ROOT\CIMV2")

' ** Object Hook to the NIC **
Set objNICadaptor = objServer.get ("win32_NetworkAdaptorConfiguration=0")

arrIPs = objNICadaptor.IPaddress

For bloop = 0 to ubound (arrIps,1 )

	wscript.echo arrIPs (bloop)
	wscript.quit
	
Next

' ** Clean Up objects **
set objServer = nothing
set objWshNetwork = nothing

wscript.quit(0)

' ****************** Functions\Subroutines *******************

Sub CheckIPs(strIP, strTypeofIP)

	Dim arrQuartet
	Dim Aloop,strUpperIP
	
	'Display IP
	'wscript.echo strIP

	'Split IP string into four quartets 
	arrQuartet = Split ( strIP, ".")

	'Loop all array values
	For aloop = 0 to Ubound (arrQuartet,1)
		
		if strTypeofIP = "mask" then strUpperIP = "255"
		if strTypeofIP = "IP" then strUpperIP = "254"
		
		
		'Check to make sure each value is between 0 and 255,254
		If arrQuartet(aloop) > strUpperIP or arrQuartet (aloop) < 0 then 
			wscript.echo " Error! IP Address is invalid, check you entered the correct detail" 
			wscript.quit(1)
		else
			'wscript.echo "IP Address " & strIP & " is valid"
		end if
		
	Next

End Sub

