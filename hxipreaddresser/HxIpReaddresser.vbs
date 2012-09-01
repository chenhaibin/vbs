' File Name	: HxIpReAddresser.vbs - Readdresses IP nos. On a specific card
' Author		: Adrian Farnell
' Version		: 1.0
' DateAmended	: 13.09.2000

Dim arripaddresses()

' Start filesystemobject, wsharguments and wshshell objects------------------------------

Set objfso = wscript.createobject ("scripting.filesystemobject")
Set objwshshell = wscript.createobject ("wscript.shell")
set objwshnetwork = wscript.createobject ("wscript.network")
Set objlogfile = objfso.getfile ("log.txt")
Set objlogfiletextstream = objlogfile.openastextstream(8)

objlogfiletextstream.write "HxIpReAddresser ran On: "
objlogfiletextstream.write Date & " @ " & Time & vbcrlf
wscript.echo ""
If wscript.arguments.count <> 5 then 

	' Display usage
	wscript.echo vbcrlf
	wscript.echo "HxIpaddresser assigns a continuous specified range of IP addresses"
	wscript.echo "To a specific Network card (On the ecommerce platform, NT4)"
	wscript.echo ""
	wscript.echo "USAGE: cscript HxIpReAddresser.vbs /Start:[IP_address] /finish:[IP_address] /subnet:[Subnet_mask] /gateway:[default_gway] /NIC:[type of card]"
	wscript.echo ""
	
	wscript.echo ""
	wscript.echo "	/Start:[IP_address]	the Start of the block of " 
	wscript.echo "				IP addresses To add"
	wscript.echo ""
	wscript.echo "	/finish:[IP_address]	the end of the block of " 
	wscript.echo "				ipaddresses To add"
	wscript.echo ""
	wscript.echo "	/subnet:[Subnet_mask]	Script assumes the same subnet wscript.echo"
	wscript.echo "				mask should be used for each IP address"
	wscript.echo ""
	wscript.echo "	/gateway:[default_gway]	Script assumes the same default"
	wscript.echo "				gateway for each IP address,use NONE If no gateway is used"
	wscript.echo ""
	wscript.echo "	/NIC:[type of NIC]	this switch can be used To specify"
	wscript.echo "				which card To reipaddress"
	wscript.echo "				(Standard in ecommerce card is"
	wscript.echo "				E100Bx where x is the no of the card"
	wscript.quit
Else

	'Assign arguments To variables---------------------------------------------------------
	
	for i=0 To 4	
		strtest = wscript.arguments(i)
		If left(strtest,4) = "/sta" then strstart = strtest
		If left(strtest,4) = "/fin" then strfinish = strtest
		If left(strtest,4) = "/sub" then strsubnet = strtest
		If left(strtest,4) = "/gat" then strgateway = strtest
		If left(strtest,4) = "/NIC" then strNIC = strtest
	next

	objlogfiletextstream.write "HxIpReAddresser.vbs: " & vbcrlf
	objlogfiletextstream.write "Parameters entered:-"& vbcrlf
	objlogfiletextstream.write	"strstart: " & strstart & vbcrlf
	objlogfiletextstream.write "strfinish: " & strfinish & vbcrlf
	objlogfiletextstream.write "strsubnet: " & strsubnet & vbcrlf
	objlogfiletextstream.write "strgateway: " & strgateway & vbcrlf
	objlogfiletextstream.write "strNIC: " & strNIC & vbcrlf
	objlogfiletextstream.write ""

end If

'process strstart and strfinish To make an array of servers To add-----------------------

'strnic process
strnic = Mid(strnic,6)

'strsubnet process
strsubnet = Mid (strsubnet,9)

'strgateway process
strgateway = Mid (strgateway,10)


'Check To see If NIC To alter exists-----------------------------------------------------



'find the last IP address quartet and put it into strIPbegin-----------------------------

StrIPbegin = Right (strstart, 3)
If left (StrIPbegin,1) = "." then StrIPbegin = right(StrIPbegin,2)
objlogfiletextstream.write "strIPbegin: " & strIPbegin & vbcrlf

'find the last IP address quartet and put it into strIPend-------------------------------

StrIPend = Right (strfinish, 3)
If left (strIPend,1) = "." then StrIPend= Right (strIPend,2)
objlogfiletextstream.write "strIPend: " & strIPend & vbcrlf

'find the class C address and put in strclassC-------------------------------------------

strclassC = Mid (strstart, 8)
StrLenminus2 = Len(strclassC) - 2
strclassC = Mid (strclassC, 1, StrLenminus2)
strlenminus1 = Len(strclassC) - 1
If Right(strclassC, 1)="." then strclassC= Mid (strclassC,1,strlenminus1) Else strclassC= Mid (strclassC,1,strlenminus1-1)

objlogfiletextstream.write "strclacssc: " & strclassC & vbcrlf

'strdifference contains a value for the number of ipaddresses to add---------------------

Strdifference = strIPend - strIPbegin
objlogfiletextstream.write "strdifference: " & strdifference & vbcrlf

'redimension the array to contain the correct no of elements-----------------------------

redim arripaddresses(strdifference)

'input the ipaddresses into the array----------------------------------------------------

for number = 0 To Strdifference

	straddress = strclassC & "." & number + strIPbegin
	arripaddresses(number) = straddress
	objlogfiletextstream.write "IP address: " & arripaddresses(number) & vbcrlf
next

'attempt to write into the registry (failed, vbscript does not support multi_sz)---------

'objwshshell.regwrite "HKEY_LOCAL_MACHINE\Software\Halifax\Test", arripaddresses

'Dump key to %computername%.txt using regdmp.exe-----------------------------------------

objwshshell.run "hxregdump.cmd " & strnic, true

wscript.echo "sleeping for 10 sec"
'objwshshell.run "sleep 10",true
for wait = 1 To 10000000
next 
'wscript.sleep 10000
wscript.echo ""

'get File for reading--------------------------------------------------------------------
set objfile = objfso.getfile ( objwshnetwork.computername & ".txt")

'backup old %computername%.txt which contains original IP addresses----------------------

If objfso.fileExists( objwshnetwork.computername & ".txt") then
	objfile.copy objwshnetwork.computername & ".old", true
Else
	wscript "pausing.."	
	for wait = 1 To 100000
	next
	objfile.copy objwshnetwork.computername & ".old", true
end if
	

'open File as textstream-----------------------------------------------------------------
Set objfiletextstream = objfile.openastextstream(2,false)

'Do While objfiletextstream.AtEndOfstream <> true
'	
'
'	strline = objfiletextstream.readline
'	strfindrightline = Left (strline, 17)
'	
'	if strfindrightline = "        IPAddress" then 
'		strwrite = "no"
'	Else
'	end if	
'	
'	if strfindrightline = "        TCPAllowe" then 
'		strwrite = "yes"
'	Else
'	end If
'
'	wscript.echo strline & " : " & strwrite
'	If strwrite="yes" then strnewfile = strnewfile & vbcrlf & strline 
'
'	
'Loop
'
'wscript.echo strnewfile

'Build Regini Script from scratch.....

'IPAddress part---------------------------------------------------------------------
Strnewreginiscript = "HKEY_LOCAL_MACHINE\System\ControlSet001\Services\" & strNIC & "\Parameters\Tcpip" & vbcrlf
Strnewreginiscript = Strnewreginiscript & "    IPAddress = REG_MULTI_SZ """ & arripaddresses(0) & """  \"& vbcrlf 

for ips = 1 To strdifference
	
	If ips = strdifference then
		Strnewreginiscript = Strnewreginiscript & "                             """ & arripaddresses(ips) & """ " & vbcrlf	
	Else
		Strnewreginiscript = Strnewreginiscript & "                             """ & arripaddresses(ips) & """  \" & vbcrlf
	end If
	
next

'subnetMask part--------------------------------------------------------------------

Strnewreginiscript = Strnewreginiscript & "    SubnetMask = REG_MULTI_SZ """ & strsubnet & """  \" & vbcrlf


for ips = 1 To strdifference

	If ips = strdifference then
		Strnewreginiscript = Strnewreginiscript & "                              """ & strsubnet & """ " & vbcrlf	
	Else
		Strnewreginiscript = Strnewreginiscript & "                              """ & strsubnet & """  \" & vbcrlf
	end If
	
next

'strgateway part---------------------------------------------------------------------

If strgateway = "NONE" then
	Strnewreginiscript = Strnewreginiscript & "    DefaultGateway = REG_MULTI_SZ " & vbcrlf
Else
	Strnewreginiscript = Strnewreginiscript & "    DefaultGateway = REG_MULTI_SZ """ & strgateway & """ " & vbcrlf

end if

objlogfiletextstream.write "Strnewreginiscript: " & vbcrlf & Strnewreginiscript 

'save File and close it--------------------------------------------------------------

objfiletextstream.write Strnewreginiscript
objfiletextstream.close
objlogfiletextstream.close
objwshshell.run "Regini.exe " & objwshnetwork.computername & ".txt"

wscript.echo "-------------------------------------------------"
wscript.echo "Please reboot Computer for changes To take effect"
wscript.echo "-------------------------------------------------"