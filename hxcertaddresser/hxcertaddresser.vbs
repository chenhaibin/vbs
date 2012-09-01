'Dimension Arrays
Dim Arrline,Arrtable()

'start the filesystemobject, and open text file
set objfso = createobject ("scripting.filesystemobject")
set objfile = objfso.opentextfile ("ecomipmap.txt", 1)

'get name of server
set wshNetwork = Wscript.createobject("Wscript.network")
'strcomputername = wshNetwork.computername

strcomputername="ntw067804"

'make servername all uppercase
strcomputername = Ucase(strcomputername)

wscript.echo strcomputername

Z=0


'read each line of file

Do While objfile.AtEndOfstream <> True
	
	
	strLine = objfile.readline
	Wscript.echo strLine 
	if left(strLine, 8)=left(strcomputername, 8) then
		arrline=split (strLine, ",")
		
	wscript.echo Arrline(0) & Arrline (1) & Arrline(2)
	Z=Z+1
	
	for x=0 to 2
		'arrtable(x,Z) = Arrline (x)
		Wscript.echo x & "," & z	
	next
	
	else
	
	end if



Loop 




