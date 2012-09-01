
set oRegSec = createobject ("ntsec.regsecurity")

set oRegProp = createobject ("ntsec.propagation")

on error resume next

if oRegSec.Lookup ("HKLM\System\CurrentControlSet\Services\EventLog\Security") <> true then

	wscript.echo "False"
	wscript.quit (1)
	
else
	
	wscript.echo "True"
	
end if 

on error goto 0

strPerms = oRegSec.GetFormattedPerms (1)

strPerms = Replace (strPerms, ",", vbcrlf)

wscript.echo strPerms
