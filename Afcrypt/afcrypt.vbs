Option Explicit

on error resume next

dim objwshshell, struser, strhash, objwshnetwork
Dim arrstring , strtoencrypt, strlength, i, cell, strencrypted

set objwshshell = createobject ("wscript.shell")
set objwshnetwork = createobject ("wscript.network")

struser = "daed"

wscript.echo "User to test against: " & struser

strhash = objwshshell.regread ("HKLM\SOFTWARE\af\authentication\" & struser)

wscript.echo "Password Hash from Registry: " & strhash

aftoencrypt strhash

wscript.echo "Decrypted Password for " & struser & ": " & strencrypted

wscript.echo "Testing....."

wscript.echo "Current user is: " & objwshnetwork.username

wscript.echo "Trying to map drive to testshare...."

objwshnetwork.mapnetworkdrive "x:", "\\127.0.0.1\testshare", true, struser, strencrypted

if err.number <> 0 then
	wscript.echo "An error occured trying to map the share"
	wscript.echo Err.number & " : " & Err.description
else
	wscript.echo "Mapped share"
end if

'-----------------------------------------------------------------------


function aftoencrypt (strtoencrypt)

	'Dim arrstring , strtoencrypt, strlength, i, cell, strencrypted
	
	'strtoencrypt = "trustno1"
	
	'wscript.echo strtoencrypt
	
	strlength = Len(strtoencrypt)
	
	'wscript.echo strlength
	
	if strlength < 8 then 
		wscript.echo " Error! Length of password is to short"
		wscript.quit
	else

	end if

	redim arrstring(strlength)
	
	for i = 1 to strlength
		
		arrstring(i) = mid (strtoencrypt, i, 1)	
	
	next
	
	for i = strlength to 1 step -1
		
		strencrypted = strencrypted & arrstring(i)
	
	next
	
	'wscript.echo strencrypted

end function