
' *****************************************
' Name: afTestPassword.vbs - Tests user password
' Author: Adrian Farnell
' Date: 23.01.02
' Version: pre alpha
' *****************************************

ReDim ArrDrivesMapped (0)

strUser = "daed"
strPassword = "potato"
strComputer = "95l085673"

' hook into wshnetwork
set objwshnetwork = createobject ("wscript.network")
set objdrives = objwshnetwork.enumnetworkdrives

'check drives m thru z to make sure they're not used.

for j = 0 to objDrives.count - 1 step 2
	'wscript.echo "Drive " & objDrives.Item(j) & " = " & objDrives.Item(j+1)
	if objDrives.count <> 0 then
		'wscript.echo "Drives Mapped"
		Redim Preserve arrDrivesMapped (Ubound (arrDrivesMapped)+1)
		arrDrivesMapped (Ubound (arrDrivesMapped)) = objDrives.Item(j)
	else
		'wscript.echo "No drives Mapped"
	end if
next

if arrDrivesMapped(0) <> "" then
	strCanUseLetter = "M:"
else
	i = 76
	strCanUseLetter = chr(i) & ":"
	Do while strCanNot = 1
	for x = 1 to Ubound (arrDrivesMapped)
		if strCanUseLetter = arrDrivesMapped(x) then
			strCanNot = 1
			If x = Ubound (arrDrivesMapped) then
				wscript.echo "No free drive letters, can not test"
				wscript.quit
			else
			end if
		else
			If x = Ubound (arrDrivesMapped) then
				strCanNot = 0
			else
			end if
		end if
	next
	
	i = i + 1
	
	Loop
end if

' try to map drive using user credentials given.
objwshnetwork.mapnetworkdrive strCanUseLetter, "\\" & strComputer & "\admin$", 0, strUser, strPassword

' remove network drive
objwshnetwork.removenetworkdrive
