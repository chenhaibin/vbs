
'==========================================================
'
' OS Version and Service Pack Detection Script
'
' Written by Adrian Farnell
' - 20/10/03
'
' Amendments
' - James Taylor 20/10/03 : Aditional code to run patchs
'
'==========================================================

'turn on proper errors
Option Explicit

'Dim objects
Dim objShell

'Dim strings
Dim strNTVer, strSPVer, strPath, intReturn

'start the objshell 
set objShell = createobject ("wscript.shell")

'allow errors to not stop script
on error resume next

' path of where script lives
strPath = Left (wscript.scriptfullname, InstrRev (wscript.scriptfullname, "\"))

'Read NT version
strNTVer = objshell.regread ("HKLM\Software\Microsoft\Windows NT\CurrentVersion\CurrentVersion")

'error checking
if err.number <> 0 then
	wscript.echo "Error: Can't determine OS version"
	wscript.quit
else

end if

'allow errors to stop script here
on error goto 0

'Can decide what to do based on os version here
Select case strNTVer

	Case "4.0"
		'Do wscript NT4 stuff here
		wscript.echo "This machine is Windows NT 4.0"
	Case "5.0"
		'Do wscript w2k stuff here
		wscript.echo "This machine is Windows 2000"
	Case "5.1"
		'Do wscript wxp stuff here
		wscript.echo "This machine is Windows XP"
	case "5.2" 
		'Do wscript 2k3 stuff here
		wscript.echo "This machine is Windows 2003"
	case Else
		'error
		wscript.echo "Unknown Operating System"
end select

'allow errors to not stop script
on error resume next
	
'Read Service Pack version
strSPver = objshell.regread ("HKLM\Software\Microsoft\Windows NT\CurrentVersion\CSDVersion")

'error checking
if err.number <> 0 then
	wscript.echo "Error: Can't determine SP version"
	wscript.quit
else

end if

'allow errors to stop script here
on error goto 0

Select case lcase (strSPver)

	case "service pack 1"
		'Do specific SP1 stuff here
		wscript.echo "This OS is on SP1"
	case "service pack 2"
		'Do specific SP2 stuff here
		wscript.echo "This OS is on SP2"
	case "service pack 3"
		'Do specific SP3 stuff here
		wscript.echo "This OS is on SP3"
	case "service pack 4"
		'Do specific SP4 stuff here
		wscript.echo "This OS is on SP4"
	case "service pack 5"
		'Do specific SP5 stuff here
		wscript.echo "This OS is on SP5"
	case "service pack 6"
		'Do specific SP6 stuff here
		wscript.echo "This OS is on SP6"
	case else
		'Do specific SP1 stuff here
		wscript.echo "Undetermined service pack level"
		
end select

'Now runs the appropriate patch

if strNTver = "4.0" and lcase (strSPver) = "service pack 6" then 
	wscript.echo "Applying hotfix for NT4 SP6"
	intReturn = objshell.run (strpath & "WindowsNT4-KB828035-x86-ENU.exe -q -z",1, TRUE) 
	wscript.echo "Script returned - " & intReturn

elseif strNTver = "5.0" and lcase (strSPver) = "service pack 2" then 
	wscript.echo "Applying hotfix for W2K SP2"
	intReturn = objshell.run (strpath & "Windows2000-KB828035-x86-ENU-sp2.exe -quiet -norestart",1, TRUE) 
	wscript.echo "Script returned - " & intReturn
	
elseif strNTver = "5.0" and lcase (strSPver) = "service pack 3" then 
	wscript.echo "Applying hotfix for W2K SP3"
	intReturn = objshell.run (strpath & "Windows2000-KB828035-x86-ENU-sp3-sp4.exe -quiet -norestart",1, TRUE) 
	wscript.echo "Script returned - " & intReturn
	
elseif strNTver = "5.0" and lcase (strSPver) = "service pack 4" then 
	wscript.echo "Applying hotfix for W2K SP4"
	intReturn = objshell.run (strpath & "Windows2000-KB828035-x86-ENU-sp3-sp4.exe -quiet -norestart",1, TRUE) 
	wscript.echo "Script returned - " & intReturn

else wscript.echo "Not at the correct service pack"

end if

