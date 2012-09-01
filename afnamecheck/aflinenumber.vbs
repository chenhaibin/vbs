
set objfso = wscript.createobject ("scripting.filesystemobject")
set wsharguments = wscript.arguments

if wsharguments.count <> 1 then
	
	wscript.echo " File not specified"
	wscript.quit

else

	strnameoffile = wsharguments(0)
	
end if

set objfile = objfso.opentextfile (strnameoffile,1)

err.clear


if Err.number <> 0 then
	WScript.echo " Failed to open " & strnameoffile & ". File does not exist in same directory." & vbcrlf
	set objFso = nothing
	WScript.quit(1)
end if

linecount = 1

Do while objfile.atendofstream <> true

	wscript.echo linecount & ": " & objfile.readline
	linecount = linecount + 1
Loop
