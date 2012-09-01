set objfso = wscript.createobject ("scripting.filesystemobject")

strother = "f:\winnt\system32\msvcrt.dll"	

set cresult = objfso.getfile(strother)
	wscript.echo strother
	'if cresult <> 1 then 
	'	wscript.echo " Error! File " & strother & " Can not be loaded."
	'else
		strsizeoffile = cresult.size
		strdatestamponfile = cresult.datecreated
		wscript.echo strsizeoffile
		wscript.echo strdatestamponfile