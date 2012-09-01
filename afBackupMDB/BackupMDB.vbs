
' ***************************************************
'
'  Name		: afBackupMDB.vbs 
'  Notes	: Backs up Pa2002.mdb to floppy
' 		  Uses pkzip.
'
'  Author	: Adrian Farnell
'  Date         : 6 Apr 2002
'  Version      : 1.0
'
' ***************************************************

on error resume next

strdrive = "e"
strFiletoBackup = "test.txt"

set objfso = createobject ("scripting.filesystemobject")
set objwshshell = createobject ("wscript.shell")

if objfso.DriveExists (strdrive) = true then

	
	set objCdrive = objfso.GetDrive (strdrive)
	
	if objCdrive.isready = true then
		
		wscript.echo " INFO: " & strdrive & ": Drive is ready"
		
		'zip up our file
		if objfso.fileexists ( strFiletoBackup ) = false then
			wscript.echo " ERR : File to backup does not exist"
			wscript.quit
		end if
		
		err.clear
		objwshshell.run "pkzip temp.zip " & strfiletoBackup, 2, true
		
		if err.number <> 0 then
		
			wscript.echo " ERR : zipping failed"
			
			if err.number = -2147024894 then wscript.echo " ERR : Can not find pkzip.exe program"
			
			wscript.quit
		
		end if
		
		'how big is our zip file
		if objfso.fileexists ("temp.zip") = true then
			wscript.echo " INFO: backup exists"
			set objZipFile = objfso.getfile ("temp.zip")
			strSizeOfZip = formatnumber (objZipFile.size/1024,1)
			strzipcreated = objzipfile.datelastmodified
			wscript.echo " INFO: backup is " & strsizeofzip & "kB in size "
			wscript.echo " INFO: backup created on " & strzipcreated
			set objZipfile = nothing
		else
			wscript.echo " ERR : backup file has not been created"
			wscript.quit
		end if
		
		
		'how much space is left?
		strFreespace = formatnumber (objcdrive.freespace/1024,1)
		wscript.echo " INFO: Space left on backup media is " & strfreespace & "kB"
		
		'can we fit backup on to drive
		if strfreespace < strsizeofzip then 
			wscript.echo " ERR : Not enough space on backup drive"
			wscript.quit
		else
			wscript.echo " INFO: Attempting Copy"
		end if
		
		
		'copy our file to a:
		err.clear 
		objfso.copyfile "temp.zip",  strdrive & ":\backupMDB.zip", true
		if err.number <> 0 then
			wscript.echo " ERR : File Copy failed"
			wscript.echo " ERR : Error Number " & err.number & " [ "& err.description & " ]"
			wscript.quit
		end if
		
		'check our file made it.
		if objfso.fileexists (strdrive & ":\backupMDB.zip") = true then
			wscript.echo " INFO: Backup file now on backup media"
			objfso.Deletefile "temp.zip", true
		else
			wscript.echo " ERR : Backup file does not exist on backup media"
		end if
		
		wscript.echo " INFO: Script ends"
		wscript.quit
		
	else
		wscript.echo " ERR : Drive is not ready"
		wscript.quit
	end if
	
else
	
	wscript.echo " ERR: Drive A: not ready"
	wscript.quit
	
end if

