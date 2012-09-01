
' backup.vbs 

funcMessage "Backup Started"
dim strBackuplog, strBackupdir

' ** first open text file **
set objfso = createobject ("scripting.filesystemobject")


if objfso.fileexists ("backup.ini") = true then 
	set objBackupFile = objfso.opentextfile ("backup.ini",1)
	funcMessage "Backup.ini found"
else
	funcMessage "Error! Backup.ini file does not exist"
	wscript.quit
end if

do while objbackupFile.Atendofstream = false

	strLine = objbackupFile.readline
	if Left (strline,1)= "#" or Len(strline) = 0 then
		'skip line
	else

		' ** set backupdir **
		if instr (strLine, "BackupDir=") <> 0 then
			'funcMessage "Found BackupDir"
			funcSetBackupDir Mid (strLine, Instr(strline, "=")+1)
		else
			
		end if
		
		if instr (strline, "[") <> 0 then
			
			funcMessage "Backup up of " & strline & " started"
			strfolder = replace (strline, "[", "")
			strfolder = replace (strfolder, "]", "")
			funcCreateBackupSubFolder strfolder
			
			strBackupfolder = strbackupdir & "\" & strfolder
			
		end if
		
		if instr (strline, "files=") <> 0 and instr (strline, "*") = 0 then
			'wscript.echo "specified files to " & strbackupfolder
			strline = replace (strline, "files=", "")
			funcCopySpecifiedFile strline
		end if
		
		if instr (strline, "files=") <> 0 and instr (strline, "*") <> 0 then
			'wscript.echo "wildcard files to " & strbackupfolder
			strline = replace (strline, "files=", "")
			funcCopyWildCardFiles strline
		end if
		
		if instr (strline, "dir=") <> 0 then
			'wscript.echo "copy directory to " & strbackupfolder
			strline = replace (strline, "dir=", "")
			funcCopyfolder strline
		end if
		
		
	end if
	
	
loop

wscript.echo strBackuplog

set objfso = nothing



' ************* functions **************

function funcMessage (strMessage)

	strdate = Now
	'wscript.echo " " & strdate & " " & strmessage
	strBackuplog = strBackuplog & vbcrlf & strdate & " " & strMessage

end function

function funcSetBackupDir (strDir)

	'set global backupdir
	strBackupDir = strDir
	
	'check if directory is there
	if objfso.folderexists (strdir) = false then
		objfso.createfolder (strDir)
		funcMessage strDir & " created"
		
	else
		'funcmessage strDir & " exists"
	end if
	

end function

function funcCopyfolder (strDir)

	if objfso.folderexists (strdir) = true then 
	
		objfso.copyfolder strDir, strBackupDir & "\" & strfolder
	
		'compare original folder and backup folder
		set objOrigfolder = objfso.getfolder (strDir)
		set objNewFolder = objfso.getfolder (strBackupDir & "\" & strfolder)
		
		if objOrigfolder.size = objNewFolder.size then
			funcMessage "Backup of "& strfolder & " Complete and OK"
		else
			funcMessage "Error! Backup of " & strfolder & " inComplete, folder size does not agree"
		end if
	else
		funcMessage "Error! Folder to backup (" & strdir & ") Does not exist"
	end if
end function

function funcCopyWildcardFiles (strfiles)

	strDirectory = Left (strfiles, InstrRev (strfiles, "\")-1)
	strFiles = Mid (strfiles, InstrRev(strfiles, "\")+1)
	
	strFiles = Replace (strfiles, "*", "")
	
	set objDirectory = objfso.getfolder (strDirectory)
	set colfiles = objDirectory.files
	
	for each file in colfiles
	
		if Instr (file.name, strfiles) <> 0 then
			'here files match our wilcard	
			objfso.copyfile file.path, strBackupDir & "\" & strFolder & "\" & file.name
			
			'wscript.echo file.path
			set objOrigFile = objfso.getfile (file.path)
			set objNewFile = objfso.getfile (strBackupDir & "\" & strFolder & "\" & file.name)

			if objOrigFile.size = objNewfile.size then 
				'file is same size
				'funcMessage "File " & Mid (strfile, InstrRev (strfile, "\")+1) & " backup matches original (size)"
			else
				'file sizes do not match
				funcMessage "Error! File " & file.name & " sizes do not match"
			end if
			
			if objOrigfile.datelastmodified = objNewfile.dateLastmodified then
				' file dates match
				'funcMessage "File " & Mid (strfile, InstrRev (strfile, "\")+1) & " backup matches original (date)"
			else
				'file dates do not match
				funcMessage "Error! File " & file.name & " Dates do not match"
			end if			
			
		else
	
		end if
	
	next

end function

function funcCopySpecifiedFile (strfile)

	if objfso.fileexists (strfile) = true then
		
		'wscript.echo strbackupdir & "\" &strfolder
		objfso.copyfile strfile, strBackupDir & "\" & strFolder & "\" & Mid (strfile, InstrRev (strfile, "\")+1)
		
		'compare original file and backup file
		set objOrigfile = objfso.getfile (strfile)
		set objNewFile = objfso.getfile (strBackupDir & "\" & strfolder & "\" & Mid (strfile, InstrRev (strfile, "\")+1))
		
		'compare by date and size
		if objOrigFile.size = objNewfile.size then 
			'file is same size
			'funcMessage "File " & Mid (strfile, InstrRev (strfile, "\")+1) & " backup matches original (size)"
		else
			'file sizes do not match
			funcMessage "Error! File " & Mid (strfile, InstrRev (strfile, "\")+1) & " sizes do not match"
		end if
		
		if objOrigfile.datelastmodified = objNewfile.dateLastmodified then
			' file dates match
			'funcMessage "File " & Mid (strfile, InstrRev (strfile, "\")+1) & " backup matches original (date)"
		else
			'file dates do not match
			funcMessage "Error! File " & Mid (strfile, InstrRev (strfile, "\")+1) & " Dates do not match"
		end if
		
	else
		funcMessage "Error! File to backup (" & strfile & ") Does not exist"
	end if

end function

function funcCreateBackupSubfolder (strDir)
	
	if objfso.folderexists ( strBackupDir & "\" & strDir ) = false then
		objfso.createfolder ( strBackupDir & "\" & strDir )
		funcmessage strBackupDir & "\" & strDir & " created"
	else
		'funcmessage strBackupDir & "\" & strDir & " exists"
	end if

end function