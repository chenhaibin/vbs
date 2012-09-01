
strFolder = "E:\sharedFiles\Complete\Emulators\gb_gbc\gbc"

set oFso = createobject ("scripting.filesystemobject")
set oShell = createobject ("wscript.shell")

if oFso.folderExists ( strfolder ) = true then

	set oFolder = oFso.getfolder ( strfolder )

	set colFiles = oFolder.files

	for each file in colFiles

		'is it a zip file?
		'wscript.echo Mid (file.path, instrRev (file.path, "."))
		
		if Mid (file.path, instrRev (file.path, ".")) = ".zip" then
			wscript.echo Now & ": Skipping file " & file.name
		else
			strZipFileName = chr(34) & Left (file.path, InstrRev (file.path, ".")) & "zip" & chr(34)	
			if ofso.filEexists ( Replace (strZipFileName, chr(34), "")) = true then
				wscript.echo Now & ": Skipping file " & file.name
			else
				wscript.echo Now & ": Zipping file " & file.name
				'wscript.echo strZipFileName
				strCommand = "pkzip25 -add " & strZipFileName & " " & chr(34) & file.path & chr(34)
				'wscript.echo strCommand
				oShell.run strCommand, 0, true
			end if
			
			
		end if
	next
else

	wscript.echo "Error! Folder does not exist"	

end if

set oFso = nothing
set oShell = nothing