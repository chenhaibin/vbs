
subEnumFolder ("e:\my Media\mp3")

sub subEnumFolder( strpath )

	set ofso = createobject ("scripting.filesystemobject")
	set ofolder = ofso.getfolder ( strpath )
	
	'Display Files in my folder
	wscript.echo vbcrlf& " ----------------------------------------- " & vbcrlf
	wscript.echo " + " & ofolder.path & vbcrlf
	For each item in ofolder.files
		wscript.echo "   * " & Replace (item.name, ".mp3", "")
	Next	
	
	' Display SubFolders
	For each item in ofolder.Subfolders
		'wscript.echo " + " & item.path
		subEnumFolder (item.path)
	Next
	
	


end sub