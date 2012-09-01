
On error Resume Next

Dim arrExclude

'folder to generate sums 
strRootFolder = "c:\system volume information"

'exclude folders here


'basis of name of log file to be generated.
strDateString = Replace (Replace (Replace (Now, "/", ""), ":", ""), " ", "_")

'start the filesystem and shell objects
set oFSO = createobject ("scripting.filesystemobject")
set oShell = createobject ("wscript.shell")

'start enumerating folders
EnumFolder strRootFolder

sub enumFolder (strFolder)
	
	if err.number <> 0 then 
		wscript.echo "Error! " & err.number & " - " & err.description & "(" & strFolder & ")"
	end if
	
	err.clear

	set oEnumFolder = ofso.getfolder (strFolder)
	
	if oEnumfolder.files.count = 0 then
		
		set ofile = ofso.opentextfile (strDateString & ".info",8) 

		ofile.write "no files in " & oEnumFolder.path & vbcrlf & vbcrlf

		ofile.close
		
		'EnumFolder item.path

		set ofile = nothing

	else 
	
	i = 0

	for each item in oEnumfolder.files
		
		i = i + 1
		
		oshell.run "shabatch.cmd " & item.path & " " & strDateString,0,true

		if err.number <> 0 then
			wscript.echo "Error! " & Err.number & " " & Err.description
		else

		end if				
	
		
	
	next 
	
	if oEnumfolder.files.count -1 = i then

	else
		'wscript.echo "Error! " & oEnumFolder.files.count - i & " files locked"
	end if

	end if
	
	
	
	for each item in oEnumfolder.SubFolders
		
		set ofile = ofso.opentextfile (strDateString & ".info",8) 

		ofile.write  vbcrlf & "Enumerating " & item.path & vbcrlf

		ofile.close
		
		EnumFolder item.path
		
				
		set ofile = nothing

	next



end sub

