Listfoldersandmp3s ("i:\mp3s\processed\cd1")

sub ListFoldersandmp3s(path)

     dim fs, folder, file, item, url

     set fs = CreateObject("Scripting.FileSystemObject")
     set folder = fs.GetFolder(path)
     set objid3 = createobject ("id3.id3get")
    
      wscript.echo(folder.name)

     'Display sub folders.

     for each item in folder.SubFolders
       ListFoldersandmp3s(item.Path)
     next

     'Display files.

     for each item in folder.Files
        'wscript.echo ( "	" & item.name)
	strname = item.name
	if right (strname,4) = ".mp3" then
	
		objid3.settargetfile(item.path)
		if objid3.hasid3 then
			
			wscript.echo "		" & objid3.gettrackname & " : " & objid3.gettrackartist
	
		else

			wscript.echo item.name & "		No Track Data"

		end if
	
	else
	end if
	
     next

end sub
