
strpath = "e:\mp3cd4"
Listfoldersandmp3s (strpath)


sub ListFoldersandmp3s(path)

     dim fs, folder, file, item, url

     set fs = CreateObject("Scripting.FileSystemObject")
     set folder = fs.GetFolder(path)
     'set objid3 = createobject ("id3.id3get")


     for each item in folder.SubFolders
       ListFoldersandmp3s(item.Path)
     next


     for each item in folder.Files
	strname = item.name
	wscript.echo item.path
     next





end sub

