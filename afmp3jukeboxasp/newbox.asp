<HTML>

<HEAD></HEAD>

<b>The mp3 jukebox v2</b>

<hr>

<font face="tahoma" size="-1">

<%

Response.write "<form action=""results.asp"" method=""POST"">" & vbcrlf
Response.write "<TABLE BORDER=""1"">" & vbcrlf & "<TR>" & vbcrlf 
strpath = "i:\mp3s\processed\cd1" 
Listfoldersandmp3s (strpath)

%>

<%

sub ListFoldersandmp3s(path)

     dim fs, folder, file, item, url

     set fs = CreateObject("Scripting.FileSystemObject")
     set folder = fs.GetFolder(path)
     set objid3 = createobject ("id3.id3get")

	
	'Write Folder Names to webpage
	if folder.files.count = 0 then
	Response.write"<td><font face=""tahoma"" size=""-1"">&nbsp;" & folder.name & "<br></td><td></td></tr>" & vbcrlf
	else
	Response.write"<td><font face=""tahoma"" size=""-1"">&nbsp;" & folder.name & "<br>" & vbcrlf
	end if
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
			
			Response.write "<li><input type=""checkbox"" name=""selection"" value=""" & item.path & """>"
			response.write "<a href=""" & item.path & """>" & objid3.gettrackname & " : " & objid3.gettrackartist & "</a></li>" & vbcrlf
			'wscript.echo "		" & objid3.gettrackname & " : " & objid3.gettrackartist
			
	
		else

			'wscript.echo item.name & "		No Track Data"
			response.write "No id3 tag<br>" & vbcrlf
		end if
	
	else

	end if
	
	if right (item.name, 4) = ".jpg" then
		strjpg = item.path
	else
	end if
	
	if right (item.name, 4) = ".pls" then
		strplsname = item.name
		strplspath = item.path
	else
	end if
	
     next
	


	if strjpg <> "" then
	response.write "</td><td><center>&nbsp<img src=""" & strjpg & """><br><a href=""" & strplspath & """>"& strplsname & "</A></center></td></tr>"
	else 
	response.write "</td><td>&nbsp</td></tr>"
	end if


end sub

%>




</font>

</HTML>