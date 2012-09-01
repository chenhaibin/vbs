<HTML>
<HEAD></HEAD>

<BODY>
<%
' Constants to be used later
Dim strOutput
commaspace = ", "
carriageReturn = vbCrLf
%>


<form method="POST" action="<%= request.servervariables("script_name") %>">
    <input type="checkbox" name="opt" value="option1" >Option 1 </p>
    <input type="checkbox" name="opt" value="option2" >Option 2</p>
    <input type="checkbox" name="opt" value="option3" >Option 3</p>
    <input type="checkbox" name="opt" value="option4" >Option 4</p>
<input type="checkbox" name="opt" value="option5" >Option 5</p>
<input type="checkbox" name="opt" value="option6" >Option 6</p>
<input type="checkbox" name="opt" value="option7" >Option 7</p>
<input type="checkbox" name="opt" value="new option" >Option 8</p>
<input type="textarea" name="name_of_file" cols="15">

    <input type="submit" name="B1" value="Submit"><input
    type="reset" name="B2" value="Reset"></p>
    <p>&nbsp;</p>
</form>

<%
'User input into variables?
opt=request.form("opt")
name_of_file=request.form("name_of_file")
'opt2=request.form("opt2")
'opt3=request.form("opt3")
'opt4=request.form("opt4")

' Creates file system object code
set fso = createobject("scripting.filesystemobject")

' creates a text file of name playlist.m3u on the web server. 
if name_of_file <> "" then 
set pllst = fso.createtextfile(server.mappath("/" & name_of_file & ""), true)
end if

if name_of_file = "" then
set pllst = fso.createtextfile(server.mappath("/playlist.m3u"), true)
end if
' write users selection into text file
if opt <> "" then
   	strOutput = opt
		strOutput = Replace(strOutput, commaspace, carriageReturn)
   	pllst.Writeline strOutput
end if
'if opt="option2" then
'   	pllst.Writeline opt2
'	pllst.Writeline chr(13)
'end if
'if opt="option3" then
'   	pllst.Writeline opt3
'	pllst.Writeline chr(13)
'end if
'if opt="option4" then
'   	pllst.Writeline opt4
'	pllst.Writeline chr(13)
'end if

'Close the test file
pllst.close



%>

</BODY>
</HTML>