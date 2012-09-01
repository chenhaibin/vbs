<%
'Check to see if title has been entered or not
u_title=request.form("u_title")
if u_title = "" then
%>
<!-- Input form area - This will only display when no Title has been entered -->
<form method="POST" action="<%= request.servervariables("script_name") %>">

<p>Document Title<br>
<input type="text" name="u_title" size="35"></p>
<p>Font Size<br>
<select size="1" name="u_text_size">
<option selected value="1">1</option>
<option value="2">2</option>
<option value="3">3</option>
<option value="4">4</option>
<option value="5">5</option>
<option value="6">6</option>
&nbsp;
</select></p>
<p>Paragraph 1<br>
<textarea rows="2" name="u_paragraph1" cols="35"></textarea></p>
<p>Paragraph 2<br>
<textarea rows="2" name="u_paragraph2" cols="35"></textarea><input type="submit" value="Submit" ></p>
</form>
<%
end if
%>

<%
if u_title <> "" then
' If there is a user inputted title
' get all of the user inputed values
u_title=request.form("u_title")
u_paragraph1=request.form("u_paragraph1")
u_paragraph2=request.form("u_paragraph2")
u_text_color=request.form("u_text_color")
u_text_size=request.form("u_text_size")
g_filename=replace(u_title," ","_")


set fso = createobject("scripting.filesystemobject")
' create the text (html) file to the server adding the -mmddyyyy after the g_title value
Set act = fso.CreateTextFile(server.mappath("/"&g_filename & ""& month(date())& day(date())& year(date()) &".htm"), true)

' write all of the user input to the text (html) document 
' The .htm extension can just as easily be .asp or .inc whatever best suits your needs
act.WriteLine "<html>"
act.WriteLine chr(13) 
act.WriteLine "<title>"& u_title &"</title>"
act.WriteLine chr(13)
act.WriteLine "<body bgcolor='#FFFFFF'>"
act.WriteLine chr(13)
act.WriteLine "<p align='center'><font face='arial' size='"& u_text_size &"'>"
act.WriteLine chr(13)
act.WriteLine u_title &"</p>"
act.WriteLine chr(13)
act.WriteLine "<p align='left'><font face='arial' size='"&u_text_size&"'>"
act.WriteLine chr(13)
act.WriteLine u_paragraph1 &"</p>"
act.WriteLine chr(13)
act.WriteLine "<p align='left'><font face='arial' size='"& u_text_size &"'>"
act.WriteLine chr(13)
act.WriteLine u_paragraph2 &"</p>"
act.WriteLine chr(13)
act.WriteLine "<p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p>"
act.WriteLine "<p align='center'><font face='arial' size='"& u_text_size &"'>"
act.WriteLine "This document was created on&nbsp;"
act.WriteLine now() &"</p>"

' close the document 
act.Close
%>
Your page has been successfully create and can be viewed by clicking 
<a href="/<%= g_filename & month(date())& day(date())& year(date()) %>.htm" target="_blank">here</a>
<br>
<br>
<% response.write "<html>"
response.write chr(13) 
response.write "<title>"& u_title &"</title>"
response.write chr(13)
response.write "<body bgcolor='#FFFFFF'>"
response.write chr(13)
response.write "<p align='center'><font face='arial' size='"& u_text_size &"'>"
response.write chr(13)
response.write u_title &"</p>"
response.write chr(13)
response.write "<p align='left'><font face='arial' size='"&u_text_size&"'>"
response.write chr(13)
response.write u_paragraph1 &"</p>"
response.write chr(13)
response.write "<p align='left'><font face='arial' size='"& u_text_size &"'>"
response.write chr(13)
response.write u_paragraph2 &"</p>"
response.write chr(13)
response.write "<p>&nbsp;</p><p>&nbsp;</p><p>&nbsp;</p>"
response.write "<p align='center'><font face='arial' size='"& u_text_size &"'>"
response.write "This document was created on&nbsp;"
response.write now() &"</p>"
end if
%>
