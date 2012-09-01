<%@ LANGUAGE="VBScript" %>

<%  'Set file i/o constants.

    Const ForReading = 1
    Const ForWriting = 2
    Const ForAppending = 8

    'Map the file name to the physical path on the server.

    filename = "test.txt"
    path = Server.MapPath("data") & "\" & filename

    'Get requested operation.

    operation = Request.Form("operation") %>

<html>
<head>
<title>Accesing Files</title>
</head>
<body bgcolor="#ffffff">
<font face="Arial,Helvetica" size=2>

<table bgcolor="#000000" border=0 cellpadding=1 cellspacing=0><tr><td>
<table bgcolor="#ffffcc" border=0 cellpadding=8 cellspacing=0><tr valign=bottom><td>
<font face="Arial,Helvetica" size=2>
<form action="<%  = Request.ServerVariables("SCRIPT_NAME") %>" method="post">
<input name="operation" type="radio" value="create"> Create
<input name="operation" type="radio" value="delete"> Delete
<input name="operation" type="radio" value="read" checked> Read
<input name="operation" type="radio" value="write"> Write
<input name="operation" type="radio" value="append"> Append
<p>
<center>
<input type="submit" value="Submit"> <input type="reset" value="Reset">
</center>
</form>
<p>
<a href="fileops.html"><b>Return</b></a>
</font>
</td></tr></table>
</td></tr></table>
<p>

<%  'Perform requested operation, if any.

    if operation = "create" then
      call CreateFile(path)
      call ReadFile(path)
    elseif operation = "delete" then
      call DeleteFile(path)
    elseif operation = "read" then
      call ReadFile(path)
    elseif operation = "write" then
      call WriteFile(path)
      call ReadFile(path)
    elseif operation = "append" then
      call AppendFile(path)
      call ReadFile(path)
    end if %>
<p>

</font>
</body>
</html>

<%  sub CreateFile(path)

 