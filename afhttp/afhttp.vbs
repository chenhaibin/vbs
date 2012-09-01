

for i = 1 to 2880
	
	strStartTime = timer

	Set oHTTP = WScript.CreateObject("Microsoft.XMLHTTP")

	Call oHTTP.Open("HEAD", "http://hww.resolvel2.hx-online.hxgroup.com/resolve/", False)

	err.clear

	on error resume next
	
	Call oHTTP.Send()
	
	'wscript.echo err.number

	if err.number = 0 then

	else
		msgbox err.description, 1, "Problem"
	end if 
	
	on error goto 0
	
	strResponse = oHTTP.GetAllResponseHeaders()
	
	if Len (strResponse) < 5 then
		msgbox "Response less than 5",1, "Problem"
	end if

	WScript.Echo "----------------" & time & "(" & err.number & ")" & vbcrlf & strResponse
	
	strFinishTime = timer
	
	wscript.Echo "Request Took: " & 100 * (strFinishTime - strStartTime) & " hundredths of a second"

	wscript.sleep 20000
	wscript.sleep 20000
	wscript.sleep 20000
	
	Set oHTTP = Nothing
	


next

