set oFso = createobject ("scripting.filesystemObject")

Set DNS = CreateObject("TCPIP.DNS")


set oFile = ofso.opentextfile ("195.92.1_244.x.txt")

i = 0

Do while oFile.atendofstream = false

	i = i + 1	

	strline = ofile.readline
	
	if Left (strline, 2) = "--" or Len (strline) = 0 then
	
		'ignore line
	
	else

		'process strline
		strline = Replace (strline, "* - ", "")

		'process strhost
		strhost = gethost (strline)		

		'process strServer

		set oHTTP = createobject ("Coalesys.CSHttpClient.1")

		oHTTP.RequestURL = strline
		oHTTP.Execute "GET"
	
		strHeader = oHTTP.ResponseHeaders

		arrHeader = Split (strheader, vbcrlf)

		for i = 0 to ubound (arrheader)
		
		if Left (arrheader(i), 6) = "Server" then
			strServer = arrHeader(i)
		else
		end if
	
		set oHTTP = nothing 

next

wscript.echo strServer

		wscript.echo "------------------------"
		wscript.echo "Host No  : " & i
		wscript.echo "Host IP  : " & strline
		wscript.echo "HostName : " & strHost
		wscript.echo "Server   : " & strServer

	end if



loop


Function GetHost(IP)
  If bDNS = "N" Then
    GetHost = "-"
  Else
    On Error Resume Next
      GetHost = DNS.GetHostByIP(IP)
      If Err Then GetHost = Err.Description
    On Error GoTo 0
  End If
End Function