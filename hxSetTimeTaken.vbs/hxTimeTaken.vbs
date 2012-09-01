
On error resume Next

strInstance = 1

Set oIIS = GetObject ("IIS://localhost/W3SVC/" & strInstance)

select case err.number

	case "-2147024893"
		wscript.echo "site doesn't exist"
		wscript.quit(1)
	case else
		wscript.echo "aok"
		On Error GoTo 0 
end select

if oIIS.LogType = 0 then
	wscript.echo "Site logs"
else
	wscript.echo "site doesn't log"
	funcEnableLoggi
end if
