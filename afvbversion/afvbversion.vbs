
' FileName	: Afvbversion.vbs - Displays current versions and build numbers of WSH
' Author	: Adrian Farnell
' Date Amended	: 06.10.2000
' Version	: Pre-Alpha

wscript.echo 	"Version of Windows Scripting host : " & wscript.version & vbcrlf & _
		"Version of Vbscript : " & ScriptengineMajorversion & "." & scriptengineminorversion & "." & Scriptenginebuildversion & vbcrlf & _
		"Version of WSH Components :  Unknown" 

wscript.quit