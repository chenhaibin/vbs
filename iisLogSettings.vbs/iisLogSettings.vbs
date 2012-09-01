Dim IISOBJ, Servername, WebSiteID

ServerName = "localhost"
WebSiteID = 1

Set IISOBJ = GetObject("IIS://" & ServerName & "/W3SVC/" & WebSiteID) 

if (IISOBJ.LogType = 0) then
	WScript.Echo "Logging                   : Disabled"
else 
	WScript.Echo "Logging                   : Enabled"
end if

WScript.Echo "LogFileDirectory          : " & IISOBJ.LogFileDirectory
WScript.Echo "LogPluginClsId            : " & IISOBJ.LogPluginClsId 
WScript.Echo "LogFileFormat             : " & ReturnLogFileFormat(IISOBJ.LogPluginClsId)
WScript.Echo
WScript.Echo "LogOdbcDataSource         : " & IISOBJ.LogOdbcDataSource
WScript.Echo "LogOdbcTableName          : " & IISOBJ.LogOdbcTableName
WScript.Echo "LogOdbcUserName           : " & IISOBJ.LogOdbcUserName
WScript.Echo "LogOdbcPassword           : " & IISOBJ.LogOdbcPassword
WScript.Echo
Select Case IISOBJ.logFilePeriod
	case 0 : if (IISOBJ.LogFileTruncateSize = -1) then
			WScript.Echo "LogFilePeriod             : Unlimited	"
		 else
			WScript.Echo "LogFilePeriod             : When LogFileTruncateSize is reached"
		 end if
	case 1 : WScript.Echo "LogFilePeriod             : Daily"
	case 2 : WScript.Echo "LogFilePeriod             : Weekly"
	case 3 : WScript.Echo "LogFilePeriod             : Monthly"
	case 4 : WScript.Echo "LogFilePeriod             : Hourly"
	case else : 
                 WScript.Echo "LogFilePeriod             : Unknown (value=" & IISOBJ.Logfileperiod & ")"
end select 
if (IISOBJ.LogFileTruncateSize = "-1") then
	WScript.Echo "LogFileTruncateSize       : Unlimited Size"
else
	WScript.Echo "LogFileTruncateSize       : " & IISOBJ.LogFileTruncateSize & " Bytes, or " & IISOBJ.LogFileTruncateSize  / 1024 & " KB, or " & IISOBJ.LogFileTruncateSize / 1048576 & " MB"
end if
WScript.Echo "LogFileLocaltimeRollover  : " & IISOBJ.LogFileLocaltimeRollover
 
WScript.Echo
WScript.Echo "LogExtFileBytesRecv       : " & IISOBJ.LogExtFileBytesRecv
WScript.Echo "LogExtFileBytesSent       : " & IISOBJ.LogExtFileBytesSent
WScript.Echo "LogExtFileClientIp        : " & IISOBJ.LogExtFileClientIp
WScript.Echo "LogExtFileComputerName    : " & IISOBJ.LogExtFileComputerName
WScript.Echo "LogExtFileCookie          : " & IISOBJ.LogExtFileCookie
WScript.Echo "LogExtFileDate            : " & IISOBJ.LogExtFileDate
WScript.Echo "LogExtFileFlags           : " & IISOBJ.LogExtFileFlags
WScript.Echo "LogExtFileHttpStatus      : " & IISOBJ.LogExtFileHttpStatus
WScript.Echo "LogExtFileMethod          : " & IISOBJ.LogExtFileMethod
WScript.Echo "LogExtFileProtocolVersion : " & IISOBJ.LogExtFileProtocolVersion
WScript.Echo "LogExtFileReferer         : " & IISOBJ.LogExtFileReferer
WScript.Echo "LogExtFileServerIp        : " & IISOBJ.LogExtFileServerIp
WScript.Echo "LogExtFileServerPort      : " & IISOBJ.LogExtFileServerPort
WScript.Echo "LogExtFileSiteName        : " & IISOBJ.LogExtFileSiteName
WScript.Echo "LogExtFileTime            : " & IISOBJ.LogExtFileTimeTaken
WScript.Echo "LogExtFileTimeTaken       : " & IISOBJ.LogExtFileTimeTaken 
WScript.Echo "LogExtFileUriQuery        : " & IISOBJ.LogExtFileUriQuery
WScript.Echo "LogExtFileUriStem         : " & IISOBJ.LogExtFileUriStem
WScript.Echo "LogExtFileUserAgent       : " & IISOBJ.LogExtFileUserAgent
WScript.Echo "LogExtFileUserName        : " & IISOBJ.LogExtFileUserName
WScript.Echo "LogExtFileWin32Status     : " & IISOBJ.LogExtFileWin32Status


' Determine the current logging method
Function ReturnLogfileFormat(CLSID)
Set LoggingModules = GetObject("IIS://" & ServerName & "/logging") 
'Loop through the installed modules to find the currently selected module at this server. 
If CLSID <> "" Then 
	for Each LogModule in LoggingModules 
		if (LogModule.Class = "IIsLogModule") then
      			If LogModule.LogModuleID = CLSID Then 
				ReturnLogfileFormat = LogModule.Name
				Exit Function
			end if
		end if
	Next 
	ReturnLogfileFormat = "Unknown Format"
	exit function
end if
ReturnLogfileFormat = "No Log File Format Configured"
end function

