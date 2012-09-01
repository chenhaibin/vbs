<% Option Explicit %>
<!-- Christoph Wille, 20.1.99 -->


<%
Const EVENTLOG_SUCCESS = 0
Const EVENTLOG_ERROR_TYPE = 1
Const EVENTLOG_WARNING_TYPE = 2
Const EVENTLOG_INFORMATION_TYPE = 4
%>

<%
' Create the AspEventlog object
Dim xObj, bResult
Set xObj = Server.CreateObject("SOFTWING.ASPEventlog")

' open the log and report the status
bResult = xObj.Open()
Response.Write "Open called with return value of: " &  bResult & "<BR>" & vbCRLF

' write a information type event and give feedback
bResult = xObj.ReportEvent(EVENTLOG_INFORMATION_TYPE, "This is a demo event from SOFTWING")
Response.Write "Event was reported successfully: " & bResult & "<BR>" & vbCRLF

' close the log (would be done automatically on object destruction)
bResult = xObj.Close()
Response.Write "Log closed successfully: " & bResult & "<BR>" & vbCRLF

' delete the object and free up memory
Set xObj = Nothing
%>
