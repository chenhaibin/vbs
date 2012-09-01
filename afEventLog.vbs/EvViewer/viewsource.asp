<!-- #include file="evlogconst.inc" -->
<HTML>
<HEAD>
<TITLE>EventLog Source</TITLE>
</HEAD>
<BODY topmargin=10 leftmargin=10 bgcolor=#ffffff>

<%
sub WriteIcon(nEvType)
Response.Write "<IMG WIDTH=16 HEIGT=15 BORDER=0 SRC="""
select case nEvType
 case EVENTLOG_SUCCESS
   Response.Write "asuccess"
 case EVENTLOG_ERROR_TYPE
   Response.Write "error"
 case EVENTLOG_WARNING_TYPE
   Response.Write "warning"
 case EVENTLOG_INFORMATION_TYPE
   Response.Write "info"
 case EVENTLOG_AUDIT_SUCCESS 
   Response.Write "afail"
 case EVENTLOG_AUDIT_FAILURE
   Response.Write "asuccess"
end select
Response.Write ".gif"">"
end sub

strSource = Trim(Request.Form("LogSource"))
nMaxRecs = Trim(Request.Form("maxrecs"))
Set x = CreateObject("Softwing.EventLogReader")
retval = x.OpenLog(vbNullChar, strSource)
n = x.GetFirst()
nPrevious = n
%>

<TABLE BORDER=0>
<TR><TH ALIGN=LEFT></TH><TH ALIGN=LEFT>Date</TH><TH ALIGN=LEFT>Time</TH><TH ALIGN=LEFT>Source</TH><TH ALIGN=LEFT>Event ID</TH></TR>

<%
While nMaxRecs > 0
    n = x.GetNext(EvId, EvType, strEvSource, EvDate)
    Response.Write "<TR><TD><A TARGET=""evdesc"" HREF=""viewentry.asp?Id=" & nPrevious & "&log=" & strSource & """>"
    WriteIcon EvType
    nPrevious = n
    Response.Write "</A></TD><TD>" & FormatDateTime(EvDate,2)
    Response.Write "</TD><TD>" & FormatDateTime(EvDate,3)
    Response.Write "</TD><TD>" & strEvSource & "</TD><TD>" & EvId 
    Response.Write "</TD></TR>" & vbCRLF
    nMaxRecs = nMaxRecs - 1
Wend
%>

</TABLE>

<%
retval = x.CloseLog()
%>

</BODY>
</HTML>