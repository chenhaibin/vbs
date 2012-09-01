<!-- #include file="evlogconst.inc" -->
<HTML>
<HEAD>
<TITLE>EventLog Source</TITLE>
</HEAD>
<BODY topmargin=10 leftmargin=10 bgcolor=#ffffff>

<%
sub WriteIcon(nEvType)
select case nEvType
 case EVENTLOG_SUCCESS
   Response.Write "Success"
 case EVENTLOG_ERROR_TYPE
   Response.Write "Error"
 case EVENTLOG_WARNING_TYPE
   Response.Write "Warning"
 case EVENTLOG_INFORMATION_TYPE
   Response.Write "Information"
 case EVENTLOG_AUDIT_SUCCESS 
   Response.Write "Audit Success"
 case EVENTLOG_AUDIT_FAILURE
   Response.Write "Audit Failure"
end select
end sub

strSource = Trim(Request.QueryString("Log"))
Set x = CreateObject("Softwing.EventLogReader")
retval = x.OpenLog(vbNullChar, strSource)
nRecord = Trim(Request.QueryString("Id"))
retval = x.getrecord(nRecord, EvId, EvType, strSource, EvDate, strComputer, strAccount, strDesc)
%>

<TABLE width=95% BORDER=0 cellspacing=2>
<TR><TH ALIGN=LEFT>Date</TH><TD><%=FormatDateTime(EvDate,2)%></TD><TH ALIGN=LEFT>Event ID</TH><TD><%=EvId%></TD></TR>
<TR><TH ALIGN=LEFT>Time</TH><TD><%=FormatDateTime(EvDate,3)%></TD><TH ALIGN=LEFT>Source</TH><TD><%=strSource%></TD></TR>
<TR><TH ALIGN=LEFT>User</TH><TD></TD><TH ALIGN=LEFT>Type</TH><TD><%WriteIcon EvType%></TD></TR>
<TR><TH ALIGN=LEFT>Computer</TH><TD><%=strComputer%></TD><TH ALIGN=LEFT>Category</TH><TD></TD></TR>
<TR><TH ALIGN=LEFT colspan=4>Description:</TH></TR>
<TR><TD colspan=4><%=strDesc%></TD></TR>
</TABLE>

<%
retval = x.CloseLog()
%>

</BODY>
</HTML>