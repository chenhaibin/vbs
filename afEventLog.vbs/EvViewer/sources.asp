<HTML>
<HEAD>
<TITLE>EventLog Sources</TITLE>
</HEAD>

<BODY topmargin=10 leftmargin=10 bgcolor=#ffffff>
<Font face=Arial>

<form action="viewsource.asp" target="evsource" method=POST>
<table><tr>
<td><B>Source:</B>&nbsp;</td>
<td>
<select name="LogSource">
<OPTION VALUE="System">System
<OPTION VALUE="Security">Security
<OPTION VALUE="Application">Application
</select>
</td>
<td>
<input size=5 name="maxrecs" value="25">
</td>
<td>
<input type=submit value="Submit">
</td>
<td valign=center>
&nbsp;&nbsp;<A HREF="./index.asp" target="_top"><img src="info.gif" width=16 height=15 border=0 alt="Information about Online Eventlog"></A>
</td valign=top>
</table>

</form>


</BODY>
</HTML>