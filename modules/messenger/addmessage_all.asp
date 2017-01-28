<%@ Language=VBScript %>
<%Response.Buffer=TRUE%>
<% Response.CacheControl="no-cache" %>
<% Response.AddHeader "Pragma", "no-cache" %>
<% Response.Expires=-1 %>
<%
fromnum=request.cookies("myid")
userfrom="user" & fromnum
%>

<%
from=application("" & userfrom & "")
application.lock
%>
<%
numusers=application("totusers")
for x=1 to numusers
%>
<%
newid="user" & x
boxid="box" & x
flagid="flag" & x
fromid="from" & x
%>
<%
if application("" & newid & "")="" then
else
%>
<%
application("" & boxid & "")="<div align='right'><font size='1'>[" & now & "]</font></div><font size='2'><b>" & from & " says:</b><br />" & request("message") & "</font>"
application("" & flagid & "")="yes"
application("" & fromid & "")=request.cookies("myid")
%>
<%
end if
next
application.unlock
%>
<html>
<body OnLoad="javascript:window.close()">
</body>
</html>

