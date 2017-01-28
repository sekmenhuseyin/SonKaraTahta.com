<%@ Language=VBScript %>
<%Response.Buffer=TRUE%>
<% Response.CacheControl="no-cache" %>
<% Response.AddHeader "Pragma", "no-cache" %>
<% Response.Expires=-1 %>
<%
newid="user" & request("recipid")
boxid="box" & request("recipid")
flagid="flag" & request("recipid")
fromid="from" & request("recipid")
fromnum=request.cookies("myid")
userfrom="user" & fromnum
%>
<%=fromnum%><br /><%=userfrom%>

<%
from=application("" & userfrom & "")

application.lock
application("" & boxid & "")="<div align='right'><font size='1'>[" & now & "]</font></div><font size='2'><b>" & from & " says:</b><br />" & request("message") & "</font>"
'application("" & boxid & "")="Same message"
application("" & flagid & "")="yes"
application("" & fromid & "")=request.cookies("myid")
application.unlock
%>
<html>
<body OnLoad="javascript:window.close()">
</body>
</html>

