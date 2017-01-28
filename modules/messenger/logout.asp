<%Response.Buffer=TRUE%>
<html>
<body onLoad="javascript:window.close()">
<%
userid="user" & request.cookies("newid")
boxid="box" & request.cookies("newid")
flagid="flag" & request.cookies("newid")
fromid="from" & request.cookies("newid")
authid="auth" & request.cookies("newid")
updateid="update" & request.cookies("newid")
%>
<%=userid%>-Kapat
<%=boxid%>
<%=flagid%>
<%
application.lock
application.contents.remove("" & userid & "")
application.contents.remove("" & boxid & "")
application.contents.remove("" & flagid & "")
application.contents.remove("" & fromid & "")
application.contents.remove("" & authid & "")
application.contents.remove("" & updateid & "")
response.cookies("newid")=""
response.cookies("boxid")=""
response.cookies("flagid")=""
response.cookies("fromid")=""
response.cookies("authid")=""
response.cookies("imlogdin")="no"
response.cookies("myid")=""
response.cookies("messagelast")=""
session("message")=""
session("flag")=""
application.unlock
%>
</body>