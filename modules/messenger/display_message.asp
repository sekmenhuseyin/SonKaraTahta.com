<%@ Language=VBScript %>
<%Response.Buffer=TRUE%>
<% Response.CacheControl="no-cache" %>
<% Response.AddHeader "Pragma", "no-cache" %>
<% Response.Expires=-1 %>
<%
fromt="from" & request.cookies("myid")
fromid=application("" & fromt & "")
newfrom="from" & fromid

message=session("message")
application("" & request.cookies("boxid") & "")=""
application("" & request.cookies("flagid") & "")="no"
application("" & newfrom & "")="no"
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Incoming Message</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="C1D8EE" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<bgsound src="beep.wav" loop="0">
<div align="center"><img src="inmessage.gif" width="150" height="35"> 
  <table width="235" height="147" border="0" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
    <tr bgcolor="C1D8EE"> 
      <td height="115" colspan="2" valign="top"><div align="center"><font size="2"> <iframe name="message" style="border:0px single white" width=100% height=100 src="message.asp?message=<%=message%>"><%=message%></iframe>
          </font></div></td>
    </tr>
    <tr bgcolor="C1D8EE"> 
      <td width="115" valign="top"><a href="send_message.asp?id=<%=fromid%>"><img src="reply.gif" width="72" height="24" border="0"></a></td>
      <td width="110" valign="top"><div align="right"><a href="javascript:window.close()"><img src="close.gif" width="72" height="24" border="0"></a></div></td>
    </tr>
  </table>
</div>
</body>
</html>

<%
response.cookies("messagelast")=message
session("flag")="no"
session("message")=""
%>