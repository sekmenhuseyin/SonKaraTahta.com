<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Mini Messenger</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<SCRIPT language="JavaScript">
<!--
var exit=true;
function out()


    {
    if (exit)
    window.open('logout.asp','40top',"toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=yes,width=100,height=100");
}
// -->
</script>
</head>
<%
newid="user" & request.cookies("myid")
%>
<body onUnload=out() bgcolor="C1D8EE" leftmargin="7" topmargin="7" marginwidth="7" marginheight="7">
<div align="center"> 
  <table width="280" height="354" border="0" bgcolor="#FFFFFF">
    <tr> 
      <td height="350" valign="top"><table width="280" height="348" border="0">
          <tr> 
            <td height="32"> <div align="center"><img src="messenger.gif" width="248" height="32"></div></td>
          </tr>
          <tr> 
            <td height="244"><iframe name="userwindow" style="border:0px single white" width=100% height=300 src="online_users.asp"></iframe></td>
          </tr>
          <tr> 
            <td height="20" bgcolor="#71A5D9"><div align="center"><strong><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif">Şuanda Bağlısınız: <%=application("" & newid & "")%></font></strong></div></td>
          </tr>
          <tr> 
            <td height="40"><iframe name="flashconnection" style="border:0px single white" width=100% height=45 src="im.asp"></iframe></td>
          </tr>
        </table></td>
    </tr>
  </table>
</div>
</body>
</html>