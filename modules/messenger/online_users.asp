<%
updateid="update" & request.cookies("myid")
application("" & updateid & "")=now
%>
<html>
<head>
<meta http-equiv="Refresh" content="30">
<script language="JavaScript">

function startPopUp(desktopURL) {
  var desktop=
  window.open(desktopURL, "_blank","width=250,height=200");
  }

</script>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254"></head>
<%
numusers=application("totusers")
%><body link="#0066CC" vlink="#0066CC" alink="#FF0000" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="265" height="20" border="0" cellpadding="0" cellspacing="0">
  <tr bgcolor="6396C6"> 
    <td width="98" bgcolor="6396C6"><img src="toolbar.gif" width="265" height="20" border="0" usemap="#Map"></td>
  
</table>
<table width="265" border="0">
  <a href="Javascript:startPopUp('send_message_all.asp')">
  <tr onMouseOver="this.bgColor='#C1D8EE'" onMouseOut ="this.bgColor='#FFFFFF'"> 
    <td> <a href="Javascript:startPopUp('send_message_all.asp')"><img src="bullet.gif" width="15" height="12" border="0"> 
      <font size="2" face="Arial, Helvetica, sans-serif">Tüm Kullanıcılar</font><br />
      </td></tr></a>
  
  </a>
</table>
<%
for x=1 to numusers
%>
<%
newid="user" & x
boxid="box" & x
flagid="flag" & x
updateid="update" & x
%>
<%
if application("" & newid & "")="" then
else
 ud=application("" & updateid & "")
 updatediff=datediff("n",now,ud)
 if updatediff < -.7 then
 else
%>
<table width="265" border="0">
<a href="Javascript:startPopUp('send_message.asp?id=<%=x%>')">
  <tr onMouseOver="this.bgColor='#C1D8EE'" onMouseOut ="this.bgColor='#FFFFFF'">
    <td><a href="Javascript:startPopUp('send_message.asp?id=<%=x%>')"><img src="bullet.gif" width="15" height="12" border="0">
      <font size="2" face="Arial, Helvetica, sans-serif"><%=application("" & newid & "")%></font><br /></td></tr></a>
  
</a>
</table>

<%
end if
end if
%>
<%
next
%>
<map name="Map">
  <area shape="rect" coords="208,1,221,18" href="online_users.asp" alt="Aktif Kişiler">
  <area shape="rect" coords="226,1,240,17" href="Javascript:startPopUp('user_info.asp')" alt="Kullanıcı Bilgisi">
  <area shape="rect" coords="243,2,261,14" href="logout.asp" target="_top" alt="ÇIKIŞ">
</map>
</body>
</html>