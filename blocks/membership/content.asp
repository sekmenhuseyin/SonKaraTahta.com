<!--#include file="../../INC/b_include.asp" -->
<% If Request.Cookies("MiniNuke_Gir")("girdi")="Evet" OR Session("Enter") = "Yes" Then
Set memInfo = Connection.Execute("Select * FROM MEMBERS WHERE uye_id = "&Session("Uye_ID")&"")
Set unR = Connection.Execute("Select Count(*) as count FROM MESSAGES WHERE TYPE = 0 AND READ = False AND SEE = True AND TO = '"&Session("Member")&"'")
Set allmsgs = Connection.Execute("Select Count(*) as count FROM MESSAGES WHERE TYPE = 0 AND SEE = True AND TO = '"&Session("Member")&"'")
If unR("count") > sett("msg_limit") Then unRd = sett("msg_limit") Else unRd = unR("count")
If allmsgs("count") > sett("msg_limit") Then allM = sett("msg_limit") Else allM = allmsgs("count")%>
<font class="td_menu">
<table border="0" cellspacing="2" cellpadding="0" width="90%" align="center" class="td_menu"><tr><td width="100%" bgcolor=<%=menu_back%>><b>&nbsp;<%=Session("Member")%></b></td>
<td width="8%" bgcolor="<%=menu_back%>" align="center"><a href="Your_Account.asp?op=Logout"><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/logout.gif" border="0"></a></td></tr></table>
<img src="IMAGES/stats/over.gif"> <a href="Your_Account.asp"><%=ya_topic%></a><br />
<img src="IMAGES/stats/over.gif"> <a href="search.asp?action=Member"><%=uyemenu_uyeara%></a><br />
<img src="IMAGES/stats/over.gif"> <a href="msgbox.asp?action=main&id=<%=Session("Uye_ID")%>"><%=uyemenu_msgbox%></a> (<%=unRd%>/<%=allM%>)<br />
<img src="IMAGES/stats/over.gif"> <a href="iletisim.asp?mn-fixed=iletisim">Ýletiþim</a><br />
<img src="IMAGES/stats/over.gif"> <a href="javascript:void window.open('modules/messenger/login.asp','','width=310,height=440,scrollbars=0')"> Messenger</a><br />
<img src="IMAGES/stats/over.gif"> <a href="Your_Account.asp?op=Logout">Güvenli Çýkýþ</a><br />
<%If Session("Level") = "1" OR Session("Level") = "2" OR Session("Level") = "3" OR Session("Level") = "4" OR Session("Level") = "5" OR Session("Level") = "6" OR Session("Level") = "7" OR Session("Level") = "8"  Then	Response.Write "» <a href=administrator/default.asp TARGET=_blank>"&control_panel&"</a>"
Response.Write "<br /><div align=""center""><img width=100 height=100 src="""&memInfo("u_avatar")&""" border=""1""></div>"%>
</font>
<%Else
the_security_code=cpass(8):session("login_gguvenlik")=the_security_code%>
<script language="JavaScript"><!-- 
function CheckfrmLogin () {
if (document.frmLogin.kuladi.value==""){document.getElementById('img_login').innerHTML="Nick eksik!";return false;}
else if (document.frmLogin.password.value==""){document.getElementById('img_login').innerHTML="Þifre eksik!";return false;}
else {document.frmLogin.lgn_sbmt.disabled=true; return true;}
}
// --></script>
<form action="enter.asp" name="frmLogin"  method="post" onsubmit="return CheckfrmLogin();">
<table width="1" border="0" cellspacing="0" cellpadding="0" align="center" class="td_menu">
<tr><td colspan="2" align="center"><div id="img_login" style="font-family: Tahoma; font-size: 7pt;"></div></td></tr>
<tr><td align="right"><%=uye_kuladi%>:&nbsp;</td><td align="left"><input type="text" name="kuladi" align="center"  class="form1" size="9" style="font-family: Tahoma; font-size: 8pt" onfocus="document.getElementById('img_login').innerHTML='';"></td></tr>
<tr><td align="right"><%=uye_password%>:&nbsp;</td><td align="left"><input type="password" name="password" align="center"  class="form1" size="9" style="font-family: Tahoma; font-size: 8pt" onfocus="document.getElementById('img_login').innerHTML='';"></td></tr>
<tr><tr><td align="right">Kod:&nbsp;</td><td align="left"><input type="text" name="guvenlik" value="<%=the_security_code%>" class="form1" size="9" style="font-family: Tahoma; font-size: 8pt"></td></tr>
<tr><td align="right"><input type="hidden" name="gguvenlik" value="<%=the_security_code%>"></td><td align="left"><b><font face="verdana" size="2" align="left"><%=the_security_code%></font></b></td></tr>
<tr><td colspan="2"></td></tr>
<tr><td colspan="2" align="center"><input name="lgn_sbmt" type="submit" value="<%=uye_submit%>" class="submit" style="width:100%"></td></tr>
</table>
</form>
<table align="center" border="0" cellspacing="1" cellpadding="1" class="td_menu2">
<tr><td width="2%"><img src="images/stats/uye.gif" border="0"></td><td width="98%"><a href="membership.asp?action=new"><%=uye_new%></a></td></tr>
<tr><td width="2%"><img src="images/stats/pass.gif" border="0"></td><td width="98%"><a href="membership.asp?action=lostpass&step=1"><%=uye_lostpwd%></a></td></tr>
</table>
<% End If %>
