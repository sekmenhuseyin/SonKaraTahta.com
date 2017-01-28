<!--#include file="view.asp" -->
<%call UST:call ORTA
Set keyifweb=Server.CreateObject("ADODB.Recordset"):sor="Select * from mesaj order by id desc":keyifweb.Open sor,Connection,1,3%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
	<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=sett("site_isim")%> Ýletiþim</font></td></tr>
	<tr> 
	    <td align="left" valign="top" bgcolor="<%=menu_bg%>" class="td_menu">
			<%islem=request.QueryString("islem")
			ekleyen=request.form("ekleyen")
			email=request.form("email")
			baslik=request.form("baslik")
			mesaj=request.form("mesaj")
			tarih=request.form("tarih")
			if Session("Enter")<>"Yes" or islem="" or ekleyen="" or email="" or baslik="" or mesaj="" or tarih="" then%>
<%if islem<>"" then Response.Write "<center>Lütfen tüm alanlarý doldurunuz...<br /><br /></center>"%>
<script language="JavaScript"><!-- 
function Checkfrmiletisim () {
if (document.FrmIletisim.baslik.value==""){document.getElementById('img_login').innerHTML="Baþlýk eksik!";return false;}
else if (document.FrmIletisim.mesaj.value==""){document.getElementById('img_login').innerHTML="Mesaj eksik!";return false;}
else {document.FrmIletisim.ilet_sbmt.disabled=true; return true;}
}
// --></script>
<form method="post" name="FrmIletisim" action="?islem=ekle" onSubmit="return Checkfrmiletisim();"><table border="0" cellspacing="2" cellpadding="2" align="center" width="100%" class="td_menu">
<tr><td colspan="2" align="center"><div id="img_login"></div></td></tr>
<tr><td><b>Baþlýk :</b> &nbsp;&nbsp; <input name="baslik" type="text" class="form1" size="60%" onFocus="document.getElementById('img_login').innerHTML=''"></td></tr>
<tr><td align="center" bgcolor=buttonface>
<script src="htmlarea/editor.js" language="Javascript1.2"></script>
<script language="JavaScript1.2" defer>editor_generate('mesaj');</script>
<textarea name="mesaj" rows="10" class="form1" style="width:100%"></textarea>
</td></tr>
<tr>
<td align="center">
<input name="ekleyen" type="hidden" value="<%=Session("Member")%>">
<input name="email" type="hidden" value="<%=Session("Email")%>">
<input name="tarih" type="hidden" value="<%=(Date)%>">
<input type="submit" name="ilet_sbmt" value="Gönder" class="submit" style="width:60%">
</td>
</tr>
</table></form>
			<%else
				Set gir=server. CreateObject("ADODB.Recordset")
				kayit="Select * from mesaj"
				gir.Open kayit,Connection,1,3
				gir.AddNew
				gir("ekleyen")=ekleyen 
				gir("email")=email
				gir("baslik")=baslik
				gir("mesaj")=mesaj
				gir("tarih")=tarih
				gir.Update
				Response.Write "<center><br />Mesajýnýz Eklenmiþtir. Teþekkürler.</center>"
			End IF%>
		</td>
	</tr>
</table>
<%call ORTA2
call ALT%>
