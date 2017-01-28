<% Response.Buffer = True %>
<!--#include file="includes.asp" -->
<%
IF Session("Level") = "1" OR Session("Level") = "7" Then

act = Request.QueryString("action")
IF act = "settings" Then
call settings
ElseIF act = "update" Then
call updSet
ElseIF act = "BackupDB" Then
call backup_db
End If

Sub backup_db
IF dbType = "personal" Then
Response.Redirect "../../DB/mn"&dbNum&".mdb"
ElseIF dbType = "normal" Then
Response.Redirect "../DB/mn"&dbNum&".mdb"
End IF
End Sub

Sub settings

Set rs = Connection.Execute("Select * FROM SETTINGS")
%>
<form action="settings.asp?action=update" method="post">
<table border="0" cellspacing="2" cellpadding="2" align="center" height="708">
<tr bgcolor="#E0E0E0">
<td width="50%" align="right" height="22"><font face=verdana size=2>Site Adý</font><font face=verdana size=1 color="#FF0000">(Örn.: 
Mini-NUKE Fixed Versiyon)</font><font face=verdana size=2> : </font>
</td>
<td width="50%" height="22"><input type="text" name="name" value="<%=rs("site_isim")%>" class="form1" size="40">
</td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="50%" align="right" height="22"><font face=verdana size=2>Site Adresi</font><font face=verdana size=1 color="#FF0000">(Örn.: 
http://freehost01.websamba.com/isim)</font><font face=verdana size=2> : </font>
</td>
<td width="50%" height="22"><input type="text" name="url" value="<%=rs("site_adres")%>" class="form1" size="40">
</td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="50%" align="right" height="22"><font face=verdana size=2>Görünüm Adresi</font><font face=verdana size=1 color="#FF0000">(Örn.: 
WwW.Mini-NUKE.De)</font><font face=verdana size=2> : </font>
</td>
<td width="50%" height="22"><input type="text" name="gorunusurl" value="<%=rs("gorunusurl")%>" class="form1" size="40">
</td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="50%" align="right" height="22"><font face=verdana size=2>Site Radyosu</font><font face=verdana size=1 color="#FF0000">(Örn.:
<a href="http://yayinonline.com/rady">http://yayinonline.com/rady</a>.... vs.</font><font face=verdana size=2>) : </font>
</td>
<td width="50%" height="22"><input type="text" name="radyo" value="<%=rs("radyo")%>" class="form1" size="40">
</td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="50%" align="right" height="22"><font face=verdana size=2>Site Logosu</font><font face=verdana size=1 color="#FF0000">(Örn.:images/logo.gif)</font><font face=verdana size=2> 
: </font>
</td>
<td width="50%" height="22"><input type="text" name="site_logo" value="<%=rs("site_logo")%>" class="form1" size="40">
</td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="50%" align="right" height="22"><font face=verdana size=2>Site Topbanneri</font><font face=verdana size=1 color="#FF0000">(Örn.:images/banner.gif</font><font face=verdana size=2> : </font>
</td>
<td width="50%" height="22">
<input type="text" name="site_topbanner" value="<%=rs("site_topbanner")%>" class="form1" size="40">
</td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="50%" align="right" height="22"><font face=verdana size=2>Site Slogani</font><font face=verdana size=1 color="#FF0000">(Örn.: 
Bu Yolda Aynen Devam)</font><font face=verdana size=2> : </font>
</td>
<td width="50%" height="22"><input type="text" name="site_slogan" value="<%=rs("site_slogan")%>" class="form1" size="40">
</td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="50%" align="right" height="22"><font face=verdana size=2>Site Maili</font><font face=verdana size=1 color="#FF0000">(Örn.:admin@mini-nuke.de</font><font face=verdana size=2> 
: </font>
</td>
<td width="50%" height="22"><input type="text" name="mail" value="<%=rs("site_mail")%>" class="form1" size="40">
</td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="50%" align="right" height="22"><font face=verdana size=2>Anasayfada Listelenecek Haber Sayýsý : </font>
</td>
<td width="50%" height="22">
<select name="news" class="form1" size="1">
<%
For I = 1 TO 10
IF rs("haber_sayisi") = I Then
opt = "selected"
ELSE
opt = ""
End IF
%>
<option value="<%=I%>" <%=opt%>><%=I%></option>
<% Next %>
</select>
</td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="50%" align="right" height="21"><font face=verdana size=2>Bir Sayfada Listelenecek Program Sayýsý : </font>
</td>
<td width="50%" height="21">
<select name="prgs" class="form1" size="1">
<%
For I = 1 TO 50
IF rs("prg_sayisi") = I Then
opt = "selected"
ELSE
opt = ""
End IF
%>
<option value="<%=I%>" <%=opt%>><%=I%></option>
<% Next %>
</td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="50%" align="right" height="21"><font face=verdana size=2>Anasayfada Son 5 Haber : </font>
</td>
<td width="50%" height="21">
<select name="ln5" class="form1" size="1">
<%
IF rs("mp_news") = True Then
opt11 = "selected"
ELSE
opt21 = "selected"
End IF
%>
<option value="True" <%=opt11%>>Olsun</option>
<option value="False" <%=opt21%>>Olmasýn</option>
</td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="50%" align="right" height="21"><font face=verdana size=2>Anasayfada Son 5 Forum : </font>
</td>
<td width="50%" height="21">
<select name="lf5" class="form1" size="1">
<%
IF rs("mp_forum") = True Then
opt12 = "selected"
ELSE
opt22 = "selected"
End IF
%>
<option value="True" <%=opt12%>>Olsun</option>
<option value="False" <%=opt22%>>Olmasýn</option>
</td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="50%" align="right" height="60"><font face=verdana size=2>Üyeler Kapalý Sayfada Çýkacak Yazý : </font>
</td>
<td width="50%" height="60"><textarea name="npm" rows="5" cols="40" class="form1"><%=rs("np_msg")%></textarea>
</td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="50%" align="right" height="22"><font face=verdana size=2>Mesaj Limiti : </font>
</td>
<td width="50%" height="22">
<select name="msg_limit" class="form1" size="1">
<%
For I = 1 TO 1000
IF rs("msg_limit") = I Then
opt = "selected"
ELSE
opt = ""
End IF
%>
<option value="<%=I%>" <%=opt%>><%=I%></option>
<% Next %>
</select>
</td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="50%" align="right" height="22"><font face=verdana size=2>Þifre Hatýrlatma Sistemi : </font>
</td>
<td width="50%" height="22">
<%
If rs("fpass") = "Site" Then
opt1 = "selected"
ElseIf rs("fpass") = "Mail" Then
opt2 = "selected"
Else
opt1 = "selected"
End If
%>
<select name="fpass" size="1" class="form1">
<option value="Site" <%=opt1%>>Site Üzerinden</option>
<option value="Mail" <%=opt2%>>Mail Yolu Ýle</font>
</select>
</td>
</tr>
<tr bgcolor="#E0E0E0" class="td_menu2">
<td width="50%" align="right" height="22"><font face=verdana size=2>Site Temasý : </font>
</td>
<td width="50%" height="22">
<%
Set themes = Server.CreateObject("ADODB.RecordSet")
themes.open "Select * FROM THEMES ORDER BY t_name ASC",Connection,3,3
%>
<select name="theme" size="1" class="form1">
<%
Do Until themes.EoF
IF rs("theme") = themes("t_dir") Then
opt = "selected"
ELSE
opt = ""
End IF
Response.Write "<option value="""&themes("t_dir")&""" "&opt&">"&themes("t_name")&"</option>"
themes.MoveNext
Loop
%>
</select>
</td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="50%" align="right" height="22"><font face="verdana" size="2">Flood Koruma Süresi (Dakika) :</font></td>
<td width="50%" height="22">
<select name="flood_time" size="1" class="form1">
<%
For I = 1 TO 60
IF I = rs("flood_time") Then
strOPT = "seLected"
ELSE
strOPT = ""
End IF
Response.Write "<option value="""&I&""" "&strOPT&">" & I & "</option>"
Next
%>
</select>
</td>
</tr>
<tr bgcolor="#E0E0E0" class="td_menu2">
<td width="50%" align="right" height="21"><font face="verdana" size="2">Sað Tuþ Menu : </font>
</td>
<td width="50%" height="21">
<select name="SA" class="form1" size="1">
<%
IF rs("sagtus") = True Then
opts12 = "selected"
ELSE
opts22 = "selected"
End IF
%>
<option value="True" <%=opts12%>>Olsun</option>
<option value="False" <%=opts22%>>Olmasýn</option>
</select></td>
</tr>
<tr bgcolor="#E0E0E0" class="td_menu2">
<td width="50%" align="right" height="21"><font face="verdana" size="2">Merlin 
Scripti : </font>
</td>
<td width="50%" height="21">
<select name="merlin" class="form1" size="1">
<%
IF rs("merlin") = True Then
opts13 = "selected"
ELSE
opts33 = "selected"
End IF
%>
<option value="True" <%=opts13%>>Olsun</option>
<option value="False" <%=opts33%>>Olmasýn</option>
</select></td>
</tr>
<tr bgcolor="#E0E0E0" class="td_menu2">
<td width="50%" align="right" height="21"><font face="verdana" size="2">Hýzlý 
Menu : </font>
</td>
<td width="50%" height="21">
<select name="hmenu" class="form1" size="1">
<%
IF rs("hmenu") = True Then
opts14 = "selected"
ELSE
opts44 = "selected"
End IF
%>
<option value="True" <%=opts14%>>Olsun</option>
<option value="False" <%=opts44%>>Olmasýn</option>
</select></td>
</tr>
<tr>
<td width="50%" align="right" height="26">&nbsp;</td>
<td width="50%" height="26"><input type="submit" value="Deðiþiklikleri Kaydet" class="submit" style="width:76%">
</td>
</tr>
</table>
</form>
<%
themes.Close
rs.Close
Set themes = Nothing
Set rs = Nothing

End Sub
Sub updSet

gorunusurl = duz(Request.Form("gorunusurl"))
radyo = duz(Request.Form("radyo"))
sname = duz(Request.Form("name"))
surl = duz(Request.Form("url"))
smail = duz(Request.Form("mail"))
snews = duz(Request.Form("news"))
sprgs = duz(Request.Form("prgs"))
sml = duz(Request.Form("msg_limit"))
snpm = Request.Form("npm")
stheme = duz(Request.Form("theme"))
sfpass = duz(Request.Form("fpass"))
site_slogan = duz(Request.Form("site_slogan"))
site_logo = duz(Request.Form("site_logo"))
sagt = duz(Request.Form("sa"))
merlin = duz(Request.Form("merlin"))
hmenu = duz(Request.Form("hmenu"))
site_topbanner = duz (request.form("site_topbanner"))

If sname="" Then
Response.Write "<center><font face=verdana size=2 color=red>Lütfen Site Ýsmini Yazýnýz...</font></center>"
ElseIf surl="" Then
Response.Write "<center><font face=verdana size=2 color=red>Lütfen Site Adresini Yazýnýz...</font></center>"
ElseIf smail="" Then
Response.Write "<center><font face=verdana size=2 color=red>Lütfen Site Mailini Yazýnýz...</font></center>"
ElseIf snews="" Then
Response.Write "<center><font face=verdana size=2 color=red>Lütfen Bir Sayfada Listelenecek Haber Sayýsýný Yazýnýz...</font></center>"
ElseIf sprgs="" Then
Response.Write "<center><font face=verdana size=2 color=red>Lütfen Bir Sayfada Listelenecek Program Sayýsýný Yazýnýz...</font></center>"
ElseIf sml="" Then
Response.Write "<center><font face=verdana size=2 color=red>Lütfen Mesaj Limitini Yazýnýz...</font></center>"
ElseIf snpm="" Then
Response.Write "<center><font face=verdana size=2 color=red>Lütfen Üyeler Kapalý Sayfada çýkacak Yazýyý Yazýnýz...</font></center>"
ELSE

Set ent = Server.CreateObject("ADODB.RecordSet")
entSQL = "Select * FROM SETTINGS"
ent.open entSQL,Connection,1,3

ent("gorunusurl")=gorunusurl
ent("radyo")=radyo
ent("site_topbanner")=site_topbanner
ent("sagtus")=sagt
ent("hmenu")=hmenu
ent("merlin")=merlin
ent("site_isim") = sname
ent("site_logo") = site_logo
ent("site_adres") = surl
ent("site_mail") = smail
ent("haber_sayisi") = snews
ent("prg_sayisi") = sprgs
ent("msg_limit") = sml
ent("np_msg") = snpm
ent("theme") = stheme
ent("fpass") = sfpass
ent("mp_news") = Request.Form("ln5")
ent("mp_forum") = Request.Form("lf5")
ent("site_slogan") = Request.Form("site_slogan")
ent("flood_time") = Request.Form("flood_time")
ent.Update
Response.Redirect Request.ServerVariables("HTTP_REFERER")
END IF

End Sub
End If
%>