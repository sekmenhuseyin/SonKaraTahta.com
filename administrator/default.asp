<% Response.Buffer = True %>
<!--#include file="includes.asp" -->
<title>Yönetim Paneli</title>
<%
IF Session("Level") = "1" OR Session("Level") = "2" OR Session("Level") = "3" OR Session("Level") = "4" OR Session("Level") = "5" OR Session("Level") = "7" OR Session("Level") = "8" OR Session("Level") = "9" OR Session("Level") = "10" Then
IF Request.QueryString("pane") = "Left" Then
%>
<center><a href="?pane=Main" target="page"><img src="IMAGES/logo.gif" border="0"></a></center>
<br>
<% IF Session("Level") = "1" OR Session("Level") = "7" Then %>
<table border="0" cellspacing="1" cellpadding="0" width="90%" bgcolor="#999999" align="center" class="td_menu">
<tr style="font-weight:bold">
  <td background="Tema/sidebox-title-bg.gif" height="20" align="center">
  <font face="Verdana">

  <img border="0" src="Tema/menu_bullet.gif" align="left"></font>SITE</td></tr>
<tr><td>
<table border="0" cellspacing="0" cellpadding="1" width="100%" bgcolor="#FFFFFF" class="td_menu2">
<tr>
<td>
<img border="0" src="Tema/simge.gif"> <a href="settings.asp?action=settings" target="page">Ayarlar</a><br>
<img border="0" src="Tema/simge.gif"> <a href="settings.asp?action=BackupDB" target="page">Veritabanýný Yedekle</a><br>
<img border="0" src="Tema/simge.gif"> <a href="upgrade.asp" target="page">DB Upgrade</a><br>
<img border="0" src="Tema/simge.gif"> <a href="ipban.asp" target="page">Banlama Ýþlemleri</a><br>
<img border="0" src="Tema/simge.gif"> <a href="sql_exec.asp" target="page">SQL Kodu Çalýþtýr</a><br><br>
<img border="0" src="Tema/simge.gif"> <a href="lock_site.asp" target="page">Siteyi Kilitle</a><br>
<img border="0" src="Tema/simge.gif"> <a href="themes.asp" target="page">Temalar</a><br>
<img border="0" src="Tema/simge.gif"> <a href="themes.asp?op=New" target="page">Tema Ekle</a>
</td>
</tr>
</table>
</table>
<br>
<p align="center">
<%
End IF
IF Session("Level") = "1" OR Session("Level") = "9" OR Session("Level") = "7" OR Session("Level") = "5" Then
%>
<table border="0" cellspacing="1" cellpadding="0" width="90%" bgcolor="#999999" align="center" class="td_menu">
<tr style="font-weight:bold">
  <td background="Tema/sidebox-title-bg.gif" height="20" align="center">
  <font face="Verdana">

  <img border="0" src="Tema/menu_bullet.gif" align="left"></font>FORUM</td></tr>
<tr><td bgcolor="#FFFFFF">
<table border="0" cellspacing="0" cellpadding="1" width="100%" bgcolor="#FFFFFF" class="td_menu2">
<tr>
<td>
<img border="0" src="Tema/simge.gif">
<a target="page" href="Mods.asp?action=ShowSettings">Forum Ayarlarý</a><br>
<img border="0" src="Tema/simge.gif">
<a target="page" href="Mods.asp?action=NewCat">Kategori Ekle</a><br>
<img border="0" src="Tema/simge.gif">
<a target="page" href="Mods.asp?action=NewForum">Forum Ekle</a><br>
<img border="0" src="Tema/simge.gif">
<a target="page" href="Mods.asp?action=NewMod">Yönetici Ata</a></td>
</tr>
</table>
</table>
<br>
<p align="center">

<%
End IF
IF Session("Level") = "1" OR Session("Level") = "4" OR Session("Level") = "7" Then
%> </p>
<table border="0" cellspacing="1" cellpadding="0" width="90%" bgcolor="#999999" align="center" class="td_menu">
<tr style="font-weight:bold">
  <td background="Tema/sidebox-title-bg.gif" height="20" align="center">
  <font face="Verdana">

  <img border="0" src="Tema/menu_bullet.gif" align="left"></font>HABERLER</td></tr>
<tr><td>
<table border="0" cellspacing="0" cellpadding="1" width="100%" bgcolor="#FFFFFF" class="td_menu2">
<tr>
<td>
<img border="0" src="Tema/simge.gif"> <a href="news.asp?action=AllNews" target="page">Haberler</a><br>
<img border="0" src="Tema/simge.gif"> <a href="news.asp?action=Cats" target="page">Kategoriler</a><br>
<img border="0" src="Tema/simge.gif"> <a href="news.asp?action=WaitApprove" target="page">Onay Bekleyenler</a><br><br>
<img border="0" src="Tema/simge.gif"> <a href="news.asp?action=new_enter" target="page">Haber Ekle</a><br>
<img border="0" src="Tema/simge.gif"> <a href="news.asp?action=Cat_New" target="page">Kategori Aç</a>
</td>
</tr>
</table>
</td></tr>
</table>
<br>
<%
End IF 
IF Session("Level") = "1" OR Session("Level") = "7" OR Session("Level") = "4" Then
%>
<table border="0" cellspacing="1" cellpadding="0" width="90%" bgcolor="#999999" align="center" class="td_menu">
<tr style="font-weight:bold">
  <td background="Tema/sidebox-title-bg.gif" height="20" align="center">
  <font face="Verdana">

  <img border="0" src="Tema/menu_bullet.gif" align="left"></font>Sabit HABERLER</td></tr>
<tr><td>
<table border="0" cellspacing="0" cellpadding="1" width="100%" bgcolor="#FFFFFF" class="td_menu2">
<tr>
<td>
<img border="0" src="Tema/simge.gif"> <a href="fixnews.asp?action=AllNews" target="page">Sabit Haberler</a><br>
<img border="0" src="Tema/simge.gif"> <a href="fixnews.asp?action=new_enter" target="page">Haber Ekle</a><br>
</td>
</tr>
</table>

</table>

<br>
<p align="center">
<%
End IF
IF Session("Level") = "1" OR Session("Level") = "2" OR Session("Level") = "7" OR Session("Level") = "9" Then
%> </p>
<table border="0" cellspacing="1" cellpadding="0" width="90%" bgcolor="#999999" align="center" class="td_menu">
<tr style="font-weight:bold">
  <td background="Tema/sidebox-title-bg.gif" height="20" align="center">
  <font face="Verdana">

  <img border="0" src="Tema/menu_bullet.gif" align="left"></font>DOSYALAR</td></tr>
<tr><td>
<table border="0" cellspacing="0" cellpadding="1" width="100%" bgcolor="#FFFFFF" class="td_menu2">
<tr>
<td>
<img border="0" src="Tema/simge.gif"> <a href="downloads.asp?action=NoCats" target="page">Dosyalar</a><br>
<img border="0" src="Tema/simge.gif"> <a href="downloads.asp?action=all" target="page">Kategoriler</a><br>
<img border="0" src="Tema/simge.gif"> <a href="downloads.asp?action=WaitApprove" target="page">Onay Bekleyenler</a><br><br>
<img border="0" src="Tema/simge.gif"> <a href="downloads.asp?action=Programs&page=New" target="page">Dosya Ekle</a><br>
<img border="0" src="Tema/simge.gif"> <a href="downloads.asp?action=New_Cat" target="page">
Kategori Ekle</a>
</td>
</tr>
</table>
</td></tr>
</table>
<br>
<%
End IF
IF Session("Level") = "1" OR Session("Level") = "3" OR Session("Level") = "7" Then
%>
<table border="0" cellspacing="1" cellpadding="0" width="90%" bgcolor="#999999" align="center" class="td_menu">
<tr style="font-weight:bold">
  <td background="Tema/sidebox-title-bg.gif" height="20" align="center">
  <font face="Verdana">

  <img border="0" src="Tema/menu_bullet.gif" align="left"></font>YAZILAR</td></tr>
<tr><td>
<table border="0" cellspacing="0" cellpadding="1" width="100%" bgcolor="#FFFFFF" class="td_menu2">
<tr>
<td>
<img border="0" src="Tema/simge.gif">
<a target="page" href="articles.asp?action=NoCats">Yazýlar</a><br>
<img border="0" src="Tema/simge.gif"> <a href="articles.asp?action=all" target="page">Kategoriler</a><br>
<img border="0" src="Tema/simge.gif"> <a href="articles.asp?action=WaitApprove" target="page">Onay Bekleyenler</a><br><br>
<img border="0" src="Tema/simge.gif">
<a target="page" href="articles.asp?action=Articles&page=New">Yazý Ekle</a><br>
<img border="0" src="Tema/simge.gif">
<a target="page" href="articles.asp?action=NewCat">Kategori Aç</a><br>

</td>
</tr>
</table>
</td></tr>
</table>
<br>
<%
End IF
IF Session("Level") = "1" OR Session("Level") = "7" OR Session("Level") = "8" Then
%>
<table border="0" cellspacing="1" cellpadding="0" width="90%" bgcolor="#999999" align="center" class="td_menu">
<tr style="font-weight:bold">
  <td background="Tema/sidebox-title-bg.gif" height="20" align="center">
  <font face="Verdana">

  <img border="0" src="Tema/menu_bullet.gif" align="left"></font>RESÝM GALERÝSÝ</td></tr>
<tr><td bgcolor="#FFFFFF">
<table border="0" cellspacing="0" cellpadding="1" width="100%" bgcolor="#FFFFFF" class="td_menu2">
<tr>
<td>
<img border="0" src="Tema/simge.gif"> <a href="resimekle.asp" target="page">Rastgele Resim</a><br>
<img border="0" src="Tema/simge.gif"> <a href="gallery.asp?secil=katekle" target="page">Kategori Ekle</a><br>
<img border="0" src="Tema/simge.gif"> <a href="gallery.asp?secil=katsil" target="page">Kategori Sil</a><br>
<img border="0" src="Tema/simge.gif"> <a href="gallery.asp?secil=kartekle" target="page">Resim Ekle</a><br>
<img border="0" src="Tema/simge.gif"> <a href="gallery.asp?secil=kartsil" target="page">Resim Sil</a><br>
</td>
</tr>
</table>
</table>
<br>
<% 
End IF
IF Session("Level") = "1" OR Session("Level") = "7" Then
%>
<table border="0" cellspacing="1" cellpadding="0" width="90%" bgcolor="#999999" align="center" class="td_menu">
<tr style="font-weight:bold">
  <td background="Tema/sidebox-title-bg.gif" height="20" align="center">
  <font face="Verdana">

  <img border="0" src="Tema/menu_bullet.gif" align="left"></font>ÜYELIK</td></tr>
<tr><td>
<table border="0" cellspacing="0" cellpadding="1" width="100%" bgcolor="#FFFFFF" class="td_menu2">
<tr>
<td>
<img border="0" src="Tema/simge.gif"> <a href="membership.asp?action=members" target="page">Üye Listesi</a><br>
<img border="0" src="Tema/simge.gif"> <a href="membership.asp?action=AdminMsgs" target="page">Toplu Mesajlar</a><br><br>
<img border="0" src="Tema/simge.gif"> <a href="membership.asp?action=MsgToMembers" target="page">Toplu Mesaj Gönder</a><br>
<img border="0" src="Tema/simge.gif"> <a href="membership.asp?action=Mail2Members" target="page">Toplu Mail Gönder</a><br>
<img border="0" src="Tema/simge.gif"> <a href="membership.asp?action=DeleteMessages" target="page">Mesajlarý Sil</a>
</td>
</tr>
</table>
</td></tr>
</table>
<br>
<table border="0" cellspacing="1" cellpadding="0" width="90%" bgcolor="#999999" align="center" class="td_menu">
<tr style="font-weight:bold">
  <td background="Tema/sidebox-title-bg.gif" height="20" align="center">
  <font face="Verdana">

  <img border="0" src="Tema/menu_bullet.gif" align="left"></font>ANKETLER</td></tr>
<tr><td>
<table border="0" cellspacing="0" cellpadding="1" width="100%" bgcolor="#FFFFFF" class="td_menu2">
<tr>
<td>
<img border="0" src="Tema/simge.gif"> <a href="poll.asp?action=all" target="page">Anketler</a><br>
<img border="0" src="Tema/simge.gif"> <a href="poll.asp?action=New" target="page">Anket Ekle</a>
</td>
</tr>
</table>
</td></tr>
</table>
<br>
<table border="0" cellspacing="1" cellpadding="0" width="90%" bgcolor="#999999" align="center" class="td_menu">
<tr style="font-weight:bold">
  <td background="Tema/sidebox-title-bg.gif" height="20" align="center">
  <font face="Verdana">

  <img border="0" src="Tema/menu_bullet.gif" align="left"></font>SMILEY YÖNETIMI</td></tr>
<tr><td>
<table border="0" cellspacing="0" cellpadding="1" width="100%" bgcolor="#FFFFFF" class="td_menu2">
<tr>
<td>
<img border="0" src="Tema/simge.gif"> <a href="smileys.asp" target="page">Tüm Smililer</a><br>
<img border="0" src="Tema/simge.gif"> <a href="smileys.asp?op=New" target="page">Smili Ekle</a>
</td>
</tr>
</table>
</td></tr>
</table>
<br>
<table border="0" cellspacing="1" cellpadding="0" width="90%" bgcolor="#999999" align="center" class="td_menu">
<tr style="font-weight:bold">
  <td background="Tema/sidebox-title-bg.gif" height="20" align="center">
  <font face="Verdana">

  <img border="0" src="Tema/menu_bullet.gif" align="left"></font>AVATAR YÖNETIMI</td></tr>
<tr><td>
<table border="0" cellspacing="0" cellpadding="1" width="100%" bgcolor="#FFFFFF" class="td_menu2">
<tr>
<td>
<img border="0" src="Tema/simge.gif"> <a href="avatars.asp" target="page">Tüm Avatarlar</a><br>
<img border="0" src="Tema/simge.gif"> <a href="avatars.asp?op=New" target="page">Avatar Ekle</a>
</td>
</tr>
</table>
</td></tr>
</table>
<br>
<table border="0" cellspacing="1" cellpadding="0" width="90%" bgcolor="#999999" align="center" class="td_menu">
<tr style="font-weight:bold">
  <td background="Tema/sidebox-title-bg.gif" height="20" align="center">
  <font face="Verdana">

  <img border="0" src="Tema/menu_bullet.gif" align="left"></font>BANNER YÖNETIMI</td></tr>
<tr><td>
<table border="0" cellspacing="0" cellpadding="1" width="100%" bgcolor="#FFFFFF" class="td_menu2">
<tr>
<td>
<img border="0" src="Tema/simge.gif"> <a href="banners.asp?action=Banners" target="page">Bannerler</a><br>
<img border="0" src="Tema/simge.gif"> <a href="banners.asp?action=Banner&op=New" target="page">Banner Ekle</a>
</td>
</tr>
</table>
</td></tr>
</table>
&nbsp;<table border="0" cellspacing="1" cellpadding="0" width="90%" bgcolor="#999999" align="center" class="td_menu">
<tr style="font-weight:bold">
  <td background="Tema/sidebox-title-bg.gif" height="20" align="center">
  <font face="Verdana">

  <img border="0" src="Tema/menu_bullet.gif" align="left"></font>TOPLIST</td></tr>
<tr><td bgcolor="#FFFFFF">
<table border="0" cellspacing="0" cellpadding="1" width="100%" bgcolor="#FFFFFF" class="td_menu2" height="42">
<tr>
<td height="40">
<img border="0" src="Tema/simge.gif">
<a target="page" href="toplistban.asp?action=all">Dost Bannerler</a><br>
<img border="0" src="Tema/simge.gif"> <a target="page" href="linkler.asp">
Linkler</a><br>
<img border="0" src="Tema/simge.gif">
<a target="page" href="ufakbanner.asp?action=all">Toplistler</a></td>
</tr>
</table>
</table>
<br>
<p align="center">
<table border="0" cellspacing="1" cellpadding="2" width="90%" bgcolor="#999999" align="center" class="td_menu">
<tr style="font-weight:bold" bgcolor="#FFFFFF">
  <td background="Tema/sidebox-title-bg.gif" height="20" align="center">
  <font face="Verdana">

  <img border="0" src="Tema/menu_bullet.gif" align="left"></font>K.S.M.B</td></tr>
<tr><td bgcolor="#FFFFFF">
<table border="0" cellspacing="0" cellpadding="1" width="100%" bgcolor="#FFFFFF" class="td_menu2">
<tr>
<td>
<b>
<font face="Verdana">

<img border="0" src="Tema/arrow.gif"></font> Menü Kategorileri</b><br>
<img border="0" src="Tema/simge.gif"> <a href="menu_cats.asp" target="page">Tüm Kategoriler</a><br>
<img border="0" src="Tema/simge.gif"> <a href="menu_cats.asp?op=New" target="page">Yeni Kategori</a>
<br><br>
<b>
<font face="Verdana">

<img border="0" src="Tema/arrow.gif"></font> Linkler</b><br>
<img border="0" src="Tema/simge.gif"> <a href="menu.asp" target="page">Menü Linkleri</a><br>
<img border="0" src="Tema/simge.gif"> <a href="menu.asp?op=New" target="page">Yeni Menü Linki</a>
<br><br>
<b>
<font face="Verdana">

<img border="0" src="Tema/arrow.gif"></font> Sayfalar</b><br>
<img border="0" src="Tema/simge.gif"> <a href="pages.asp?action=all" target="page">Tüm Sayfalar</a><br>
<img border="0" src="Tema/simge.gif"> <a href="pages.asp?action=NewPage" target="page">Yeni Sayfa</a>
<br><br>
<b>
<font face="Verdana">

<img border="0" src="Tema/arrow.gif"></font> Modüller</b><br>
<img border="0" src="Tema/simge.gif"> <a href="modules.asp?action=all" target="page">Tüm Modüller</a><br>
<img border="0" src="Tema/simge.gif"> <a href="modules.asp?action=New" target="page">Yeni Modül</a>
<br><br>
<b>
<font face="Verdana">

<img border="0" src="Tema/arrow.gif"></font> Bloklar</b><br>
<img border="0" src="Tema/simge.gif"> <a href="blocks.asp?action=all" target="page">Tüm Bloklar</a><br>
<img border="0" src="Tema/simge.gif"> <a href="blocks.asp?action=New" target="page">Yeni Blok</a>
</td>
</tr>
</table>
</td></tr>
</table>
<br>
<table border="0" cellspacing="1" cellpadding="2" width="90%" bgcolor="#999999" align="center" class="td_menu">
<tr style="font-weight:bold" bgcolor="#FFFFFF">
  <td background="Tema/sidebox-title-bg.gif" height="20" align="center">
  <font face="Verdana">

  <img border="0" src="Tema/menu_bullet.gif" align="left"></font>DIGERLERI...</td></tr>
<tr><td bgcolor="#FFFFFF">
<table border="0" cellspacing="0" cellpadding="1" width="100%" bgcolor="#FFFFFF" class="td_menu2">
<tr>
<td>
<img border="0" src="Tema/simge.gif"> <a href="upload/default.asp" target="page">UPLOAD</a><br>
<img border="0" src="Tema/simge.gif"> <a href="notices.asp" target="page">Duyurular</a><br>
<img border="0" src="Tema/simge.gif"> <a href="iletisim.asp" target="page">Iletiþim</a><br>
<img border="0" src="Tema/simge.gif"> <a href="swearwords.asp" target="page">Küfür Korumasý</a>
</td>
</tr>
</table>
</td></tr>
</table>
<%
End IF
ElseIF Request.QueryString("pane") = "Main" Then
Session.LCID = 1033
liste = DateAdd("n", -1 * 1, Now())
Set om = Connection.Execute("SELECT * FROM MEMBERS where son_tarih >= #"&liste&"# ORDER BY son_tarih DESC")
Session.LCID = 2048

Set getInfo = Connection.Execute("Select * FROM SETTINGS")

Set total_login = Connection.Execute("SELECT SUM(login_sayisi) AS count FROM MEMBERS")
Set mems = Connection.Execute("Select Count(*) AS count FROM MEMBERS")
Set mem_msg = Connection.Execute("Select Count(*) AS count FROM MESSAGES WHERE SEE = True AND TYPE = 0")
Set admins = Connection.Execute("Select Count(*) AS count FROM MEMBERS WHERE seviye = 1")
Set eprograms = Connection.Execute("Select Count(*) AS count FROM MEMBERS WHERE seviye = 2")
Set earticle = Connection.Execute("Select Count(*) AS count FROM MEMBERS WHERE seviye = 3")
Set enews = Connection.Execute("Select Count(*) AS count FROM MEMBERS WHERE seviye = 4")
Set eforum = Connection.Execute("Select Count(*) AS count FROM MEMBERS WHERE seviye = 5")

Set news_cats = Connection.Execute("Select Count(*) AS count FROM NEWS_CATS")
Set news = Connection.Execute("Select Count(*) AS count FROM NEWS WHERE onay = True")
Set wa_news = Connection.Execute("Select Count(*) AS count FROM NEWS WHERE onay = False")
Set cnews = Connection.Execute("Select Count(*) AS count FROM COMMENTS WHERE page = 'News'")
Set news_read = Connection.Execute("Select SUM(okunma) AS count FROM NEWS")

Set downs = Connection.Execute("Select Count(*) AS count FROM DOWNLOADS WHERE p_approved = True")
Set wa_downs = Connection.Execute("Select Count(*) AS count FROM DOWNLOADS WHERE p_approved = False")
Set totald = Connection.Execute("SELECT SUM(p_hit) AS count FROM DOWNLOADS")

Set articles = Connection.Execute("Select Count(*) AS count FROM ARTICLES WHERE a_approved = True")
Set wa_articles = Connection.Execute("Select Count(*) AS count FROM ARTICLES WHERE a_approved = False")
Set articlesh = Connection.Execute("SELECT SUM(hit) AS count FROM ARTICLES")

Set downs_comments = Connection.Execute("Select Count(*) AS count FROM COMMENTS WHERE page = 'programs'")
Set articles_comments = Connection.Execute("Select Count(*) AS count FROM COMMENTS WHERE page = 'articles'")
%>
<div class="td_menu"><font face="Verdana" style="font-size: 10pt"><b>» Site Hakkýnda</b></font></div>
<table border="0" cellspacing="0" cellpadding="0" width="100%" class="td_menu2" style="font-size:12px; border-collapse:collapse" bordercolor="#111111">
<tr style="font-weight:bold" height="20">
<td width="25%" background="Tema/Arkaplan.gif">
<font face="Verdana" style="font-size: 9pt">
<img border="0" src="Tema/istatistik.gif" width="24" height="16">Istatistik</font></td>
<td width="25%" align="center" background="Tema/Arkaplan.gif">
<b>
<font face="Verdana">

<img border="0" src="Tema/arrow.gif"></font></b><font face="Verdana">Deðer</font></td>
<td width="25%" background="Tema/Arkaplan.gif">
<font face="Verdana" style="font-size: 9pt">
<img border="0" src="Tema/istatistik.gif" width="24" height="16"></font><font face="Verdana">Istatistik</font></td>
<td width="25%" align="center" background="Tema/Arkaplan.gif">
<b>
<font face="Verdana">

<img border="0" src="Tema/arrow.gif"></font></b><font face="Verdana">Deðer</font></font></td>
</tr>
<tr bgcolor="#CCCCCC" style="font-weight:bold">
<td colspan="4">
<b>
<font face="Verdana">

<img border="0" src="Tema/arrow.gif"></font></b><font face="Verdana" style="font-size: 9pt"> 
Üyelik</font></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="25%">
<font face="Verdana">

<img border="0" src="Tema/arrow2.gif"></font><b><font face="Verdana" style="font-size: 9pt">Üye Sayýsý</font></b></td>
<td width="25%" align="center"><%=mems("count")%></td>
<td width="25%"></font>
<font face="Verdana">

<img border="0" src="Tema/arrow2.gif"></font><b><font face="Verdana" style="font-size: 9pt">Giriþ Yapýlma Sayýsý</font></b></td>
<td width="25%" align="center"><%=total_login("count")%></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="25%">
<font face="Verdana">

<img border="0" src="Tema/arrow2.gif"></font><b><font face="Verdana" style="font-size: 9pt">Yönetici Sayýsý</font></b></td>
<td width="25%" align="center"><%=admins("count")%></td>
<td width="25%"></font>
<font face="Verdana">

<img border="0" src="Tema/arrow2.gif"></font><b><font face="Verdana" style="font-size: 9pt">Download Editörü Sayýsý</font></b></td>
<td width="25%" align="center"><%=eprograms("count")%></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="25%">
<font face="Verdana">

<img border="0" src="Tema/arrow2.gif"></font><b><font face="Verdana" style="font-size: 9pt">Makale Editörü Sayýsý</font></b></td>
<td width="25%" align="center"><%=earticle("count")%></td>
<td width="25%"></font>
<font face="Verdana">

<img border="0" src="Tema/arrow2.gif"></font><b><font face="Verdana" style="font-size: 9pt">Haber Editörü Sayýsý</font></b></td>
<td width="25%" align="center"><%=enews("count")%></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="25%">
<font face="Verdana">

<img border="0" src="Tema/arrow2.gif"></font><b><font face="Verdana" style="font-size: 9pt">Forum Yöneticisi Sayýsý</font></b></td>
<td width="25%" align="center"><%=eforum("count")%></td>
<td width="25%" bgcolor="#FFFFFF"></td>
<td width="25%" bgcolor="#FFFFFF"></td>
</tr>
<tr><td colspan="4"></td></tr>
<tr bgcolor="#CCCCCC" style="font-weight:bold">
<td colspan="4">
<b>
<font face="Verdana">

<img border="0" src="Tema/arrow.gif"></font></b><font face="Verdana" style="font-size: 9pt"> 
Haberler</font></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="25%">
<font face="Verdana">

<img border="0" src="Tema/arrow2.gif"></font><b><font face="Verdana" style="font-size: 9pt">Kategori Sayýsý</font></b></td>
<td width="25%" align="center"><%=news_cats("count")%></td>
<td width="25%"></font>
<font face="Verdana">

<img border="0" src="Tema/arrow2.gif"></font><b><font face="Verdana" style="font-size: 9pt">Haber Sayýsý</font></b></td>
<td width="25%" align="center"><%=news("count")%></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="25%">
<font face="Verdana">

<img border="0" src="Tema/arrow2.gif"></font><b><font face="Verdana" style="font-size: 9pt">Yorum Sayýsý</font></b></td>
<td width="25%" align="center"><%=cnews("count")%></td>
<td width="25%"></font>
<font face="Verdana">

<img border="0" src="Tema/arrow2.gif"></font><b><font face="Verdana" style="font-size: 9pt">Okunma Sayýsý</font></b></td>
<td width="25%" align="center"><%=news_read("count")%></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="25%">
<font face="Verdana">

<img border="0" src="Tema/arrow2.gif"></font><b><font face="Verdana" style="font-size: 9pt">Onay Bekleyen Haber Sayýsý</font></b></td>
<td width="25%" align="center"><%=wa_news("count")%></td>
<td width="25%" bgcolor="#FFFFFF"></td>
<td width="25%" bgcolor="#FFFFFF"></td>
</tr>
<tr><td colspan="4"></td></tr>
<tr bgcolor="#CCCCCC" style="font-weight:bold">
<td colspan="4">
<b>
<font face="Verdana">

<img border="0" src="Tema/arrow.gif"></font></b><font face="Verdana" style="font-size: 9pt"> 
Dosyalar</font></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="25%">
<font face="Verdana">

<img border="0" src="Tema/arrow2.gif"></font><b><font face="Verdana" style="font-size: 9pt">Toplam Dosya Sayýsý</font></b></td>
<td width="25%" align="center"><%=downs("count")%></td>
<td width="25%"></font>
<font face="Verdana">

<img border="0" src="Tema/arrow2.gif"></font><b><font face="Verdana" style="font-size: 9pt">Ýndirilme Sayýsý</font></b></td>
<td width="25%" align="center"><% IF Len(totald("count")) >= 1 Then %><%=totald("count")%><%ELSE%><%Response.Write "0"%><% End IF %></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="25%">
<font face="Verdana">

<img border="0" src="Tema/arrow2.gif"></font><b><font face="Verdana" style="font-size: 9pt">Yorum Sayýsý</font></b></td>
<td width="25%" align="center"><%=downs_comments("count")%></td>
<td width="25%"></font>
<font face="Verdana">

<img border="0" src="Tema/arrow2.gif"></font><b><font face="Verdana" style="font-size: 9pt">Onay Bekleyen Dosya Sayýsý</font></b></td>
<td width="25%" align="center"><%=wa_downs("count")%></td>
</tr>
<tr><td colspan="4"></td></tr>
<tr bgcolor="#CCCCCC" style="font-weight:bold">
<td colspan="4">
<b>
<font face="Verdana">

<img border="0" src="Tema/arrow.gif"></font></b><font face="Verdana" style="font-size: 9pt"> Makeleler</font></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="25%">
<font face="Verdana">

<img border="0" src="Tema/arrow2.gif"></font><b><font face="Verdana" style="font-size: 9pt">Toplam Makale Sayýsý</font></b></td>
<td width="25%" align="center"><%=articles("count")%></td>
<td width="25%"></font>
<font face="Verdana">

<img border="0" src="Tema/arrow2.gif"></font><b><font face="Verdana" style="font-size: 9pt">Okunma Sayýsý</font></b></td>
<td width="25%" align="center"><% IF Len(articlesh("count")) >= 1 Then %><%=articlesh("count")%><%ELSE%><%Response.Write "0"%><% End IF %></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="25%">
<font face="Verdana">

<img border="0" src="Tema/arrow2.gif"></font><b><font face="Verdana" style="font-size: 9pt">Yorum Sayýsý</font></b></td>
<td width="25%" align="center"><%=articles_comments("count")%></td>
<td width="25%"></font>
<font face="Verdana">

<img border="0" src="Tema/arrow2.gif"></font><b><font face="Verdana" style="font-size: 9pt">Onay 
Bekleyen Makale Sayýsý</font></b></td>
<td width="25%" align="center"><%=wa_articles("count")%></td>
</tr>
</table>
<hr size="1">
<div class="td_menu"><font face="Verdana" style="font-size: 10pt"><b>» Kimler Online ?</b></font></div>
<table border="0" cellspacing="0" cellpadding="0" width="100%" class="td_menu2" style="font-size:12px; border-collapse:collapse" bordercolor="#111111">
<tr style="font-weight:bold" height="20">
<td width="25%" background="Tema/arkaplan.jpg">
<font face="Verdana" style="font-size: 9pt"><img border="0" src="Tema/uye.gif">Kullanýcý Adý</font></td>
<td width="20%" align="center" background="Tema/arkaplan.jpg">
<font face="Verdana" style="font-size: 9pt">
<img border="0" src="Tema/oturum.gif" width="18" height="20"></font><font face="Verdana">Son Oturum Tarihi</font></td>
<td width="15%" align="center" background="Tema/arkaplan.jpg">
<font face="Verdana" style="font-size: 9pt">
<img border="0" src="Tema/IP.gif">
 </font><font face="Verdana">IP Adresi</font></td>
<td width="15%" align="center" background="Tema/arkaplan.jpg">
<font face="Verdana" style="font-size: 9pt">
<img border="0" src="Tema/arrow.gif">
 </font><font face="Verdana">Ýþletim Sistemi</font></td>
<td width="25%" align="center" background="Tema/arkaplan.jpg"></font>
<font face="Verdana" style="font-size: 9pt">
<img border="0" src="Tema/arrow.gif"> Browser</font></td>
</tr>
<% Do Until om.EoF %>
<tr bgcolor="#E0E0E0">
<td width="25%"><a href="membership.asp?action=member_details&mid=<%=om("uye_id")%>"><%=om("kul_adi")%></a></td>
<td width="20%" align="center"><%=om("son_tarih")%></td>
<td width="15%" align="center"><%=om("last_ip")%></td>
<td width="15%" align="center"><%=om("u_ws")%></td>
<td width="25%" align="center"><%=om("u_browser")%></td>
</tr>
<%
om.MoveNext
Loop
%>
</table>
<%
End IF
%>
<frameset rows="*" border="1" framespacing="0" frameborder="yes">
  <frameset cols="201,*" border="1" framespacing="0" frameborder="yes">
  <frame src="?pane=Left" name="nav" marginwidth="3" marginheight="3" scrolling="auto">
  <frame src="?pane=Main" name="page" marginwidth="10" marginheight="10" scrolling="auto">
  </frameset>
</frameset>
<% End If %>