
<!--#include file="includes.asp" -->
<%
If Session("Level") = "1" OR Session("Level") = "7" Then
islem = Request.QueryString("action")
if islem = "all" then
call hepsi
elseif islem = "duzenle" then
call duzenle
elseif islem = "guncelle" then
call guncelle
elseif islem = "yeni" then
call yeni
elseif islem = "sil" then
call sil
elseif islem = "kayit" then
call kayit
elseif islem = "change" then
call degis
end if
%>
<style>
a:link { font-family:tahoma; font-size:11px }
a:visited { font-family:tahoma; font-size:11px }
a:hover { font-family:tahoma; font-size:11px }
a:aktif { font-family:tahoma; font-size:11px }
</style>
<%
Sub degis

banner = Request.QueryString("bid")
atraksiyon = Request.QueryString("atraksiyon")
Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM UFAKBAN where banid = "&banner
rs.open SQL,Connection,1,3

if atraksiyon = "aktif" then
rs("aktif") = True
elseif atraksiyon = "pasif" then
rs("aktif") = False
end if
rs.Update
response.redirect "?action=all"

End Sub
Sub hepsi

set rs = Server.CreateObject("ADODB.RecordSet")
sorgu = "select * from UFAKBAN where aktif= True"
rs.open sorgu,Connection,1,3
%>
<center><a href="?action=yeni">Yeni Toplist Banner Ekle</a></center><br><br>
<center><font face="Tahoma" size="3" color="navy"><b>Yayýndaki </b></font><b>
<font face="Tahoma" color="#000080">Toplist Bannerler</font></b><br>
<table border="0" cellspacing="1" cellpading="1" bgcolor="#000000" align="center" width="238">
<% if rs.eof or rs.bof then
response.write "<font face=verdana size=1>Kayýtlý Banner Yok !</font>"
end if %>
<% do while not rs.eof %>
<tr><td bgcolor="#F4F4F4" width="110">
  <img src="<%=rs("resurl")%>" width="88" height="31"></td>
  <td align="center" bgcolor="#F4F4F4" width="374"><a href="?action=duzenle&bid=<%=rs("banid")%>">Düzenle</a><br><a href="?action=sil&bid=<%=rs("banid")%>">Sil</a><br><a href="?action=change&atraksiyon=pasif&bid=<%=rs("banid")%>">Pasifleþtir</a></td></tr>
<% rs.MoveNext
Loop %>
</table>
<%
rs.Close
set rs = Nothing
%><p><br><br>
<%
set rs = Server.CreateObject("ADODB.RecordSet")
sorgu = "select * from UFAKBAN where aktif= False"
rs.open sorgu,Connection,1,3
%> </p>
<center><font face="Tahoma" size="3" color="navy"><b>Pasif </b></font><b>
<font face="Tahoma" color="#000080">Toplist Bannerler</font></b><br>
<table border="0" cellspacing="1" cellpading="1" bgcolor="#000000" align="center" width="236">
<% if rs.eof or rs.bof then
response.write "<font face=verdana size=1>Yayýnlanmayan Banner Yok !</font>"
end if
Do while not rs.eof %>
<tr><td bgcolor="#F4F4F4" width="91">
  <img src="<%=rs("resurl")%>" width="88" height="31"></td>
  <td align="center" bgcolor="#F4F4F4" width="138"><a href="?action=duzenle&bid=<%=rs("banid")%>">Düzenle</a> - <a href="?action=sil&bid=<%=rs("banid")%>">Sil</a><br><a href="?action=change&atraksiyon=aktif&bid=<%=rs("banid")%>">Aktifleþtir</a></td></tr>
<% rs.MoveNext
Loop %>
</table>
<%
rs.Close
set rs = Nothing
%><hr size="2" color="#CC0000"><center>&nbsp;</center>
<%
End Sub
Sub sil

bid = Request.QueryString("bid")
set del = Server.CreateObject("ADODB.RecordSet")
sorgu = "DELETE * FROM UFAKBAN where banid = "&bid
del.open sorgu,Connection,1,3

response.redirect "?action=all"

End Sub
Sub duzenle

bid = Request.QueryString("bid")
set rs = Server.CreateObject("ADODB.RecordSet")
sorgu = "Select * FROM UFAKBAN where banid = "&bid
rs.open sorgu,Connection,1,3

if rs("aktif") = True then
aktif = "checked"
else
aktif = ""
end if
If rs("pozisyon") = "top" Then
ust = "selected"
ElseIf rs("pozisyon") = "bottom" Then
alt = "selected"
Else
ust = "selected"
End If
%>
<form action="?action=guncelle&bid=<%=bid%>" method="post">
<table border="0" cellspacing="1" cellpading="1" width="691" align="center">
<tr><td width="151"><font face="Arial" size="2"><b>Banner URL : </b></font></td>
  <td width="533">
  <input type="text" name="resurl" value="<%=rs("resurl")%>" class="form1" size="20"><b><font face="Tahoma" size="1">(Baþýna 
  &quot;http://&quot; 
  Koyunuz)</font></b></td></tr>
<tr><td width="151"><font face="Arial" size="2"><b>Gidilecek URL : </font> </b></td>
  <td width="533">
  <input type="text" name="giturl" value="<%=rs("giturl")%>" class="form1" size="20"><b><font face="Tahoma" size="1">(Baþýna 
  &quot;http://&quot; 
  Koyunuz)</font></b></td></tr>
<tr><td width="151"><font face="Arial" size="2"><b>Aktif / Pasif : </b></font></td>
  <td width="533">
  <input type="checkbox" name="aktif" <%=aktif%> value="ON"></td></tr>
<tr><td colspan="2" width="687"><input type="submit" value="Güncelle" class="submit" style="width: 100%"></td></tr>
</table></form><br><br><center><a href="?action=all"><< Geri</a></center>
<%
rs.Close
set rs = Nothing

End Sub
Sub guncelle

bid = Request.QueryString("bid")
set rs = Server.CreateObject("ADODB.RecordSet")
sorgu = "Select * FROM UFAKBAN where banid = "&bid
rs.open sorgu,Connection,1,3

resurl = Request.Form("resurl")
giturl = Request.Form("giturl")
restik = Request.Form("restik")
konum = Request.Form("konum")
sifre = Request.Form("siffre")
kontor = Request.Form("kontor")
gosterim = Request.Form("gosterim")

if Request.Form("aktif") = "ON" then
aktif = True
else
aktif = False
end if

rs("resurl") = resurl
rs("giturl") = giturl
rs("restik") = restik
rs("pozisyon") = konum
rs("aktif") = aktif
rs("siffre") = sifre
rs("sinir") = kontor
rs("gor") = gosterim
rs.Update

response.write "<center><font face=tahoma size=2>Banner Bilgileri Güncellendi !!!</font><br><br><a href=?action=all><< Geri</a></center>"

End Sub
Sub yeni
%>
<form action="?action=kayit" method="post">
<table border="0" cellspacing="1" cellpading="1" width="692" align="center">
<tr><td width="238"><font face="Arial" size="2"><b>Banner URL : </b></font></td>
  <td width="530">
  <input type="text" name="resurl" class="form1" size="20"><b><font face="Tahoma" size="1">(Baþýna 
  &quot;http://&quot; 
  Koyunuz)</font></b></td></tr>
<tr><td width="238"><font face="Arial" size="2"><b>Gidilecek URL : </b></font></td>
  <td width="530">
  <input type="text" name="giturl" class="form1" size="20"><b><font face="Tahoma" size="1">(Baþýna 
  &quot;http://&quot; 
  Koyunuz)</font></b></td></tr>
<tr><td width="238"><font face="Arial" size="2"><b>Aktif / Pasif : </b></font></td>
  <td width="530">
  <input type="checkbox" name="aktif" value="ON" checked></td></tr>
<tr><td colspan="2" width="752"><input type="submit" value="Kaydet" class="submit" style="width: 100%"></td></tr>
</table></form><br><br><center><a href="?action=all"><< Geri</a></center>
<%
End Sub
Sub kayit

set rs = Server.CreateObject("ADODB.RecordSet")
sorgu = "Select * FROM UFAKBAN"
rs.open sorgu,Connection,1,3

resurl = Request.Form("resurl")
giturl = Request.Form("giturl")
konum = Request.Form("konum")
sifre = Request.Form("siffre")
kontor = Request.Form("kontor")

if Request.Form("aktif") = "ON" then
aktif = True
else
aktif = False
end if

rs.AddNew
rs("resurl") = resurl
rs("giturl") = giturl
rs("restik") = 0
rs("pozisyon") = konum
rs("aktif") = aktif
rs("siffre") = sifre
rs("gor") = 0
rs("sinir") = kontor
rs.Update

response.write "<center><font class=""td_menu"">Banner Eklendi !!!</font><br><br>Banner ID : <b>"&rs("banid")&"</b><br>Þifre : <b>"&rs("siffre")&"</b><br><br><a href=?action=all><< Geri</a></center>"
End Sub
End If
%>