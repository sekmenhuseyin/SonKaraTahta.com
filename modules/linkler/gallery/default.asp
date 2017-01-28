<!--#include file="includes.asp"-->
<% IF Request.QueryString("menu")="category" Then %>
<% kat=request.querystring("kat") %>
<%
set liste=Server.CreateObject("ADODB.RecordSet")
SQL_L="Select * from fotograf where kategori ='" & kat &"' ORDER by fno desc"
liste.open SQL_L,Connection,1,3
%>
<%
liste.pagesize=ksay
liste.cachesize=ksay
%>
<% 
If Request.QueryString("Kayit")="" Then
	gosterilen_kayit=1
Else
	gosterilen_kayit=CInt(Request.QueryString("Kayit"))
End If %>
<%
toplam_kayit=liste.PageCount
msg_sayisi= liste.recordcount
If gosterilen_kayit > Toplam_Kayit Then gosterilen_kayit=toplam_kayit
If gosterilen_kayit < 1 Then gosterilen_kayit=1
If Toplam_Kayit=0 Then
	Response.Write "Kayýt bulunamadý!" %>
	<br /><br /><a href='javascript:history.go(-1);'><- Geri</a>
	<%
Else
	liste.AbsolutePage=gosterilen_kayit 
%>
<%
End If
%>
<% 
page_bilgi=(gosterilen_kayit - 1) * ksay %>
<table width="100%" cellspacing="0" cellpadding="10" align="center">
<% do while kayitsayi < ksay and not liste.eof %>
	<tr>
		<td width="20%">
<% if page_bilgi < msg_sayisi then %><div align="center">
<a href="modules/gallery/show.asp?menu=show&imgid=<%=liste("fno")%>" target="_blank"><img src="<%=liste("kucuk")%>" width=100 height=100 border="0"></a><br /><font face="Arial, Helvetica, sans-serif" size="1"><%=liste("hit")%> / <%=liste("tarih")%> / <%=liste("isim")%> / <a href="modules.asp?name=<%=Request.QueryString("name")%>&menu=yorum&imgid=<%=liste("fno")%>" title="Yorum Oku-Yaz">+</a></font></div>
<%  kayitsayi=kayitsayi+1
page_bilgi=page_bilgi+1
liste.movenext %>
<% else exit do %>
&nbsp;
<% end if %>
		</td>
		<td width="20%">
<% if page_bilgi < msg_sayisi then %><div align="center">
<a href="modules/gallery/show.asp?menu=show&imgid=<%=liste("fno")%>" target="_blank"><img src="<%=liste("kucuk")%>" width=100 height=100 border="0"></a><br /><font face="Arial, Helvetica, sans-serif" size="1"><%=liste("hit")%> / <%=liste("tarih")%> / <%=liste("isim")%> / <a href="modules.asp?name=<%=Request.QueryString("name")%>&menu=yorum&imgid=<%=liste("fno")%>" title="Yorum Oku-Yaz">+</a></font></div>
<%  kayitsayi=kayitsayi+1
page_bilgi=page_bilgi+1
liste.movenext
%>
<% else exit do %>
&nbsp;
<% end if %>		
		</td>
		<td width="20%">
<% if page_bilgi < msg_sayisi then %><div align="center">
<a href="modules/gallery/show.asp?menu=show&imgid=<%=liste("fno")%>" target="_blank"><img src="<%=liste("kucuk")%>" width=100 height=100 border="0"></a><br /><font face="Arial, Helvetica, sans-serif" size="1"><%=liste("hit")%> / <%=liste("tarih")%> / <%=liste("isim")%> / <a href="modules.asp?name=<%=Request.QueryString("name")%>&menu=yorum&imgid=<%=liste("fno")%>" title="Yorum Oku-Yaz">+</a></font></div>
<%  kayitsayi=kayitsayi+1
page_bilgi=page_bilgi+1
liste.movenext
%>
<% else exit do %>
&nbsp;
<% end if %>			
		</td>
		<td width="20%">
<% if page_bilgi < msg_sayisi then %>
<a href="modules/gallery/show.asp?menu=show&imgid=<%=liste("fno")%>" target="_blank"><img src="<%=liste("kucuk")%>" width=100 height=100 border="0"></a><div align="center"><font face="Arial, Helvetica, sans-serif" size="1"><%=liste("hit")%> / <%=liste("tarih")%> / <%=liste("isim")%> / <a href="modules.asp?name=<%=Request.QueryString("name")%>&menu=yorum&imgid=<%=liste("fno")%>" title="Yorum Oku-Yaz">+</a></font></div>
<%  kayitsayi=kayitsayi+1
page_bilgi=page_bilgi+1
liste.movenext
%>
<% else exit do %>
&nbsp;
<% end if %>			
		</td>
	</tr>
<% loop %>
</table>
<br />

<center>
<% For k=1 To Toplam_Kayit %> 
<% If k=Gosterilen_Kayit Then %> 
<%=k %> 
<% Else %> 
<b><a href='?name=<%=Request.QueryString("name")%>&menu=category&Kayit=<%=k %>&kayitsayi=<%=kayitsayi %>&kat=<%= kat %>'><%=k %></a></b> 
<% End If %> 
<% Next %>
<% liste.close
set liste=nothing
Connection.close
set Connection=nothing
 %><br /><a href='javascript:history.go(-1);'><- Geri</a></center>
<% ElseIF Request.QueryString("menu")="yorum" Then %>
<% islem=request.querystring("islem") %>
<% imgid=request.querystring("imgid") %>
<% if islem="gir" then call gir
if islem="hata" then call hata
if islem="" then %>
<% sec_code=CPASS(4):Session("EnterSecurityCode")=sec_code %>
    <div align="center"><center>
<form method="POST" action="?name=<%=Request.QueryString("name")%>&menu=yorum&islem=gir">
    <table border="0"  
  cellspacing="0" cellpadding="2" width="300" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td width="25%"><small>Ýsim :</small></td>
      <td><input type="text" name="isim" size="20"
class="form1"></td>
    </tr>
<tr>
      <td valign="top"><small>Yorum :</small></td>
      <td><textarea rows="7" cols="28" name="mesaj" class="form1""></textarea></td>
    </tr> 
</table>
<br />
<input type="hidden" value="<%=imgid%>" name="imgid">
<input type="submit" value="Yaz" class="submit" Onclick="javascript:this.form.submit();this.disabled=true; return true;">
</form>
<br />
<%
set liste=Server.CreateObject("ADODB.RecordSet")
sql_l="Select * from yorum where onay =" & 1 &" and hangi ="& imgid &""
liste.open sql_l,Connection,1,3
%>
<% for i=1 to liste.recordcount %>
<b><%= liste("isim") %></b> / <%= liste("tarih") %>
<br />
<%= liste("yorum") %>
<hr width="300" align=center>
<%
liste.movenext
Next
liste.close
set liste=Nothing
%>
<% end if %>

<% sub gir %>
<% isim=request.form("isim")
imgid=request.form("imgid")
mesaj=request.form("mesaj")
mesaj=Replace(mesaj,Chr(13), "<br />")
que=Request.QueryString("name")
scode=duz(Request.Form("security_code"))
if isim="" then 
response.redirect "?name="& que &"&menu=yorum&islem=hata" 
ELSE

set kayit=Server.CreateObJect("ADODB.RecordSet")
SQL_k="Select * from yorum WHERE yno"
kayit.open SQL_k,Connection,1,3
kayit.addnew
kayit("isim")=isim
kayit("yorum")=mesaj
kayit("tarih")=Date
kayit("hangi")=imgid
kayit("onay")=0
kayit.update
kayit.close
END IF
%><table height='100%' align='center'><tr><td align='center'><font size="2">Ýletiniz Kaydedildi.<br />Onaylandýktan sonra eklenecektir.<br />
  </font><a href='?name=<%=Request.QueryString("name")%>'><font size="2">Geri</font></a></td></tr></table>"
<% end sub %>

<% sub hata %>
<table height='100%' align='center'><tr><td align="center"><font size="2">Hata! Ýsim belirtmediniz.<br />
  </font><a href="javascript:history.go(-1);"><font size="2">Düzelt</font></a></td></tr></table>
<% end sub %>
<%
Else
set liste=Server.CreateObject("ADODB.RecordSet")
SQL_L="Select * from kategori order by kategori asc"
liste.open SQL_L,Connection,1,3
%>
<% msg_sayisi= liste.recordcount %>
<table cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" id="AutoNumber1">
<tr>
    <td width="100%" colspan="3">
    <p align="center"><b>Kategoriler</b></td>
  </tr> <% do while not liste.eof %> 
  <tr>
    <td width="34%" align="center"><% if kayitsayi < msg_sayisi then %>
<img src="Images/stats/over.gif">&nbsp<font size="2" color="#000000"><a href="?name=<%=Request.QueryString("name")%>&menu=category&kat=<%=liste("kategori")%>"><%=liste("kategori")%></a><br /></font><font size="2"><%=liste("statement")%></font>
<%  kayitsayi=kayitsayi+1
liste.movenext %> &nbsp;</td>
<% else %>
<% exit do %>
<% end if %>
    <td width="33%" align="center"><% if kayitsayi < msg_sayisi then %>
<img src="Images/stats/over.gif">&nbsp<font size="2" color="#000000"><a href="?name=<%=Request.QueryString("name")%>&menu=category&kat=<%=liste("kategori")%>"><%=liste("kategori")%></a><br /></font><font size="2"><%=liste("statement")%></font>
<%  kayitsayi=kayitsayi+1
liste.movenext
%> &nbsp;</td><% else %>
<% end if %><td width="33%" align="center">
    <% if kayitsayi < msg_sayisi then %>
<img src="Images/stats/over.gif">&nbsp<font size="2" color="#000000"><a href="?name=<%=Request.QueryString("name")%>&menu=category&kat=<%=liste("kategori")%>"><%=liste("kategori")%></a><br /></font><font size="2"><%=liste("statement")%></font>
<%  kayitsayi=kayitsayi+1
liste.movenext
%> &nbsp;</td>
<% else %>
<% exit do %>
<% end if %>
<% loop %>
  </tr>
</table>

<% end if %>