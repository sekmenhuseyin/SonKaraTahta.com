<!--#include file="includes.asp" -->
<% Response.Buffer = True %>
<%IF Session("Level") = "1" OR Session("Level") = "7" OR Session("Level") = "8" Then%>

<html>
<head>
<title>Gallery manage</title> 
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<meta http-equiv="Content-Language" content="tr">

</head>
<body>
<p>

<br>
<%
secil = request.querystring("secil")
if secil = "kartekle" then call kartekle
if secil = "kartekle2" then call kartekle2
if secil = "kartsil" then call kartsil
if secil = "kartsil2" then call kartsil2
if secil = "katekle" then call katekle
if secil = "katekle2" then call katekle2
if secil = "katsil" then call katsil
if secil = "katsil2" then call katsil2
if secil = "onay1" then call onay1
if secil = "onay2" then call onay2
if secil = "onay3" then call onay3
%>

<% sub kartekle %>
<%
set liste = Server.CreateObject("ADODB.RecordSet")
SQL_L = "Select * from kategori"
liste.open SQL_L,Connection,1,3
%> </p>
<form method="post" action="gallery.asp?secil=kartekle2">
<table width="300" align="center">
<tr> 
<td width="25%"> 
<div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>İsim:</b></font></div>
</td>
<td><input type="text" name="kart_adi" size="20" class="form1"></td>
</tr>
<tr> 
<td width="25%"> 
<div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Küçük</b></font></div>
</td>
<td><font face="Verdana"><input type="text" name="kucuk_resim" size="20" class="form1"></font><font size="2" face="Verdana"> URL olarak</font></td>
</tr>
<tr> 
<td width="25%"> 
<div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Büyük</b></font></div>
</td>
<td><font face="Verdana"><input type="text" name="buyuk_resim" size="20" class="form1"></font><font size="2" face="Verdana"> URL olarak</font></td>
</tr>
<tr> 
<td width="25%"> 
<div align="right"><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><b>Bilgi:</b></font></div>
</td>
<td><input type="text" name="aciklama" size="20" class="form1"></td>
</tr>
<tr> 
<td width="25%"> 
<div align="right"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><b>Kategori</b></font></div>
</td>
<td><select name="kategori" size="1" class="form1">
<%
while not liste.eof
%> 
<option value="<%= liste("kategori")%>"><%= liste("kategori")%> 
<%
liste.movenext
wend
%> 
</select>
</td>
</tr>
<tr> 
<td colspan="2"><div align="center"><input type="submit" name="Submit" value="ekle" class="submit"></div></td>
</tr>
</table>
</form>
<% liste.close
Connection.close
set Connection = Nothing %>
<% end sub %>


<% sub kartekle2 %>
<%
set liste = Server.CreateObject("ADODB.RecordSet")
SQL_L = "Select * from fotograf"
liste.open SQL_L,Connection,1,3
liste.AddNew 
liste("isim")=request("kart_adi")
liste("kucuk")=request("kucuk_resim")
liste("buyuk")=request("buyuk_resim")
liste("kategori")=request("kategori")
liste("aciklama")=request("aciklama")
liste("tarih") = date
liste.Update
liste.Close
Set liste = Nothing
Connection.close
set Connection = Nothing
response.redirect "gallery.asp?secil=kartekle"
%>
<% end sub %>

<% sub kartsil %>
<%
set liste = Server.CreateObject("ADODB.RecordSet")
SQL_L = "Select * from fotograf"
liste.open SQL_L,Connection,1,3
if liste.eof then
response.write "<font face=verdana size=2>Hiç Kayıt Yok</font>"
else
%>
<% do while not liste.eof %>
<br>
<font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=liste("fno")%> &nbsp; <%=liste("isim")%> &nbsp;  <%=liste("kucuk")%> &nbsp; [<a href="gallery.asp?secil=kartsil2&kart_id=<%=liste("fno")%>">Sil</a>]</font>
<%
liste.movenext
Loop
%>
<% 
liste.close
set liste = Nothing
%>
<% end if end sub %>


<% sub kartsil2 %>
<%
kart_id = Request.QueryString("kart_id")
if kart_id="" then
response.write "<font face=verdana size=2>Bol veri yollamışsınız...</font>"
response.end
end if
Set kayit = Server.CreateObject("ADODB.RecordSet")
SQL_sil = "DELETE FROM fotograf WHERE fno="&kart_id&""
kayit.open SQL_sil , Connection , 1 , 3
Connection.close
set Connection = Nothing
response.redirect "gallery.asp?secil=kartsil"
%>
<% end sub %>


<% sub katekle %>
<div align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><b>Kategori Ekleme Formu</font></b></div>
            <form method="post" action="gallery.asp?secil=katekle2">
              <table width="100%" border="0" align="center" height="122">
                <tr> 
                  <td width="24%" height="19"> 
                    <div align="right"><font face="Verdana, Arial, Helvetica, sans-serif"><b><font size="1">Kategori 
                      Adı :</font></b></font></div>
                  </td>
                  <td width="51%" height="19"> <font face="Verdana, Arial, Helvetica, sans-serif"><b><font size="1"> 
                    <input type="text" name="kategori" size="20" class="form1">
                    </font></b></font></td>
                </tr>
                <tr> 
                  <td colspan="2" height="2"> 
                    <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif"><b><font size="1"> 
                      <input type="submit" name="Submit" value="ekle" class="submit">
                      </font></b></font><font face="Verdana, Arial, Helvetica, sans-serif"><b><font size="1"></font></b></font></div>
                  </td>
                </tr>
              </table>
            </form>
<% end sub %>

<% sub katekle2 %>
<%
set liste = Server.CreateObject("ADODB.RecordSet")
SQL_L = "Select * from kategori"
liste.open SQL_L,Connection,1,3
liste.AddNew 
liste("kategori")=request("kategori")
liste.Update
liste.Close
Set liste = Nothing
Connection.close
set Connection = Nothing
response.redirect"gallery.asp?secil=katekle"
%>
<% end sub %>

<% sub katsil %>
<%
set kayit = Server.CreateObject("ADODB.RecordSet")
SQL_L = "Select * from kategori"
kayit.open SQL_L,Connection,1,3
if kayit.eof then
response.write "<font face=verdana size=2>Hiç Kayıt Yok</font>"
else
%>
<% do while not kayit.eof %>
<br>
<b><font face="Verdana, Arial, Helvetica, sans-serif" size="1"><%=kayit("kno")%> &nbsp; <%=kayit("kategori")%> &nbsp; [<a href="gallery.asp?secil=katsil2&kat_id=<%=kayit("kno")%>">Sil</a>]</font>
<%
kayit.movenext
Loop
%>
<% 
kayit.close
set kayit = Nothing
Connection.close
set Connection = Nothing
end if
end sub
%>

<% sub katsil2 %>
<%
kat_id = Request.QueryString("kat_id")
if kat_id="" then
response.write "<font face=verdana size=2>Bol veri yollamışsınız...</font>"
response.end
end if
Set kayit = Server.CreateObject("ADODB.RecordSet")
SQL_sil = "DELETE FROM kategori WHERE kno="&kat_id&""
kayit.open SQL_sil , Connection , 1 , 3
Connection.close
set Connection = Nothing
response.redirect "gallery.asp?secil=katsil"
%>
<% end sub %>

<% sub onay1 %>
<%
set liste = Server.CreateObJect("ADODB.RecordSet")
SQL_i = "Select * from yorum where onay=" & 0
liste.open SQL_i,Connection,1,3
if liste.eof or liste.bof then response.write "<center>Onaylanacak Yeni Basvuru Yok.</center>"
%>
<% for i = 1 to liste.recordcount %><center><font face="Arial, Helvetica, sans-serif" size="2">
<b><%= liste("isim") %></b> / <%= liste("tarih") %> &nbsp; &nbsp; [ <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?secil=onay2&onayno=<%=liste("yno")%>">+</a>
 &nbsp; &nbsp; <a href="<%=Request.ServerVariables("SCRIPT_NAME")%>?secil=onay3&onayno=<%=liste("yno")%>">-</a> ]
<br>
<%= liste("yorum") %></font>
<hr width="300">
</center>
<%
liste.movenext
Next
liste.close
set liste = Nothing
%>
<% end sub %>

<% sub onay2 %>
<%
onayno = request.querystring("onayno")
sehir = request.querystring("sehir")
set liste = Server.CreateObJect("ADODB.RecordSet")
SQL_i = "Select * from yorum where yno=" & onayno
liste.open SQL_i,Connection,3,3
liste("onay") = "1"
liste.update
liste.close
response.redirect Request.ServerVariables("SCRIPT_NAME")
%>
<% end sub %>

<% sub onay3 %>
<%
onayno = request.querystring("onayno")
Set sil = Server.CreateObJect("ADODB.RecordSet")
Session.LCID=1033 
sql="delete * from yorum where yno=" & onayno
sil.open SQL,Connection,1,3
response.redirect Request.ServerVariables("SCRIPT_NAME")
%>
<% end sub %> </b>              
</body>
</html>
<% 
End If %>