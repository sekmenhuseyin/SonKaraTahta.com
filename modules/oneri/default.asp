<html>
<head>
<meta http-equiv="Content-Type" content="text/html;charset=ISO-8859-9">
<meta http-equiv="Content-Type" content="text/html;charset=windows-1254">
</head>
<% @ LANGUAGE=VBScript Codepage=1254 %>
<%
oneri=Server.HTMLEncode(Request.Form("oneri"))
ad=Server.HTMLEncode(Request.Form("ad"))
soyad=Server.HTMLEncode(Request.Form("soyad"))
tarih=date()
%>
<body>
<FORM Action="modules.asp?name=�nerileriniz" method="post">
<B>L�tfen Sayfam�z Hakk�ndaki De�erlendirme ve �nerilerinizi Belirtiniz.</B> :<br />
Ad�n�z:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<INPUT Name="ad" size="10" maxlength=12><br />
Soyad�n�z:&nbsp;&nbsp;&nbsp; <INPUT Name="soyad" size="10" maxlenght=12><br />
�nerileriniz:&nbsp;<INPUT Name="oneri" size="10" maxlength=50 style="HEIGHT: 22px; WIDTH: 470px">
<INPUT type="submit" value="G�nder" id=submit1 name=submit1>
<hr />
</FORM>

<%
Set dn=Server.CreateObject ("Scripting.FileSystemObject")
Set Setdn=dn.OpenTextFile(Server.MapPath("db/zdefter.txt"), 8)

if oneri="" then
Response.Write "<b>L�tfen �neri ve Yorumlar�n�z� Yukar�daki Alana Yaz�n�z.</b>"
else
if ad="" then
Response.Write "<b>L�tfen Ad�n�z� Yaz�n�z</b>"
else
if soyad="" then
Response.Write "<b>L�tfen Soyad�n�z� Yaz�n�z.</b>"
else

Setdn.WriteLine "<b>"&ad&" "&soyad&"</b><br />"
Setdn.WriteLine "<b>"&tarih&"</b><br />"
Setdn.WriteLine oneri
Setdn.WriteLine "<br /><hr />"
Response.Redirect "Modules/oneri/oneri.asp"
Response.Write"Te�ekk�r Ederiz"

Setdn.Close
Set Setdn=Nothing
Set dn=Nothing

end if
end if
end if
%>

<p><FONT color=orangered size=4><STRONG>Daha �nceki G�r�� ve �neriler</STRONG></FONT></p>
<!--#include file="../../db/zdefter.txt" -->
</BODY>
</HTML>

