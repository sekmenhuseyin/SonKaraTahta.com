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
<FORM Action="modules.asp?name=Önerileriniz" method="post">
<B>Lütfen Sayfamız Hakkındaki Değerlendirme ve Önerilerinizi Belirtiniz.</B> :<br />
Adınız:&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<INPUT Name="ad" size="10" maxlength=12><br />
Soyadınız:&nbsp;&nbsp;&nbsp; <INPUT Name="soyad" size="10" maxlenght=12><br />
Önerileriniz:&nbsp;<INPUT Name="oneri" size="10" maxlength=50 style="HEIGHT: 22px; WIDTH: 470px">
<INPUT type="submit" value="Gönder" id=submit1 name=submit1>
<hr />
</FORM>

<%
Set dn=Server.CreateObject ("Scripting.FileSystemObject")
Set Setdn=dn.OpenTextFile(Server.MapPath("db/zdefter.txt"), 8)

if oneri="" then
Response.Write "<b>Lütfen Öneri ve Yorumlarınızı Yukarıdaki Alana Yazınız.</b>"
else
if ad="" then
Response.Write "<b>Lütfen Adınızı Yazınız</b>"
else
if soyad="" then
Response.Write "<b>Lütfen Soyadınızı Yazınız.</b>"
else

Setdn.WriteLine "<b>"&ad&" "&soyad&"</b><br />"
Setdn.WriteLine "<b>"&tarih&"</b><br />"
Setdn.WriteLine oneri
Setdn.WriteLine "<br /><hr />"
Response.Redirect "Modules/oneri/oneri.asp"
Response.Write"Teşekkür Ederiz"

Setdn.Close
Set Setdn=Nothing
Set dn=Nothing

end if
end if
end if
%>

<p><FONT color=orangered size=4><STRONG>Daha Önceki Görüş ve Öneriler</STRONG></FONT></p>
<!--#include file="../../db/zdefter.txt" -->
</BODY>
</HTML>

