<% @Language=VBScript %>
<%
If session("Member") = "" then
Response.Write "<center><font class=""tabloalt_font"">Lütfen Giriş Yapınız.</font></center>"
Else
%>
<!-- #include file="includes.asp" -->
<%
IF Session("Level") = "1" OR Session("Level") = "7" OR Session("Level") = "11" Then
%>
<%
'   ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
'   Keyif Web Mesajlasma Scripti
'   Bu Script Keyifweb.com Admini Vatanay Özbeyli Tarafindan  yazilmiştir.
'   Scripti İstediğiniz Gibi Değiştirip Kullanabillirsiniz.<br>
'   Daha Fazlasi için WwW.Keyifweb.CoM 'a Uğramayi Unutmayiniz.
'   Not : Bu scripti para ile satmak yasaktir.
'   _______________________________________________________
%>
<html>

<head>
<meta http-equiv="Content-Language" content="tr">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Keyifweb Mesajlasma : : www.keyifweb.com</title>
<link rel="stylesheet" type="text/css" href="stil.css">
</head>

<body text="#000000" link="#000000" vlink="#000000" alink="#000000">
<div align="center">
  <center>
     <p><strong><font size="6" face="Verdana">Mini-NUKE</font><font size="7" face="Courier New, Courier, mono"><br>
      </font></strong><font size="7" face="Courier New, Courier, mono"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> <font color="#CCCCCC">--</font><font size="1"><font color="#999999">---</font><font color="#333333">------</font>--- 
      &lt; % <a href="http://Www.Mini-NUKE.De">Www.Mini-NUKE.De</a> iLeTiŞiM % &gt; <font color="#000000" face="Courier New, Courier, mono">---</font></font><font size="1" face="Courier New, Courier, mono"><font color="#333333">---</font><font color="#666666">---</font><font color="#999999">---</font><font color="#CCCCCC">---</font></font></font></font></p>
  </center>
</div>
<div align="center"> <font color="#000000"><font color="#333399" size="1" face="Verdana, Arial, Helvetica, sans-serif"><center> </font> 
  </font>
  <table border="0" cellpadding="0" cellspacing="0" width="600">
    <tr> 
      <td> <form method="POST" action="?islem=ekle">
          <div align="center"></div>
        </form></td>
    </tr>
  </table>
</div>
<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
<%
Set Sur = Server.CreateObject("ADODB.Connection")
Sur.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../db/mn"&dbNum&".mdb")

islem = Request.QueryString("islem")
if islem = "sil" then
call sil
elseif islem="yil" then
call yil
elseif islem="ekle" then
call ekle
else
end if
%>
<%
sub sil
id = Request("id")
Set keyifweb = Server.CreateObject("ADODB.RecordSet")
SQL_delete = "DELETE from mesaj WHERE id="&id&""
keyifweb.open SQL_delete,Sur,1,3
response.redirect Request.ServerVariables("HTTP_REFERER")
end sub
%>
</font> 
<div align="center"></div>
<div align="center"> 
  <center>
  <p></p>
  </center>
</div>
<div align="center"> 
  <center>
    <font color="#000000" face="Verdana, Arial, Helvetica, sans-serif"><br>
    <font size="1"><strong>DB'de kayitli Olan Mesajler </strong></font></font> 
    <table border="0" cellpadding="0" width="304" bgcolor="#FFFFFF">
      <tr> 
        <td width="300" bgcolor="#FFFFFF" valign="top"> <div align="center"> <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
            <%
Set keyifweb = Server.CreateObject("ADODB.Recordset")
sor = "Select * from mesaj order by id desc"
keyifweb.Open sor,Sur,1,3

if keyifweb.eof or keyifweb.bof then
Response.Write "<p align=left><b><font color=#1D5A87>»</font></b> Hiç mesaj yok...</p>"
end if
%>
            </font> 
            <table border="0" cellpadding="0" cellspacing="0" width="300">
              <%
for i = 1 to 5
if keyifweb.eof then exit for
%>
              <tr> 
                <td width="250" height="15"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <%=keyifweb("baslik")%> </font></td>
                <td width="250" <%=keyifweb("mesaj")%></td>
                <td width="50" height="15"><font color="#000000"><a href="?islem=sil&id=<%=keyifweb("id")%>">
                <img src="images/delete_icon.gif" border="0"></a></font></td>
              </tr>
              <%
keyifweb.movenext
Next
keyifweb.close
set keyifweb = Nothing
%>
            </table>
          </div></td>
      </tr>
    </table>
    <p>&nbsp;</p>
  </center>
</div>
<p align="center">&nbsp;</p>
<font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
<%sub ekle
Set Sur = Server.CreateObject("ADODB.Connection")
Sur.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../db/mn"&dbNum&".mdb")

Set gir = server. CreateObject("ADODB.Recordset")
kayit = "Select * from mesaj"
gir.Open kayit,sur,1,3

Dim ekleyen, email, baslik, mesaj, tarih

ekleyen = request.form("ekleyen")
email = request.form("email")
baslik = request.form("baslik")
mesaj = request.form("mesaj")
tarih = request.form("tarih")

if ekleyen="" or email="" or baslik="" or mesaj="" or tarih="" then
Response.Write "<center>Lütfen tüm alanları doldurunuz...</center>"
Response.End

else
gir.AddNew
gir("ekleyen") = ekleyen 
gir("email") = email
gir("baslik") = baslik
gir("mesaj") = mesaj
gir("tarih") = tarih
gir.Update

Response.Redirect Request.ServerVariables("HTTP_REFERER")

end if
end sub
%>
<%
End IF
%>
</font> 
<p align="center">&nbsp;<head>
<link rel="stylesheet" type="text/css" href="stil.css">
<%
'Baglantı dosyamızı include ettiğimize göre artık veritabanından bilgilerimizi alabiliriz
Set keyifweb = Server.CreateObject("ADODB.Recordset")
sor = "Select * from mesaj order by id desc"
keyifweb.Open sor,Sur,1,3

'Yapacağımız sayfalama için tanımlama
shf = Request.QueryString("shf")
if shf="" then 
shf=1
end if

%>
<meta http-equiv="Content-Language" content="tr">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Keyifweb Mesajlasma : : www.keyifweb.com</title></head>

<body text="#000000" link="#000000" vlink="#000000" alink="#000000">
<div align="center"> 
  <center>
    <font color="#1D4196" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
  <%'Veritabanında kayit yoksa eğer bunu belirtmemiz gerekir değil mi?
If keyifweb.eof then
Response.Write "<center>Veritabanında kayıtlı mesaj yok...</center>"
Response.End
End If
%>
    </font> 
    <table border="0" cellpadding="1" cellspacing="0" width="240">
      <% keyifweb.pagesize = 10 '1 sayfada görüntülemek istediğini mesaj sayısı (değiştirebilirsiniz)
keyifweb.absolutepage = shf
sayfa = keyifweb.pagecount
for i=1 to keyifweb.pagesize
if keyifweb.eof then exit for
%>
      <tr> 
        <td width="517" bgcolor="#CCCCCC">
<table width="387" border="0" cellpadding="2" cellspacing="0" bgcolor="#FFFFFF" height="123">
            <tr> 
              <td bgcolor="#F5F5F5" height="119" width="383">
<div align="center"> 
                  <p align="justify"><b>
                  <font face="Verdana, Arial, Helvetica, sans-serif" style="font-size: 9pt">
                  Mesaj Başlığı</font><font color="#1D4196" size="1" face="Verdana, Arial, Helvetica, sans-serif"><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif" style="font-size: 9pt">:</font><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> <%=keyifweb("baslik")%> 
                    <br></font></font></b><font color="#1D4196" size="1" face="Verdana, Arial, Helvetica, sans-serif"><font color="#1D4196" size="1" face="Verdana, Arial, Helvetica, sans-serif"><font color="#000000"><br>
                        </font>
                  </font></font><b>
                  <font face="Verdana, Arial, Helvetica, sans-serif" style="font-size: 9pt">
                  Mesaj</font></b><font color="#1D4196" size="1" face="Verdana, Arial, Helvetica, sans-serif"><font color="#000000" face="Verdana, Arial, Helvetica, sans-serif" style="font-size: 9pt"><b>:</b></font><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> <%=keyifweb("mesaj")%><br><br></font></font><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"><br>
                    </font>
                  <font color="#000000" face="Verdana, Arial, Helvetica, sans-serif" style="font-size: 9pt">
                  <b>Gönd.&amp; Maili:</b></font><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> <a href="mailto:<%=keyifweb("email")%>"><em><%=keyifweb("ekleyen")%></em></a><em>&nbsp;&nbsp; 
                    <strong>// </strong>    </font>
                  <font color="#000000" face="Verdana, Arial, Helvetica, sans-serif" style="font-size: 9pt">
                  <b>Tarih:</b></font><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">  <%=keyifweb("tarih")%> </font> </em><font color="#1D4196" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                    </font></p>
                  </div></td>
            </tr>
          </table>
        </td>
      </tr>
      <%keyifweb.movenext%>
      <% next %>
    </table>
  </center>
</div>
<div align="center"> 
  <center>
    <table width="216" height="22" border="0" cellpadding="0" cellspacing="0">
      <tr> 
        <td width="351"> <p align="left"> <font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">Sayfa 
            : 
            <%
for y=1 to sayfa 
if shf=y then
response.write y
else
response.write "<b> <a href=""mesajlar.asp?shf="&y&""">"&y&"</a></b>"
end if
next
%>
            <style>
<!--
a{text-decoration:none}
//-->
        </style>
            </font> </td>
      </tr>
    </table>
  </center>
</div>
</body>


</body>
</html></p>
<p align="center"><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif">
Katkılarından Dolayı <a target="_blank" href="http://www.keyifweb.com">
Keyifweb.com</a>'a Teşekkürler..
<a href="mailto:malibay@mini-nuke.info?subject=Meraba">ßy MaLiBaY</a><br>
  © <a href="http://www.Keyifweb.com">Www.Mini-NUKE.De</a></font></p>
</html>
<%
END IF
%>