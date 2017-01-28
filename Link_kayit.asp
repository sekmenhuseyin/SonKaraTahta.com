<!--#include file="db/data.asp" -->
<!--#include file="inc/includes.asp" -->
<html>
<%
Set baglanti=Server.CreateObject("ADODB.Connection")
baglanti.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("db/mn"&dbNum&".mdb")
%>
<head>
<title>Yeni Kayıt İşlemleri</title>
<META content=tr http-equiv=Content-Language>
<META content="text/html; charset=windows-1254" http-equiv=Content-Type>
<style type="text/css">
</style>
<body bgcolor="<%=menu_back%>" text="<%=text%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"<%=body_have_new_msg%>>
<link rel="stylesheet" type="text/css" href="../THEMES/<%=Session("SiteTheme")%>/style.css">
</head>
<%
if Request("aks")="kaydet" then
kaydet
else
%>
<div align="center">
  <table width="380" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="<%=menu_back%>" >
    <tr> 
      <td> 
        <table width="380" border="0" cellspacing="1" cellpadding="3" align="center">
          <form action="<%=Request.ServerVariables("SCRIPT_NAME")%>?aks=kaydet" method="post">
            <tr > 
            <td width="100"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Email 
              Adresiniz: </font></td>
              <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif">
                <input type="text" name="email" size="22" class="form1" >
                </font></td>
          </tr>
          <tr>
            <td width="100"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Sitenizin 
              Adı: </font></td>
              <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif">
                <input type="text" name="adi" size="23" class="form1" >
                </font></td>
          </tr>
          <tr > 
            <td width="100"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Sitenizin 
              Adresi: </font></td>
              <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif">
                <input type="text" name="url" value="http://www." size="28" class="form1" >
                </font></td>
          </tr>
          <tr > 
            <td width="100">&nbsp;</td>
            <td>
                <br />
                <input type="submit" value="     Kaydet     " class="submit">
              </td>
          </tr>
		 </form>
        </table>
      </td>
    </tr>
  </table>
</div>

<%
end if
sub kaydet
email=Request.Form("email")
adi=Request.Form("adi")
url=Request.Form("url")

if email="" then
Response.Write "<font face='Verdana,Arial' size='1'><b>HATA!</b> Email kısmı boş bırakılmış<br />"
hata
end if
if adi="" then
Response.Write "<font face='Verdana,Arial' size='1'><b>HATA!</b> Sitenizin adı kısmı boş bırakılmış<br />"
hata
end if
if url="" then
Response.Write "<font face='Verdana,Arial' size='1'><b>HATA!</b> Sitenizi adresi kısmı boş bırakılmış<br />"
hata
end if

set kayit=Server.CreateObJect("ADODB.RecordSet")
SQL_k="Select * from linklist WHERE email='"&email&"'"
kayit.open SQL_k,baglanti,1,3
if kayit.eof then
kayit.addnew
kayit("email")=email
kayit("adi")=adi
kayit("url")=url
kayit("tekil")=0
kayit("cogul")=0
kayit("onay")=0
kayit("kayit_tarihi")=DateAdd("h",7,now)
kayit.update
%>
<table width="311" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr> 
      <td width="318"> 
<p align="center"> 
<font size="1" face="Verdana, Arial, Helvetica, sans-serif"><i>Sizi tebrik ederiz! Kayıt işleminiz başarı ile tamamlanmış ve veritabanımıza kaydınız eklenmiş durumdadır. En geç 24 
saat içerisinde siteniz onaylanacak ve listemizde yerinizi alacaksınız.Artık 
sizin tek yapmanız gereken aşağıdaki kodu sitenizin kodları arasına 
yerleştirmek. Aksi Takdirde Galiye <u>Alınmayacaktır.</u></i><textarea name="aciklama" wrap="VIRTUAL" cols="60" rows="5" class="form" style="border-style:solid; border-width:1; font-size: 10px; font-family: Verdana"><!--TopList Başlangıç -->
 <a href="<%=sett("site_adres")%>" target="_blank"><img src="<%=sett("site_isim")%>/<%=sett("site_topbanner")%>" border="0" height="31" width="88" alt="TopListimize Siz de katılın"></a>
 <!-- TopList Bitiş--></textarea> </p>
</p>
</td></tr></table>
<%
else
Response.Write "<font face='Verdana,Arial' size='1'><b>HATA!</b> Bu E-posta Adresi bir başkası tarafından kullanılmaktadır.<br />"
hata
end if
end sub

sub hata
Response.Write "Lütfen <a href='javascript:history.back(1)'><b><< Geri</b></a> dönüp düzeltin ve yeniden deneyin</font>"
Response.End
end sub%>

</body>
</html>