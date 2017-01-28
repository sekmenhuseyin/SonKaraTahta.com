<!--#include file="includes.asp" -->

<%
dim parola
parola = ""&a_pass&""
if session("administrator") <> parola then
if request.form("administrator") <> parola then
call parolaform
else
session("administrator") = parola
end if
end if
%>

<%'-----------------------------------------------------------------------------------------------------------'%>
<%
sub parolaform
aynen = "http://" & Request.ServerVariables("HTTP_HOST") & Request.ServerVariables("URL")
%>

<head>

<link rel="stylesheet" type="text/css" href="stil.css">

<title></title>
</head>

<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="550">
    <tr>
      <td>
    </td>
    </tr>
    <tr>
      <td height="5">
        <hr style="border: 1 dashed #0059B3" color="#81BAE5" size="0">
      </td>
    </tr>
  </table>
  </center>
</div>
<p align="center"><b><font face="Tahoma" size="2">Lütfen Admin $ifrenizi Giriniz ...</font></b>
<form method=post action="<%=aynen%>">
<p align="center"><font face="Verdana" size="1">
<input type="password" name="administrator" value="" size="20" style="background-color: #FFFFFF; font-family: Verdana; font-size: 8pt; border: 1 ridge #525454"></font>
<font face="Tahoma" color="#FF6600" size="1">- -</font> <input type="submit" value="Login" style="background-color: #FFFFFF; font-family: Verdana; font-size: 8pt; border: 1 outset #525454">
</form>
<p align="center">
&nbsp;
        <hr style="border: 1 dashed #0059B3" color="#81BAE5" size="0" width="550">
<p align="center"><font face="Tahoma" size="1" color="#525454">CODED ßy Mini-NUKE TEAM<br>
</font><a href="http://www.mini-nuke.info"><font face="Tahoma" size="1" color="#525454">http://www.mini-nuke.info</font></a><font face="Tahoma" size="1" color="#525454">
</font><br>
</p>

<%
response.end
end sub
%>
<%'-------------------------------------------------------------------------------------------------------------'%>

<%
Set Sur = Server.CreateObject("ADODB.Connection")
Sur.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../db/mn"&dbNum&".mdb")
%>
<%
Set liste = Server.CreateObject("ADODB.Recordset")
SQL = "Select * from RASTGELE ORDER BY id desc"
liste.Open SQL,Sur,1,3
%>

<%
shf = Request.QueryString("shf")
if shf="" then 
shf=1
end if
%>


<%
sil = Request.QueryString("sil")
if sil = "delete" then
call delete
else
end if
%>


<%
sub delete
id = Request("id")
Set liste = Server.CreateObject("ADODB.RecordSet")
kayit_delete = "DELETE from RASTGELE WHERE id="&id&""
liste.open kayit_delete,Sur,1,3
response.redirect Request.ServerVariables("HTTP_REFERER")
end sub
%>
<html>
<head>
<meta http-equiv="Content-Language" content="tr">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>Yeni Resim Ekle</title>
 


</head>
<body text="#545454" link="#545454" vlink="#545454" alink="#545454" bgcolor="#FFFFFF" topmargin="0">
          <div align="center">
            <center>
          <table border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="600" valign="top">
                <div align="center">
                  <center>
                <table border="0" width="100%" cellspacing="0" cellpadding="0">
                  <tr>
    <td valign="top" bgcolor="#FFFFFF">
    <table border="0" width="100%" cellspacing="1" cellpadding="0" bgcolor="#FFFFFF">
      <tr>
        <td width="100%" align="center" bgcolor="#FFFFFF" height="21">&nbsp;


<%
islem=Request.QueryString("islem")
if islem="ekle" then
call ekle
end if
%>

<%
sub ekle

Set liste = server. CreateObject("ADODB.Recordset")
kayit = "Select * from RASTGELE"
liste.Open kayit,sur,1,3

Dim resim

resim = request.form("resim")

if resim="" then
Response.Write "<font face=verdana size=2><center>Lütfen Ekliyeceðiniz Resmin Adresini Giriniz...!!!</center></font>"
Response.End
end if


liste.AddNew
liste("resim") = resim
liste.Update

Response.Write "<font face=verdana color=#000000 size=2>"
Response.Write "<center>"
Response.Write "Resim Eklendi."
Response.Write "<br>"
Response.Write "<br>"
Response.Write "<br>"
Response.Write "<br>"
Response.Write "<br>"
Response.Write "<br>"
Response.Write "</center>"
Response.Write "</font>"

end sub
%>
          <br>
          <font face="Verdana"><b>Yeni Resim Ekle</b></font>
<form method="POST" action="resimekle.asp?islem=ekle">
      <tr>
        <td align="center" bgcolor="#FFFFFF"><font face="Verdana" size="1"><b>&nbsp;<br>
          Resim Adresi (Resmin URLUNU Yazýnýz...)
          :&nbsp; 
        <input type="text" name="resim" size="50" style="background-color: #FFFFFF; border: 1 solid #525454" value="http://"><br>
          &nbsp;</b></font></td>
      </tr>
      <tr>
        <td align="center" bgcolor="#FFFFFF">
          <p align="center"><font face="Verdana" size="1"><b><br>
          <input type="submit" value="--- Ekle ---" name="B1" style="background-color: #FFFFFF; border: 1 solid #525454">
          <input type="reset" value="--- Temizle ---" name="B2" style="background-color: #FFFFFF; border: 1 solid #525454"></b></font>&nbsp;<br>
          &nbsp;</td>
      </tr></form>

      <tr>
        <td align="center" bgcolor="#FFFFFF">
          &nbsp;</td>
      </tr>
      <tr>
        <td align="center" bgcolor="#FFFFFF">
          &nbsp;<%
for i = 1 to 5000000
if liste.eof then exit for
%>
            <tr>
              <td height="25" bgcolor="#FFFFFF" style="padding-top: 2; padding-bottom: 2">
                <div align="left">
                  <table border="0" cellpadding="0" cellspacing="0">
                    <tr>
                      <td width="555"><font face="Verdana" size="2">&nbsp;<b><font color="#1D5A87">»</font> </b>
                <%=liste("resim")%></font></td>





            </center>



                  </center>
                  <td>
                    <p align="right"><a href="resimekle.asp?sil=delete&amp;id=<%=liste("id")%>"><font face="Verdana" size="2" color="#525454"><b>Sil</b></font></a></td>
                </tr>
              </table>
        </div>
              </td>
            <center>
            </tr>
                  <center>
<%
liste.movenext
Next
liste.close
set liste= Nothing
%>
    </table>





            </center>



</body>