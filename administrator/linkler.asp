<%
If session("Member") = "" then
Response.Write "<center><font class=""tabloalt_font"">Lütfen Giriş Yapınız.</font></center>"
Else
%>
<!--#include file="../db/data.asp" -->
<% response.buffer = True %>
<%
Set baglanti = Server.CreateObject("ADODB.Connection")
baglanti.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../db/mn"&dbNum&".mdb")
%>

<html>
<head>
<title>Tüm Siteler</title>
<META content=tr http-equiv=Content-Language>
<META content="text/html; charset=windows-1254" http-equiv=Content-Type>
<style type="text/css">
<!--
.form {  font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; color: #000000}
-->
</style>
</head>

<body bgcolor="#FFFFFF" text="#000000" link="#000000" vlink="#000000" alink="#666666">
<%
islem = Request("islem")
if islem = "onama" then
onay
elseif islem="silme" then
silme
else

set liste = Server.CreateObject("ADODB.RecordSet")
SQL_L = "Select * from linklist WHERE onay=0 ORDER by uye_no ASC"
liste.open SQL_L,baglanti,1,3
%>
<br>
<table width="400" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#CCCCCC">
  <tr>
    <td>
      <div align="center">
        <center>
      <table width="600" border="1" cellspacing="0" cellpadding="3" style="border-collapse: collapse" bordercolor="#C0C0C0">
	  <form action="<%=Request.ServerVariables("SCRIPT_NAME")%>?islem=onama" method="post">
        <tr bgcolor="#b0c4e5"> 
          <td width="95"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Site Adı</font></td>
          <td width="95"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Site Url</font></td>
          <td width="132"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">E-posta</font></td>
          <td width="347"> 
            <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif" size="1">Onay Ver </font></div>
          </td>
        </tr>
		<% if liste.eof then %>
		<tr><td bgcolor="#ffffff" colspan="4" width="392"><font face="Verdana, Arial, Helvetica, sans-serif" size="1">
          Onaylanmamış site yok.</font></td></tr>
        <%
		else
		for i = 1 to liste.recordcount
		if liste.eof then exit for
		if i mod 2 = 0 then
		bg = "#efefef"
		else
		bg = "#ffffff"
		end if
		%> 
        <tr bgcolor="<%=bg%>"> 
          <td width="900"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="<%=liste("url")%>" target="_blank"><%=liste("adi")%></a></font>&nbsp;</td>
          <td width="900"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="<%=liste("url")%>" target="_blank"><%=liste("url")%></a></font>&nbsp;</td>
          <td width="900"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="mailto:<%=liste("email")%>"> <%=Liste("email")%></a></font>&nbsp;</td>
          <td width="30"> 
            <div align="center"><input type="checkbox" name="<%=liste("uye_no")%>" value="<%=liste("uye_no")%>"></div>
          </td>
        </tr>
        <%
		liste.movenext
		Next
		liste.close
		set liste = nothing
		%> 
		<tr>
		<td colspan="4" bgcolor="#b0c4e5" width="392">
        <p align="center"><input type="submit" value="     Onayla     " class="form"></td></tr>
		<% end if %>
		</form>
      </table>
        </center>
      </div>
    </td>
  </tr>
</table>
<br>
<table width="400" border="0" cellspacing="0" cellpadding="0" align="center" bgcolor="#CCCCCC">
  <tr> 
    <td> 
      <div align="center">
        <center> 
      <table width="600" border="0" cellspacing="0" cellpadding="3" style="border-collapse: collapse" bordercolor="#111111">
<form method="post" action="<%=Request.ServerVariables("SCRIPT_NAME")%>?islem=silme">
        <tr bgcolor="#b0c4e5"> 
          <td><font size="1" face="Verdana, Arial, Helvetica, sans-serif">Listedeki 
            Tüm Siteler</font></td>
        </tr>
        <tr bgcolor="#FFFFFF"> 
            <td>
            <p align="left"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;&nbsp; Adı &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Url 
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Mail

            <select name="uye_no" class="form" size="20" style="width: 183; height: 300; border-style: solid; border-width: 0">
              
			  <%
			  set onayli = Server.CreateObJect("ADODB.RecordSet")
			  SQL_li = "Select * from linklist WHERE onay = 1 ORDER BY uye_no DESC"
			  onayli.open SQL_li,baglanti,1,3
			  for i = 1 to onayli.recordcount
			  if onayli.eof then exit for
			  %>
			  <option value="<%=onayli("uye_no")%>"><%=onayli("adi")%></option>
			  <%
			  onayli.movenext
			  Next
			  onayli.close
			  set onayli = Nothing
			  %>
              </select><select name="uye_no1" class="form" size="20" style="width: 183; height: 300; border-style: solid; border-width: 0">
              
			  <%
			  set onayli = Server.CreateObJect("ADODB.RecordSet")
			  SQL_li = "Select * from linklist WHERE onay = 1 ORDER BY uye_no DESC"
			  onayli.open SQL_li,baglanti,1,3
			  for i = 1 to onayli.recordcount
			  if onayli.eof then exit for
			  %>
			  <option value="<%=onayli("uye_no")%>"><%=onayli("url")%></option>
			  <%
			  onayli.movenext
			  Next
			  onayli.close
			  set onayli = Nothing
			  %>
              </select><select name="uye_no2" class="form" size="20" style="width: 183; height: 300; border-style: solid; border-width: 0">
              
			  <%
			  set onayli = Server.CreateObJect("ADODB.RecordSet")
			  SQL_li = "Select * from linklist WHERE onay = 1 ORDER BY uye_no DESC"
			  onayli.open SQL_li,baglanti,1,3
			  for i = 1 to onayli.recordcount
			  if onayli.eof then exit for
			  %>
			  <option value="<%=onayli("uye_no")%>"><%=onayli("email")%></option>
			  <%
			  onayli.movenext
			  Next
			  onayli.close
			  set onayli = Nothing
			  %>
              </select><br>
            <br>
              <input type="submit" value="   Sil   " class="form"><br>
            <br>
&nbsp;</font></td>
        </tr>
</form>
      </table>
        </center>
      </div>
    </td>
  </tr>
</table>
<br>
<div align="center"><br>
  <font size="1" face="Verdana, Arial, Helvetica, sans-serif">&lt;
  <a target="_blank" href="http://www.mini-nuke.com">http://www.mini-nuke.de &gt;</a></font></div>

  <%
  end if
  sub onay
  ' formdan gelen onay işaretli her eleman bu döngü ile onaylanır.
  for each eleman In Request.Form
  neyi = Request.Form(eleman)
  Set ona = Server.CreateObject("ADODB.RecordSet")
  SQL_ona = "Select * from linklist WHERE uye_no = "&neyi&""
  ona.open SQL_ona,baglanti,1,3
  ona("onay") = true
  ona.update
  ona.close
  set ona = Nothing
  Next
  Response.Redirect "Linkler.asp"
  end sub

  sub silme
  'silme işlemi
  uye_no = Request.Form("uye_no")
  Set sil = Server.CreateObJect("ADODB.RecordSet")
  SQL_S = "DELETE from linklist WHERE uye_no = "&uye_no&""
  sil.open SQL_S,baglanti,1,3
  response.redirect "Linkler.asp"
  end sub
  %>

<p></p>
<p>&nbsp;</p>
</body>
<% END IF %>
</html>