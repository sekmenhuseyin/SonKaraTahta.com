<!--#include file="../../db/data.asp" -->
<% IF Request.QueryString("menu")="show" Then %>
<%
IF dbType="personal" Then
Set Connection=Server.CreateObject("ADODB.Connection")
Connection.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../../../db/mn"&dbNum&".mdb")
ELSE
Set Connection=Server.CreateObject("ADODB.Connection")
Connection.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../../db/mn"&dbNum&".mdb")
END IF
 %>
<% 
aces=Request.QueryString("ace")
IF aces="sily" then call siliver


imgid=request.querystring("imgid")

set liste=Server.CreateObject("ADODB.RecordSet")
sql_l="select * from fotograf where fno=" & imgid
liste.open sql_l,Connection,3,3
liste("hit")=liste("hit") + 1
liste.update
liste.close
set liste=nothing

set liste=Server.CreateObject("ADODB.RecordSet")
sql_l="select * from fotograf where fno=" & imgid
liste.open sql_l,Connection,1,3
%>
<div align="center"><img src="<%=liste("buyuk") %>" border="0"><br /><font face="Arial, Helvetica, sans-serif" size="2"><%=liste("aciklama") %></font></div>
<div align="center">
  <center>
<table border="0" cellpadding="2" cellspacing="3" style="border-collapse: collapse" bordercolor="#111111" width="64%" id="AutoNumber1">
  <tr>
    <td width="100%" colspan="3">
    <p align="center"><b><font face="Tahoma" color="#003300">Yorumlar</font></b></td>
  </tr>
  <%
IF sayfa= "" then sayfa=1 ELSE sayfa=CInt(Sayfa) END IF
set RS=Server.CreateObject("ADODB.RecordSet")
sql_l="select * from yorum where hangi=" & imgid
RS.open sql_l,Connection,1,3
IF RS.eof THEN %>
<tr><td bgcolor="FFFFFF" colspan="5" width="100%">
  <p align="center"><font face="Verdana" size="1">Kayýtlý bilgi bulunmamaktadýr...</font></td></tr>
<% ELSE
FOR i=1 to 50
IF RS.eof THEN EXIT FOR
IF i mod 2=0 THEN bg="#FFFFFF" ELSE bg="#f7f7f7" END IF %>
  <tr>
    <td width="11%" bgcolor="<%=bg%>">
    <p align="right"><font face="Verdana" size="1">Yazan :</font><font size="1">
    </font> </td>
    <td width="14%" bgcolor="<%=bg%>"><p align="left">
    <font face="Verdana" size="1"><%=RS("isim")%></font></p></td>
    <td width="75%" rowspan="2" bgcolor="<%=bg%>"><font face="Verdana" size="1"><%=RS("yorum")%></font>
<%    If Session("Level")="1" Then%>
<font face="Verdana" size="1">
<a href="sil.asp?silyorum=<%=rs("yno")%>">
//Yorumu Sil</a>
</font>
<%end if%></td></tr><td width="11%">
    <p align="right" bgcolor="<%=bg%>"><font face="Verdana" size="1">Tarih :</font><font size="1">
    </font> </td>
    <td width="14%" bgcolor="<%=bg%>"><p align="left">
    <font face="Verdana" size="1"><%=RS("tarih")%></font></p></td>
  </tr>
  <% RS.Movenext
Next
RS.Close
Set RS=Nothing
END IF %>


</table>
  </center>
</div>
<% end if %>
<%
sub siliver
yorid=request.querystring("silyorum")
Set dc=Server.CreateObject("ADODB.RecordSet")
dcSQL="DELETE * FROM yorum WHERE yno="&yorid&""
dc.open dcSQL,Connection,3,3
Response.Redirect Request.ServerVariables("HTTP_REFERER")
end sub
%>