<% Response.Buffer = True %>
<!--#include file="includes.asp" -->
<%
IF Session("Level") = "1" OR Session("Level") = "7" Then
x = Request.QueryString("action")
	IF x = "New" Then
%>
<table border="0" cellspacing="2" cellpadding="1" align="center" class="td_menu" style="font-weight:bold">
<form action="?action=Register" method="post">
<tr bgcolor="#EEEEEE">
<td width="30%">Sözcük : </td>
<td width="70%"><input type="text" name="text" size="80" class="form1"></td>
</tr>
<tr bgcolor="#EEEEEE">
<td width="30%">Deðer : </td>
<td width="70%"><input type="text" name="value" size="80" class="form1"></td>
</tr>
<tr bgcolor="#FFFFFF">
<td width="40%"></td>
<td width="60%"><input type="submit" value="Ekle +" class="submit" style="width:20%"></td>
</tr>
</form>
</table>
<br><br><div align="center" class="td_menu"><a href="<%=Request.ServerVariables("HTTP_REFERER")%>"><< Geri</a></div>
<%
	ElseIF x = "Register" Then
f_f1 = Request.Form("text")
f_f2 = Request.Form("value")
IF f_f1="" OR f_f2="" Then
Response.Write "<div align=center class=td_menu>Lütfen Tüm Alanlarý Doldurunuz...</div>"
ElseIF f_f1<>"" AND f_f2<>"" Then
Connection.Execute("INSERT INTO SWEARWORDS (s_text,s_value) VALUES ('"&f_f1&"','"&f_f2&"')")
Response.Redirect "swearwords.asp"
End IF
	ElseIF x = "Edit" Then
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM SWEARWORDS WHERE id = "&Request.QueryString("id")&"",Connection,3,3
%>
<table border="0" cellspacing="2" cellpadding="1" align="center" class="td_menu" style="font-weight:bold">
<form action="?action=Update&id=<%=Request.QueryString("id")%>" method="post">
<tr bgcolor="#EEEEEE">
<td width="30%">Sözcük : </td>
<td width="70%"><input type="text" name="text" size="80" class="form1" value="<%=rs("s_text")%>"></td>
</tr>
<tr bgcolor="#EEEEEE">
<td width="30%">Deðer : </td>
<td width="70%"><input type="text" name="value" size="80" class="form1" value="<%=rs("s_value")%>"></td>
</tr>
<tr bgcolor="#FFFFFF">
<td width="40%"></td>
<td width="60%"><input type="submit" value="Güncelle" class="submit" style="width:20%"></td>
</tr>
</form>
</table>
<br><br><div align="center" class="td_menu"><a href="<%=Request.ServerVariables("HTTP_REFERER")%>"><< Geri</a></div>
<%
	ElseIF x = "Update" Then
f1 = Request.Form("text")
f2 = Request.Form("value")
IF f1="" OR f2="" Then
Response.Write "<div align=center class=td_menu>Lütfen Tüm Alanlarý Doldurunuz</div>"
ElseIF f1<>"" AND f2<>"" Then
Connection.Execute("UPDATE SWEARWORDS SET s_text = '"&f1&"' WHERE id = "&Request.QueryString("id")&"")
Connection.Execute("UPDATE SWEARWORDS SET s_value = '"&f2&"' WHERE id = "&Request.QueryString("id")&"")
Response.Redirect "swearwords.asp"
End IF
	ElseIF x = "Delete" Then
Connection.Execute("DELETE FROM SWEARWORDS WHERE id = "&Request.QueryString("id")&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
	ELSE
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM SWEARWORDS ORDER BY s_text DESC",Connection,3,3
%>
<div class="td_menu" style="font-weight:bold"><a href="?action=New">+ Yeni Ekle</a></div>
<table border="1" cellspacing="0" cellpadding="1" width="90%" align="center" class="td_menu">
<tr bgcolor="#FFCC77" style="font-weight:bold">
<td width="40%">Sözcük</td>
<td width="40%">Deðer</td>
<td width="20%" align="center">Iþlem</td>
</tr>
<% Do Until rs.Eof %>
<tr>
<td width="40%"><%=rs("s_text")%>&nbsp;</td>
<td width="40%"><%=rs("s_value")%>&nbsp;</td>
<td width="20%" align="center"><a href="?action=Edit&id=<%=rs("id")%>">Düzenle</a> - <a href="?action=Delete&id=<%=rs("id")%>">Sil</a></td>
</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
	End IF
%>
<% End IF %>