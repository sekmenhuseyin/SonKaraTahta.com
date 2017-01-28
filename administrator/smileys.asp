<% Response.Buffer = True %>
<!--#include file="includes.asp" -->
<%
IF Session("Level") = "1" OR Session("Level") = "7" Then

	IF Request.QueryString("op") = "Delete" Then
Connection.Execute("DELETE FROM SMILEYS WHERE s_id = "&Request.QueryString("id")&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
	ElseIF Request.QueryString("op") = "Edit" Then
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM SMILEYS WHERE s_id = "&Request.QueryString("id")&"",Connection,3,3
%>
<table border="0" cellpadding="2" cellspacing="1" width="60%" align="center" class="td_menu2" style="font-weight:bold">
<form method="post" action="?op=Update&id=<%=Request.QueryString("id")%>">
<tr bgcolor="#EEEEEE">
<td width="50%" align="right">Smili Ýþareti : </td>
<td width="50%" align="right">[<input type="text" name="information" class="form1" style="width:10%" value="<%=rs("s_info")%>" size="20">]</td>
</tr>
<tr bgcolor="#EEEEEE">
<td width="50%" align="right">Smili  : </td>
<td width="50%" align="right">IMAGES/smileys/<input type="text" name="smiley" class="form1" style="width:50%" value="<%=rs("s_img")%>" size="20"></td>
</tr>
<tr bgcolor="#EEEEEE">
<td width="50%" align="right"></td>
<td width="50%"><input type="submit" value="Güncelle »" class="submit"></td>
</tr>
</form>
</table>
<%
rs.Close
Set rs = Nothing
	ElseIF Request.QueryString("op") = "Update" Then

f_info = duz(Request.Form("information"))
f_smiley = Request.Form("smiley")
IF f_info = "" OR f_smiley = "" Then
Response.Write "<div align=center class=td_menu><b>Tüm Alanlarý Doldurunuz !</b></div>"
ELSE
Connection.Execute("UPDATE SMILEYS Set s_info = '"&f_info&"' WHERE s_id = "&Request.QueryString("id")&"")
Connection.Execute("UPDATE SMILEYS Set s_img = '"&f_smiley&"' WHERE s_id = "&Request.QueryString("id")&"")
Response.Redirect "smileys.asp"
End IF

	ElseIF Request.QueryString("op") = "New" Then
%>
<table border="0" cellpadding="2" cellspacing="1" width="60%" align="center" class="td_menu2" style="font-weight:bold">
<form method="post" action="?op=Register">
<tr bgcolor="#EEEEEE">
<td width="50%" align="right">Smili Ýþareti : </td>
<td width="50%" align="right">[<input type="text" name="information" class="form1" style="width:10%" size="20">]</td>
</tr>
<tr bgcolor="#EEEEEE">
<td width="50%" align="right">Smili  : </td>
<td width="50%" align="right">IMAGES/smileys/<input type="text" name="smiley" class="form1" style="width:50%" size="20"></td>
</tr>
<tr bgcolor="#EEEEEE">
<td width="50%" align="right"></td>
<td width="50%"><input type="submit" value="Ekle   +" class="submit"></td>
</tr>
</form>
</table>
<%
	ElseIF Request.QueryString("op") = "Register" Then

f_info = duz(Request.Form("information"))
f_smiley = Request.Form("smiley")
IF f_info = "" OR f_smiley = "" Then
Response.Write "<div align=center class=td_menu><b>Tüm Alanlarý Doldurunuz !</b></div>"
ELSE
Connection.Execute("INSERT INTO SMILEYS (s_info,s_img) VALUES ('"&f_info&"','"&f_smiley&"')")
Response.Redirect "smileys.asp"
End IF

	ELSE
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM SMILEYS ORDER BY s_info ASC",Connection,3,3
%>
<table border="0" cellpadding="2" cellspacing="1" width="30%" align="center" class="td_menu2">
<tr><td colspan="3">» <a href="?op=New">Smili Ekle</a></td></tr>
<tr bgcolor="#FFCC00" style="font-weight:bold">
<td width="1%" align="center">Smili</td>
<td width="20%" align="center">Yazý</td>
<td width="50%" align="center">Ýþlem</td>
</tr>
<%
Do Until rs.EoF
number=number + 1
select_color=right(number,1) 
Select Case select_color
case "1","3","5","7","9"
color = forum_bg_1
Case Else
color = forum_bg_2
End Select
%>
<tr bgcolor="<%=color%>">
<td width="1%" align="center"><img src="../IMAGES/smileys/<%=rs("s_img")%>"></td>
<td width="20%" align="center">[<%=rs("s_info")%>]</td>
<td width="50%" align="center"><a href="?op=Edit&id=<%=rs("s_id")%>">Düzenle</a> - <a href="?op=Delete&id=<%=rs("s_id")%>">Sil</a></td>
</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
rs.Close
Set rs = Nothing
	End IF

End IF
%>