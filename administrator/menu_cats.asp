<% Response.Buffer = True %>
<!--#include file="includes.asp" -->
<%
IF Session("Level") = "1" OR Session("Level") = "7" Then

	IF Request.QueryString("op") = "New" Then
Set rs = Connection.Execute("Select * FROM MENU_CATS")
%>
<table border="0" cellspacing="0" cellpadding="2" width="100%" class="td_menu" style="font-weight:bold">
<tr><td colspan="2" bgcolor="#000000"></td></tr>
<form action="?op=Register" method="post">
<tr bgcolor="#EEEEEE">
<td width="35%" align="right">Kategori Adý : </td>
<td width="65%" align="left"><input type="text" name="title" size="50" class="form1"></td>
</tr>
<tr bgcolor="#EEEEEE">
<td width="35%" align="right">Sýra : </td>
<td width="65%" align="left">
<select name="menu_queue" size="1" class="form1">
<% For I = 1 TO rs.RecordCount+2 %>
<option value="<%=I%>"><%=I%></option>
<% Next %>
</select>
</td>
</tr>
<tr bgcolor="#EEEEEE">
<td width="35%" align="right"></td>
<td width="65%" align="left"><input type="submit" value="Ekle »" class="submit"></td>
</tr>
</form>
<tr><td colspan="2" bgcolor="#000000"></td></tr>
</table>
<%
	ElseIF Request.QueryString("op") = "Register" Then
cat_title = duz(Request.Form("title"))
IF cat_title = "" Then
Response.Write "<div align=center class=td_menu><b>Lütfen Kategori Adýný yazýnýz...</b></div>"
ELSE
Set ne = Server.CreateObject("ADODB.RecordSet")
ne.open "Select * FROM MENU_CATS",Connection,3,3
ne.AddNew
ne("mc_title") = cat_title
ne("mc_queue") = Request.Form("menu_queue")
ne.Update
ne.Close
Set ne = Nothing
Response.Redirect "menu_cats.asp"
End IF
	ElseIF Request.QueryString("op")  = "Edit" Then
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM MENU_CATS WHERE mc_id = "&Request.QueryString("id")&"",Connection,3,3
%>
<table border="0" cellspacing="0" cellpadding="2" width="100%" class="td_menu">
<tr><td colspan="2" bgcolor="#000000"></td></tr>
<form action="?op=Update&id=<%=Request.QueryString("id")%>" method="post">
<tr bgcolor="#EEEEEE">
<td width="35%" align="right">Kategori Adý : </td>
<td width="65%" align="left"><input type="text" name="title" size="50" class="form1" value="<%=rs("mc_title")%>"></td>
</tr>
<tr bgcolor="#EEEEEE">
<td width="35%" align="right">Sýra : </td>
<td width="65%" align="left">
<select name="queue" size="1" class="form1">
<%
For I = 1 TO rs.RecordCount+2
IF I = rs("mc_queue") Then
opt = "selected"
ELSE
opt = ""
End IF
%>
<option value="<%=I%>" <%=opt%>><%=I%></option>
<% Next %>
</select>
</td>
</tr>
<tr bgcolor="#EEEEEE">
<td width="35%" align="right"></td>
<td width="65%" align="left"><input type="submit" value="Güncelle »" class="submit"></td>
</tr>
</form>
<tr><td colspan="2" bgcolor="#000000"></td></tr>
</table>
<%
rs.Close
Set rs = Nothing
	ElseIF Request.QueryString("op") = "Update" Then
cat_title = duz(Request.Form("title"))
IF cat_title = "" Then
Response.Write "<div align=center class=td_menu><b>Lütfen Kategori Adýný yazýnýz...</b></div>"
ELSE
Connection.Execute("UPDATE MENU_CATS Set mc_title = '"&cat_title&"' WHERE mc_id = "&Request.QueryString("id")&"")
Connection.Execute("UPDATE MENU_CATS Set mc_queue = '"&Request.Form("queue")&"' WHERE mc_id = "&Request.QueryString("id")&"")
Response.Redirect "menu_cats.asp"
End IF
	ElseIF Request.QueryString("op") = "Delete" Then
Connection.Execute("DELETE FROM MENU_CATS WHERE mc_id ="&Request.QueryString("id")&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
	ELSE
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM MENU_CATS ORDER BY mc_title ASC",Connection,3,3
%>
<table border="1" cellspacing="0" cellpadding="2" width="50%" align="center" class="td_menu">
<tr><td colspan="2"><b>» <a href="?op=New">Yeni Kategori</a></b></td></tr>
<tr bgcolor="#FFFFCC" style="font-weight:bold">
<td width="60%">Kategori Adý</td>
<td width="40%" align="center">Ýþlem</td>
</tr>
<% Do Until rs.EoF %>
<tr bgcolor="#F0F0F0">
<td width="60%"><%=rs("mc_title")%>&nbsp;</td>
<td width="40%" align="center"><a href="?op=Edit&id=<%=rs("mc_id")%>">Düzenle</a> - <a href="?op=Delete&id=<%=rs("mc_id")%>">Sil</a></td>
</tr>
<%
rs.MoveNext
Loop
%>
</table>
<br><br><br>
<table border="0" cellspacing="0" cellpadding="2" width="100%" class="td_menu" style="font-weight:bold">
<tr><td colspan="2" bgcolor="#000000"></td></tr>
<form action="?op=Register" method="post">
<tr bgcolor="#EEEEEE">
<td width="35%" align="right">Kategori Adý : </td>
<td width="65%" align="left"><input type="text" name="title" size="50" class="form1"></td>
</tr>
<tr bgcolor="#EEEEEE">
<td width="35%" align="right">Sýra : </td>
<td width="65%" align="left">
<select name="menu_queue" size="1" class="form1">
<% For I = 1 TO rs.RecordCount+2 %>
<option value="<%=I%>"><%=I%></option>
<% Next %>
</select>
</td>
</tr>
<tr bgcolor="#EEEEEE">
<td width="35%" align="right"></td>
<td width="65%" align="left"><input type="submit" value="Ekle »" class="submit"></td>
</tr>
</form>
<tr><td colspan="2" bgcolor="#000000"></td></tr>
</table>
<%
	End IF

End IF
%>