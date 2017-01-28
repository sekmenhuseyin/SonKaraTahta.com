<% Response.Buffer = True %>
<!--#include file="includes.asp" -->
<%
IF Session("Level") = "1"  OR Session("Level") = "7" Then

	IF Request.QueryString("op") = "New" Then
%>
<table border="0" cellspacing="2" cellpadding="2" width="75%" align="center" class="td_menu2" style="font-weight:bold">
<form action="?op=Register" method="post">
<tr bgcolor="#E0E0E0">
<td width="50%" align="right">Tema Adý : </td>
<td width="50%" align="left">
<input type="text" name="theme_name" class="form1" style="width:100%" size="20"></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="50%" align="right">Tema Klasörü : </td>
<td width="50%" align="left">THEMES/<input type="text" name="theme_dir" class="form1" style="width:83%" size="20"></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="50%" align="right">Aktif / Pasif : </td>
<td width="50%" align="left">
<input type="checkbox" name="theme_status" value="ON"></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="50%" align="right"></td>
<td width="50%" align="left"><input type="submit" value="Ekle »" class="submit"></td>
</tr>
</form>
</table>
<%
	ElseIF Request.QueryString("op") = "Register" Then

name = Request.Form("theme_name")
dir = Request.Form("theme_dir")
status = Request.Form("theme_status")

IF status = "on" Then
theme_status = "True"
ELSE
theme_status = "False"
End IF

IF name = "" OR dir = "" Then
Response.Write "<div align=center class=td_menu><b>Tüm Alanlarý Doldurunuz...</b></div>"
ELSE
Connection.Execute("INSERT INTO THEMES (t_name,t_dir,t_active) VALUES ('"&name&"','"&dir&"',"&theme_status&")")
Response.Redirect "themes.asp"
End IF

	ElseIF Request.QueryString("op") = "Edit" Then
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM THEMES WHERE t_id = "&Request.QueryString("id")&"",Connection,3,3
%>
<table border="0" cellspacing="2" cellpadding="2" width="75%" align="center" class="td_menu2" style="font-weight:bold">
<form action="?op=Update&id=<%=Request.QueryString("id")%>" method="post">
<tr bgcolor="#E0E0E0">
<td width="50%" align="right">Tema Adý : </td>
<td width="50%" align="left">
<input type="text" name="theme_name" class="form1" style="width:100%" value="<%=rs("t_name")%>" size="20"></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="50%" align="right">Tema Klasörü : </td>
<td width="50%" align="left">THEMES/<input type="text" name="theme_dir" class="form1" style="width:83%" value="<%=rs("t_dir")%>" size="20"></td>
</tr>
<%
IF rs("t_active") = True Then
opt = "checked"
ELSE
opt = ""
End IF
%>
<tr bgcolor="#E0E0E0">
<td width="50%" align="right">Aktif / Pasif : </td>
<td width="50%" align="left">
<input type="checkbox" name="theme_status" <%=opt%> value="ON"></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="50%" align="right"></td>
<td width="50%" align="left"><input type="submit" value="Güncelle »" class="submit"></td>
</tr>
</form>
</table>
<br><br>
<div align="center" class="td_menu2"><a href="<%=Request.ServerVariables("HTTP_REFERER")%>">« Geri</a></div>
<%
rs.Close
Set rs = Nothing
	ElseIF Request.QueryString("op") = "Update" Then

t_name = Request.Form("theme_name")
t_dir = Request.Form("theme_dir")
t_status = Request.Form("theme_status")

IF t_status = "on" Then
status = "True"
ELSE
status = "False"
End IF

IF t_name = "" OR t_dir = "" Then
Response.Write "<div align=center class=td_menu><b>Tüm Alanlarý Doldurunuz...</b></div>"
ELSE
Connection.Execute("UPDATE THEMES SET t_name = '"&t_name&"' WHERE t_id = "&Request.QueryString("id")&"")
Connection.Execute("UPDATE THEMES SET t_dir = '"&t_dir&"' WHERE t_id = "&Request.QueryString("id")&"")
Connection.Execute("UPDATE THEMES SET t_active = "&status&" WHERE t_id = "&Request.QueryString("id")&"")
Response.Redirect "themes.asp"
End IF

	ElseIF Request.QueryString("op") = "Delete" Then

		Connection.Execute("DELETE FROM THEMES WHERE t_id = "&Request.QueryString("id")&"")
		Response.Redirect Request.ServerVariables("HTTP_REFERER")

	ElseIF Request.QueryString("op") = "Change" Then

		Connection.Execute("UPDATE THEMES SET t_active = "&Request.QueryString("isl")&" WHERE t_id = "&Request.QueryString("id")&"")
		Response.Redirect Request.ServerVariables("HTTP_REFERER")

	ELSE
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM THEMES ORDER BY t_id",Connection,3,3
%>
<table border="0" cellspacing="2" cellpadding="2" width="90%" align="center" class="td_menu2">
<tr>
<td colspan="4">» <b><a href="?op=New">Tema Ekle</a></td>
</tr>
<tr bgcolor="#FFBB00" style="font-weight:bold">
<td width="1%" align="center">ID#</td>
<td width="30%" align="center">Tema Adý</td>
<td width="24%" align="center">Tema Klasörü</td>
<td width="15%" align="center">Durum</td>
<td width="30%" align="center">Ýþlemler</td>
</tr>
<%
sayi = "0"
Do Until rs.EoF
sayi = sayi + 1
ssayi = Right(sayi,1)
IF ssayi = "0" OR ssayi = "2" OR ssayi = "4" OR ssayi = "6" OR ssayi = "8" Then
tdbg = "#E0E0E0"
ELSE
tdbg = "#F0F0F0"
End IF

IF rs("t_active") = True Then
islem = "Pasifleþtir"
status = "Aktif"
isl = "False"
ELSe
islem = "Aktifleþtir"
status = "Pasif"
isl = "True"
End IF
%>
<tr bgcolor="<%=tdbg%>">
<td width="1%" align="center"><%=rs("t_id")%></td>
<td width="30%" align="center"><%=rs("t_name")%></td>
<td width="24%" align="center"><%=rs("t_dir")%></td>
<td width="15%" align="center"><%=status%></td>
<td width="30%" align="center"><a href="?op=Change&isl=<%=isl%>&id=<%=rs("t_id")%>"><%=islem%></a> - <a href="?op=Edit&id=<%=rs("t_id")%>">Düzenle</a> - <a href="?op=Delete&id=<%=rs("t_id")%>">Sil</a></td>
</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
	End IF

End IF
%>