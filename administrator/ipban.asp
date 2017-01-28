<% Response.Buffer = True %>
<!--#include file="includes.asp" -->
<%
IF Session("Level") = "1" OR Session("Level") = "7" Then

	IF Request.QueryString("op") = "New" Then
%>
<table border="0" cellspacing="2" cellpadding="2" width="50%" align="center" class="td_menu2">
<form method="post" action="?op=Register">
<tr bgcolor="#E0E0E0">
<td width="40%" align="right">IP Adresi : </td>
<td width="60%" align="left">
<input type="text" name="ip" class="form1" style="width:100%" size="20"></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="40%" align="right">Tarih : </td>
<td width="60%" align="left">
<input type="text" value="<%=Now()%>" class="form1" style="width:100%" disabled size="20"></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="40%" align="right"></td>
<td width="60%" align="left"><input type="submit" value="Tamam" class="submit"></td>
</tr>
</form>
</table>
<%
	ElseIF Request.QueryString("op") = "Register" Then

ip = duz(Request.Form("ip"))

IF ip="" Then
Response.Write "<div align=center class=td_menu2><b>Lütfen IP Adresi Yazýnýz...</b></div>"
ElseIF IsNumeric(ip) = False Then
Response.Write "<div align=center class=td_menu2><b>IP Adresi sayýlardan oluþmalýdýr</b></div>"
ELSE
Connection.Execute("INSERT INTO BANNED_IPS (b_ip,b_date) VALUES ('"&ip&"','"&Now()&"')")
Response.Redirect "ipban.asp"
End IF

	ElseIF Request.QueryString("op") = "Edit" Then
Set rs = Connection.Execute("Select * FROM BANNED_IPS WHERE b_id = "&Request.QueryString("id")&"")
%>
<table border="0" cellspacing="2" cellpadding="2" width="50%" align="center" class="td_menu2">
<form method="post" action="?op=Update&id=<%=Request.QueryString("id")%>">
<tr bgcolor="#E0E0E0">
<td width="40%" align="right">
<%
IF Request.QueryString("Type") = "IP" Then
Response.Write "IP Adresi : "
ELSE
Response.Write "Üye Seçin : "
End IF
%>
</td>
<td width="60%" align="left">
<% IF Request.QueryString("Type") = "IP" Then %>
<input type="text" name="ip" class="form1" style="width:100%" value="<%=rs("b_ip")%>" size="20">
<%
ELSE
Set mems = Connection.Execute("Select * FROM MEMBERS ORDER BY kul_adi ASC")
%>
<select name="ip" size="1" class="form1">
<%
Do Until mems.EoF
IF rs("b_ip") = ""&mems("uye_id")&"" Then
opt = "selected"
ELSE
opt = ""
End IF
%>
<option value="<%=mems("uye_id")%>"><%=mems("kul_adi")%></option>
<%
mems.MoveNext
Loop
%>
</select>
<% End IF %>
</td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="40%" align="right">Tarih : </td>
<td width="60%" align="left">
<input type="text" value="<%=rs("b_date")%>" class="form1" style="width:100%" disabled size="20"></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="40%" align="right"></td>
<td width="60%" align="left"><input type="submit" value="Tamam" class="submit"></td>
</tr>
</form>
</table>
<%
	ElseIF Request.QueryString("op") = "Update" Then

ip = duz(Request.Form("ip"))

IF ip="" Then
Response.Write "<div align=center class=td_menu2><b>Lütfen IP Adresi Yazýnýz...</b></div>"
ElseIF IsNumeric(ip) = False Then
Response.Write "<div align=center class=td_menu2><b>IP Adresi sayýlardan oluþmalýdýr</b></div>"
ELSE
Connection.Execute("UPDATE BANNED_IPS Set b_ip = '"&ip&"' WHERE b_id = "&Request.QueryString("id")&"")
Response.Redirect "ipban.asp"
End IF

	ElseIF Request.QueryString("op") = "Delete" Then

	Connection.Execute("DELETE FROM BANNED_IPS WHERE b_id = "&Request.QueryString("id")&"")
	Response.Redirect Request.ServerVariables("HTTP_REFERER")

	ElseIF Request.QueryString("op") = "NewMember" Then
Set mems = Connection.Execute("Select * FROM MEMBERS")
%>
<table border="0" cellspacing="2" cellpadding="2" width="90%" align="center" class="td_menu2">
<form action="?op=Register" method="post">
<tr bgcolor="#E0E0E0">
<td width="50%" align="right"><b>Üye Seçin : </b></td>
<td width="50%" align="left">
<select name="ip" size="1" class="form1">
<% Do Until mems.EoF %>
<option value="<%=mems("uye_id")%>"><%=mems("kul_adi")%></option>
<%
mems.MoveNext
Loop
%>
</select>
</td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="50%" align="right"><b>Banlanma Tarihi : </b></td>
<td width="50%" align="left">
<input type="text" value="<%=Now()%>" class="form1" style="width:100%" disabled size="20"></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="50%"></td>
<td width="50%"><input type="submit" value="Tamam" class="submit"></td>
</tr>
</form>
</table>
<%
	ELSE

Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM BANNED_IPS ORDER BY b_ip DESC",Connection,3,3
%>
<table border="0" cellspacing="2" cellpadding="2" width="90%" align="center" class="td_menu2">
<tr>
<td colspan="3">
» <a href="?op=New">Yeni IP Gir</a> <b>|</b> <a href="?op=NewMember">Üye Banla</a>
</td>
</tr>
<tr bgcolor="#CCCCCC" style="font-weight:bold">
<td width="40%" align="center">IP</td>
<td width="40%" align="center">Banlanma Tarihi</td>
<td width="20%" align="center">Ýþlem</td>
</tr>
<%
Do Until rs.EoF
number=number + 1
select_color=right(number,1) 
Select Case select_color
case "1","3","5","7","9"
color = "#E0E0E0"
Case Else
color = "#F0F0F0"
End Select
IF InStr(1,rs("b_ip"),".",1) Then
text = "IP"
l_text = "IP"
ELSE
text = "Üye"
l_text = "Member"
End IF
%>
<tr bgcolor="<%=color%>">
<td width="40%" align="center"><%=rs("b_ip")%> <i>(<%=text%>)</i></td>
<td width="40%" align="center"><%=rs("b_date")%></td>
<td width="20%" align="center"><a href="?op=Edit&Type=<%=l_text%>&id=<%=rs("b_id")%>">Düzenle</a> - <a href="?op=Delete&id=<%=rs("b_id")%>">Sil</a></td>
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