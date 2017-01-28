<% Response.Buffer = True %>
<!--#include file="includes.asp" -->
<%
IF Session("Level") = "1" OR Session("Level") = "7" Then

Function DelExtension(MN_STR)
MN_STR = Replace(MN_STR,".gif","",1,-1,1)
MN_STR = Replace(MN_STR,".jpg","",1,-1,1)
MN_STR = Replace(MN_STR,".png","",1,-1,1)
DelExtension = MN_STR
End Function

	IF Request.QueryString("op") = "Delete" Then
Connection.Execute("DELETE FROM AVATARS WHERE a_id = "&Request.QueryString("id")&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
	ElseIF Request.QueryString("op") = "Edit" Then
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM AVATARS WHERE a_id = "&Request.QueryString("id")&"",Connection,3,3

IF Right(rs("a_img"),3) = "gif" Then
opt1 = "selected"
ElseIF Right(rs("a_img"),3) = "jpg" Then
opt2 = "selected"
ElseIF Right(rs("a_img"),3) = "png" Then
opt3 = "selected"
End IF
%>
<table border="0" cellspacing="2" cellpadding="2" width="60%" align="center" class="td_menu2">
<form method="post" action="?op=Update&id=<%=Request.QueryString("id")%>">
<tr bgcolor="#E0E0E0">
<td width="30%" align="right"><b>Ýsim : </b></td>
<td width="70%" align="right">
<input type="text" name="avtr_name" class="form1" style="width:100%" value="<%=rs("a_name")%>" size="20"></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="30%" align="right"><b>Avatar : </b></td>
<td width="70%" align="right">IMAGES/avatars/<input type="text" name="avtr_img" class="form1" style="width:50%" value="<%=DelExtension(rs("a_img"))%>" size="20">
<select name="avtr_img_ex" size="1" class="form1">
<option value=".gif" <%=opt1%>>GIF</option>
<option value=".jpg" <%=opt2%>>JPG</option>
<option value=".png" <%=opt3%>>PNG</option>
</select>
</td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="30%" align="right"></td>
<td width="70%" align="left"><input type="submit" value="Güncelle »" class="submit" style="width:100%"></td>
</tr>
</form>
</table>
<%
	ElseIF Request.QueryString("op") = "Update" Then

name = duz(Request.Form("avtr_name"))
avatar = duz(Request.Form("avtr_img")) & duz(Request.Form("avtr_img_ex"))

IF name="" OR avatar="" Then
Response.Write "<div align=center class=td_menu><b>Lütfen Tüm Alanlarý Doldurunuz</b></div>"
ELSE
Connection.Execute("UPDATE AVATARS Set a_name = '"&name&"' WHERE a_id = "&Request.QueryString("id")&"")
Connection.Execute("UPDATE AVATARS Set a_img = '"&avatar&"' WHERE a_id = "&Request.QueryString("id")&"")
Response.Redirect "avatars.asp"
End IF

	ElseIF Request.QueryString("op") = "New" Then
%>
<table border="0" cellspacing="2" cellpadding="2" width="60%" align="center" class="td_menu2">
<form method="post" action="?op=Register">
<tr bgcolor="#E0E0E0">
<td width="30%" align="right"><b>Ýsim : </b></td>
<td width="70%" align="right">
<input type="text" name="avtr_name" class="form1" style="width:100%" size="20"></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="30%" align="right"><b>Avatar : </b></td>
<td width="70%" align="right">IMAGES/avatars/<input type="text" name="avtr_img" class="form1" style="width:50%" size="20">
<select name="avtr_img_ex" size="1" class="form1">
<option value=".gif">GIF</option>
<option value=".jpg">JPG</option>
<option value=".png">PNG</option>
</select>
</td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="30%" align="right"></td>
<td width="70%" align="left"><input type="submit" value="Ekle   +" class="submit" style="width:100%"></td>
</tr>
</form>
</table>
<%
	ElseIF Request.QueryString("op") = "Register" Then

name = duz(Request.Form("avtr_name"))
avatar = duz(Request.Form("avtr_img")) & duz(Request.Form("avtr_img_ex"))

IF name="" OR avatar="" Then
Response.Write "<div align=center class=td_menu><b>Lütfen Tüm Alanlarý Doldurunuz</b></div>"
ELSE
Connection.Execute("INSERT INTO AVATARS (a_name,a_img) VALUES ('"&name&"','"&avatar&"')")
Response.Redirect "avatars.asp"
End IF

	ELSE
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM AVATARS ORDER BY a_name ASC",Connection,3,3
%>
<table border="0" cellspacing="2" cellpadding="2" width="50%" align="center" class="td_menu2">
<tr><td colspan="3">» <a href="?op=New">Avatar Ekle</a></td></tr>
<tr style="font-weight:bold" bgcolor="#FFCC00">
<td width="30%" align="center">Avatar</td>
<td width="40%" align="center">Ýsim</td>
<td width="30%" align="center">Ýþlem</td>
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
<td width="30%" align="center"><img src="../IMAGES/avatars/<%=rs("a_img")%>"></td>
<td width="40%" align="center"><%=rs("a_name")%></td>
<td width="30%" align="center"><a href="?op=Edit&id=<%=rs("a_id")%>">Düzenle</a> - <a href="?op=Delete&id=<%=rs("a_id")%>">Sil</a></td>
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