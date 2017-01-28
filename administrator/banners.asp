<% Response.Buffer = True %>
<!--#include file="includes.asp" -->
<%
IF Session("Level") = "1" OR Session("Level") = "7" Then

x = Request.QueryString("action")

	IF x = "Banners" Then
		Set rs = Server.CreateObject("ADODB.RecordSet")
		rs.Open "Select * FROM BANNERS",Connection,3,3
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" align="center" class="td_menu2" style="font-size:11px">
<tr bgcolor="#CCCCCC" style="font-weight:bold">
<td width="1%" align="center"></td>
<td width="1%" align="center">ID</td>
<td width="15%">Banner URL</td>
<td width="15%">Gidilecek URL</td>
<td width="10%" align="center">Tip</td>
<td width="5%" align="center">Konum</td>
<td width="10%" align="center">Gösterim</td>
<td width="10%" align="center">Limit</td>
<td width="10%" align="center">Týklanma</td>
<td width="19%" align="center">Ýþlem</td>
</tr>
<%
Do Until rs.EoF
IF rs("active") = True Then
b_opt1 = "False"
b_opt2 = "Pasifleþtir"
b_img = "Active"
ElseIF rs("active") = False Then
b_opt1 = "True"
b_opt2 = "Aktifleþtir"
b_img = "Passive"
End IF
IF rs("position") = "top" Then
b_pos = "Üst"
ElseIF rs("position") = "bottom" Then
b_pos = "Alt"
ELSE
b_pos = "Tanýmsýz"
End IF
IF rs("ban_type") = "Flash" Then
b_type = "Flash"
ELSE
b_type = "Normal"
End IF
%>
<tr bgcolor="#E9E9E9">
<td width="1%" align="center" bgcolor="#FFFFFF"><img src="IMAGES/<%=b_img%>.gif"></td>
<td width="1%" align="center"><%=rs("banner_id")%></td>
<td width="15%"><%=Left(rs("ban_url"),20)%>...</td>
<td width="15%"><%=Left(rs("go_url"),20)%>...</td>
<td width="10%" align="center"><%=b_type%></td>
<td width="5%" align="center"><%=b_pos%></td>
<td width="10%" align="center"><%=rs("show")%></td>
<td width="10%" align="center"><%=rs("limit")%></td>
<td width="10%" align="center"><%=rs("hit")%></td>
<td width="19%" align="center"><a href="?action=Banner&op=Change&option=<%=b_opt1%>&id=<%=rs("banner_id")%>"><%=b_opt2%></a> | <a href="?action=Banner&op=Edit&id=<%=rs("banner_id")%>">Düzenle</a> | <a href="?action=Banner&op=Delete&id=<%=rs("banner_id")%>">Sil</a></td>
</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
rs.Close
Set rs = Nothing
	ElseIF x = "Banner" Then
y = Request.QueryString("op")

		IF y = "New" Then
%>
<form action="?action=Banner&op=Create" method="post">
<table border="0" cellspacing="2" cellpading="2" width="75%" align="center" class="td_menu">
<tr bgcolor="#E0E0E0">
<td width="40%" align="right"><b>Banner URL : </b></td>
<td width="60%">
<input type="text" name="ban_url" class="form1" style="width:100%" value="http://" size="20"></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="40%" align="right"><b>Gidilecek URL : </b></td>
<td width="60%">
<input type="text" name="go_url" class="form1" style="width:100%" value="http://" size="20"></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="40%" align="right"><b>Banner Tipi : </b></td>
<td width="60%">
<select name="type" size="1" class="form1">
<option value="Flash">Flash</option>
<option value="Normal">Normal</option>
</select>
</td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="40%" align="right"><b>Konum : </b></td>
<td width="60%"><select name="konum" class="form1">
<option value="top">Üst</option>
<option value="bottom">Alt</option>
</select></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="40%" align="right"><b>Þifre : </b></td>
<td width="60%">
<input type="text" name="password" class="form1" style="width:100%" size="20"></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="40%" align="right"><b>Kontör : </b></td>
<td width="60%">
<input type="text" name="kontor" class="form1" style="width:100%" size="20"></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="40%" align="right"><b>Aktif / Pasif : </b></td>
<td width="60%"><input type="checkbox" name="aktif" value="ON"></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="40%"></td>
<td width="60%" align="right"><input type="submit" value="  Kaydet  +" class="submit" style="width: 30%"></td>
</tr>
</table>
</form>
<%
		ElseIF y = "Create" Then
ban_url = duz(Request.Form("ban_url"))
go_url = duz(Request.Form("go_url"))
btype = duz(Request.Form("type"))
konum = duz(Request.Form("konum"))
password = duz(Request.Form("password"))
limit = duz(Request.Form("kontor"))
IF Request.Form("aktif") = "on" Then
active = "True"
ELSE
active = "False"
End IF
			IF ban_url="" OR go_url="" OR password="" OR limit="" Then
			Response.Write "<div align=center class=td_menu><b>Tüm Alanlarý Doldurunuz !</b></div>"
			ELSE
			Connection.Execute("INSERT INTO BANNERS (ban_url,go_url,ban_type,hit,position,active,password,limit,show) values ('"&ban_url&"','"&go_url&"','"&btype&"','0','"&konum&"',"&active&",'"&password&"','"&limit&"','0')")
			Response.Redirect "?action=Banners"
			End IF
		ElseIF y = "Edit" Then
Set rs = Connection.Execute("Select * FROM BANNERS WHERE banner_id = "&Request.QueryString("id")&"")

IF rs("ban_type") = "Flash" Then
opt1 = "selected"
ELSE
opt2 = "selected"
End IF
%>
<form action="?action=Banner&op=Update&id=<%=Request.QueryString("id")%>" method="post">
<table border="0" cellspacing="2" cellpading="2" width="75%" align="center" class="td_menu">
<tr bgcolor="#E0E0E0">
<td width="40%" align="right"><b>Banner URL : </b></td>
<td width="60%">
<input type="text" name="ban_url" class="form1" style="width:100%" value="<%=rs("ban_url")%>" size="20"></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="40%" align="right"><b>Gidilecek URL : </b></td>
<td width="60%">
<input type="text" name="go_url" class="form1" style="width:100%" value="<%=rs("go_url")%>" size="20"></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="40%" align="right"><b>Banner Tipi : </b></td>
<td width="60%">
<select name="type" size="1" class="form1">
<option value="Flash" <%=opt1%>>Flash</option>
<option value="Normal" <%=opt2%>>Normal</option>
</select>
</td>
</tr>
<%
IF rs("position") = "top" Then
opt1 = "selected"
ElseIF rs("position") = "bottom" Then
opt2 = "selected"
End IF
%>
<tr bgcolor="#E0E0E0">
<td width="40%" align="right"><b>Konum : </b></td>
<td width="60%"><select name="konum" class="form1">
<option value="top" <%=opt1%>>Üst</option>
<option value="bottom" <%=opt2%>>Alt</option>
</select></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="40%" align="right"><b>Þifre : </b></td>
<td width="60%">
<input type="text" name="password" class="form1" style="width:100%" value="<%=rs("password")%>" size="20"></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="40%" align="right"><b>Kontör : </b></td>
<td width="60%">
<input type="text" name="kontor" class="form1" style="width:100%" value="<%=rs("limit")%>" size="20"></td>
</tr>
<% IF rs("active") = True Then
opt21 = "checked"
End IF %>
<tr bgcolor="#E0E0E0">
<td width="40%" align="right"><b>Aktif / Pasif : </b></td>
<td width="60%"><input type="checkbox" name="aktif" <%=opt21%> value="ON"></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="40%"></td>
<td width="60%" align="right"><input type="submit" value="  Güncelle  »" class="submit" style="width: 30%"></td>
</tr>
</table>
</form>
<%
		ElseIF y = "Update" Then
ban_url = duz(Request.Form("ban_url"))
go_url = duz(Request.Form("go_url"))
btype = duz(Request.Form("type"))
konum = duz(Request.Form("konum"))
password = duz(Request.Form("password"))
limit = duz(Request.Form("kontor"))
IF Request.Form("aktif") = "on" Then
active = "True"
ELSE
active = "False"
End IF
			IF ban_url="" OR go_url="" OR password="" OR limit="" Then
			Response.Write "<div align=center class=td_menu><b>Tüm Alanlarý Doldurunuz !</b></div>"
			ELSE
				SET upd = Server.CreateObject("ADODB.RecordSet")
				upd.open "Select * FROM BANNERS WHERE banner_id = "&Request.QueryString("id")&"",Connection,3,3
				upd("ban_url") = ban_url
				upd("go_url") = go_url
				upd("position") = konum
				upd("active") = active
				upd("password") = password
				upd("limit") = limit
				upd("ban_type") = btype
				upd.Update
			Response.Redirect "?action=Banners"
			End IF
		ElseIF y = "Delete" Then
			Connection.Execute("DELETE FROM BANNERS WHERE banner_id = "&Request.QueryString("id")&"")
			Response.Redirect Request.ServerVariables("HTTP_REFERER")
		ElseIF y = "Change" Then
			Connection.Execute("UPDATE BANNERS SET active = "&Request.QueryString("option")&" WHERE banner_id = "&Request.QueryString("id")&"")
			Response.Redirect Request.ServerVariables("HTTP_REFERER")
		ELSE
		End IF

	End IF

End IF
%>