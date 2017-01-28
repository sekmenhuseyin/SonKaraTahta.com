<!--#include file="includes.asp" -->
<%
	IF Request.QueryString("op")="Cikis" Then
	Session("BanStats_Enter")="No"
	Session("BanStats_ID")=""
	Response.Redirect "?name="&Request.QueryString("name")&""
	ElseIF Request.QueryString("op")="Giris" Then
%>
<table border="0" cellspacing="2" cellpadding="1" width="100%" align="center" class="td_menu">
<form method="post" action="?name=<%=Request.QueryString("name")%>&op=GirisYap">
<tr>
<td width="40%" align="right"><%=f_bannerid%> : </td>
<td width="60%"><input type="text" name="banner_id" size="20" class="form1"></td>
</tr>
<tr>
<td width="40%" align="right"><%=f_bannerpwd%> : </td>
<td width="60%"><input type="password" name="banner_pass" size="20" class="form1"></td>
</tr>
<tr>
<td width="40%"></td>
<td width="60%"><input type="submit" value="<%=f_submit%>" class="submit"></td>
</tr>
</form>
</table>
<%
	ElseIF Request.QueryString("op")="GirisYap" Then
user_banid=Request.Form("banner_id")
user_pass=duz(Request.Form("banner_pass"))

	IF user_banid="" OR user_pass="" Then
	Response.Write err0
	ELSE
set rs=Connection.Execute("SELECT * FROM BANNERS")
Do Until rs.Eof
If ucase(user_banid)=ucase(rs("banner_id")) then
ban_bid=True
If ucase(user_pass)=ucase(rs("password")) then 
ban_pass=True
Session("BanStats_ID")=rs("banner_id")
Session("BanStats_Enter")="Yes"

Response.Redirect "?name="&Request.QueryString("name")&""
End If
End If
rs.MoveNext
Loop

If ban_bid <> True then
Response.Write err1
ElseIf ban_pass <> True then
Response.Write err2
End If

rs.Close
Set rs=Nothing
	End IF
	ELSE
		IF Session("BanStats_Enter")="Yes" Then
		Set rs=Server.CreateObject("ADODB.RecordSet")
		rs.open "Select * FROM BANNERS WHERE banner_id="&Session("BanStats_ID")&"",Connection,3,3
%>
<table border="0" cellspacing="1" cellpadding="1" width="90%" align="center" class="td_menu" bgcolor="<%=td_border_color%>" style="font-weight:bold">
<tr>
<td colspan="2"><img src="<%=rs("ban_url")%>" border="0" width="400"></td>
</tr>
<tr bgcolor="<%=menu_bg%>" align="right">
<td width="50%"><%=b_bno%> : </td>
<td width="50%"><%=rs("banner_id")%></td>
</tr>
<tr bgcolor="<%=menu_bg%>" align="right">
<td width="50%"><%=b_gosterim%> : </td>
<td width="50%"><%=rs("show")%></td>
</tr>
<tr bgcolor="<%=menu_bg%>" align="right">
<td width="50%"><%=b_tiklanma%> : </td>
<td width="50%"><%=rs("hit")%></td>
</tr>
<tr bgcolor="<%=menu_bg%>" align="right">
<td width="50%"><%=b_akontor%> : </td>
<td width="50%"><%=rs("limit")%></td>
</tr>
<tr bgcolor="<%=menu_bg%>" align="right">
<td width="50%"><%=b_hkontor%> : </td>
<td width="50%"><%=rs("show")%></td>
</tr>
<tr bgcolor="<%=menu_bg%>" align="right">
<td width="50%"><%=b_kkontor%> : </td>
<td width="50%"><%=Int(rs("limit")-rs("show"))%></td>
</tr>
</table>
<br />
<div align="right" style="font-weight:bold"><a href="?name=<%=Request.QueryString("name")%>&op=Cikis">Çýkýþ Yap [X]</a></div>
<%
		ELSE
		Response.Redirect "modules.asp?name="&Request.QueryString("name")&"&op=Giris"
		End IF
	End IF
%>