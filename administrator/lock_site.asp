<!--#include file="includes.asp" -->
<% Response.Buffer = True %>
<%
IF Session("Level") = "1" OR Session("Level") = "7" Then

	IF Request.QueryString("op") = "Lock" Then
locked_msg = Replace(Request.Form("locked_msg"),"'","´",1,-1,1)
		Connection.Execute("UPDATE SETTINGS SET site_locked_msg = '"&locked_msg&"'")
		Connection.Execute("UPDATE SETTINGS SET site_locked = '1'")
		Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	ElseIF Request.QueryString("op") = "UnLock" Then
		Connection.Execute("UPDATE SETTINGS SET site_locked = '0'")
		Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	ELSE
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.Open "Select * FROM SETTINGS",Connection,3,3

IF rs("site_locked") = 1 Then
strAction = "UnLock"
ELSE
strAction = "Lock"
End IF
%>
<table border="0" cellspacing="2" cellpadding="0" width="50%" align="center" class="td_menu">
<form action="?action=Lock/UnLock&op=<%=strAction%>" method="post">
<% IF rs("site_locked") = 1 Then %>
<tr>
<td align="center"><input type="submit" value="Kilidi Aç" style="width:100;height:30" class="submit"></td>
</tr>
<% ELSE %>
<tr>
<td width="25%" valign="top" align="right"><b>Kilit Mesajı : </b></td>
<td width="75%">
<textarea name="locked_msg" rows="10" style="width:100%" class="form1" cols="20">Site Yöneticiler tarafından kilitlenmiştir !</textarea></td>
</tr>
<tr>
<td>&nbsp;</td>
<td><input type="submit" value="Kilitle" style="width:100" class="submit"></td>
</tr>
<% End IF %>
</form>
</table>
<%
	End IF

ELSE
Response.Redirect "http://www.minigirls.com"
End IF
%>