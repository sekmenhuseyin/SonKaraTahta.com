<!--#include file="includes.asp" -->
<%
IF Request.QueryString("op")="Reg" Then
Set Fsys=Server.CreateObject("Scripting.FileSystemObject")
Set Fdeg=Fsys.OpenTextFile(""&dataPath&"", 8)
msg=duz(Request.Form("msg"))
IF msg="" Then
Response.Write err0
ELSE
IF Session("Enter")="Yes" Then
user=Session("Member")
ELSE
user=guest
End IF
Fdeg.WriteLine "<b>"&user&" : </b>"&msg&"<br />"
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End IF
ELSE
Set Fsys=Server.CreateObject("Scripting.FileSystemObject")
Set Fdeg=Fsys.OpenTextFile(""&dataPath&"", 1)
%>
<table border="0" cellspacing="1" cellpadding="1" width="70%" class="td_menu" bgcolor="<%=td_border_color%>">
<tr bgcolor="<%=menu_bg%>">
<td><%=Fdeg.ReadAll%></td>
</tr>
</table>
<%
Fdeg.Close
Set Fdeg=Nothing
Set Fsys=Nothing
%>
<form action="?name=<%=Request.QueryString("name")%>&op=Reg" method="post">
<input type="text" name="msg" size="50" class="form1">&nbsp;<input type="submit" value="<%=msg_submit%>" class="submit">
</form>
<%
End IF
%>