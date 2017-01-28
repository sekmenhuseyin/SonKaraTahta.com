<!--#include file="view.asp" -->
<%call UST:call ORTA
page_id=Request.QueryString("id"):page_id=QS_CLEAR(page_id,"false")
Set pageRs=Server.CreateObject("ADODB.RecordSet"):pageRs.open "SELECT * FROM PAGES WHERE p_id="&page_id&"",Connection,3,3
IF pageRs.Eof OR PageRs.Bof OR PageRs.recordcount=0 Then
Response.Write WriteMsg(module_eof,100)
ELSE%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=pageRs("p_title")%></font></td></tr>
<tr><td align="left" valign="top" bgcolor="<%=menu_bg%>" class="td_menu">
<%page_content=pageRs("p_content")
page_content=Replace(page_content, vbCrLf, "<br />", 1, -1, 1)
If pageRs("p_members")=True Then
If Session("Enter")="Yes" Then Response.Write page_content Else Response.Write "<center>"&sett("np_msg")&"</center>"
ELSE
Response.Write page_content
End IF%>
</td>
</tr>
</table>
<%End IF
call ORTA2:call ALT %>