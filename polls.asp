<!--#include file="view.asp" -->
<% call UST:call ORTA
action=Request.QueryString("action"):action=QS_CLEAR(action,"")%>
<table border="0" cellspacing="0" cellpadding="2" width="100%" style="<%=TableBorder%>" class="td_menu">
<tr bgcolor="<%=menu_back%>"><td align="center" style="font-size:18px"><%=sett("site_isim")%> : <%=top_menu6%></td></tr>
</table><br /><table border="0" cellspacing="0" cellpadding="2" width="100%" style="<%=TableBorder%>" class="td_menu">
<tr bgcolor="<%=menu_back%>"><td align="center"><a href="?action=polls"><img src="images/logos/anket.gif" border="0"></a><br /><br /></td></tr>
</table><br />
<%if action="polls" then
call all
elseif action="poll_details" then
call details
elseif action="poll_process" then
call poll_process
end if

Sub all
Set rs=Server.CreateObject("ADODB.RecordSet"):SQL="Select * FROM POLLS ORDER BY p_date DESC":rs.open SQL,Connection,3,3
if rs.recordcount>0 then%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=top_menu6%></font></td></tr>
<tr><td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu">
<table border="0" cellspacing="1" cellpadding="2" align="center" width="100%" bgcolor="<%=td_border_color%>" class="td_menu">
<tr style="font-weight:bold;">
<td width="55%" align="left" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=h_baslik%></td>
<td width="15%" align="center" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=poll_totalvotes%></td>
<td width="30%" align="center" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=the_tarih%></td>
</tr>
<%Do Until rs.EoF
Set total_votes=Connection.Execute("SELECT SUM(a_counter) AS count FROM POLL_ALTERNATIVES WHERE pid="&rs("poll_id")&"")%>
<tr bgcolor="<%=menu_bg%>">
<td width="55%"><a href="polls.asp?action=poll_details&pid=<%=rs("poll_id")%>"><%=rs("p_question")%></a></td>
<td width="15%" align="center"><%=total_votes("count")%></td>
<td width="30%" align="center"><%=rs("p_date")%></td>
</tr>
<%rs.MoveNext
Loop%>
</table>
</td>
</tr>
</table>
<%else
response.Write(writemsg(error14,100))
end if
rs.Close:Set rs=Nothing
End Sub

Sub details
pollid=Request.QueryString("pid"):pollid=QS_CLEAR(pollid,"false")
Set prs=Server.CreateObject("ADODB.RecordSet"):pSQL="Select * FROM POLLS WHERE poll_id="&pollid&"":prs.open pSQL,Connection,3,3
Set prs2=Server.CreateObject("ADODB.RecordSet"):pSQL2="Select * FROM POLL_ALTERNATIVES WHERE pid="&pollid&"":prs2.open pSQL2,Connection,3,3
Set prss=Connection.Execute("Select * FROM POLL_ALTERNATIVES WHERE pid="&pollid&"")%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
<tr> <td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=poll_details%></font></td></tr>
<tr> <td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu">
<center><font size="2"><%=prs("p_question")%></font></center><br /><table border="0" cellspacing="3" cellpadding="0" width="100%" align="center" class="td_menu">
<%do while not prs2.eof
Set p_tpl=Connection.Execute("SELECT SUM(a_counter) AS count FROM POLL_ALTERNATIVES WHERE pid="&prs2("pid")&"")
If p_tpl("count")=0 Then percent="0" Else percent=Int((prs2("a_counter") / p_tpl("count")) * 100)%>
<tr><td width="40%" align="right"><b><%=prs2("alternative")%></b></td><td width="60%"><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/poll_result.gif" width="<%=percent%>" height="12">&nbsp;&nbsp;( <%=prs2("a_counter")%> - <%=percent%>% )</td></tr>
<%prs2.MoveNext
Loop %>
</table><br /><center><font class="td_menu" size="1"><b><%=poll_katilan%> : </b><%=p_tpl("count")%></font></center>
</td></tr>
</table>
<%prs.Close:set prs=Nothing
prs2.Close:set prs2=Nothing
End Sub

Sub poll_process
pid=duz(Request.QueryString("poll_id")):pid=QS_CLEAR(pid,"false")
alter=duz(Request("alternative"))
IF Request.Cookies("MiniNuke_POLL")("VOTE")=""&pid&"" Then
Response.Write WriteMsg(poll_error0,100)
ELSE
Set pro=Server.CreateObject("ADODB.RecordSet"):proSQL="Select * FROM POLL_ALTERNATIVES WHERE pid="&pid&" AND a_id="&alter&"":pro.open proSQL,Connection,3,3
pro("a_counter")=pro("a_counter") + 1
pro.Update
Response.Cookies("MiniNuke_POLL")("VOTE")=""&pid&""
Response.Cookies("MiniNuke_POLL").Expires=Date() + 365
pro.Close
Set pro=Nothing
Response.Redirect "polls.asp?action=poll_details&pid="&pid&""
End IF
End Sub
call ORTA2
call ALT %>