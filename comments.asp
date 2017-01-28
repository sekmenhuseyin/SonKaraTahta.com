<!--#include file="view.asp" -->
<%call UST:call ORTA 
action=Request.QueryString("action")
page=Request.QueryString("page")
IF action="comments" then
call comments
ELSEIF action="new_comment" then
call new_comment
ELSEIF action="comment_register" then
call comment_reg
ELSEIF action="delete" then
call cdel
END IF

Sub new_comment
If Session("Enter")="Yes" Then
id=Request.QueryString("id"):id=QS_CLEAR(id,"false")
If page="news" then
Set rs=Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM NEWS WHERE hid="&id&"",Connection,3,3
ElseIf page="programs" then
Set rs=Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM DOWNLOADS WHERE cid="&id&"",Connection,3,3
sql_conn="Connection"
ElseIf  page="articles" then
Set rs=Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM ARTICLES WHERE aid="&id&"",Connection,3,3
End If%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=comment_new%></font></td></tr>
<tr><td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu">
<form action="comments.asp?page=<%=page%>&action=comment_register&id=<%=id%>" method="post">
<table border="0" cellspacing="0" cellpadding="0" width="90%" align="center" class="td_menu2">
<tr><td width="40%" align="right" valign="top"><%=comment%> : </td><td width="60%"><textarea name="comment" rows="8" cols="40" class="form1"></textarea></td></tr>
<tr><td width="40%">&nbsp;</td><td width="60%" align="left"><input type="submit" value="<%=lang_button%>" class="form1"></td></tr>
</table></form>
</td></tr>
</table>
<% Else
Response.Write WriteMsg(no_entry,100)
End If
End Sub

Sub comment_reg
If Session("Enter")="Yes" Then
id=Request.QueryString("id"):id=QS_CLEAR(id,",false")
com=Request.Form("comment")
Set enter=Server.CreateObject("ADODB.RecordSet"):entSQL="Select * FROM COMMENTS":enter.open entSQL,Connection,3,3
IF com="" Then
Response.Write WriteMsg(empty_comment,100)
ELSE
enter.AddNew
enter("writer")=Session("Member")
enter("comment")=com
enter("cdate")=Now()
enter("nid")=id
enter("page")=page
enter.Update
Response.Redirect "comments.asp?action=comments&page="&page&"&id="&id&""
End If
End If
End Sub

Sub comments
id=Request.QueryString("id"):id=QS_CLEAR(id,"false")
If page="news" then
Set n=Server.CreateObject("ADODB.RecordSet"):n.open "Select * FROM NEWS WHERE hid="&id&"",Connection,3,3
ElseIf page="programs" then
Set n=Server.CreateObject("ADODB.RecordSet"):n.open "Select * FROM DOWNLOADS WHERE pid="&id&"",Connection,3,3
ElseIf  page="articles" then
Set n=Server.CreateObject("ADODB.RecordSet"):n.open "Select * FROM ARTICLES WHERE aid="&id&"",Connection,3,3
End If
Set rs=Server.CreateObject("ADODB.RecordSet"):rs.Open "Select * FROM COMMENTS WHERE nid="&id&" AND page='"&page&"' ORDER BY cdate DESC",Connection,3,3
If page="news" Then
cp_title=n("baslik"):p="news.asp?action=Read&hid="&id
ElseIf page="programs" Then
cp_title=n("p_name"):p="programs.asp?action=programs&catid="&id
ElseIf page="articles" Then
cp_title=n("a_title"):p="articles.asp?action=Read&aid="&id
End If%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=duz(cp_title)%></font></td></tr>
<tr><td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu">
<b><%=news_comments%></b><div align="left"><a href="comments.asp?page=<%=page%>&action=new_comment&id=<%=id%>"><b>» <%=comment_new%></b></a></div><br />
<% do while not rs.eof
Set mem=Connection.Execute("Select * FROM MEMBERS Where kul_adi='"&rs("writer")&"'")%>
<table border="0" cellspacing="1" cellpadding="2" width="90%" bgcolor="<%=td_border_color%>" class="td_menu2">
<tr><td bgcolor="<%=menu_color%>" colspan="2"><%=duz(rs("comment"))%></td></tr>
<tr><td width="60%" bgcolor="<%=menu_color%>"><b><%=comment_writer%> : </b></font><font face=verdana size=1><% If NOT mem.Eof Then%><a href="members.asp?action=member_details&uid=<%=mem("uye_id")%>"><%end if%><%=duz(rs("writer"))%><%if NOT mem.eof Then %></a><%end if%> // <%=rs("cdate")%><% IF Session("Level")="1" OR Session("Level")="4" Then%> // <a href="?action=delete&cid=<%=rs("cid")%>">Yorumu Sil</a><% End If %></td></tr>
</table><br />
<%rs.MoveNext
Loop%>
</td></tr>
<tr><td bgcolor="<%=menu_bg%>" class="td_menu" align="center"><a href='<%=p%>'><%=p_back%></a></td></tr>
</table>
<%End Sub

Sub cdel
IF Session("Level")="1" OR Session("Level")="4" Then
id=Request.QueryString("cid"):id=QS_CLEAR(id,,"false")
If id<>"" then Set dc=Server.CreateObject("ADODB.RecordSet"):dcSQL="DELETE * FROM COMMENTS WHERE cid="&id&"":dc.open dcSQL,Connection,3,3
End If
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub

call ORTA2:call ALT %>