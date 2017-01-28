<!--#include file="view.asp" -->
<%call UST:call ORTA%>
<table border="0" cellspacing="0" cellpadding="2" width="100%" style="<%=TableBorder%>" class="td_menu">
<tr bgcolor="<%=menu_back%>"><td align="center" style="font-size:18px"><%=sett("site_isim")%> : <%=top_menu3%></td></tr>
</table><br /><table border="0" cellspacing="0" cellpadding="2" width="100%" style="<%=TableBorder%>" class="td_menu">
<tr bgcolor="<%=menu_back%>">
<td align="center"><a href="?action=cats"><img src="IMAGES/logos/dosya.gif" border="0"></a><br />
<b>[</b> <a href="?action=HitPrograms"><%=title_popular%></a> <b>|</b> <a href="?action=LikePrograms"><%=title_like%></a> <b>|</b> <a href="?action=NewPrograms"><%=title_new%></a> <b>|</b> <a href="?action=SendPrg"><%=title_prgs_send%></a> <b>|</b> <a href="?action=Stats"><%=title_statistics%></a> <b>]</b>
</td>
</tr>
</table><br />
<%act=Request.QueryString("action"):act=QS_CLEAR(act,"")
If act="programs" Then
call prgs
ElseIf act="p_details" Then
call details
ElseIf act="download" Then
call down
ElseIf act="SendPrg" Then
call prgsend
ElseIf act="RegPrg" Then
call prgreg
ElseIf act="Vote" Then
call pvote
ElseIF act="Stats" Then
call stats
ElseIF act="HitPrograms" or act="NewPrograms" or act="LikePrograms" Then
call HitLikeNew_prgs
Else
call cats
End If

Sub HitLikeNew_prgs
if act="NewPrograms" then
Set n5=Connection.Execute("Select * FROM DOWNLOADS WHERE p_approved=True ORDER BY p_date DESC"):thepagetitle=title_new
elseif act="HitPrograms" then
Set n5=Connection.Execute("Select * FROM DOWNLOADS WHERE p_approved=True ORDER BY p_hit DESC"):thepagetitle=title_popular
elseif act="LikePrograms" then
Set n5=Connection.Execute("Select * FROM DOWNLOADS WHERE p_approved=True ORDER BY point DESC"):thepagetitle=title_like
end if%>
<table border="0" cellspacing="1" cellpadding="2" width="100%" bgcolor="<%=td_border_color%>" align="center">
<tr class="td_menu" style="font-weight:bold;"><td colspan="3" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=thepagetitle%></font></td></tr>
<%For p=1 To 5
if n5.eof Then exit For
Set cat=Connection.Execute("Select * FROM DOWN_CATS WHERE cid="&n5("cid")&"")%>
<tr class="td_menu" bgcolor="<%if p=2 or p=4 then response.write forum_bg_2 else response.write forum_bg_1%>">
<td valign="top"><font face=webdings color="<%=text%>" size=1>4</font> <b><%=duz(n5("p_name"))%></b> &nbsp; - &nbsp; <%=n5("p_info")%><br /><b><%=a_writer%> : </b><%=n5("p_author")%><div><br /><b>[<a class="td_menu2" href="?action=p_details&pid=<%=n5("pid")%>"><%=p_more%></a>]</b></div></td>
<td valign="top"><br /><b><%=the_site_cat%> : </b><%=cat("c_name")%><br /><b><%=the_tarih%> : </b><%=n5("p_date")%></td>
<td valign="top"><br /><b><%=h_averagepoint%> : </b><%=left(n5("point"),4)%><br /><b><%=p_hit%> : </b><%=n5("p_hit")%></td></tr>
<%n5.MoveNext
Next%>
<tr class="td_menu"><td bgcolor="#FFFFFF" align="center" colspan="3"><br /><b>[</b> <a href="<%=Request.ServerVariables("HTTP_REFERER")%>"><%=p_back%></a> <b>]</b><br /><br /></td></tr>
</table>
<%End Sub

Sub stats
Set cnews=Connection.Execute("Select Count(*) AS count FROM COMMENTS WHERE page='programs'")
Set t_cats=Connection.Execute("Select Count(*) AS count FROM DOWN_CATS")
Set t_prgs=Connection.Execute("Select Count(*) AS count FROM DOWNLOADS WHERE p_approved=True")
Set prg_hits=Connection.Execute("SELECT SUM(p_hit) AS count FROM DOWNLOADS WHERE p_approved=True")
Set wait_approve=Connection.Execute("Select Count(*) AS count FROM DOWNLOADS WHERE p_approved=False")%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=title_statistics%></font></td></tr>
<tr><td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu"><table border="0" cellspacing="0" cellpadding="1" width="100%" bgcolor="<%=td_border_color%>" class="td_menu">
<tr bgcolor="<%=menu_bg%>" class="td_menu2">
<td width="10%" align="center">
<img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/stats.gif">
</td>
<td width="90%">
<%=pstats_totalcats%> : <b><%=t_cats("count")%></b><br />
<%=pstats_totalprgs%> : <b><%=t_prgs("count")%></b><br />
<%=pstats_totaldownload%> : <b><%=prg_hits("count")%></b><br />
<%=pstats_totalwa%> : <b><%=wait_approve("count")%></b><br />
<%=title_total_comments%> : <b><%=cnews("count")%></b><br />
</td>
</tr>
</table>
<br />
<b>[</b> <a class="td_menu2" href="<%=Request.ServerVariables("HTTP_REFERER")%>"><%=p_back%></a> <b>]</b>
</td>
</tr>
</table>
<%
End Sub

Sub cats %>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=top_menu3%></font></td></tr>
<tr><td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu"><br /><table width="90%" align="center" cellspacing="1" bgcolor="<%=td_border_color%>" class="td_menu" style="<%=TableShadow%>">
<%Set rs=Server.CreateObject("ADODB.RecordSet"):SQL="Select * FROM DOWN_CATS ORDER BY c_name DESC":rs.open SQL,Connection,3,3
rsrecordcount=rs.recordcount:k=1
for i=1 to rsrecordcount
Set progs=Connection.Execute("Select Count(*) AS count FROM DOWNLOADS WHERE cid="&rs("cid")&" AND p_approved=True")
if k=1 then response.write "<tr>"%>
<td width="50%" height="15" bgcolor="<%=menu_back%>"><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/folder.gif" align="absmiddle">&nbsp;<a href="?action=programs&catid=<%=rs("cid")%>"><b><%=rs("c_name")%></b></a> (<%=progs("count")%>)</td>
<%if k=2 then response.write "</tr>":k=1 else k=2%>
<%rs.MoveNext
next
if k=1 then response.write "<td>&nbsp;</td></tr>"%>
</table><br />
</td></tr>
</table>
<%rs.Close:Set rs=Nothing
End Sub

Sub prgs
id=Request.QueryString("catid"):id=QS_CLEAR(id,"false")
Set crs=Server.CreateObject("ADODB.RecordSet"):cSQL="Select * FROM DOWN_CATS WHERE cid="&id&"":crs.open cSQL,Connection,3,3
Set prs=Server.CreateObject("ADODB.RecordSet"):prSQL="Select * FROM DOWNLOADS WHERE cid="&id&" AND p_approved=True ORDER BY pid DESC":prs.open prSQL,Connection,3,3
Page=Request.QueryString("Page"):Page=QS_CLEAR(Page,"false")
If Page="" Then Page="1"%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=crs("c_name")%></font></td></tr>
<tr><td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu">
<% If prs.eof Then
Response.Write WriteMsg(no_prg,100)
ELSE %>
<table border="0" cellspacing="1" cellpadding="1" width="90%" bgcolor="<%=td_border_color%>">
<%prs.pagesize=sett("prg_sayisi")
prs.absolutepage=Page
sayfa=prs.pagecount
For pr=1 To prs.pagesize
If prs.eof Then Exit For
Set comments=Connection.Execute("Select Count(*) AS count FROM COMMENTS WHERE nid="&prs("pid")&" AND page='programs'")%>
<tr>
<td bgcolor="<%=menu_back%>" class="td_menu">
<b><a href="?action=p_details&pid=<%=prs("pid")%>"><%=prs("p_name")%></a></b><br />
<font class="td_menu2"><b><%=p_author%> : </b><%=prs("p_author")%> | <b><%=p_size%> : </b><%=prs("p_size")%> | <b><%=p_hit%> : </b><%=prs("p_hit")%> | <a href="comments.asp?action=comments&page=programs&id=<%=prs("pid")%>"><%=news_comments%>(<%=comments("count")%>)</a></font>
</td>
</tr>
<%prs.MoveNext
Next%>
</table>
<% End If %>
<br /><%
bp=Page-1
IF Page <> 1 Then Response.Write "<a href=""?action=programs&catid="&id&"&Page="&bp&""">« "&previous_page&"</a>&nbsp;-&nbsp;"
for y=1 to sayfa
if s=y then
response.write y
else
response.write "<a href=""?action=programs&catid="&id&"&Page="&y&"""><font class=""td_menu"">["&y&"]</font></a>"
end if
next
np=Page+1
IF NOT prs.Eof Then Response.Write "&nbsp;-&nbsp;<a href=""?action=programs&catid="&id&"&Page="&np&""">"&next_page&" »</a>"%></font>
</td>
</tr>
</table>
<%End Sub

Sub details
id=Request.QueryString("pid"):id=QS_CLEAR(id,"false")
Set rs=Server.CreateObject("ADODB.RecordSet"):rs.open "Select * FROM DOWNLOADS WHERE p_approved=True AND pid="&id&"",Connection,3,3
Set comments=Connection.Execute("Select Count(*) AS count FROM COMMENTS WHERE nid="&rs("pid")&" AND page='programs'")
Set total_votes=Connection.Execute("Select SUM(point) AS count FROM DOWNLOADS")%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=rs("p_name")%></font></td></tr>
<tr><td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu"><table border="0" cellspacing="2" cellpadding="2" width="100%" align="center" class="td_menu">
<tr><td width="5%" valign="top">
<% IF Len(rs("p_img")) <= 0 Then
Response.Write ""
ELSE %><img src="<%=rs("p_img")%>" border="1"><% End IF %></td><td width="95%" valign="top"><%=duz(rs("p_info"))%></td></tr>
</table>
<hr size="1" color="<%=td_border_color%>">
<b><%=p_author%> : </b><%=rs("p_author")%> / <b><%=p_size%> : </b><%=rs("p_size")%> / <b><%=p_hit%> : </b><%=rs("p_hit")%> / <a href="comments.asp?action=comments&page=programs&id=<%=rs("pid")%>"><%=news_comments%>(<%=comments("count")%>)</a>
<br /><br /><a href="?action=download&pid=<%=id%>"><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/download.gif" border="0"></a><br /><br />
<% If Session("Enter")="Yes" and Request.Cookies("MiniNUKE_NV")("op")<>"OK" AND Request.Cookies("MiniNUKE_NV")("pid")<>""&id&"" then %>
<form action="?action=Vote&id=<%=id%>" method="post"><font class="td_menu2"><%=vote_program%> : </font>&nbsp;<select name="point" class="form1">
<option value="1"><%=vote1%></option><option value="2"><%=vote2%></option><option value="3"><%=vote3%></option><option value="4"><%=vote4%></option><option value="5"><%=vote5%></option>
</select>&nbsp;<input type="submit" value="<%=poll_button%>" class="submit">
</form><% End If %><font class="td_menu"><%=h_averagepoint%> : <b><%=left(rs("point"),4)%></b><br><%=h_totalvote%> : <b><%=rs("pointer")%></font>
</td>
</tr>
</table>
<%
rs.Close
Set rs=Nothing
End Sub

Sub down
IF Session("Enter")="Yes" Then
id=Request.QueryString("pid"):id=QS_CLEAR(id,"false")
Set pdown=Server.CreateObject("ADODB.RecordSet"):pdSQL="Select * FROM DOWNLOADS WHERE p_approved=True AND pid="&id&"":pdown.open pdSQL,Connection,3,3
IF NOT pdown.Eof Then
pdown("p_hit")=pdown("p_hit") + 1
pdown.Update
Response.Redirect ""&pdown("p_url")&""
ELSE
Response.Redirect ""&sett("site_adres")&""
END IF
ELSE
Response.Write WriteMsg(no_entry,100)
End IF
END SUB

Sub prgsend
Set lcats=Connection.Execute("Select * FROM DOWN_CATS")
IF Session("Enter")="Yes" then%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=title_prgs_send%></font></td></tr>
<tr><td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu">
<form action="?action=RegPrg" method="post"><table border="0" cellspacing="0" cellpadding="2" width="90%" class="td_menu2">
<tr><td width="40%" align="right"><%=prg_name%> : </td><td width="60%"><input type="text" name="pname" size="30" class="form1"></td></tr>
<tr><td width="40%" align="right" valign="top"><%=prg_info%> : </td><td width="60%"><textarea name="pinfo" rows="5" cols="30" class="form1"></textarea></td></tr>
<tr><td width="40%" align="right"><%=prg_size%> : </td><td width="60%"><input type="text" name="psize" size="30" class="form1" value="0 KB"></td></tr>
<tr><td width="40%" align="right"><%=prg_url%> : </td><td width="60%"><input type="text" name="purl" size="30" class="form1" value="http://"></td></tr>
<tr><td width="40%" align="right"><%=prg_author%> : </td><td width="60%"><input type="text" name="pauthor" size="30" class="form1"></td></tr>
<tr><td width="40%" align="right"><%=the_site_cat%> : </td><td width="60%"><select name="pcat" size="1" class="form1">
<% do while not lcats.eof %><option value=<%=lcats("cid")%>><%=lcats("c_name")%></option><%lcats.MoveNext
Loop%></td></tr>
<tr><td width="40%" align="right">&nbsp;</td><td width="60%"><input type="submit" class="submit" value="<%=uye_kayit%>"></td></tr>
</table></form>
</td></tr>
</table>
<%ELSE
Response.Write WriteMsg(no_entry,100)
END IF
End Sub ''prgsend

Sub prgreg
name=duz(Request.Form("pname"))
info=duz(Request.Form("pinfo"))
size=duz(Request.Form("psize"))
url=duz(Request.Form("purl"))
author=duz(Request.Form("pauthor"))
cat=duz(Request.Form("pcat"))
Set nameCheck=Connection.Execute("Select * FROM DOWNLOADS WHERE p_name='"&name&"'")
If NOT nameCheck.Eof Then
Response.Write WriteMsg(prg_error1,100)
ElseIf name="" Then
Response.Write WriteMsg(prg_error2,100)
ElseIf info="" Then
Response.Write WriteMsg(prg_error3,100)
ElseIf url="" Then
Response.Write WriteMsg(prg_error4,100)
ElseIf author="" Then
Response.Write WriteMsg(prg_error5,100)
ElseIF cat="" Then
Response.Write WriteMsg(prg_error6,100)
ELSE
Set pent=Server.CreateObject("ADODB.RecordSet"):peSQL="Select * FROM DOWNLOADS":pent.open peSQL,Connection,3,3
pent.AddNew
pent("p_name")=name
pent("p_info")=info
pent("p_size")=size
pent("p_hit")=0
pent("p_url")=url
pent("cid")=cat
pent("p_date")=Date()
pent("p_author")=author
pent("point")=0
pent("p_approved")=False
pent("p_img")=""
pent.Update
Response.Write WriteMsg(send_all_success,100)
END IF
END SUB ''prgreg

Sub pvote
pid=Request.QueryString("id"):pid=QS_CLEAR(pid,"false")
IF Request.Cookies("MiniNUKE_NV")("op")="OK" AND Request.Cookies("MiniNUKE_NV")("pid")=""&pid&"" Then
Response.Write WriteMsg(nv_errmsg,100)
ELSE
voteprg=Request.Form("point")
set entvote=Server.CreateObject("ADODB.RecordSet"):evSQL="Select * FROM DOWNLOADS WHERE pid="&pid&"":entvote.open evSQL,Connection,3,3
thenewpoint=((entvote("point")*entvote("pointer"))+voteprg)/(entvote("pointer")+1)
entvote("point")=thenewpoint
entvote("pointer")=entvote("pointer")+1
entvote.Update
Response.Cookies("MiniNUKE_NV")("op")="OK":Response.Cookies("MiniNUKE_NV")("pid")=""&pid&""
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End IF
End Sub ''pvote

call ORTA2
call ALT
%>