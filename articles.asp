<!--#include file="view.asp" -->
<%call UST:call ORTA%>
<table border="0" cellspacing="0" cellpadding="2" width="100%" style="<%=TableBorder%>" class="td_menu">
<tr bgcolor="<%=menu_back%>"><td align="center" style="font-size:18px"><%=sett("site_isim")%> : <%=top_menu4%></td></tr>
</table><br />
<table border="0" cellspacing="0" cellpadding="2" width="100%" style="<%=TableBorder%>" class="td_menu">
<tr bgcolor="<%=menu_back%>"><td align="center">
<a href="?action=cats"><img src="images/logos/yazi.gif" border="0"></a><br />
<b>[</b> <a href="?action=HitArticles"><%=title_popular%></a> <b>|</b> <a href="?action=LikeArticles"><%=title_like%></a> <b>|</b> <a href="?action=NewArticles"><%=title_new%></a> <b>|</b> <a href="?action=SendArticle"><%=title_art_send%></a> <b>|</b> <a href="?action=Stats"><%=title_statistics%></a> <b>]</b>
</td></tr>
</table><br />
<%act=Request.QueryString("action"):act=QS_CLEAR(act,"")
If act="Articles" Then
call arts
ElseIf act="Read" Then
call read
ElseIf act="SendArticle" Then
call artsend
ElseIf act="RegArt" Then
call artreg
ElseIf act="Vote" Then
call avote
ElseIF act="HitArticles" or act="NewArticles" or act="LikeArticles" Then
call arts_hitlikenew
ElseIF act="Stats" Then
call stats
Else
call cats
End If

Sub arts_hitlikenew
if act="NewArticles" then
Set n5=Connection.Execute("Select * FROM ARTICLES WHERE a_approved=True ORDER BY a_date DESC"):thepagetitle=title_new
elseif act="HitArticles" then
Set n5=Connection.Execute("Select * FROM ARTICLES WHERE a_approved=True ORDER BY hit DESC"):thepagetitle=title_popular
elseif act="LikeArticles" then
Set n5=Connection.Execute("Select * FROM ARTICLES WHERE a_approved=True ORDER BY point DESC"):thepagetitle=title_like
end if%>
<table border="0" cellspacing="1" cellpadding="2" width="100%" bgcolor="<%=td_border_color%>" align="center">
<tr class="td_menu" style="font-weight:bold;"><td colspan="3" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=thepagetitle%></font></td></tr>
<%For p=1 To 5
IF n5.EoF Then Exit For
Set cat=Connection.Execute("Select * FROM ARTICLE_CATS WHERE cat_id="&n5("cat_id")&"")
Set mem=Connection.Execute("Select * FROM MEMBERS Where kul_adi='"&n5("a_writer")&"'")%>
<tr class="td_menu" bgcolor="<%if p=2 or p=4 then response.write forum_bg_2 else response.write forum_bg_1%>">
<td valign="top"><font face=webdings color="<%=text%>" size=1>4</font> <b><%=duz(n5("a_title"))%></b><br /><b><%=a_writer%> : </b><%if not mem.eof then response.write duz(mem("kul_adi")) else response.write "XXX"%><br /><br /></td>
<td valign="top"><br /><b><%=the_site_cat%> : </b><%=cat("cat_name")%><br /><b><%=the_tarih%> : </b><%=n5("a_date")%></td>
<td valign="top"><br /><b><%=h_averagepoint%> : </b><%=left(n5("point"),4)%><br /><b><%=a_hit%> : </b><%=n5("hit")%></td>
</tr>
<%n5.MoveNext
Next%>
<tr class="td_menu"><td bgcolor="#FFFFFF" align="center" colspan="3"><br /><b>[</b> <a href="<%=Request.ServerVariables("HTTP_REFERER")%>"><%=p_back%></a> <b>]</b><br /><br /></td></tr>
</table>
<%End Sub

Sub stats
Set cnews=Connection.Execute("Select Count(*) AS count FROM COMMENTS WHERE page='articles'")
Set t_cats=Connection.Execute("Select Count(*) AS count FROM ARTICLE_CATS")
Set t_arts=Connection.Execute("Select Count(*) AS count FROM ARTICLES WHERE a_approved=True")
Set art_hits=Connection.Execute("SELECT SUM(hit) AS count FROM ARTICLES WHERE a_approved=True")
Set wait_approve=Connection.Execute("Select Count(*) AS count FROM ARTICLES WHERE a_approved=False")%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=title_statistics%></font></td></tr>
<tr><td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu"><table border="0" cellspacing="0" cellpadding="1" width="100%" bgcolor="<%=td_border_color%>" class="td_menu">
<tr bgcolor="<%=menu_bg%>" class="td_menu2">
<td width="10%" align="center"><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/stats.gif"></td>
<td width="90%"><%=astats_totalcats%> : <b><%=t_cats("count")%></b><br /><%=astats_totalarts%> : <b><%=t_arts("count")%></b><br /><%=astats_totalread%> : <b><%=art_hits("count")%></b><br /><%=astats_totalwa%> : <b><%=wait_approve("count")%></b><br /><%=title_total_comments%> : <b><%=cnews("count")%></b><br /></td>
</tr></table><br /><b>[</b> <a href="<%=Request.ServerVariables("HTTP_REFERER")%>"><%=p_back%></a> <b>]</b></td></tr></table>
<%End Sub

Sub cats%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=top_menu4%></font></td></tr>
<tr><td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu">
<br />
<%Set rs=Connection.Execute("Select * FROM ARTICLE_CATS ORDER BY cat_name ASC")%>
<table width="90%" align="center" cellspacing="1" cellpadding="1" bgcolor="<%=td_border_color%>" class="td_menu" style="<%=TableShadow%>">
  <tr>
<%page_i=0
While Not rs.EOF
Set articles=Connection.Execute("Select Count(*) AS count FROM ARTICLES WHERE cat_id="&rs("cat_id")&" AND a_approved=True")%>
<td width="50%" height="15" bgcolor="<%=menu_back%>"><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/folder.gif" align="absmiddle">&nbsp;<a href="?action=Articles&catid=<%=rs("cat_id")%>"><b><%=duz(rs("cat_name"))%></b></a> (<%=articles("count")%>)</td>
<%rs.MoveNext
page_i=page_i+1
if cint(page_i)=2 then
response.Write("</tr><tr>")
page_i=0
end if%>
<% Wend
if cint(page_i)=1 then response.Write("<td width='50%' height='15' bgcolor="&menu_back&">&nbsp;</td>")%>
</tr>
</table>
<%rs.Close
Set rs=Nothing%>
<br />
</td>
</tr>
</table>
<%End Sub

Sub arts
id=Request.QueryString("catid"):id=QS_CLEAR(id,"false")
Set crs=Server.CreateObject("ADODB.RecordSet"):cSQL="Select * FROM ARTICLE_CATS WHERE cat_id="&id&"":crs.open cSQL,Connection,3,3
Set ars=Server.CreateObject("ADODB.RecordSet"):arSQL="Select * FROM ARTICLES WHERE cat_id="&id&" AND a_approved=True ORDER BY a_date DESC":ars.open arSQL,Connection,3,3
Page=Request.QueryString("Page"):Page=QS_CLEAR(Page,"false")
If Page="" Then Page="1"%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=duz(crs("cat_name"))%></font></td></tr>
<tr><td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu">
<% If ars.eof Then
Response.Write WriteMsg(no_art,100)
ELSE %>
<table border="0" cellspacing="1" cellpadding="1" width="90%" bgcolor="<%=td_border_color%>" class="td_menu">
<%ars.pagesize=sett("prg_sayisi")
ars.absolutepage=Page
sayfa=ars.pagecount
For ar=1 To ars.pagesize
If ars.eof Then Exit For
Set comments=Connection.Execute("Select Count(*) AS count FROM COMMENTS WHERE nid="&ars("aid")&" AND page='articles'")
%>
<tr>
<td bgcolor="<%=menu_back%>">
<% Set mem=Connection.Execute("Select * FROM MEMBERS Where kul_adi='"&ars("a_writer")&"'") %>
<b><a href="?action=Read&aid=<%=ars("aid")%>"><%=duz(ars("a_title"))%></a></b><br />
<font class="td_menu2"><b><%=a_writer%> : </b><% If NOT mem.Eof Then %><a href="members.asp?action=member_details&uid=<%=mem("uye_id")%>"><%=duz(ars("a_writer"))%></a><%Else%><%=duz(ars("a_writer"))%><%End If%> | <b><%=a_hit%> : </b><%=ars("hit")%> | <a href="comments.asp?action=comments&page=articles&id=<%=ars("aid")%>"><%=news_comments%>(<%=comments("count")%>)</a></font>
</td>
</tr>
<%ars.MoveNext
Next%>
</table>
<% End If %>
<br />
  <%
IF Page <> 1 Then Response.Write "<a href=""?action=Articles&catid="&id&"&Page="&(cint(Page)-1)&""">« "&previous_page&"</a>&nbsp;-&nbsp;"
for y=1 to sayfa
if cint(Page)=cint(y) then
response.write "<font class=""td_menu"">["&y&"]</font>"
else
response.write "<a href=""?action=Articles&catid="&id&"&Page="&y&""">["&y&"]</a>"
end if
next
IF NOT ars.Eof Then Response.Write "&nbsp;-&nbsp;<a href=""?action=Articles&catid="&id&"&Page="&(cint(Page)+1)&""">"&next_page&" »</a>"
%></font>
</td>
</tr>
</table>
<%End Sub

Sub read
IF Session("Enter")="Yes" Then
id=Request.QueryString("aid"):id=QS_CLEAR(id,"false")
Set rs=Server.CreateObject("ADODB.RecordSet"):rSQL="Select * FROM ARTICLES WHERE a_approved=True AND aid="&id&"":rs.open rSQL,Connection,3,3
If rs.recordcount=0 then
response.write WriteMsg(art_not_found,100)
Else
rs("hit")=rs("hit") + 1
rs.Update
Set crs=Connection.Execute("Select * FROM ARTICLE_CATS WHERE cat_id="&rs("cat_id")&"")
Set comments=Connection.Execute("Select Count(*) AS count FROM COMMENTS WHERE nid="&rs("aid")&" AND page='articles'")%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<a href="articles.asp?action=Articles&catid=<%=rs("cat_id")%>"><%=duz(crs("cat_name"))%></a> » <%=duz(rs("a_title"))%></font></td></tr>
<tr><td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu">
<div align="left"><%=SmiLey(rs("a_article"))%></div><hr size="1" color="<%=menu_back%>">
<% Set mem=Connection.Execute("Select * FROM MEMBERS Where kul_adi='"&rs("a_writer")&"'") 
if rs("pointer")<>0 then art_total_vote=rs("point")/rs("pointer") else art_total_vote=0%>
<b><%=a_writer%> : </b><% If NOT mem.Eof Then %><a href="members.asp?action=member_details&uid=<%=mem("uye_id")%>"><%=duz(rs("a_writer"))%></a><%Else%><%=duz(rs("a_writer"))%><%End If%> / <b><%=a_hit%></b> : </b><%=rs("hit")%> / <a href="comments.asp?action=comments&page=articles&id=<%=rs("aid")%>"><%=news_comments%>(<%=comments("count")%>)</a>
<br /><br />
<% If Session("Enter")="Yes" and Request.Cookies("MiniNUKE_NV")("op")<>"OK" AND Request.Cookies("MiniNUKE_NV")("aid")<>""&id&"" then %>
<form action="?action=Vote&id=<%=id%>" method="post"><font class="td_menu"><%=vote_article%> : </font>&nbsp;<select name="point" class="form1">
<option value="1"><%=vote1%></option><option value="2"><%=vote2%></option><option value="3"><%=vote3%></option><option value="4"><%=vote4%></option><option value="5"><%=vote5%></option>
</select>&nbsp;<input type="submit" value="<%=poll_button%>" class="submit">
</form><%end if%><font class="td_menu"><%=h_averagepoint%> : <b><%=left(art_total_vote,4)%></b><br><%=h_totalvote%> : <b><%=rs("pointer")%></font>
</td>
</tr>
</table>
<%End If
rs.Close:Set rs=Nothing
ELSE
Response.Write WriteMsg(no_entry,100)
End If
End Sub

Sub artsend
IF Session("Enter")="Yes" Then
Set lcats=Connection.Execute("Select * FROM ARTICLE_CATS ORDER BY cat_name ASC")%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=title_art_send%></font></td></tr>
<tr><td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu">
<script language="JavaScript"><!-- 
function CheckfrmArticle () {
if (document.frmArticle.atitle.value==""){document.getElementById('img_login').innerHTML="Baþlýk eksik!";return false;}
else if (document.frmArticle.aarticle.value==""){document.getElementById('img_login').innerHTML="Mesaj eksik!";return false;}
else {document.frmArticle.art_sbmt.disabled=true; return true;}
}
// --></script>
<form action="?action=RegArt" method="post" name="frmArticle" onSubmit="return CheckfrmArticle();">
<table border="0" cellspacing="0" cellpadding="2" width="90%" class="td_menu" style="font-weight:bold">
<tr><td colspan="2" align="center"><div id="img_login"></div></td></tr>
<tr>
<td width="30%" align="right"><%=art_title%> : </td>
<td width="70%"><input type="text" name="atitle" size="55" class="form1" onFocus="document.getElementById('img_login').innerHTML=''"></td>
</tr>
<tr>
<td width="30%" align="right" valign="top"><%=art_article%> : </td>
<td width="70%">
<script src="htmlarea/editor.js" language="Javascript1.2"></script>
<script language="JavaScript1.2" defer>editor_generate('aarticle');</script>
<textarea name="aarticle" rows="10" cols="55" class="form1"></textarea>
</td>
</tr>
<tr>
<td width="30%" align="right"><%=the_site_cat%> : </td>
<td width="70%"><select name="acat" size="1" class="form1">
<% do while not lcats.eof %>
<option value=<%=lcats("cat_id")%>><%=lcats("cat_name")%></option>
<%lcats.MoveNext
Loop%>
</select></td>
</tr>
<tr><td width="30%" align="right">&nbsp;</td><td width="70%"><input type="submit" name="art_sbmt" class="submit" value="<%=uye_kayit%>"></td></tr>
</table></form>
</td></tr>
</table>
<%ELSE
Response.Write WriteMsg(no_entry,100)
End IF
End Sub ''artsend

Sub artreg
title=Request.Form("atitle")
article=Request.Form("aarticle")
cat=Request.Form("acat")
Set nameCheck=Connection.Execute("Select * FROM ARTICLES WHERE a_title='"&name&"'")
If NOT nameCheck.Eof Then
Response.Write WriteMsg(art_error1,100)
ElseIf title="" Then
Response.Write WriteMsg(art_error2,100)
ElseIf article="" Then
Response.Write WriteMsg(art_error3,100)
ElseIF cat="" Then
Response.Write WriteMsg(art_error4,100)
ELSE
Set pent=Server.CreateObject("ADODB.RecordSet"):peSQL="Select * FROM ARTICLES":pent.open peSQL,Connection,3,3
pent.AddNew
pent("a_title")=title
pent("a_article")=article
pent("hit")=0
pent("cat_id")=cat
pent("a_date")=Date()
pent("a_writer")=Session("Member")
pent("point")=0
pent("a_approved")=False
pent.Update
Response.Write WriteMsg(send_all_success,100)
END IF
END SUB ''regart

Sub avote
aid=Request.QueryString("id"):aid=QS_CLEAR(aid,"false")
IF Request.Cookies("MiniNUKE_NV")("op")="OK" AND Request.Cookies("MiniNUKE_NV")("aid")=""&aid&"" Then
Response.Write WriteMsg(nv_errmsg,100)
ELSE
voteprg=Request.Form("point")
set entvote=Server.CreateObject("ADODB.RecordSet"):evSQL="Select * FROM ARTICLES WHERE aid="&aid&"":entvote.open evSQL,Connection,3,3
thenewpoint=((entvote("point")*entvote("pointer"))+voteprg)/(entvote("pointer")+1)
entvote("point")=thenewpoint
entvote("pointer")=entvote("pointer")+1
entvote.Update
Response.Cookies("MiniNUKE_NV")("op")="OK":Response.Cookies("MiniNUKE_NV")("aid")=""&aid&""
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End IF
End Sub ''pvote

call ORTA2
call ALT %>