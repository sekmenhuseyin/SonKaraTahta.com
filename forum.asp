<!--#include file="view.asp" -->
<%Set sett2=Connect.Execute("Select * FROM SETTINGS"):forum_topics=sett2("f_posts"):forum_replies=sett2("f_replies"):set sett2=Nothing
strFloodTime=sett("flood_time")
IF Session("Enter")="Yes" Then Session("Forum_LastTime")=Request.Cookies("MiniNuke_Forum")("LastTime") ELSE Session("Forum_LastTime")=Now()
act=Request.QueryString("action"):msgid=Request.QueryString("id"):msgid=QS_CLEAR(msgid,"false")
call UST%>
<table border="0" cellspacing="0" cellpadding="2" width="70%" style="<%=TableBorder%>" class="td_menu">
<tr bgcolor="<%=menu_back%>"><td align="center" style="font-size:18px"><%=sett("site_isim")%> : <%=top_menu7%></td></tr>
</table><br />
<table border="0" cellspacing="0" cellpadding="2" width="70%" style="<%=TableBorder%>" class="td_menu">
<tr bgcolor="<%=menu_back%>"><td align="center"><a href="forum.asp"><img src="images/logos/forum.gif" border="0"></a><br /><b>[</b> <a href="?action=HitForums"><%=title_popular%></a> <b>|</b> <a href="?action=LikeForums"><%=title_like%></a> <b>|</b> <a href="?action=NewForums"><%=title_new%></a> 
<%if act="Topics" then%><b>|</b> <a href="?action=Msg&id=<%=msgid%>"><%=f_new_msg%></a> <%elseif act="Topic" then%><b>|</b> <a href="?action=Rep&id=<%=msgid%>"><%=f_reply_msg%></a> <%end if%><b>|</b> <a href="?action=Stats"><%=title_statistics%></a> <b>]</b>
</td></tr>
</table><br />
<b><font class="td_menu"><a href="Forum.asp"><%=sett("site_isim")%>&nbsp;<%=top_menu7%></a></font></b>
<%IF Session("Enter")<>"Yes" Then
Response.Write WriteMsg(no_entry,90)
ELSE

Session.LCID=S_LCID
If act="" Then
call Main''themain''
ElseIf act="Topics" Then
call Topics
ElseIf act="Topic" Then
call Topic
ElseIf act="Control" Then
call Control
ElseIf act="Msg" Then
call msg
ElseIf act="MsgReg" Then
call reg
ElseIf act="Rep" Then
call rep
ElseIf act="RepReg" Then
call repreg
ElseIF act="Search" Then
call f_search
ElseIF act="Results" Then
call f_results
ElseIF act="EditPost" Then
call post_edit
ElseIF act="UpdatePost" Then
call post_update
ElseIF act="RegQuickReply" Then
call reg_qr
ElseIF act="HitForums" or act="LikeForums" or act="NewForums" Then
call forums_HitLikeNew
ElseIF act="Stats" Then
call forum_stats
End If

Sub themain
Set l_forum=Server.CreateObject("ADODB.RecordSet"):l_forum.open "Select * FROM MESSAGES WHERE topic=True",Connect,3,3:therecordcount=l_forum.recordcount
for i=1 to therecordcount
Set reps=Connect.Execute("Select Count(*) AS rCount FROM MESSAGES WHERE mhit="&l_forum("mid")&" AND topic=False")
if reps("rCount")>0 then Set reps2=Server.CreateObject("ADODB.RecordSet"):reps2.open "Select * FROM MESSAGES WHERE mhit="&l_forum("mid")&" AND topic=False order by mdate DESC",Connect,3,3:thetarih=reps2("mdate"):reps2.close:set reps2=Nothing else thetarih=l_forum("mdate")
Set theupdate=Server.CreateObject("ADODB.RecordSet"):theupdate.open "Select * FROM MESSAGES WHERE mid="&l_forum("mid")&"",Connect,3,3:theupdate("cvp_sayisi")=reps("rCount"):theupdate("last_post")=thetarih:theupdate.update:theupdate.close:set theupdate=Nothing
l_forum.movenext:next
l_forum.close:set l_forum=Nothing
Set l_forum=Server.CreateObject("ADODB.RecordSet"):l_forum.open "Select * FROM MESSAGES WHERE topic=False",Connect,3,3:therecordcount=l_forum.recordcount
for i=1 to therecordcount
l_forum("cvp_sayisi")=0:l_forum("last_post")=l_forum("mdate"):l_forum.update
l_forum.movenext:next
l_forum.close:set l_forum=Nothing
end sub

Sub forums_HitLikeNew
if act="NewForums" then
Set rs=Server.CreateObject("ADODB.RecordSet"):rs.open "SELECT MESSAGES.*, CATEGORIES.cname FROM CATEGORIES INNER JOIN MESSAGES ON CATEGORIES.cid = MESSAGES.cid WHERE (((MESSAGES.topic)=True) AND ((MESSAGES.locked)=False)) ORDER BY MESSAGES.mdate DESC;",Connect,3,3:thepagetitle=title_new
elseif act="LikeForums" then
Set rs=Server.CreateObject("ADODB.RecordSet"):rs.open "SELECT MESSAGES.*, CATEGORIES.cname FROM CATEGORIES INNER JOIN MESSAGES ON CATEGORIES.cid = MESSAGES.cid WHERE (((MESSAGES.topic)=True) AND ((MESSAGES.locked)=False)) ORDER BY MESSAGES.mhit DESC",Connect,3,3:thepagetitle=title_like
elseif act="HitForums" then
Set rs=Server.CreateObject("ADODB.RecordSet"):rs.open "SELECT MESSAGES.*, CATEGORIES.cname FROM CATEGORIES INNER JOIN MESSAGES ON CATEGORIES.cid = MESSAGES.cid WHERE (((MESSAGES.topic)=True) AND ((MESSAGES.locked)=False)) ORDER BY MESSAGES.cvp_sayisi DESC",Connect,3,3:thepagetitle=title_popular
end if%>
<table border="0" cellspacing="1" cellpadding="2" width="90%" bgcolor="<%=td_border_color%>" align="center">
<tr class="td_menu" style="font-weight:bold;"><td colspan="3" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=thepagetitle%></font></td></tr><%For p=1 To 10
IF rs.EoF Then Exit For
%><tr class="td_menu" bgcolor="<%if p=2 or p=4 then response.write forum_bg_2 else response.write forum_bg_1%>">
<td valign="top"><font face=webdings color="<%=text%>" size=1>4</font> <b><%=duz(rs("mtitle"))%></b><br /><b><%=a_writer%> : </b><%=rs("mwriter")%>
<div><br /><b>[<a href="?action=Topic&id=<%=rs("mid")%>"><%=a_read%></a>]</b></div>
</td>
<td valign="top"><br /><b><%=the_site_cat%> : </b><%=rs("cname")%><br /><b><%=the_tarih%> : </b><%=rs("mdate")%></td>
<td valign="top"><br /><b><%=forum_cevap_count%> : </b><%=rs("cvp_sayisi")%><br /><b><%=a_hit%> : </b><%=rs("mhit")%></td>
</tr><%rs.MoveNext
Next%><tr class="td_menu"><td bgcolor="#FFFFFF" align="center" colspan="3"><br /><b>[</b> <a href="<%=Request.ServerVariables("HTTP_REFERER")%>"><%=p_back%></a> <b>]</b><br /><br /></td></tr>
</table><%
End Sub

Sub forum_stats
Set mforum=Connect.Execute("Select Count(*) AS count FROM MAIN_CATS")
Set forum=Connect.Execute("Select Count(*) AS count FROM CATEGORIES")
Set msgs=Connect.Execute("Select Count(*) AS count FROM MESSAGES WHERE topic=True")
Set reps=Connect.Execute("Select Count(*) AS count FROM MESSAGES WHERE topic=False")%>
<table border="0" cellpadding="2" cellspacing="0" width="90%" style="<%=TableBorder%>">
<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=title_statistics%></font></td></tr>
<tr><td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu"><table border="0" cellspacing="0" cellpadding="1" width="100%" bgcolor="<%=td_border_color%>" class="td_menu">
<tr bgcolor="<%=menu_bg%>" class="td_menu2"><td width="10%" align="center"><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/stats.gif"></td><td width="90%">
Toplam Kategori : <b><%=mforum("count")%></b><br />
Toplam Forum : <b><%=forum("count")%></b><br />
Toplam Konu : <b><%=msgs("count")%></b><br />
Toplam Yanýt : <b><%=reps("count")%></b>
</td></tr>
</table><br /><b>[</b> <a href="<%=Request.ServerVariables("HTTP_REFERER")%>"><%=p_back%></a> <b>]</b></td></tr>
</table>
<%End Sub

Sub reg_qr
IF DateDiff("N" ,Session("FloodTimer"),Now()) > strFloodTime Then
amsg=Request.Form("f_rep")
IF amsg="" Then
Response.Write WriteMsg(f_error2,90)
ELSE%>
<script Language="JavaScript1.2">if (win_ie_ver < 5.5) {<% html_allow=0 %>} else {<% html_allow=1 %>}</script>
<%FOR I=1 TO UBound(saryHTMLtags)
amsg=Replace(amsg,"<"&saryHTMLtags(""&I&"")&">","&lt;"&saryHTMLtags(""&I&"")&"&gt;",1,-1,1)
Next
Set ms=Connect.Execute("Select * FROM MESSAGES WHERE mid="&msgid&"")
Set rent=Server.CreateObject("ADODB.RecordSet"):rSQL="Select * FROM MESSAGES":rent.open rSQL,Connect,3,3
rent.AddNew
rent("mmsg")=amsg
rent("mwriter")=Session("Member")
rent("mdate")=Now()
rent("last_post")=Now()
rent("mhit")=msgid
rent("cid")=ms("cid")
rent("html_allow")=html_allow
rent("topic")=False
rent("locked")=False	
rent.Update
Set Rmsg1=Connection.Execute("Select * FROM MEMBERS WHERE uye_id="&Session("Uye_ID")&"")
Set Rmsg=Connection.Execute("UPDATE MEMBERS SET msg_sayisi="&Rmsg1("msg_sayisi")+1&" WHERE uye_id="&Session("Uye_ID")&"")
Set Rmsg=Connect.Execute("UPDATE MESSAGES SET last_post=now(),cvp_sayisi='"&ms("cvp_sayisi")+1&"' WHERE mid="&msgid&"")
Session("FloodTimerStart")="Yes"
Session("FloodTimer")=Rmsg1("son_tarih")
Set Rmsg1=Nothing:Set Rmsg=Nothing
Response.Redirect("Forum.asp?action=Topic&id=" & msgid)
END IF
ELSE
Response.Write WriteMsg(f_error6,90)
End IF
End Sub

Sub post_update
amsg=Request.Form("amsg")
sc=Request.Form("f_ep_sc")
IF amsg="" Then
Response.Write "<hr color="&td_border_color&"><div align=center class=td_menu2>"&f_error2&"</div><hr color="&td_border_color&">"
ElseIF sc="" Then
Response.Write "<hr color="&td_border_color&"><div align=center class=td_menu2>"&sc_err1&"</div><hr color="&td_border_color&">"
ElseIF Session("F_EP_SC") <> sc Then
Response.Write "<hr color="&td_border_color&"><div align=center class=td_menu2>"&sc_err2&"</div><hr color="&td_border_color&">"
ELSE%>
<script Language="JavaScript1.2">if (win_ie_ver < 5.5) {<% html_allow=False %>} else {<% html_allow=True %>}</script>
<%FOR I=1 TO UBound(saryHTMLtags)
amsg=Replace(amsg,"<"&saryHTMLtags(""&I&"")&">","&lt;"&saryHTMLtags(""&I&"")&"&gt;",1,-1,1)
Next
Set rent=Server.CreateObject("ADODB.RecordSet"):rSQL="Select * FROM MESSAGES WHERE mid="&msgid&"":rent.open rSQL,Connect,3,3
IF rent("mwriter")=Session("Member") Then
rent("mmsg")=amsg
rent("mdate")=Now()
rent("html_allow")=html_allow
rent.Update
IF rent("topic")=false Then msgid=rent("mhit")
Response.Redirect "Forum.asp?action=Topic&id="&msgid&""
End IF
END IF
End Sub

Sub post_edit
Set rs=Server.CreateObject("ADODB.RecordSet"):rs.open "Select * FROM MESSAGES WHERE mid="&msgid&"",Connect,3,3
IF rs("mwriter")=Session("Member") Then
Session("F_EP_SC")=CPASS(4)%>
<script src="htmlarea/editor.js" language="Javascript1.2"></script><script language="JavaScript1.2" defer>editor_generate('amsg');</script>
<script language="JavaScript"><!--
function CheckForm(){
if (document.ForumEdit.amsg.value==""){document.getElementById('hint_forum').innerHTML="<b><%=f_error2%></b><br />";return false;}
else if (document.ForumEdit.f_ep_sc.value!=<%=Session("F_EP_SC")%>){document.getElementById('hint_forum').innerHTML="<b><%=sc_err2%></b><br />";return false;}
else {document.ForumEdit.forum_sbmt.disabled=true;return true;}
}
// --></script>
<form action="?action=UpdatePost&id=<%=msgid%>" method="post" name="ForumEdit" onSubmit="return CheckForm();">
<table border="0" cellspacing="5" cellpadding="3" width="90%" align="center" class="td_menu" style="font-weight:bold" bgcolor="<%=background%>">
<tr><td colspan="2"><span id="hint_forum"></span></td></tr>
<tr>
<td width="30%" align="right" bgcolor="<%=menu_back%>" valign="top"><%=nm_msg%> :</td>
<td bgcolor="<%=menu_back%>"><textarea name="amsg" rows="15" class="form1" style="width:100%" cols="20"><%=rs("mmsg")%></textarea></td>
</tr>
<tr>
<td align="right" bgcolor="<%=menu_back%>"><%=security_code%> : </td>
<td bgcolor="<%=menu_back%>"><%=Session("F_EP_SC")%></td>
</tr>
<tr>
<td align="right" bgcolor="<%=menu_back%>"><%=security_code_type%> : </td>
<td bgcolor="<%=menu_back%>"><input type="text" name="f_ep_sc" size="4" class="form1"></td>
</tr>
<tr><td colspan="2" bgcolor="<%=background%>" align="center"><input type="submit" class="submit" value="<%=uye_kayit%>" style="width:40%" name="forum_sbmt">&nbsp;&nbsp;&nbsp;
<input type="button" class="submit" value="Ý P T A L" style="width:40%" onClick="window.location.href='?action=Topic&id=<%=msgid%>';"></td></tr>
</table>
</form>
<%ELSE
Response.Write "<hr color="&td_border_color&"><div align=center>" & f_error4 & "</div><hr color="&td_border_color&">"
End IF
rs.Close:Set rs=Nothing
End Sub

Sub f_results
id=duz(Request.Form("forums"))
if left(id,1)="x" then
id=right(id,len(id)-1)
Set cats=Server.CreateObject("ADODB.RecordSet"):cats.open "Select * FROM CATEGORIES WHERE mcid="&id&"",Connect,3,3
therecordcount=cats.recordcount
else
therecordcount=1
end if%>
<table border="0" cellspacing="1" cellpadding="1" width="90%" class="td_menu" bgcolor="<%=td_border_color%>" align="center">
<tr height="20">
<td colspan="2" width="50%" background="THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><font class="block_title"><%=nm_title%></font></td>
<%if therecordcount>1 then%><td align="center" background="THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><font class="block_title"><%=msg_subject%></font></td><%end if%>
<td align="center" background="THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><font class="block_title"><%=fr_pop5%></font></td>
<td align="center" background="THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><font class="block_title"><%=fw_pop5%></font></td>
</tr>
<%for i=1 to therecordcount
if therecordcount>1 then id=cats("cid")
if id="all" then thesql="Select * FROM MESSAGES WHERE mtitle LIKE '%"&duz(Request.Form("word"))&"%'" else thesql="Select * FROM MESSAGES WHERE cid="&id&" AND mtitle LIKE '%"&duz(Request.Form("word"))&"%' AND topic=True"
Set rs=Server.CreateObject("ADODB.RecordSet"):rs.open thesql,Connect,3,3
Do Until rs.EoF
Set replies=Connect.Execute("Select Count(*) AS count FROM MESSAGES WHERE mhit="&rs("mid")&" AND topic=False")
Set memInfo=Connection.Execute("Select * FROM MEMBERS WHERE kul_adi='"&rs("mwriter")&"'")%>
<tr bgcolor="<%=menu_color%>">
<td width="1%"><%IF rs("locked")=True then %><img src="THEMES/<%=Session("SiteTheme")%>/locked_msg.gif"><% ElseIF replies("count") >= 10 Then %><img src="THEMES/<%=Session("SiteTheme")%>/hot_msg.gif"><% Else %><img src="THEMES/<%=Session("SiteTheme")%>/normal_msg.gif"><% End If %></td>
<td width="50%"><a href="?action=Topic&id=<%=rs("mid")%>"><%=rs("mtitle")%></a></td>
<%if therecordcount>1 then%><td align="center"><%=cats("cname")%></td><%end if%>
<td align="center"><%=replies("count")%></td>
<td align="center"><%IF memInfo.EoF OR memInfo.BoF Then Response.Write rs("mwriter") ELSE Response.Write "<a href=""Members.asp?action=member_details&uid="&memInfo("uye_id")&""">" & rs("mwriter") & "</a>"%></td>
</tr>
<%rs.MoveNext
Loop
if therecordcount>1 then cats.movenext
next%>
</table><br /><div align=center class=td_menu2><%Response.Write fs_result_1 & "<b>" & rs.RecordCount & "</b>" & fs_result_2%></div>
<%rs.Close:Set rs=Nothing
if therecordcount>1 then cats.Close:Set cats=Nothing
End Sub

Sub f_search
Set mcats=Server.CreateObject("ADODB.RecordSet"):mcats.open "Select * FROM MAIN_CATS ORDER BY mcid ASC",Connect,3,3%>
<table border="0" cellspacing="1" cellpadding="1" width="40%" align="center" class="td_menu"><form method="post" action="?action=Results">
<tr bgcolor="<%=menu_back%>"><td width="50%" align="right"><%=fs_word%> : </td><td width="50%"><input type="text" name="word" size="40" class="form1"></td></tr>
<tr bgcolor="<%=menu_back%>"><td width="50%" align="right"><%=fs_forums%> : </td><td width="50%"><select name="forums" size="1" class="form1" style="text-align:left;">
<option value="all">Tüm forumda ara</option><%mcatsrecordcount=mcats.recordcount
for i=1 to mcatsrecordcount%><option value="x<%=mcats("mcid")%>"><%=mcats("mc_name")%></option><%
Set fRs=Server.CreateObject("ADODB.RecordSet"):fRs.open "Select * FROM CATEGORIES WHERE mcid="&mcats("mcid")&" ORDER BY cname DESC",Connect,3,3:fRsrecordcount=fRs.recordcount
for j=1 to fRsrecordcount%><option value="<%=fRs("cid")%>">---<%=fRs("cname")%></option><%fRs.MoveNext
next
mcats.MoveNext
next%>
</select></td></tr>
<tr bgcolor="<%=menu_back%>"><td width="50%" align="right"></td><td width="50%"><input type="submit" value="   <%=fs_button%>   " class="submit"></td></tr>
</form></table>
<%fRs.Close:Set fRs=Nothing
mcats.Close:Set mcats=Nothing
End Sub

Sub Main
Set mcats=Server.CreateObject("ADODB.RecordSet"):mcats.open "Select * FROM MAIN_CATS ORDER BY mcid ASC",Connect,3,3:mcatsrecordcount=mcats.recordcount
for i=1 to mcatsrecordcount
Set cats=Server.CreateObject("ADODB.RecordSet"):cats.open "Select * FROM CATEGORIES WHERE mcid="&mcats("mcid")&" ORDER BY cname",Connect,3,3:catsrecordcount=cats.recordcount%>
<table border="0" cellspacing="1" cellpadding="2" width="90%" bgcolor="<%=td_border_color%>" align="center">
<tr class="td_menu" style="font-weight:bold;">
<td width="2%" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif">&nbsp;</td>
<td width="48%" height="20" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><div class="block_title"><%=mcats("mc_name")%><%IF Session("Level")="1" OR Session("Level")="5" Then%> (<a href="Mods.asp?action=EditMain&id=<%=mcats("mcid")%>" style="color:<%=text%>;"><%=f_edit%></a><%IF Session("Level")="1" Then%> - <a href="Mods.asp?action=DelMain&id=<%=mcats("mcid")%>" style="color:<%=text%>;" OnClick="return confirm('<%=onay_for_silme%>')"><%=f_del%></a><%End IF%>)<%End IF%></div></td>
<td width="5%" align="center" height="20" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><font class="block_title"><%=m_topics%></font></td>
<td width="5%" align="center" height="20" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><font class="block_title"><%=m_answers%></font></td>
<td width="50%" align="center" height="20" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><font class="block_title"><%=m_lastpost%></font></td>
</tr>
<%for j=1 to catsrecordcount
Select Case j
Case "1","3","5","7","9":color=forum_bg_1
Case Else:color=forum_bg_2
End Select
Set tcount=Connect.Execute("Select Count(*) AS count FROM MESSAGES WHERE cid="&cats("cid")&" AND topic=True")
Set acount=Connect.Execute("Select Count(*) AS count FROM MESSAGES WHERE cid="&cats("cid")&" AND topic=False")
Set last=Connect.Execute("Select * FROM MESSAGES WHERE cid="&cats("cid")&" ORDER BY mdate DESC")
If last.Eof Then
last_post="-":last_pmember="-":last_pmid=""
ELSE
last_post=last("last_post"):Set last3=Connection.Execute("Select * FROM MEMBERS WHERE kul_adi='"&last("mwriter")&"'")
IF last3.Eof OR last3.Bof Then last_pmember="###":last_pmid="###" ELSE last_pmember=last3("kul_adi"):last_pmid=last3("uye_id")
End If%>
<tr class="td_menu"><td width="2%" bgcolor="<%=background%>">
<%if last_post<>"-" then
if year(last_post)=>year(Session("LastLogged")) then
if month(last_post)=>month(Session("LastLogged")) then
if day(last_post)=>day(Session("LastLogged")) then FLT=True  else FLT=False
else:FLT=False:end if
else:FLT=False:end if
end if
IF cats("locked")=True then
Response.Write "<img src="&sett("site_adres")&"/THEMES/"&Session("SiteTheme")&"/locked.gif>"
ElseIF cats("noentry")=True Then
Response.Write "<img src="&sett("site_adres")&"/THEMES/"&Session("SiteTheme")&"/no_entry_forum.gif>"
ElseIF FLT=True Then
Response.Write "<img src="&sett("site_adres")&"/THEMES/"&Session("SiteTheme")&"/forum_open_new.gif>"
ELSE
Response.Write "<img src="&sett("site_adres")&"/THEMES/"&Session("SiteTheme")&"/forum_open.gif>"
End If%>
</td><td width="48%" bgcolor="<%=color%>"><b><a href="?action=Control&id=<%=cats("cid")%>"><%=cats("cname")%></a></b>
<%Set ismoderator=Connect.Execute("Select * FROM MODERATORS WHERE cid="&cats("cid")&" and uye_id="&Session("Uye_ID")&""):if ismoderator.eof=false and ismoderator.bof=false then ismod=true else ismod=false
If Session("Level")="1" OR Session("Level")="5" or ismod=true Then%> &nbsp; (<a href="Mods.asp?action=EditCat&id=<%=cats("cid")%>"><%=f_edit%></a>
<% If cats("locked")=True Then%> - <a href="Mods.asp?action=UnLock&id=<%=cats("cid")%>"><%=f_unlock%></a><%ElseIf cats("locked")=False Then%> - <a href="Mods.asp?action=Lock&id=<%=cats("cid")%>"><%=f_lock%></a><% End If %>
<% If cats("noentry")=True Then%> - <a href="Mods.asp?action=SetEntry&id=<%=cats("cid")%>"><%=f_entry%></a><%ElseIf cats("noentry")=False Then%> - <a href="Mods.asp?action=SetNoEntry&id=<%=cats("cid")%>"><%=f_noentry%></a><% End If %>
<%If Session("Level")="1" Then%> - <a href="Mods.asp?action=DeleteCat&id=<%=cats("cid")%>" OnClick="return confirm('<%=onay_for_silme%>')"><%=f_del%></a><% End IF %>)
<% End IF%>
<br /><%=cats("cinfo")%><%Set rsmods=Server.CreateObject("ADODB.RecordSet"):rsmods.open "Select * FROM MODERATORS where cid="&cats("cid"),Connect,3,3:rsmodsrecordcount=rsmods.recordcount
if rsmodsrecordcount>0 then response.write "<br /><font class='td_menu2'>"&forum_mods&": "
FOR k=1 TO rsmodsrecordcount
Set frm_mem=Connection.Execute("Select * FROM MEMBERS WHERE uye_id="&rsmods("uye_id"))
IF frm_mem.EOF OR frm_mem.BOF Then Response.Write frm_mem("kul_adi") ELSE Response.Write "<a href=""members.asp?action=member_details&uid="&frm_mem("uye_id")&""">"&frm_mem("kul_adi")&"</a>"
if rsmodsrecordcount<>k then Response.Write", ":rsmods.movenext
Next%></font></td>
<td width="5%" align="center" bgcolor="<%=color%>"><%=tcount("count")%></td>
<td width="5%" align="center"bgcolor="<%=color%>"><%=acount("count")%></td>
<td width="50%" align="center" bgcolor="<%=color%>"><%=last_post%><br /><%if last_pmember="-" then response.Write(last_pmember) else response.Write("<a href='members.asp?action=member_details&uid="&last_pmid&"'>"&last_pmember&"</a>")%></td>
</tr><%cats.MoveNext:next%>
</table><br /><%mcats.MoveNext:next%>
<table border="0" cellspacing="0" cellpaddin="2" align="center" width="90%" class="td_menu">
<tr>
<td><img src="THEMES/<%=Session("SiteTheme")%>/forum_open.gif" alt="<%=bf_normal%>" align="absmiddle">&nbsp;<%=bf_normal%></td>
<td><img src="THEMES/<%=Session("SiteTheme")%>/forum_open_new.gif" alt="<%=bf_new%>" align="absmiddle">&nbsp;<%=bf_new%></td>
<td><img src="THEMES/<%=Session("SiteTheme")%>/locked.gif" alt="<%=bf_locked%>" align="absmiddle">&nbsp;<%=bf_locked%></td>
<td><img src="THEMES/<%=Session("SiteTheme")%>/no_entry_forum.gif" alt="<%=bf_noentry%>" align="absmiddle">&nbsp;<%=bf_noentry%></td>
</tr>
</table>
<%End Sub

Sub Topics
If Application("Per")="OK" Then
Set msgs=Server.CreateObject("ADODB.RecordSet"):msgs.open "Select * FROM MESSAGES WHERE cid="&msgid&" AND topic=True ORDER BY mdate DESC",Connect,3,3
Set cat=Connect.Execute("Select * FROM CATEGORIES WHERE cid="&msgid&"")
Page=Request.QueryString("Page")
IF Page="" Then Page="1"%>
 » <%=cat("cname")%></font></b><br />
<%If Session("Enter")="Yes" Then%>
<table border="0" cellspacing="0" cellpadding="0" width="97%">
<tr><td align="right"><a href="?action=Msg&id=<%=msgid%>"><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/f_newpost.gif" border="0"></a></td></tr>
</table>
<% End If
IF msgs.Eof Then
Response.Write "<center><font face=verdana size=2>"&no_topics&"</font></center>"
ELSE %>
<table border="0" cellspacing="1" cellpadding="2" width="90%" bgcolor="<%=td_border_color%>" align="center">
<tr bgcolor="<%=td_border_color%>" class="td_menu" style="font-weight:bold">
<td width="2%" height="20" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif">&nbsp;</td>
<td width="40%" height="20" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=cat("cname")%></td>
<td width="15%" align="center" height="20" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=m_writer%></td>
<td width="5%" align="center" height="20" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=m_answers%></td>
<td width="5%" align="center" height="20" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=m_read%></td>
<td width="43%" align="center" height="20" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=m_lastpost%></td>
</tr>
<%msgs.pagesize=forum_topics
msgs.absolutepage=Page
sayfa=msgs.pagecount
For I=1 To msgs.pagesize
IF msgs.eof Then Exit For
number=number + 1
select_color=right(number,1) 
Select Case select_color
case "1","3","5","7","9"
color=""&background&""
Case Else
color=""&top_menu_color&""
End Select
Set acount=Connect.Execute("Select Count(*) AS count FROM MESSAGES WHERE mhit="&msgs("mid")&" AND topic=False")
Set last=Connect.Execute("Select * FROM MESSAGES WHERE mid="&msgs("mid")&" OR mhit="&msgs("mid")&" ORDER BY mdate DESC")
If last.Eof Then
last_post="###"
last_pmember="###"
ELSE
last_post=last("last_post")
Set last3=Connection.Execute("Select * FROM MEMBERS WHERE kul_adi='"&last("mwriter")&"'")
IF last3.Eof OR last3.Bof Then last_pmember="###":last_pmemid="" ELSE last_pmember=last3("kul_adi"):last_pmemid=last3("uye_id")
End If
Set wMem=Connection.Execute("Select * FROM MEMBERS WHERE kul_adi='"&msgs("mwriter")&"'")
IF wMem.Eof OR wMem.Bof Then memberid="" ELSE memberid=wMem("uye_id")%>
<tr class="td_menu">
<td width="2%" bgcolor="<%=background%>">
<%If msgs("locked")=True then %>
<img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/locked_msg.gif">
<% ElseIf acount("count") >= 10 Then %>
<img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/hot_msg.gif">
<% Else %>
<img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/normal_msg.gif">
<% End If %>
</td>
<td width="40%" bgcolor="<%=color%>"><b><a href="?action=Topic&id=<%=msgs("mid")%>"><%=duz(msgs("mtitle"))%></a></b>
<%Set ismoderator=Connect.Execute("Select * FROM MODERATORS WHERE cid="&msgid&" and uye_id="&Session("Uye_ID")&""):if ismoderator.eof=false and ismoderator.bof=false then ismod=true else ismod=false
If Session("Level")="1" OR Session("Level")="5" or ismod=true Then%><br />(<a href="Mods.asp?action=EditMsg&id=<%=msgs("mid")%>"><%=f_edit%></a> - <a href="Mods.asp?action=SafeEditMsg&id=<%=msgs("mid")%>"><%=f_edit_safe%></a>
<%If Session("Level")="1" Then%> - <a href="Mods.asp?action=DeleteMsg&id=<%=msgs("mid")%>" OnClick="return confirm('<%=onay_for_silme%>')"><%=f_del%></a><%end if%>
<% If msgs("Locked")=False Then %> - <a href="Mods.asp?action=LockMsg&id=<%=msgs("mid")%>"><%=f_lock%></a><% Else %> - <a href="Mods.asp?action=UnLockMsg&id=<%=msgs("mid")%>"><%=f_unlock%></a><% End If %>
)<% End IF %></td>
<td width="15%" align="center" bgcolor="<%=color%>"><a href="members.asp?action=member_details&uid=<%=memberid%>"><%=duz(msgs("mwriter"))%></a></td>
<td width="5%" align="center" bgcolor="<%=color%>"><%=msgs("cvp_sayisi")%></td>
<td width="5%" align="center" bgcolor="<%=color%>"><%=msgs("mhit")%></td>
<td width="43%" align="center" bgcolor="<%=color%>"><%=last_post%><br /><a href="members.asp?action=member_details&uid=<%=last_pmemid%>"><%=last_pmember%></a></td>
</tr>
<%msgs.MoveNext
Next%>
</table>
<br />
<div align="right" class="td_menu">
<%IF Page <> 1 Then Response.Write "<a href=""?action=Topics&id="&msgid&"&Page="&(cint(Page)-1)&""">« "&previous_page&"</a>&nbsp;-&nbsp;"
for y=1 to sayfa
if cint(Page)=cint(y) then response.write "<font class=""td_menu"">["&y&"]</font>" else response.write "<a href=""?action=Topics&id="&msgid&"&Page="&y&""">["&y&"]</a>"
next
IF NOT msgs.Eof Then Response.Write "&nbsp;-&nbsp;<a href=""?action=Topics&id="&msgid&"&Page="&(cint(Page)+1)&""">"&next_page&" »</a>"%>
</div><br />
<table border="0" cellspacing="0" cellpadding="2" align="center" width="90%" class="td_menu">
<tr>
<td width="20%"><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/normal_msg.gif" alt="<%=bm_normal%>">&nbsp;<%=bm_normal%></td>
<td width="20%"><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/hot_msg.gif" alt="<%=bm_hot%>">&nbsp;<%=bm_hot%></td>
<td width="20%"><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/locked_msg.gif" alt="<%=bm_locked%>">&nbsp;<%=bm_locked%></td>
</tr>
</table>
<%End If
End If
End Sub

Sub Topic
Set rs=Server.CreateObject("ADODB.RecordSet"):rSQL="Select * FROM MESSAGES WHERE mid="&msgid&" AND topic=True":rs.open rSQL,Connect,3,3
Set checkCat=Connect.Execute("Select * FROM CATEGORIES WHERE cid="&rs("cid")&"")
If checkCat("locked")=True Then
Response.Write "<center><font face=arial size=3 color="&text&"><b>"&f_locked&"</b></font></center>"
ELSE
Page=Request.QueryString("Page"):Page=QS_CLEAR(Page,"false")
IF Page="" Then Page="1"
rs("mhit")=rs("mhit") + 1
rs.Update
Set frep=Server.CreateObject("ADODB.RecordSet")
frep.open "Select * FROM MESSAGES WHERE mhit="&msgid&" AND topic=False ORDER BY mid",Connect,3,3
Set cat=Connect.Execute("Select * FROM CATEGORIES WHERE cid="&rs("cid")&"")
Set mem=Connection.Execute("Select * FROM MEMBERS WHERE kul_adi='"&rs("mwriter")&"'")
IF mem.Eof Then
f_mem_name="###"
f_mem_msg=""
f_mem_age=""
f_mem_icq=detail_empty
f_mem_msn=detail_empty
f_mem_signature=""
f_mem_id=""
f_mem_url=""
f_mem_level="0"
f_mem_avatar="IMAGES/avatars/blank.gif"
f_mem_online="?"
ELSE
f_mem_ozellik=mem("ozellik")
f_mem_name=mem("kul_adi")
f_mem_msg=mem("msg_sayisi")
strAge=Int(Right(mem("yas"),4))
strNow=Int(Right(Date(),4))
f_mem_age=strNow - strAge
f_mem_icq=mem("icq")
f_mem_msn =mem("msn")
if f_mem_icq="0" then f_mem_icq="Belirtilmemiþ" 
if f_mem_msn="" then f_mem_msn="Belirtilmemiþ" 
f_mem_signature=mem("imza")
f_mem_id=mem("uye_id")
f_mem_url=mem("url")
f_mem_level=mem("seviye")
f_mem_avatar=mem("u_avatar")
IF mem("u_online")=True Then f_mem_online=member_online ELSE f_mem_online=member_offline
End IF%>
<!--#include file="INC/forum_ratings.asp" -->
 » <a href="?action=Topics&id=<%=rs("cid")%>"><%=cat("cname")%></a> » <%=SmiLey(duz(rs("mtitle")))%></font></b><br /><br />
<table border="0" cellspacing="1" cellpadding="2" align="center" width="90%" bgcolor="<%=td_border_color%>" class="td_menu">
<tr class="td_menu" bgcolor="<%=menu_back%>">
<td width="25%" align="center" valign="top">
<b><%=f_mem_name%></b><font face="Verdana"><b><font color="#FF0000"><br />
- <%=fm_level%> -</font><br /></b></font><font color="#FF0000">
<font style="font-size: 10pt; font-weight: 700"><font face="Marlett">h</font><font face="Verdana"><%=f_mem_ozellik%></font></font><span style="font-weight: 700"><font face="Marlett" size="2">h</font></span><font face="Verdana"><b><br /><%=fm_include_img%><br /></b></font></font><font face="Verdana"><br /></font><p>
<font face="Verdana"><b><br /><img src="<%=f_mem_avatar%>" border="0"><br /><br />
</b></font></p>
<table border="0" cellspacing="0" cellpadding="0" width="100%" class="td_menu" style="border-collapse: collapse" bordercolor="#111111">
<tr>
<td width="100%" align="right" style="font-family: Verdana; font-size: 10px">
<p align="center"><font style="font-size: 11px"><b>Mesaj : <%=f_mem_msg%><br />
Durum : <%=f_mem_online%></font><br />
MSN : <%=f_mem_msn%></font><br />
ICQ : <%=f_mem_icq%></font></b>
</td>
</tr>
</table>
</td>
<td bgcolor="<%=menu_back%>" width="75%" valign="top">
<%IF rs("html_allow")=True Then Response.Write SmiLey(rs("mmsg"))  ELSE Response.Write SmiLey(duz(rs("mmsg")))%>
<br />-------------------------------------<br /><%=f_mem_signature%>
</td>
</tr>
<tr>
<td bgcolor="<%=menu_back%>" align="center"><%=rs("mdate")%></td>
<td bgcolor="<%=menu_back%>" align="left">
<a href="members.asp?action=member_details&uid=<%=f_mem_id%>"><img src="THEMES/<%=Session("SiteTheme")%>/mem_info.gif" border="0"></a> <a href="msgbox.asp?action=new_msg&who=<%=f_mem_name%>"><img src="THEMES/<%=Session("SiteTheme")%>/mem_msg.gif" border="0"></a>
<% IF f_mem_url<>"" Then response.Write("<a href='"&f_mem_url&"' TARGET='_new'><img src='THEMES/"&Session("SiteTheme")&"/mem_web.gif' border='0'></a>")%>
<a href="members.asp?action=add_friendlist&id=<%=f_mem_id%>"><img src="THEMES/<%=Session("SiteTheme")%>/mem_addfl.gif" border="0"></a>
<%IF f_mem_name=Session("Member") then response.Write("<a href='?action=EditPost&id="&rs("mid")&"'><img src='THEMES/"&Session("SiteTheme")&"/mem_edit.gif' border='0'></a>")%>
<a href="?action=Rep&Type=Quote&id=<%=msgid%>&qid=<%=rs("mid")%>"><img src="Themes/<%=Session("SiteTheme")%>/forum_quote.gif" border="0"></a>
</td>
</tr>
</table>
<hr size="1" color="<%=td_border_color%>">
<%mem.Close:Set mem=Nothing
IF frep.Eof OR frep.Bof Then
Response.Write "<div align=center class=td_menu><b>"&f_error3&"</b></div>"
ELSE
frep.pagesize=forum_replies
frep.absolutepage=Page
sayfa=frep.pagecount
For I=1 To frep.pagesize
IF frep.eof Then Exit For
Set mem=Connection.Execute("Select * FROM MEMBERS WHERE kul_adi='"&frep("mwriter")&"'")
IF mem.EoF Then
f_mem_name="###"
f_mem_msg=""
f_mem_age=""
f_mem_signature=""
f_mem_icq=detail_empty
f_mem_msn=detail_empty 
f_mem_id=""
f_mem_url=""
f_mem_level="0"
f_mem_avatar="IMAGES/avatars/blank.gif"
f_mem_online="?"
ELSE
f_mem_ozellik=mem("ozellik")
f_mem_name=mem("kul_adi")
f_mem_msg=mem("msg_sayisi")
strAge=Int(Right(mem("yas"),4))
strNow=Int(Right(Date(),4))
f_mem_age=strNow - strAge
f_mem_icq=mem("icq")
f_mem_msn =mem("msn")
if f_mem_icq="0" then f_mem_icq="Belirtilmemiþ" 
if f_mem_msn="" then f_mem_msn="Belirtilmemiþ" 
f_mem_signature=mem("imza")
f_mem_id=mem("uye_id")
f_mem_url=mem("url")
f_mem_level=mem("seviye")
f_mem_avatar=mem("u_avatar")
IF mem("u_online")=True Then f_mem_online=member_online ELSE f_mem_online=member_offline
End IF%>
<!--#include file="INC/forum_ratings.asp" -->
<table border="0" cellspacing="1" cellpadding="2" align="center" width="90%" bgcolor="<%=td_border_color%>" class="td_menu">
<tr class="td_menu" bgcolor="<%=menu_back%>">
<td width="25%" align="center" valign="top">
<b><%=f_mem_name%></b><font face="Verdana"><b><br />
<font color="#FF0000">- <%=fm_level%> -</font><br />
</b></font>
<font color="#FF0000"><span style="font-weight: 700">
<font face="Marlett" size="2">h</font></span><font face="Verdana" style="font-size: 10pt; font-weight: 700"><%=f_mem_ozellik%></font><font style="font-size: 10pt; font-weight: 700"><font face="Marlett">h</font></font><font face="Verdana"><b><br /><%=fm_include_img%><br />
</b></font>
</font><font face="Verdana"><b><br /><br /><img src="<%=f_mem_avatar%>" border="0"><br /><br />
</b></font>
<table border="0" cellspacing="0" cellpadding="0" width="100%" class="td_menu" style="border-collapse: collapse" bordercolor="#111111">
<tr>
<td width="100%" align="right" style="font-family: Verdana; font-size: 10px">
<p align="center"><font style="font-size: 11px"><b>Mesaj : <%=f_mem_msg%><br />
Durum : <%=f_mem_online%></font><br />
MSN : <%=f_mem_msn%></font><br />
ICQ : <%=f_mem_icq%></font></b>
</td>
</tr>
<% IF Session("Level")="1" OR Session("Level")="5" Then %>
<tr><td colspan="2" align="center"><font face="Verdana"><b>[ <a href="Mods.asp?action=Edit-Rep&id=<%=frep("mid")%>">Düzenle</a> | <a href="Mods.asp?action=Del-Rep&id=<%=frep("mid")%>">Sil</a> ] </b></font></td></tr>
<% End IF %>
</table>
</td>
<td bgcolor="<%=menu_back%>" width="75%" valign="top" class="td_menu">
<%IF frep("html_allow")=True Then
strReplyMsg=SmiLey(frep("mmsg"))
ELSE
strReplyMsg=frep("mmsg")
strReplyMsg=Replace(strReplyMsg,"<P>","[P]")
strReplyMsg=Replace(strReplyMsg,"</P>","[/P]")
strReplyMsg=duz(strReplyMsg)
strReplyMsg=Replace(strReplyMsg,"[P]","<P>")
strReplyMsg=Replace(strReplyMsg,"[/P]","</P>")
strReplyMsg=Smiley(strReplyMsg)
End IF
IF InStr(strReplyMsg,"[Quote:") Then strReplyMsg=Quote(strReplyMsg) ELSE strReplyMsg=strReplyMsg
Response.Write strReplyMsg%>
<br />-------------------------------------<br /><%=f_mem_signature%>
</td>
</tr>
<tr>
<td bgcolor="<%=menu_back%>" align="center"><%=frep("mdate")%></td>
<td bgcolor="<%=menu_back%>" align="left">
<a href="members.asp?action=member_details&uid=<%=f_mem_id%>"><img src="THEMES/<%=Session("SiteTheme")%>/mem_info.gif" border="0"></a> <a href="msgbox.asp?action=new_msg&who=<%=f_mem_name%>"><img src="THEMES/<%=Session("SiteTheme")%>/mem_msg.gif" border="0"></a> <a href="<%=f_mem_url%>" TARGET="_new"><img src="THEMES/<%=Session("SiteTheme")%>/mem_web.gif" border="0"></a> <a href="members.asp?action=add_friendlist&id=<%=f_mem_id%>"><img src="THEMES/<%=Session("SiteTheme")%>/mem_addfl.gif" border="0"></a>
<%IF f_mem_name=Session("Member") Then
Session("EditPost")="Yes"%>
 <a href="?action=EditPost&id=<%=frep("mid")%>"><img src="THEMES/<%=Session("SiteTheme")%>/mem_edit.gif" border="0"></a>
<% End IF %>
<a href="?action=Rep&Type=Quote&id=<%=msgid%>&qid=<%=rs("mid")%>"><img src="Themes/<%=Session("SiteTheme")%>/forum_quote.gif" border="0"></a>
</td>
</tr>
</table><br />
<%frep.MoveNext
Next
End IF%>
<br /><div align="right" class="td_menu">
<%IF Page <> 1 Then Response.Write "<a href=""?action=Topic&id="&msgid&"&Page="&(cint(Page)-1)&""">« "&previous_page&"</a>&nbsp;-&nbsp;"
for y=1 to sayfa
if cint(Page)=cint(y) then response.write "<font class=""td_menu"">["&y&"]</font>" else response.write "<a href=""?action=Topic&id="&msgid&"&Page="&y&""">["&y&"]</a>"
next
IF NOT frep.Eof Then Response.Write "&nbsp;-&nbsp;<a href=""?action=Topic&id="&msgid&"&Page="&(cint(Page)+1)&""">"&next_page&" »</a>"%>
</div><br />
<center>
<% If rs("locked")=True Then
Response.Write "<center><font face=tahoma style=font-size:11px;><b>"&m_locked&"</b></font></center>"
ELSE %>
<a href="?action=Rep&id=<%=msgid%>" method="post">
<img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/f_newreply.gif" border="0"></a>
</center><br />
<hr size="1" width="50%" color="<%=td_border_color%>">
<table border="0" cellspacing="0" cellpadding="10" width="50%" class="td_menu" align="center" style="<%=TableBorder%>"><form action="?action=RegQuickReply&id=<%=msgid%>" method="post">
<tr align="center"><td><b><%=f_quickreply%></b></td></tr>
<tr><td><textarea name="f_rep" rows="10" style="width:100%" class="form1" cols="20"></textarea></td></tr>
<tr align="right"><td><input type="submit" class="submit" value="<%=uye_kayit%>" style="width:100%" onClick="this.form.submit();this.disabled=true; return true;"></td></tr>
</form></table>
<hr size="1" width="50%" color="<%=td_border_color%>">
<% End If
END IF
End Sub

Sub msg
IF Session("Enter")="Yes" Then
Set cat=Connect.Execute("Select * FROM CATEGORIES WHERE cid="&msgid&"")
Response.Write "» " & cat("cname")
IF cat("locked")=True Then
Response.Write "<hr size=1 color="&td_border_color&"><div align=center>" & f_locked & "</div><hr size=1 color="&td_border_color&">"
ELSE
f1_sec_code=CPASS(4)
Session("Forum1SecurityCode")=f1_sec_code%>
<script language="JavaScript1.2" defer>editor_generate('mmsg');</script>
<form action="?action=MsgReg&id=<%=msgid%>" method="post">
<table border="0" cellspacing="3" cellpadding="3" width="80%" align="center" class="td_menu" style="font-weight:bold" bgcolor="<%=background%>">
<tr>
<td width="20%" align="right" bgcolor="<%=menu_back%>"><%=nm_title%> : </td>
<td bgcolor="<%=menu_back%>"><input type="text" name="mtitle" class="form1" style="width:100%" size="20"></td>
</tr>
<tr>
<td align="right" bgcolor="<%=menu_back%>" valign="top"><%=nm_msg%> : </td>
<td bgcolor=buttonface><textarea name="mmsg" rows="15" class="form1" style="width:100%" cols="20"></textarea></td>
</tr>
<tr>
<td align="right" bgcolor="<%=menu_back%>"><%=security_code%> : </td>
<td bgcolor="<%=menu_back%>"><%=Session("Forum1SecurityCode")%></td>
</tr>
<tr>
<td align="right" bgcolor="<%=menu_back%>"><%=security_code_type%> : </td>
<td bgcolor="<%=menu_back%>"><input type="text" name="f1_sec_code" size="7" class="form1"></td>
</tr>
<tr>
<td colspan="2" bgcolor="<%=background%>"><input type="submit" class="submit" value="<%=uye_kayit%>" style="width:100%" onClick="this.form.submit();this.disabled=true; return true;"></td>
</tr>
</table>
</form>
<%End IF
Else
Response.Write WriteMsg(no_entry,90)
End IF
End Sub 

Sub reg
IF DateDiff("N" ,Session("FloodTimer"),Now()) > strFloodTime Then
title=Request.Form("mtitle")
mmsg=Request.Form("mmsg")
scode=duz(Request.Form("f1_sec_code"))
If title="" Then
Response.Write WriteMsg(f_error1,90)
ElseIf mmsg="" Then
Response.Write WriteMsg(f_error2,90)
ElseIF scode="" Then
Response.Write WriteMsg(sc_err1,90)
ElseIF ucase(scode)<>ucase(Session("Forum1SecurityCode")) Then
Response.Write WriteMsg(sc_err2,90)
ELSE%>
<script Language="JavaScript1.2">if (win_ie_ver < 5.5) {<% html_allow=False %>} else {<% html_allow=True %>}</script>
<%FOR I=1 TO UBound(saryHTMLtags)
mmsg=Replace(mmsg,"<"&saryHTMLtags(""&I&"")&">","{"&saryHTMLtags(""&I&"")&"}",1,-1,1)
Next
Set fent=Server.CreateObject("ADODB.RecordSet"):fSQL="Select * FROM MESSAGES":fent.open fSQL,Connect,3,3
fent.AddNew
fent("mtitle")=title
fent("mmsg")=mmsg
fent("mwriter")=Session("Member")
fent("mdate")=Now()
fent("mhit")=0
fent("cid")=msgid
fent("topic")=True
fent("html_allow")=html_allow
fent("last_post")=now()
fent.Update
Set Rmsg1=Connection.Execute("Select * FROM MEMBERS WHERE uye_id="&Session("Uye_ID")&"")
Set Rmsg=Connection.Execute("UPDATE MEMBERS SET msg_sayisi="&Rmsg1("msg_sayisi")+1&" WHERE uye_id="&Session("Uye_ID")&"")
Session("FloodTimerStart")="Yes"
Session("FloodTimer")=Rmsg1("son_tarih")
Set Rmsg1=Nothing:Set Rmsg=Nothing
Response.Redirect "Forum.asp?action=Topics&id="&msgid&""
END IF
ELSE
Response.Write WriteMsg(f_error5,90)
End IF
End Sub

Sub rep
Set rmsg=Connect.Execute("Select * FROM MESSAGES WHERE mid="&msgid&"")
If rmsg("locked")=True Then
Response.Write "<hr size=1 color="&td_border_color&"><div align=center><b>"&m_locked&"</b></div><hr size=1 color="&td_border_color&">"
ELSE
IF Session("Enter")="Yes" Then
f2_sec_code=CPASS(4):Session("Forum2SecurityCode")=f2_sec_code
IF Request.QueryString("qid") <> "" Then strRequestID=Request.QueryString("qid") ELSE strRequestID=msgid
Set iqmRs=Server.CreateObject("ADODB.RecordSet"):iqmRs.Open "Select * FROM MESSAGES WHERE mid=" & strRequestID,Connect,3,3
IF Request.QueryString("Type")="Quote" Then strReplyMsg="[Quote:" & iqmRs("mwriter") & "]" & iqmRs("mmsg") & "[/Quote]" ELSE strReplyMsg=""
iqmRs.Close:Set iqmRs=Nothing%>
<form action="?action=RepReg&id=<%=msgid%>" method="post">
<% IF Request.QueryString("Type")="Quote" Then response.write("<input type='hidden' name='quote' value='Yes'>")%>
<table border="0" cellspacing="5" cellpadding="3" width="80%" align="center" class="td_menu" style="font-weight:bold" bgcolor="<%=background%>">
<tr>
<td width="20%" align="right" bgcolor="<%=menu_back%>" valign="top"><%=nm_msg%>:</td>
<td bgcolor=buttonface><script src="htmlarea/editor.js" language="Javascript1.2"></script><script language="JavaScript1.2" defer>editor_generate('amsg');</script>
<textarea name="amsg" rows="15" class="form1" style="width:100%" cols="20"><%=strReplyMsg%></textarea></td>
</tr>
<tr><td align="right" bgcolor="<%=menu_back%>"><%=security_code%>:</td><td bgcolor="<%=menu_back%>"><%=Session("Forum2SecurityCode")%></td></tr>
<tr><td align="right" bgcolor="<%=menu_back%>"><%=security_code_type%>:</td><td bgcolor="<%=menu_back%>"><input type="text" name="f2_sec_code" size="7" class="form1"></td></tr>
<tr><td colspan="2" bgcolor="<%=background%>"><input type="submit" class="submit" value="<%=uye_kayit%>" style="width:100%" onClick="this.form.submit();this.disabled=true; return true;"></td></tr>
</table></form>
<%Else
Response.Write WriteMsg(no_entry,90)
End If
End If
End Sub

Sub repreg
IF DateDiff("N" ,Session("FloodTimer"),Now()) > strFloodTime Then
amsg=Request.Form("amsg")
scode=duz(Request.Form("f2_sec_code"))
IF amsg="" Then
Response.Write WriteMsg(f_error2,90)
ElseIF scode="" Then
Response.Write WriteMsg(sc_err1,90)
ElseIF ucase(scode)<>ucase(Session("Forum2SecurityCode")) Then
Response.Write WriteMsg(sc_err2,90)
ELSE%>
<script Language="JavaScript1.2">if (win_ie_ver < 5.5) {<% html_allow=False %>} else {<% html_allow=True %>}</script>
<%FOR I=1 TO UBound(saryHTMLtags)
amsg=Replace(amsg,"<"&saryHTMLtags(""&I&"")&">","&lt;"&saryHTMLtags(""&I&"")&"&gt;",1,-1,1)
Next
Set ms=Connect.Execute("Select * FROM MESSAGES WHERE mid="&msgid&"")
Set rent=Server.CreateObject("ADODB.RecordSet"):rSQL="Select * FROM MESSAGES":rent.open rSQL,Connect,3,3
rent.AddNew
rent("mmsg")=amsg
rent("mwriter")=Session("Member")
rent("mdate")=Now()
rent("last_post")=Now()
rent("mhit")=msgid
rent("cid")=ms("cid")
rent("html_allow")=html_allow
rent("topic")=False
rent("locked")=False
rent.Update
Set Rmsg1=Connection.Execute("Select * FROM MEMBERS WHERE uye_id="&Session("Uye_ID")&"")
Set Rmsg=Connection.Execute("UPDATE MEMBERS SET msg_sayisi="&Rmsg1("msg_sayisi")+1&" WHERE uye_id="&Session("Uye_ID")&"")
Set Rmsg=Connect.Execute("UPDATE MESSAGES SET last_post=now(),cvp_sayisi='"&ms("cvp_sayisi")+1&"' WHERE mid="&msgid&"")
Session("FloodTimerStart")="Yes"
Session("FloodTimer")=Rmsg1("son_tarih")
Set Rmsg1=Nothing:Set Rmsg=Nothing
Response.Redirect "Forum.asp?action=Topic&id="&msgid&""
END IF
ELSE
Response.Write WriteMsg(f_error6,90)
End IF
End Sub

Sub Control
Set cat=Connect.Execute("Select * FROM CATEGORIES WHERE cid="&msgid&"")
If Session("Level") <> "1" AND Session("Level") <> "2" AND Session("Level") <> "3" AND Session("Level") <> "4" AND Session("Level") <> "5" AND Session("Level") <> "6" AND Session("Level") <> "7" AND Session("Level") <> "8" AND Session("Level") <> "9"  AND Session("Level") <> "10" AND Session("Level") <> "11" Then
If cat("noentry")=True Then
Response.Write "<center><font face=arial size=3 color="&text&"><b>"&f_notpermission&"</b></font></center>"
ElseIf cat("locked")=True Then
Response.Write "<center><font face=arial size=3 color="&text&"><b>"&f_locked&"</b></font></center>"
ELSE
Application("Per")="OK"
Response.Redirect "?action=Topics&id="&msgid&""
END IF
ELSE
Application("Per")="OK"
Response.Redirect "?action=Topics&id="&msgid&""
END IF
End Sub%>

<%End IF %>

<br /><table border="0" cellspacing="1" cellpadding="1" width="90%" align="center" class="td_menu" bgcolor="<%=td_border_color%>">
<tr><td background="THEMES/<%=Session("SiteTheme")%>/td_bg.gif" height="20"><span class="block_title" style="font-weight:bold"><%=f_whoonline%>&nbsp;&nbsp;&nbsp;[ <span style="color:<%=online_admin%>">GeneL Sorumlu - Site Yöneticisi</span> ][ <span style="color:<%=online_editor%>">Site Editörü</span> ]</span></td></tr>
<tr bgcolor="<%=menu_color%>"><td><!--#include file="inc/online.asp" --><%IF om.EoF Then
Response.Write no_online
ELSE
Do Until om.Eof
IF om("seviye")="1" Then
om_title=""&level1&""
ElseIf om("seviye")="2" Then
om_title=""&level2&""
ElseIf om("seviye")="3" Then
om_title=""&level3&""
ElseIf om("seviye")="4" Then
om_title=""&level4&""
ElseIf om("seviye")="5" Then
om_title=""&level5&""
ElseIf om("seviye")="6" Then
om_title=""&level6&""
ElseIf om("seviye")="8" Then
om_title=""&level8&""
ElseIf om("seviye")="0" Then
om_title=""&normal&""
End If
IF om("seviye")="1"  Then
om_t_c=online_admin
om_t_style="bold"
ElseIF om("seviye")="2" OR om("seviye")="3" OR om("seviye")="4"  OR om("seviye")="5"  OR om("seviye")="6"  OR om("seviye")="8"   Then
om_t_c=online_editor
om_t_style="bold"
ELSE
om_t_c=text
om_t_style="regular"
End IF
Response.Write "<a href=""members.asp?action=member_details&uid="&om("uye_id")&""" title="""&om_title&"""><font color="&om_t_c&" style=""font-weight:"&om_t_style&""">"&om("kul_adi")&"</font></a> , "
om.MoveNext
Loop
End IF%>
</td></tr>
</table><br />
<table border="0" cellspacing="1" cellpadding="1" width="90%" align="center" class="td_menu" bgcolor="<%=td_border_color%>">
<tr><td background="THEMES/<%=Session("SiteTheme")%>/td_bg.gif" height="20"><span class="block_title" style="font-weight:bold">Üyelik</span></td></tr>
<tr bgcolor="<%=background%>"><td>
<%IF Session("Enter")="Yes" Then
Set unR=Connection.Execute("Select Count(*) as count FROM MESSAGES WHERE TYPE=0 AND READ=False AND SEE=True AND TO='"&Session("Member")&"'")
Set allmsgs=Connection.Execute("Select Count(*) as count FROM MESSAGES WHERE TYPE=0 AND SEE=True AND TO='"&Session("Member")&"'")
If unR("count") > sett("msg_limit") Then unRd=sett("msg_limit") Else unRd=unR("count")
If allmsgs("count") > sett("msg_limit") Then allM=sett("msg_limit") Else allM=allmsgs("count")%>
<table width="100%" border="0" cellpadding="0" cellspacing="0"><tr><td width="200" valign="top"><img src="IMAGES/stats/arrow.gif"> <img border="0" src="Themes/<%=Session("SiteTheme")%>/ara.gif" align="absmiddle"> <a href="?action=Search">Forum'da Arama Yap </a></td>
<td><div class=td_menu><b>| <%=Session("Name")%></b> » <a href="<%=sett("site_adres")%>/Your_Account.asp"><%=ya_topic%></a> <b>|</b> <a href="<%=sett("site_adres")%>/search.asp?action=Member"><%=uyemenu_uyeara%></a> <b>|</b> <a href="<%=sett("site_adres")%>/msgbox.asp?action=main&id=<%=Session("Uye_ID")%>"><%=uyemenu_msgbox%></a> (<%=unRd%>/<%=allM%>)
<%If Session("Level")="1" OR Session("Level")="5"  OR Session("Level")="2" OR Session("Level")="3" OR Session("Level")="4" OR Session("Level")="8"  Then %><b>|</b> <a href=administrator/default.asp TARGET=_blank><span style="color:<%=online_admin%>"><%=control_panel%></span></a><% End IF %> 
<b>|</b> <a href="<%=sett("site_adres")%>/Your_Account.asp?op=Logout"> <%=ya_logout%> [ <%=Session("Member")%> ]</a><b> |</b></div>
<% IF Session("Level")="1" or Session("Level")="5" Then %>
<br /></td></tr><tr bgcolor="<%=background%>"><td width="200">&nbsp;</td><td><div class=td_menu>| Forum Yönetimi : 
<a href="#" onClick="window.name='forum_settings'; window.open('administrator/Mods.asp?action=ShowSettings','FS', 'toolbar=no,location=no,directories=no,status=yes,menubar=no,scrollbars=yes, resizable=yes,copyhistory=no,width=400,height=200'); return false;"><b>Forum Ayarlarý</b></a> | 
<a href="#" onClick="window.name='new_cat'; window.open('administrator/Mods.asp?action=NewCat','NewCat', 'toolbar=no,location=no,directories=no,status=yes,menubar=no,scrollbars=yes, resizable=yes,copyhistory=no,width=400,height=200'); return false;"><b>Kategori Ekle</b></a> | 
<a href="#" onClick="window.name='new_forum'; window.open('administrator/Mods.asp?action=NewForum','NewForum', 'toolbar=no,location=no,directories=no,status=yes,menubar=no,scrollbars=yes, resizable=yes,copyhistory=no,width=400,height=200'); return false;"><b>Forum Ekle</b></a> | 
<a href="#" onClick="window.name='new_forum'; window.open('administrator/Mods.asp?action=NewMod','NewModerator', 'toolbar=no,location=no,directories=no,status=yes,menubar=no,scrollbars=yes, resizable=yes,copyhistory=no,width=400,height=200'); return false;"><b>Moderatör Ata</b></a>
&nbsp;|</div>
<% End IF %>
</td></tr></table>
<%ELSE
the_security_code=cpass(8):session("login_gguvenlik")=the_security_code%>
<form action="<%=sett("site_adres")%>/Enter.asp" method="post">
<center><%=uye_kuladi%>:<input type="text" name="kuladi" size="10" class="form1">&nbsp; &nbsp;<%=uye_password%>:<input type="password" name="password" size="10" class="form1">&nbsp; &nbsp;Kod:<input type="text" name="guvenlik" value="<%=the_security_code%>" size="8" class="form1" value=""><input type="hidden" name="gguvenlik" value="<%=the_security_code%>"><b><font face="verdana" size="1"><%=the_security_code%></font></b>&nbsp; &nbsp;<input type="submit" value="<%=uye_submit%>" class="submit" onClick="this.form.submit();this.disabled=true; return true;"></center>
</form>
<% End IF %>
<br /><br /></td></tr>
</table>
<% call ALT %>