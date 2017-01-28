<!--#include file="view.asp" -->
<%Action=QS_CLEAR(Request.QueryString("Action"),"")
if Action<>"Print" then
call UST:call ORTA%>
<table border="0" cellspacing="0" cellpadding="2" width="100%" style="<%=TableBorder%>" class="td_menu">
<tr bgcolor="<%=menu_back%>"><td align="center" style="font-size:18px"><%=sett("site_isim")%> : <%=top_menu2%></td></tr>
</table><br />
<table border="0" cellspacing="0" cellpadding="2" width="100%" style="<%=TableBorder%>" class="td_menu">
<tr bgcolor="<%=menu_back%>"><td align="center">
<a href="news.asp"><img src="images/logos/haber.gif" border="0"></a><br />
<b>[</b> <a href="?action=HitNews"><%=title_popular%></a> <b>|</b> <a href="?action=LikeNews"><%=title_like%></a> <b>|</b> <a href="?action=NewNews"><%=title_new%></a> <b>|</b> <a href="?action=SendNews"><%=title_send_news%></a> <b>|</b> <a href="?action=Stats"><%=title_statistics%></a> <b>]</b>
</td></tr>
</table><br />
<%End IF
if Action="" Then
Call TheAllNews
Elseif Action="News" Then
Call JustNews
Elseif Action="Read" Then
Call ReadTheNews
Elseif Action="Print" Then
Call PrintTheNews
Elseif Action="Vote" Then
Call VoteTheNews
Elseif Action="SendFriend" Then
Call SendFriendTheNews
Elseif Action="Read_sbt" Then
Call Read_sbtNews
Elseif Action="HitNews" or Action="NewNews" or Action="LikeNews" Then
Call HitNewLikeNews
Elseif Action="Stats" Then
Call Stats
Elseif Action="SendNews" Then
Call SendTheNews
Elseif Action="RegNews" Then
Call RegSendedTheNews
End if

Sub TheAllNews
Set cats=Server.CreateObject("ADODB.RecordSet"):cats.open "Select * FROM NEWS_CATS ORDER BY cat_name ASC",Connection,3,3%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
<tr><td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu">
<%Do Until cats.Eof
Set rs=Server.CreateObject("ADODB.RecordSet"):rsSQL="Select * FROM NEWS WHERE cat="&cats("cat_id")&" AND onay=True ORDER BY hid DESC":rs.open rsSQL,Connection,3,3
if not rs.Eof then
Set reads=Connection.Execute("Select SUM(okunma) AS r_count FROM NEWS WHERE cat="&cats("cat_id")&"")
IF rs.RecordCount <= 0 Then n_reads="0" ELSe n_reads=reads("r_count")%>
<table border="0" cellspacing="1" cellpadding="2" align="center" width="100%" bgcolor="<%=td_border_color%>" class="td_menu">
<tr class="td_menu" style="font-weight:bold;"><td colspan="2" align="center" height="15" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><font class="block_title"><%=cats("cat_name")%></font></td></tr>
<tr bgcolor="<%=menu_bg%>">
<td width="40%" valign="top"><div align="center"><a href="?action=News&id=<%=cats("cat_id")%>"><img src="<%=cats("cat_image")%>" border="0" title="<%=cats("cat_name")%>"></a></div>
<%=cats("cat_info")%><br /><br /><b>· <%=hc_news%> : </b><%=rs.RecordCount%><br /><b>· <%=hc_reads%> : </b><%=n_reads%></td>
<td width="60%" valign="top">
<%IF rs.Eof Then
Response.Write hc_eof
ELSE
For I=1 TO 10
IF rs.Eof Then Exit For
Response.Write "» <a href=""?action=Read&hid="&rs("hid")&""">"&rs("baslik")&"</a><br />"
rs.MoveNext
Next
IF I >= 10 Then Response.Write "<div align=""right"" style=""font-weight:bold""><a href=""?action=News&id="&cats("cat_id")&""">" & hc_all & "</a></div>"
End IF%>
</td></tr>
</table><br />
<%End IF
cats.MoveNext
Loop%>
</td></tr>
</table>
<%End Sub

Sub JustNews
Set rs=Server.CreateObject("ADODB.RecordSet"):rs.open "Select * FROM NEWS WHERE cat="&QS_CLEAR(Request.QueryString("id"),"false")&"",Connection,3,3%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>"><tr> 
<td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=top_menu2%></font></td></tr>
<tr><td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu">
<table border="0" cellspacing="1" cellpadding="1" align="center" width="100%" bgcolor="<%=td_border_color%>">
<tr class="td_menu" style="font-weight:bold;">
<td width="70%" align="left" height="15" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=h_baslik%></td>
<td width="30%" align="center" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif" height="15"><%=the_tarih%></td>
</tr>
<%Do Until rs.Eof
Set nmem=Connection.Execute("Select * FROM MEMBERS WHERE kul_adi='"&rs("ekleyen")&"'")%>
<tr class="td_menu">
<td width="70%" bgcolor="<%=menu_bg%>"><a href="?Action=Read&hid=<%=rs("hid")%>"><%=rs("baslik")%></a> ( <a href="members.asp?action=member_details&uid=<%=nmem("uye_id")%>"><%=rs("ekleyen")%></a> )</td>
<td align="center" bgcolor="<%=menu_bg%>" width="30%"><%=rs("tarih")%>
</td>
</tr>
<% rs.MoveNext
Loop %>
</table>
<br /><br /><b><%=hc_news%> : </b><%=rs.RecordCount%>
</td></tr>
</table>
<%End Sub

Sub ReadTheNews
id=Request.QueryString("hid"):id=QS_CLEAR(id,"false")
Set rs=Server.CreateObject("ADODB.RecordSet"):rsSQL="Select * FROM NEWS WHERE hid="&id&" AND onay=True":rs.open rsSQL,Connection,3,3
rs("okunma")=rs("okunma") + 1
rs.Update
Set comments=Connection.Execute("Select Count(*) AS count FROM COMMENTS WHERE nid="&rs("hid")&" AND page='news'")
Set cats=Connection.Execute("Select * FROM NEWS_CATS WHERE cat_id="&rs("cat")&"")%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=rs("baslik")%></font></td></tr>
<tr><td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu">
<table border="0" cellspacing="5" cellpadding="0" align="center" width="100%" class="td_menu">
<tr><td valign="top"><img src="<%=cats("cat_image")%>" border="0" align="left" hspace="5" vspace="0">
<%IF Len(rs("resim")) > 0 Then Response.Write "<img src='"&rs("resim")&"' align='right'>"&rs("metin")&""%>
</td></tr>
</table>
<br /><br /><a href='javascript:history.go(-1)'>« Geri</a>
<hr size="1" color="<%=td_border_color%>"><center>
<% Set nwmem=Connection.Execute("Select * FROM MEMBERS WHERE kul_adi='"&rs("ekleyen")&"'") %>
<% If NOT nwmem.Eof then %><a href="members.asp?action=member_details&uid=<%=nwmem("uye_id")%>"><%=rs("ekleyen")%></a><%Else%><%=rs("ekleyen")%><% End If %> <b>|</b> <%=rs("tarih")%> <b>|</b> <%=h_readed%> : <%=rs("okunma")%><br />
<% If Session("Enter")="Yes" and Request.Cookies("MiniNUKE_NV")("op")<>"OK" AND Request.Cookies("MiniNUKE_NV")("hid")<>""&id&"" then%>
<form action="?action=Vote&id=<%=id%>" method="post"><font class="td_menu2"><%=h_nvmsg%> : </font>&nbsp;<select name="point" class="form1">
<option value="1"><%=vote1%></option><option value="2"><%=vote2%></option><option value="3"><%=vote3%></option><option value="4"><%=vote4%></option><option value="5"><%=vote5%></option>
</select>&nbsp;<input type="submit" value="<%=poll_button%>" class="submit">
</form><% End If %><font class="td_menu"><%=h_averagepoint%> : <b><%=left(rs("puan"),4)%></b><br><%=h_totalvote%> : <b><%=rs("katilimci")%></font>
<b>[</b> <a href="comments.asp?action=comments&page=news&id=<%=id%>"><%=news_comments%>(<%=comments("count")%>)</a> <b>|</b> 
<a href="?Action=Print&hid=<%=id%>" target="_blank"><img src="images/site/print.gif" border="0" align="absmiddle"> <%=h_print%></a> <b>|</b> 
<a href="?Action=SendFriend&ST=1&hid=<%=id%>"><img src="images/site/sendfriend.gif" border="0" align="absmiddle"> <%=h_sendfriend%></a> <b>]</b></center>
</td></tr>
</table>
<%End Sub

Sub PrintTheNews
id=Request.QueryString("hid"):id=QS_CLEAR(id,"false")
Set rs=Server.CreateObject("ADODB.RecordSet"):SQL="Select * FROM NEWS WHERE hid="&id&"":rs.open SQL,Connection,3,3%>
<br />
<table border="0" cellspacing="1" cellspacing="10" align="center" width="90%" bgcolor="<%=td_border_color%>">
<tr bgcolor="<%=background%>"><td>
<center><img src="Images/logo.gif"><br /><b><%=top_menu2%></b><br /><br />
<b><%=rs("baslik")%></b><br /><b><%=the_tarih%> : </b><%=rs("tarih")%><br /><br /></center><%=DelHTML(rs("metin"))%>
</td></tr>
</table>
<br /><center><%=h_from%> : <%=sett("site_isim")%> (<a href="<%=sett("site_adres")%>"><%=sett("site_adres")%></a>)<br /><%=h_address%> : <a href="<%=sett("site_adres")%>/news.asp?Action=Read&hid=<%=id%>"><%=sett("site_adres")%>/news.asp?Action=Read&hid=<%=id%></a></center>
<%rs.Close:Set rs=Nothing%>
<script>document.all.yukleme.style.visibility='hidden';</script>
<%End Sub

Sub VoteTheNews
hid=Request.QueryString("id"):hid=QS_CLEAR(hid,"false")
IF Request.Cookies("MiniNUKE_NV")("op")="OK" AND Request.Cookies("MiniNUKE_NV")("hid")=""&hid&"" Then
Response.Write WriteMsg(nv_errmsg,100)
ELSE
point=Request.Form("point")
Set eV=Server.CreateObject("ADODB.RecordSet"):eSQL="Select * FROM NEWS WHERE hid="&hid&"":eV.open eSQL,Connection,3,3
thenewpoint=((eV("puan")*eV("katilimci"))+point)/(eV("katilimci")+1)
eV("puan")=thenewpoint
eV("katilimci")=eV("katilimci") + 1
eV.Update
Response.Cookies("MiniNUKE_NV")("op")="OK":Response.Cookies("MiniNUKE_NV")("hid")=""&hid&""
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End IF
End Sub

Sub SendFriendTheNews%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=h_sendfriend%></font></td></tr>
<tr><td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu">
<% IF QS_CLEAR(Request.QueryString("ST"),"false")="1" Then %>
<form action="?Action=SendFriend&ST=2&hid=<%=QS_CLEAR(Request.QueryString("hid"),"false")%>" method="post"><table border="0" cellspacing="2" cellpadding="2" width="90%" align="center" class="td_menu">
<tr><td width="50%" align="right"><%=h_yn%> : </td><td width="50%"><input type="text" name="yn" size="30" class="form1" value="<%=Session("Name")%>"></td></tr>
<tr><td width="50%" align="right"><%=h_ym%> : </td><td width="50%"><input type="text" name="ym" size="30" class="form1" value="<%=Session("Email")%>"></td></tr>
<tr><td width="50%" align="right"><%=h_fn%> : </td><td width="50%"><input type="text" name="fn" size="30" class="form1"></td></tr>
<tr><td width="50%" align="right"><%=h_fm%> : </td><td width="50%"><input type="text" name="fm" size="30" class="form1"></td></tr>
<td width="50%"></td><td width="50%"><input type="submit" value="<%=msg_ok%>" class="submit"></td></tr>
</table></form>
<%ElseIF QS_CLEAR(Request.QueryString("ST"),"false")="2" Then
yn=duz(Request.Form("yn"))
ym=duz(Request.Form("ym"))
fn=duz(Request.Form("fn"))
fm=duz(Request.Form("fm"))
If yn="" OR ym="" OR fn="" OR fm="" Then
Response.Write advise_err
ELSE
Set nRs=Server.CreateObject("ADODB.RecordSet"):nSQL="Select * FROM NEWS WHERE hid="&QS_CLEAR(Request.QueryString("hid"),"false")&"":nRs.open nSQL,Connection,3,3:
Set objCDO=Server.CreateObject("CDONTS.NewMail")
objCDO.To=""&fm&""
objCDO.From=""&yn&""
objCDO.Subject=""&h_subject&""
objCDO.BodyFormat=0
objCDO.MailFormat=0
govde=govde &" <meta http-equiv=""Content-Language"" content=""tr""> "& chr(10)
govde=govde &" <meta http-equiv=""Content-Type"" content=""text/html; charset=windows-1254""> "& chr(10)
govde=govde &" <font class=""td_menu""> "& chr(10)
govde=govde &" <b>"&yn&"</b> "&h_msg1&""& chr(10)
govde=govde &" <br /><br /><br /> "& chr(10)
govde=govde &" "&nRs("baslik")&"<br />("&the_tarih&" : "&nRs("tarih")&") "& chr(10)
govde=govde &" <br /><br /> "& chr(10)
govde=govde &" "&h_msg2&" : <a href="""&sett("site_adres")&"/news.asp?Action=Read&hid="&nRs("hid")&""">"&sett("site_adres")&"/news.asp?Action=Read&hid="&nRs("hid")&"</a> "& chr(10)
govde=govde &" <br /><br /> "& chr(10)
govde=govde &" "&h_msg3&""&sett("site_adres")&" "& chr(10)
objCDO.Body=govde
objCDO.Send
Set objCDO=Nothing
Response.Write "<font class=""td_menu"">"&h_success&"</font>"
End IF
End IF%>
</td></tr>
</table>
<%End Sub

Sub Read_sbtNews
fid=Request.QueryString("fid"):fid=QS_CLEAR(fid,"false")
Set rs=Server.CreateObject("ADODB.RecordSet"):rs.open "Select * FROM FIXED WHERE fid="&fid&"",Connection,3,3
IF rs.EoF OR rs.BoF Then
Response.Write "<div align=center class=td_menu><b>Kayýtlý haber yok !</b></div>"
ELSE%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=rs("sf_baslik")%></font></td></tr>
<tr><td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu"><table border="0" cellspacing="5" cellpadding="0" align="center" width="100%" class="td_menu">
<tr><td valign="top"><%Response.Write DelHTML(rs("f_metin"))%></td></tr></table>
<br /><br /><a href='javascript:history.go(-1)'>« Geri</a>
</td></tr>
</table>
<%End IF
End Sub

Sub HitNewLikeNews
if Action="NewNews" then
Set n5=Server.CreateObject("ADODB.RecordSet"):n5.open "Select * FROM NEWS WHERE onay=True order by tarih desc",Connection,3,3:thepagetitle=title_new
elseif Action="HitNews" then
Set n5=Server.CreateObject("ADODB.RecordSet"):n5.open "Select * FROM NEWS WHERE onay=True order by okunma desc",Connection,3,3:thepagetitle=title_popular
elseif Action="LikeNews" then
Set n5=Server.CreateObject("ADODB.RecordSet"):n5.open "Select * FROM NEWS WHERE onay=True order by puan desc",Connection,3,3:thepagetitle=title_like
end if%>
<table border="0" cellspacing="1" cellpadding="2" width="100%" bgcolor="<%=td_border_color%>" align="center">
<tr class="td_menu" style="font-weight:bold;"><td colspan="3" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=thepagetitle%></font></td></tr>
<%For J=1 TO 5
IF n5.Eof Then Exit For
Set cat=Connection.Execute("Select * FROM NEWS_CATS WHERE cat_id="&n5("cat")&"")%>
<tr class="td_menu" bgcolor="<%if J=2 or J=4 then response.write forum_bg_2 else response.write forum_bg_1%>">
<td valign="top"><font face=webdings color="<%=text%>" size=1>4</font> <b><%=duz(n5("baslik"))%></b><br /><b><%=a_writer%> : </b><%=n5("ekleyen")%><div><br /><b>[<a href="?action=Read&hid=<%=n5("hid")%>"><%=a_read%></a>]</b></div></td>
<td valign="top"><br /><b><%=the_site_cat%> : </b><%=cat("cat_name")%><br /><b><%=the_tarih%> : </b><%=n5("tarih")%></td>
<td valign="top"><br /><b><%=h_averagepoint%> : </b><%=left(n5("puan"),4)%><br /><b><%=a_hit%> : </b><%=n5("okunma")%></td>
</tr>
<%n5.MoveNext
Next%>
<tr class="td_menu"><td bgcolor="#FFFFFF" align="center" colspan="3"><br /><b>[</b> <a href="<%=Request.ServerVariables("HTTP_REFERER")%>"><%=p_back%></a> <b>]</b><br /><br /></td></tr>
</table>
<%End Sub

Sub Stats
Set cnews=Connection.Execute("Select Count(*) AS count FROM COMMENTS WHERE page='News'")
Set t_cats=Connection.Execute("Select Count(*) AS count FROM NEWS_CATS")
Set t_arts=Connection.Execute("Select Count(*) AS count FROM NEWS WHERE onay=True")
Set art_hits=Connection.Execute("SELECT SUM(okunma) AS count FROM NEWS WHERE onay=True")
Set wait_approve=Connection.Execute("Select Count(*) AS count FROM NEWS WHERE onay=False")%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=title_statistics%></font></td></tr>
<tr><td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu"><table border="0" cellspacing="0" cellpadding="1" width="100%" bgcolor="<%=td_border_color%>" class="td_menu">
<tr bgcolor="<%=menu_bg%>" class="td_menu2"><td width="10%" align="center"><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/stats.gif"></td><td width="90%">
<%=astats_totalcats%> : <b><%=t_cats("count")%></b><br />
<%=astats_totalarts%> : <b><%=t_arts("count")%></b><br />
<%=astats_totalread%> : <b><%=art_hits("count")%></b><br />
<%=astats_totalwa%> : <b><%=wait_approve("count")%></b><br />
<%=title_total_comments%> : <b><%=cnews("count")%></b><br />
</td></tr>
</table><br /><b>[</b> <a href="<%=Request.ServerVariables("HTTP_REFERER")%>"><%=p_back%></a> <b>]</b></td></tr>
</table>
<%End Sub

Sub SendTheNews
Set cats=Connection.Execute("Select * FROM NEWS_CATS ORDER BY cat_name ASC")
IF Session("Enter")="Yes" then%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=title_send_news%></font></td></tr>
<tr><td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu">
<form action="?action=RegNews" method="post">
<table border="0" cellspacing="1" cellpadding="1" width="100%" align="center" class="td_menu">
<tr><td width="100" align="right"><%=h_baslik%> : </td><td><input type="text" name="ntitle" size="50" class="form1" style="width:100%"></td></tr>
<tr><td align="right" valign="top"><%=h_text%> : </td><td><textarea name="ntext" rows="8" cols="50" class="form1" style="width:100%"></textarea></td></tr>
<tr><td align="right" valign="top"><%=the_site_cat%> : </td><td><select name="ncat" size="1" class="form1"><%Do Until cats.Eof
Response.Write "<option value='"&cats("cat_id")&"'>"&cats("cat_name")&"</option>"
cats.MoveNext
Loop%></select></td></tr>
<tr><td colspan="2" align="center"><input type="submit" value="   <%=title_send_news%>   " class="submit" style="width:80%"></td></tr>
</table></form>
<script src="htmlarea/editor.js" language="Javascript1.2"></script><script language="JavaScript1.2" defer>editor_generate('ntext');</script>
</td></tr>
</table>
<%cats.Close:Set cats=Nothing
ELSE
Response.Write WriteMsg(no_entry,100)
END IF
End Sub

Sub RegSendedTheNews
title=duz(Request.Form("ntitle"))
news=duz(Request.Form("ntext"))
cat=Request.Form("ncat")
IF title="" Then
Response.Write WriteMsg(art_error2,100)
ElseIF news="" Then
Response.Write WriteMsg(art_error4,100)
ELSE
Set ent=Server.CreateObject("ADODB.RecordSet"):entSQL="Select * FROM NEWS":ent.open entSQL,Connection,1,3
ent.AddNew
ent("baslik")=title
ent("metin")=news
ent("ekleyen")=Session("Member")
ent("tarih")=Now()
ent("onay")=False
ent("puan")=0
ent("katilimci")=0
ent("okunma")=0
ent("cat")=cat
ent.Update
Response.Write WriteMsg(send_all_success,100)
ent.Close:Set ent=Nothing
END IF
End Sub

if Action<>"Print" then call ORTA2:call ALT%>