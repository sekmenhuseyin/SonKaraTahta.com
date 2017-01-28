<!--#include file="view.asp" -->
<%call UST:call ORTA
act=Request.QueryString("action")
IF act="HitMembers" Then
pg_title=title_popular
ElseIF act="NewMembers" Then
pg_title=title_new
ElseIF act="Stats" Then
pg_title=title_statistics
Else
pg_title=top_menu5
End If%>
<table border="0" cellspacing="0" cellpadding="2" width="100%" style="<%=TableBorder%>" class="td_menu">
<tr bgcolor="<%=menu_back%>"><td align="center" style="font-size:18px"><%=sett("site_isim")%> : <%=top_menu5%></td></tr>
</table><br />
<table border="0" cellspacing="0" cellpadding="2" width="100%" style="<%=TableBorder%>" class="td_menu">
<tr bgcolor="<%=menu_back%>"><td align="center">
<a href="?action=members"><img src="images/logos/uye.gif" border="0"></a><br />
<b>[</b> <a href="?action=HitMembers"><%=title_popular%></a> <b>|</b> <a href="?action=NewMembers"><%=title_new%></a> <b>|</b> <a href="search.asp?action=Member"><%=uyemenu_uyeara%></a> <b>|</b> <a href="?action=Stats"><%=title_statistics%></a> <b>]</b>
</td></tr>
</table><br /><table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25">
<font class="block_title" face="Verdana">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=pg_title%></font>
</td></tr><tr><td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu">
<%If act="members" Then
pg_title=top_menu5:call mems
ElseIf act="member_details" Then
pg_title=top_menu5:call mem_details
ElseIf act="add_friendlist" Then
pg_title=top_menu5:call addfl
ElseIF act="OnlineUsers" Then
pg_title=top_menu5:call online_users
ElseIF act="HitMembers" Then
pg_title=title_popular:call hit_users
ElseIF act="NewMembers" Then
pg_title=title_new:call new_users
ElseIF act="Stats" Then
pg_title=title_statistics:call stats
End If

Sub online_users
Session.LCID=1033
liste=DateAdd("n", -1 * 1, Now())
Set ouRs=Server.CreateObject("ADODB.RecordSet")
ouRs.open "Select * FROM MEMBERS WHERE son_tarih >= #"&liste&"# ORDER BY son_tarih DESC",Connection,3,3
Session.LCID=S_LCID%>
<table border="0" cellspacing="1" cellpadding="2" width="100%" align="center" class="td_menu2" bgcolor="<%=td_border_color%>">
<tr height="20" style="font-weight:bold" class="td_menu">
<td width="20%" align="center" background="THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=OU_UserName%></td>
<td width="30%" align="center" background="THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=OU_LastSession%></td>
<td width="20%" align="center" background="THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=OU_Browser%></td>
<td width="30%" align="center" background="THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=OU_Workstation%></td>
</tr>
<%Do Until ouRs.EoF
number=number + 1
select_color=right(number,1) 
Select Case select_color
case "1","3","5","7","9"
color=menu_color
Case Else
color=menu_back
End Select
%><font face="Verdana">
<tr bgcolor="<%=color%>">
<td width="20%" align="center"><%=ouRs("kul_adi")%></td>
<td width="30%" align="center"><%=ouRs("son_tarih")%></td>
<td width="30%" align="center"><%=ouRs("u_browser")%></td>
<td width="20%" align="center"><%=ouRs("u_ws")%></td>
</tr>
<%ouRs.MoveNext
Loop%>
</table>
<%ouRs.Close:Set ouRs=Nothing
End Sub

Sub addfl
uid=Request.QueryString("id"):uid=QS_CLEAR(uid,"false")
If uid="" Then
Response.Redirect Request.ServerVariables("HTTP_REFERER")
Else
Set chkF=Connection.Execute("Select * FROM FRIENDS WHERE MEMBER="&Session("Uye_ID")&" AND FRIEND="&uid&"")
If chkF.Eof Then
Set nf=Server.CreateObject("ADODB.RecordSet")
nfSQL="Select * FROM FRIENDS"
nf.open nfSQL,Connection,3,3
nf.AddNew
nf("MEMBER")=Session("Uye_ID")
nf("FRIEND")=uid
nf.Update
Response.Write "<font class=td_menu><b>"&add_fl_success&"</b></font>"
Else
Response.Write "<font class=td_menu><b>"&added_once&"</b></font>"
End If
Response.Write "<br /><br /><font class=td_menu><a href=javascript:history.back();>"&p_back&"</a></font>"
End If
End Sub

Sub mems
op=Request.QueryString("x"):op=QS_CLEAR(op,"")
IF op="Arrange" Then
rsSQL="Select * FROM MEMBERS WHERE Left(kul_adi,1)='"&QS_CLEAR(Request.QueryString("y"),"false")&"' ORDER BY kul_adi ASC"
ElseIF op="signs" Then
rsSQL="Select * FROM MEMBERS WHERE cint(asc(Left(kul_adi,1)))<65 ORDER BY kul_adi ASC"
ElseIF op="All" OR op="" Then
rsSQL="Select * FROM MEMBERS ORDER BY kul_adi ASC"
End IF
Set rs=Server.CreateObject("ADODB.RecordSet"):rs.open rsSQL,Connection,3,3
Page=Request.QueryString("Page"):Page=QS_CLEAR(Page,"false")
If Page="" Then Page="1"%>
<table border="0" cellspacing="1" cellpadding="1" width="100%" align="center" class="td_menu" bgcolor="<%=td_border_color%>" style="<%=TableShadow%>">
<tr bgcolor="<%=menu_back%>"><td align="center" style="font-weight:bold">
<%Response.Write "<a href='?action=members&x=All'>"&m_all&"</a>"
Response.Write "&nbsp;-&nbsp;<a href='?action=members&x=signs'>#</a>"
For b=65 To 90 Step 1
Response.Write "&nbsp;-&nbsp;<a href='?action=members&x=Arrange&y=" & Chr(b) &"'>" & Chr(b) &"</a>"
Next%>
</td></tr>
</table><font face="Verdana"><br />
<table border="0" cellspacing="1" cellpadding="1" width="100%" align="center" style="<%=TableShadow%>" class="td_menu" bgcolor="<%=td_border_color%>">
<tr style="font-weight:bold" height="17">
<td width="15%" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=detail_memname%></td>
<td width="15%" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=detail_name%></td>
<td width="15%" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=detail_mail%></td>
</tr>
<%IF rs.Eof OR rs.Bof Then
Response.Write ""
ELSE
rs.pagesize="20"
rs.absolutepage=Page
sayfa=rs.PageCount
For ar=1 To rs.pagesize
If rs.eof Then Exit For
IF UCase(Left(rs("kul_adi"),1)) <> UCase(QS_CLEAR(Request.QueryString("y"),"")) AND op="Arrange" Then
Response.Write ""
ELSE
If rs("mail_goster")=True Then email=rs("email") Else email=""&detail_invisiblemail&""
Response.Write "<tr bgcolor="&menu_bg&"><td width=""15%""><a href=""?action=member_details&uid="&rs("uye_id")&""">"&rs("kul_adi")&"</a></td><td width=""15%"">"&rs("isim")&"</td><td width=""15%"">"&email&"</td></tr>"
End IF
rs.MoveNext
Next
End IF
Response.Write "</table><br /><br /><b>*</b> " & look_details & "<br /><br />"
bp=Page-1
IF Page <> 1 Then
IF QS_CLEAR(Request.QueryString("x"),"") <> "Arrange" Then Response.Write "<a href=""?action=members&Page="&bp&""" style=""font-size:10px"">« " ELSE Response.Write "<a href=""?action=members&x=Arrange&y="&QS_CLEAR(Request.QueryString("y"),"")&"&Page="&bp&""" style=""font-size:10px"">« "
Response.Write previous_page & "</a>&nbsp;-&nbsp;"
End IF
for y=1 to sayfa
IF cint(Page)=cint(y) then
Response.Write " <font class=""td_menu"" style=""font-size:10px"">["&y&"]</font> "
ELSE
IF QS_CLEAR(Request.QueryString("x"),"") <> "Arrange" Then Response.Write " <a href=""?action=members&Page="&y&""" style=""font-size:10px"">" ELSE Response.Write " <a href=""?action=members&x=Arrange&y="&QS_CLEAR(Request.QueryString("y"),"")&"&Page="&y&""" style=""font-size:10px"">"
Response.Write "["&y&"]</a> "
End IF
Next
IF NOT rs.Eof Then
Response.Write "&nbsp;-&nbsp;"
IF QS_CLEAR(Request.QueryString("x"),"") <> "Arrange" Then Response.Write "<a href=""?action=members&Page="&(cint(Page)+1)&""" style=""font-size:10px"">" ELSE Response.Write "<a href=""?action=members&x=Arrange&y="&QS_CLEAR(Request.QueryString("y"),"")&"&Page="&(cint(Page)+1)&""" style=""font-size:10px"">"
Response.Write next_page & " »</a>"
End IF
End Sub

Sub mem_details
uyeid=Request.QueryString("mid"):uyeid=QS_CLEAR(uyeid,"false")
IF uyeid="" Then uyeid=Request.QueryString("uid"):uyeid=QS_CLEAR(uyeid,"false")
Set rs=Server.CreateObject("ADODB.RecordSet"):SQL="Select * FROM MEMBERS WHERE uye_id="&uyeid&"":rs.open SQL,Connection,3,3
IF rs.EoF Then
Response.Write WriteMsg(error17,100)
ELSE
Session.LCID=S_LCID
kadi=rs("kul_adi")
name=rs("isim")
login=rs("login_sayisi")
lastlogin=rs("son_tarih")
regdate=rs("uyelik_tarihi")
msg_count=rs("msg_sayisi")
If rs("mail_goster")=True Then email=rs("email") Else email=""&detail_invisiblemail&""
If rs("icq")="0" Then icq=""&detail_empty&"" Else icq=rs("icq")
If rs("msn")="" Then msn=""&detail_empty&"" Else msn=rs("msn")
If rs("aim")="" Then aim=""&detail_empty&"" Else aim=rs("aim")
If rs("sehir")="" Then city=""&detail_empty&"" Else city=rs("sehir")
If rs("meslek")="" Then job=""&detail_empty&"" Else job=rs("meslek")
If rs("cinsiyet")="" Then 
sex=""&detail_empty&""
ElseIf rs("cinsiyet")="a" Then
sex=""&male&""
ElseIf rs("cinsiyet")="b" Then
sex=""&female&""
Else
sex=rs("cinsiyet")
End If
If rs("url")="" Then
url=""&detail_empty&""
Else
IF InStr(1,rs("url"),"http://",1) Then
url=rs("url")
ELSE
url="http://" & rs("url")
End IF
End If
If rs("imza")="" Then signature=detail_empty Else signature=rs("imza")
IF rs("u_avatar")="" Then member_avatar="IMAGES/avatars/blank.gif" ELSE member_avatar=rs("u_avatar")
Flag="f-tr.gif"
Set frms=Server.CreateObject("ADODB.RecordSet"):frms.Open "Select * FROM CATEGORIES WHERE mods LIKE '%"&kadi&"%'",Connect,3,3
Do Until frms.EoF
IF NOT frms.EoF Then mem_mod="Yes" ELSE mem_mod="No"
frms.MoveNext
Loop
IF rs("seviye")="1" Then
seviye=level1
ElseIF rs("seviye")="2" Then
seviye=level2
ElseIF rs("seviye")="3" Then
seviye=level3
ElseIF rs("seviye")="4" Then
seviye=level4
ElseIF rs("seviye")="5" Then
seviye=level5
ElseIF rs("seviye")="6" Then
seviye=level6
ElseIF rs("seviye")="8" Then
seviye=level8
ElseIF mem_mod="Yes" Then
seviye=level6
ElseIF mem_mod="Yes" Then
seviye=level8
ELSE
seviye=normal
End IF
%></font><center><b><%=rs("kul_adi")%><font face="Verdana"> </font></b><font face="Verdana"><b>(<%=seviye%>)</b></font></center>
<font color="#FF0000"><b><font size="2" face="Marlett">h</font><font face="Verdana" size="2"><%=rs("ozellik")%></font><font size="2" face="Marlett">h</font></b></font><table border="0" cellspacing="1" cellpadding="0" align="center" width="100%" class="td_menu" bgcolor="<%=td_border_color%>">
<tr bgcolor="<%=menu_bg%>"><td width="40%" valign="top"><table border="0" cellspacing="1" cellpadding="0" align="center" width="100%" class="td_menu">
<tr align="center" height="18" style="font-weight:bold" bgcolor="<%=menu_bg%>"><td colspan="2" background="THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=detail_avatar%></td></tr>
<tr align="center" bgcolor="<%=menu_bg%>"><td colspan="2"><font face="Verdana"><img height="100" width="100" src="<%=member_avatar%>" border="1"></font></td></tr>
<tr align="center" height="18" style="font-weight:bold" bgcolor="<%=menu_bg%>"><td colspan="2" background="THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=detail_contact%></td></tr>
<tr align="center" bgcolor="<%=menu_bg%>"><td width="25%" align="right"><%=detail_mail%><font face="Verdana"> </font><font face="Verdana">: </font> </td><td width="75%" align="left"><%=email%></td></tr>
<tr align="center" bgcolor="<%=menu_bg%>"><td width="25%" align="right"><%=detail_icq%><font face="Verdana"> </font><font face="Verdana">: </font> </td><td width="75%" align="left"><%=icq%></td></tr>
<tr align="center" bgcolor="<%=menu_bg%>"><td width="25%" align="right"><%=detail_msn%><font face="Verdana"> </font><font face="Verdana">: </font> </td><td width="75%" align="left"><%=msn%></td></tr>
<tr align="center" bgcolor="<%=menu_bg%>"><td width="25%" align="right"><%=detail_aim%><font face="Verdana"> </font><font face="Verdana">: </font> </td><td width="75%" align="left"><%=aim%></td></tr>
<tr bgcolor="<%=menu_bg%>"><td colspan="2" align="center"><font face="Verdana"><br />
<%IF rs("u_online")=True Then detail_user_image="user_online.gif" ELSE detail_user_image="user_offline.gif"%>
<img src="IMAGES/stats/<%=detail_user_image%>" border="0"> </font></td></tr>
</table></td>
<td width="60%" valign="top"><table border="0" cellspacing="1" cellpadding="0" align="center" width="100%" class="td_menu">
<tr align="center" height="18" style="font-weight:bold" bgcolor="<%=menu_bg%>"><td colspan="2" background="THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=detail_about%></td></tr>
<tr align="center" bgcolor="<%=menu_bg%>"><td width="50%" align="right"><%=detail_city%><font face="Verdana">: </font> </td><td width="50%" align="left"><%=city%></td></tr>
<tr align="center" bgcolor="<%=menu_bg%>"><td width="50%" align="right" valign="top"><%=detail_country%><font face="Verdana">: </font> </td><td width="50%" align="left"><font face="Verdana"><img src="IMAGES/site/<%=Flag%>"></font></td></tr>
<tr align="center" bgcolor="<%=menu_bg%>"><td width="50%" align="right"><%=detail_job%><font face="Verdana">: </font> </td><td width="50%" align="left"><%=job%></td></tr>
<tr align="center" bgcolor="<%=menu_bg%>"><td width="50%" align="right"><%=detail_sex%><font face="Verdana">: </font> </td><td width="50%" align="left"><%=sex%></td></tr>
<tr align="center" bgcolor="<%=menu_bg%>"><td width="50%" align="right"><%=detail_url%><font face="Verdana">: </font> </td><td width="50%" align="left"><a href="<%=url%>" target="_blank"><%=url%></a></td></tr>
<tr align="center" bgcolor="<%=menu_bg%>"><td width="50%" align="right"><%=detail_logcount%><font face="Verdana">: </font> </td><td width="50%" align="left"><%=login%></td></tr>
<tr align="center" bgcolor="<%=menu_bg%>"><td width="50%" align="right"><%=detail_msgcount%><font face="Verdana">: </font> </td><td width="50%" align="left"><%=msg_count%></td></tr>
<tr align="center" bgcolor="<%=menu_bg%>"><td width="50%" align="right"><%=detail_regdate%><font face="Verdana">: </font> </td><td width="50%" align="left"><%=rs("uyelik_tarihi")%></td></tr>
<tr align="center" bgcolor="<%=menu_bg%>"><td width="50%" align="right"><%=detail_lastlogin%><font face="Verdana">: </font> </td><td width="50%" align="left"><%=rs("son_tarih")%></td></tr>
<tr align="center" bgcolor="<%=menu_bg%>"><td width="50%" align="right" valign="top"><%=detail_signature%><font face="Verdana">: </font> </td><td width="50%" align="left"><%=signature%></td></tr>
</table></td></tr>
</table>
<%
If Session("Enter")="Yes" Then
If Session("Member") <> rs("kul_adi") Then
Set chkF=Connection.Execute("Select * FROM FRIENDS WHERE MEMBER="&Session("Uye_ID")&" AND FRIEND="&QS_CLEAR(Request.QueryString("uid"),"false")&"")
IF NOT chkF.EoF Then
addfl_button=" disabled"
End IF
%> </font><font face="Verdana">
<br />
</font>
<table border="0" cellspacing="1" cellpadding="1" align="right">
<tr>
<form method="post" action="msgbox.asp?action=new_msg&who=<%=rs("kul_adi")%>">
<td align="right"><font face="Verdana"><input type="submit" value="<%=msg_send%>" class="submit"></font></td>
</form>
<form method="post" action="?action=add_friendlist&id=<%=rs("uye_id")%>">
<td align="left"><font face="Verdana"><input type="submit" value="<%=add_fl%>" class="submit" <%=addfl_button%>></font></td>
</form>
</tr>
</table>
<%End IF
End IF
End IF
End Sub

Sub hit_users
Set rs=Server.CreateObject("ADODB.RecordSet"):rs.open "Select * FROM MEMBERS order by msg_sayisi desc",Connection,3,3%>
<table border="0" cellspacing="1" cellpadding="1" width="100%" align="center" style="<%=TableShadow%>" class="td_menu" bgcolor="<%=td_border_color%>">
<tr style="font-weight:bold" height="17">
<td width="15%" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=detail_memname%></td>
<td width="15%" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=detail_name%></td>
<td width="15%" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=detail_mail%></td>
<td width="15%" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=detail_msgcount%></td>
</tr>
<%For J=1 TO 15
IF rs.Eof Then Exit For
If rs("mail_goster")=True Then email=rs("email") Else email=detail_invisiblemail%>
<tr bgcolor="<%=menu_bg%>"><td width="15%"><a href="?action=member_details&uid=<%=rs("uye_id")%>"><%=rs("kul_adi")%></a></td>
<td width="15%"><%=rs("isim")%></td><td width="15%"><%=email%></td><td width="15%"><%=rs("msg_sayisi")%></td></tr>
<%rs.MoveNext
Next%>
</td></tr>
</table><br /><b>[</b> <a href="<%=Request.ServerVariables("HTTP_REFERER")%>"><%=p_back%></a> <b>]</b>
<%End Sub

Sub new_users
Set rs=Server.CreateObject("ADODB.RecordSet"):rs.open "Select * FROM MEMBERS order by uyelik_tarihi desc",Connection,3,3%>
<table border="0" cellspacing="1" cellpadding="1" width="100%" align="center" style="<%=TableShadow%>" class="td_menu" bgcolor="<%=td_border_color%>">
<tr style="font-weight:bold" height="17">
<td width="15%" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=detail_memname%></td>
<td width="15%" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=detail_name%></td>
<td width="15%" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=detail_mail%></td>
<td width="15%" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=detail_regdate%></td>
</tr>
<%For J=1 TO 15
IF rs.Eof Then Exit For
If rs("mail_goster")=True Then email=rs("email") Else email=detail_invisiblemail%>
<tr bgcolor="<%=menu_bg%>"><td width="15%"><a href="?action=member_details&uid=<%=rs("uye_id")%>"><%=rs("kul_adi")%></a></td>
<td width="15%"><%=rs("isim")%></td><td width="15%"><%=email%></td><td width="15%"><%=rs("uyelik_tarihi")%></td></tr>
<%rs.MoveNext
Next%>
</td></tr>
</table><br /><b>[</b> <a href="<%=Request.ServerVariables("HTTP_REFERER")%>"><%=p_back%></a> <b>]</b>
<%End Sub

Sub Stats
Session.LCID=2048
Set members=Connection.Execute("Select Count(*) AS count FROM MEMBERS")
Set mem_msg=Connection.Execute("Select Count(*) AS count FROM MESSAGES WHERE SEE=True AND TYPE=0")
Set admins=Connection.Execute("Select Count(*) AS count FROM MEMBERS WHERE seviye>0")
Set total_login=Connection.Execute("SELECT SUM(login_sayisi) AS count FROM MEMBERS")%>
<table border="0" cellspacing="0" cellpadding="1" width="100%" bgcolor="<%=td_border_color%>" class="td_menu">
<tr bgcolor="<%=menu_bg%>" class="td_menu2"><td width="10%" align="center"><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/stats.gif"></td><td width="90%">
Toplam Üye Sayýsý : <b><%=members("count")%></b><br />
Toplam Giriþ Yapýlma Sayýsý : <b><%=total_login("count")%></b><br />
Toplam Yönetici Sayýsý : <b><%=admins("count")%></b><br />
Toplam Mesaj Sayýsý : <b><%=mem_msg("count")%></b><br />
</td></tr>
</table><br /><b>[</b> <a href="<%=Request.ServerVariables("HTTP_REFERER")%>"><%=p_back%></a> <b>]</b>
<%End Sub%>
</td></tr>
</table>
<%call ORTA2:call ALT%>