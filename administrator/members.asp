<!--#include file="view.asp" -->
<%
call UST
call ORTA
%>
              <table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
                <tr>
                  <td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=top_menu5%></font></td>
                </tr>
                <tr>
<td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu">
<%
act = Request.QueryString("action")
If act = "members" Then
call mems
ElseIf act = "member_details" Then
call mem_details
ElseIf act = "add_friendlist" Then
call addfl
ElseIF act = "OnlineUsers" Then
call online_users
End If

Sub online_users
Session.LCID = 1033
liste = DateAdd("n", -1 * 1, Now())
Set ouRs = Server.CreateObject("ADODB.RecordSet")
ouRs.open "Select * FROM MEMBERS WHERE son_tarih >= #"&liste&"# ORDER BY son_tarih DESC",Connection,3,3
Session.LCID = S_LCID
%>
<table border="0" cellspacing="1" cellpadding="2" width="100%" align="center" class="td_menu2" bgcolor="<%=td_border_color%>">
<tr height="20" style="font-weight:bold" class="td_menu">
<td width="20%" align="center" background="THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=OU_UserName%></td>
<td width="30%" align="center" background="THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=OU_LastSession%></td>
<td width="20%" align="center" background="THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=OU_Browser%></td>
<td width="30%" align="center" background="THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=OU_Workstation%></td>
</tr>
<%
Do Until ouRs.EoF
number=number + 1
select_color=right(number,1) 
Select Case select_color
case "1","3","5","7","9"
color = menu_color
Case Else
color = menu_back
End Select
%>
<tr bgcolor="<%=color%>">
<td width="20%" align="center"><%=ouRs("kul_adi")%></td>
<td width="30%" align="center"><%=ouRs("son_tarih")%></td>
<td width="30%" align="center"><%=ouRs("u_browser")%></td>
<td width="20%" align="center"><%=ouRs("u_ws")%></td>
</tr>
<%
ouRs.MoveNext
Loop
%>
</table>
<%
ouRs.Close
Set ouRs = Nothing
End Sub

Sub addfl
uid = Request.QueryString("id")
uid = QS_CLEAR(uid)

If uid="" Then
Response.Redirect Request.ServerVariables("HTTP_REFERER")
Else
Set chkF = Connection.Execute("Select * FROM FRIENDS WHERE MEMBER = "&Session("Uye_ID")&" AND FRIEND = "&uid&"")
If chkF.Eof Then
Set nf = Server.CreateObject("ADODB.RecordSet")
nfSQL = "Select * FROM FRIENDS"
nf.open nfSQL,Connection,3,3

nf.AddNew
nf("MEMBER") = Session("Uye_ID")
nf("FRIEND") = uid
nf.Update
Response.Write "<font class=td_menu><b>"&add_fl_success&"</b></font>"
Else
Response.Write "<font class=td_menu><b>"&added_once&"</b></font>"
End If
Response.Write "<br><br><font class=td_menu><a href=javascript:history.back();>"&p_back&"</a></font>"
End If
End Sub

Sub mems

op = Request.QueryString("x")
op = QS_CLEAR(op)
IF op = "Arrange" Then
rsSQL = "Select * FROM MEMBERS WHERE Left(kul_adi,1) = '"&QS_CLEAR(Request.QueryString("y"))&"' ORDER BY kul_adi ASC"
ElseIF op = "All" OR op = "" Then
rsSQL = "Select * FROM MEMBERS ORDER BY kul_adi ASC"
End IF
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open rsSQL,Connection,3,3

Page = Request.QueryString("Page")
Page = QS_CLEAR(Page)
If Page="" Then
Page = "1"
End If
%>
<table border="0" cellspacing="1" cellpadding="1" width="100%" align="center" class="td_menu" bgcolor="<%=td_border_color%>" style="<%=TableShadow%>">
<tr bgcolor="<%=menu_back%>">
<td align="center" style="font-weight:bold">
<%
	Response.Write "<a href='?action=members&x=All'>"&m_all&"</a>"
	For b = 65 To 90 Step 1
	Response.Write "&nbsp;-&nbsp;<a href='?action=members&x=Arrange&y=" & Chr(b) &"'>" & Chr(b) &"</a>"
	Next
%>
</td>
</tr>
</table>
<br>
<%
Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""1"" width=""100%"" align=""center"" style="""&TableShadow&""" class=""td_menu"" bgcolor="&td_border_color&">"
Response.Write "<tr style=""font-weight:bold"" height=""17""><td width=15% background="&sett("site_adres")&"/THEMES/"&Session("SiteTheme")&"/td_bg.gif>"&detail_memname&"</td><td width=15% background="&sett("site_adres")&"/THEMES/"&Session("SiteTheme")&"/td_bg.gif>"&detail_name&"</td><td width=15% background="&sett("site_adres")&"/THEMES/"&Session("SiteTheme")&"/td_bg.gif>"&detail_mail&"</td></tr>"
IF rs.Eof OR rs.Bof Then
Response.Write ""
ELSE
rs.pagesize = "20"
rs.absolutepage = Page
sayfa = rs.PageCount
For ar = 1 To rs.pagesize
If rs.eof Then Exit For
IF UCase(Left(rs("kul_adi"),1)) <> UCase(QS_CLEAR(Request.QueryString("y"))) AND op = "Arrange" Then
Response.Write ""
ELSE
If rs("mail_goster") = 1 Then
email = rs("email")
Else
email = ""&detail_invisiblemail&""
End If
Response.Write "<tr bgcolor="&menu_bg&"><td width=""15%""><a href=""?action=member_details&uid="&rs("uye_id")&""">"&rs("kul_adi")&"</a></td><td width=""15%"">"&rs("isim")&"</td><td width=""15%"">"&email&"</td></tr>"
End IF
rs.MoveNext
Next
End IF
Response.Write "</table>"
Response.Write "<br><br>"
Response.Write "<b>*</b> " & look_details & "<br>"
%>
<br>
<%
bp = Page-1
IF Page <> 1 Then
With Response
		IF QS_CLEAR(Request.QueryString("x")) <> "Arrange" Then
	.Write "<a href=""?action=members&Page="&bp&""">« "
		ELSE
	.Write "<a href=""?action=members&x=Arrange&y="&QS_CLEAR(Request.QueryString("y"))&"&Page="&bp&""">« "
		End IF
	.Write previous_page & "</a>"
	.Write "&nbsp;-&nbsp;"
End With
End IF
for y=1 to sayfa
IF s=y then
Response.Write y
ELSE
With Response
		IF QS_CLEAR(Request.QueryString("x")) <> "Arrange" Then
	.Write "<a href=""?action=members&Page="&y&""">"
		ELSE
	.Write "<a href=""?action=members&x=Arrange&y="&QS_CLEAR(Request.QueryString("y"))&"&Page="&y&""">"
		End IF
	.Write "<font class=""td_menu"">["&y&"]</font></a>"
End With
End IF
Next
np = Page+1
IF NOT rs.Eof Then
With Response
	.Write "&nbsp;-&nbsp;"
		IF QS_CLEAR(Request.QueryString("x")) <> "Arrange" Then
	.Write "<a href=""?action=members&Page="&np&""">"
		ELSE
	.Write "<a href=""?action=members&x=Arrange&y="&QS_CLEAR(Request.QueryString("y"))&"&Page="&np&""">"
		End IF
	.Write next_page & " »</a>"
End With
End IF

End Sub
Sub mem_details
uyeid = Request.QueryString("mid")
uyeid = QS_CLEAR(uyeid)
IF uyeid="" Then
uyeid = Request.QueryString("uid")
uyeid = QS_CLEAR(uyeid)
END IF
Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM MEMBERS WHERE uye_id = "&uyeid&""
rs.open SQL,Connection,3,3

IF rs.EoF Then
Response.Write WriteMsg(error17)
ELSE
Session.LCID = S_LCID
kadi = rs("kul_adi")
name = rs("isim")
login = rs("login_sayisi")
lastlogin = rs("son_tarih")
regdate = rs("uyelik_tarihi")
msg_count = rs("msg_sayisi")

If rs("mail_goster") = 1 Then
email = rs("email")
Else
email = ""&detail_invisiblemail&""
End If

If rs("icq") = "" Then
icq = ""&detail_empty&""
Else
icq = rs("icq")
End If

If rs("msn") = "" Then
msn = ""&detail_empty&""
Else
msn = rs("msn")
End If

If rs("aim") = "" Then
aim = ""&detail_empty&""
Else
aim = rs("aim")
End If

If rs("sehir") = "" Then
city = ""&detail_empty&""
Else
city = rs("sehir")
End If

If rs("meslek") = "" Then
job = ""&detail_empty&""
Else
job = rs("meslek")
End If

If rs("cinsiyet") = "" Then
sex = ""&detail_empty&""
ElseIf rs("cinsiyet") = "a" Then
sex = ""&male&""
ElseIf rs("cinsiyet") = "b" Then
sex = ""&female&""
Else
sex = rs("cinsiyet")
End If

strAge = Int(Right(rs("yas"),4))
strNow = Int(Right(Date(),4))

age = strNow - strAge

If rs("url") = "" Then
url = ""&detail_empty&""
Else
IF InStr(1,rs("url"),"http://",1) Then
url = rs("url")
ELSE
url = "http://" & rs("url")
End IF
End If

If rs("imza") = "" Then
signature = detail_empty
Else
signature = rs("imza")
End IF

IF rs("u_avatar") = "" Then
member_avatar = "IMAGES/avatars/blank.gif"
ELSE
member_avatar = rs("u_avatar")
End IF

ip = rs("last_ip")
M_FLAG_URL = "http://www.ripe.net/perl/whois?searchtext="&ip&""
veri = Get_CF(M_FLAG_URL)

IF Err <> 0 Then
Flag = "f-TI.gif"
End IF

IF Err = 0 Then
gelenveri = Instr(1,veri,"country")
ulke = Mid(veri,gelenveri+14,2)
Flag = "f-" & ulke & ".gif"
End IF

Set frms = Server.CreateObject("ADODB.RecordSet")
frms.Open "Select * FROM CATEGORIES WHERE mods LIKE '%"&kadi&"%'",Connect,3,3
Do Until frms.EoF
IF NOT frms.EoF Then
mem_mod = "Yes"
ELSE
mem_mod = "No"
End IF
frms.MoveNext
Loop

IF rs("seviye") = "1" Then
seviye = level1
ElseIF rs("seviye") = "2" Then
seviye = level2
ElseIF rs("seviye") = "3" Then
seviye = level3
ElseIF rs("seviye") = "4" Then
seviye = level4
ElseIF rs("seviye") = "5" Then
seviye = level5
ElseIF mem_mod = "Yes" Then
seviye = level6
ELSE
seviye = normal
End IF
%>
<center><b><%=rs("kul_adi")%></b> (<%=seviye%>)</center>
<table border="0" cellspacing="1" cellpadding="0" align="center" width="100%" class="td_menu" bgcolor="<%=td_border_color%>">
<tr bgcolor="<%=menu_bg%>">
<td width="40%" valign="top">
<table border="0" cellspacing="1" cellpadding="0" align="center" width="100%" class="td_menu">
<tr align="center" height="18" style="font-weight:bold" bgcolor="<%=menu_bg%>">
<td colspan="2" background="THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=detail_avatar%></td>
</tr>
<tr align="center" bgcolor="<%=menu_bg%>">
<td colspan="2"><img src="<%=member_avatar%>" border="1"></td>
</tr>
<tr align="center" height="18" style="font-weight:bold" bgcolor="<%=menu_bg%>">
<td colspan="2" background="THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=detail_contact%></td>
</tr>
<tr align="center" bgcolor="<%=menu_bg%>">
<td width="25%" align="right"><%=detail_mail%> : </td>
<td width="75%" align="left"><%=email%></td>
</tr>
<tr align="center" bgcolor="<%=menu_bg%>">
<td width="25%" align="right"><%=detail_icq%> : </td>
<td width="75%" align="left"><%=icq%></td>
</tr>
<tr align="center" bgcolor="<%=menu_bg%>">
<td width="25%" align="right"><%=detail_msn%> : </td>
<td width="75%" align="left"><%=msn%></td>
</tr>
<tr align="center" bgcolor="<%=menu_bg%>">
<td width="25%" align="right"><%=detail_aim%> : </td>
<td width="75%" align="left"><%=aim%></td>
</tr>
<tr bgcolor="<%=menu_bg%>">
<td colspan="2" align="center">
<br>
<%
IF rs("u_online") = 1 Then
detail_user_image = "user_online.gif"
ELSE
detail_user_image = "user_offline.gif"
End IF
%>
<img src="IMAGES/stats/<%=detail_user_image%>" border="0">
</td>
</tr>
</table>
</td>
<td width="60%" valign="top">
<table border="0" cellspacing="1" cellpadding="0" align="center" width="100%" class="td_menu">
<tr align="center" height="18" style="font-weight:bold" bgcolor="<%=menu_bg%>">
<td colspan="2" background="THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><%=detail_about%></td>
</tr>
<tr align="center" bgcolor="<%=menu_bg%>">
<td width="50%" align="right"><%=detail_city%> : </td>
<td width="50%" align="left"><%=city%></td>
</tr>
<tr align="center" bgcolor="<%=menu_bg%>">
<td width="50%" align="right" valign="top"><%=detail_country%> : </td>
<td width="50%" align="left"><img src="IMAGES/flags/<%=Flag%>"></td>
</tr>
<tr align="center" bgcolor="<%=menu_bg%>">
<td width="50%" align="right"><%=detail_job%> : </td>
<td width="50%" align="left"><%=job%></td>
</tr>
<tr align="center" bgcolor="<%=menu_bg%>">
<td width="50%" align="right"><%=detail_sex%> : </td>
<td width="50%" align="left"><%=sex%></td>
</tr>
<tr align="center" bgcolor="<%=menu_bg%>">
<td width="50%" align="right"><%=detail_age%> : </td>
<td width="50%" align="left"><%=age%></td>
</tr>
<tr align="center" bgcolor="<%=menu_bg%>">
<td width="50%" align="right"><%=detail_url%> : </td>
<td width="50%" align="left"><a href="<%=url%>" target="_blank"><%=url%></a></td>
</tr>
<tr align="center" bgcolor="<%=menu_bg%>">
<td width="50%" align="right"><%=detail_logcount%> : </td>
<td width="50%" align="left"><%=login%></td>
</tr>
<tr align="center" bgcolor="<%=menu_bg%>">
<td width="50%" align="right"><%=detail_msgcount%> : </td>
<td width="50%" align="left"><%=msg_count%></td>
</tr>
<tr align="center" bgcolor="<%=menu_bg%>">
<td width="50%" align="right"><%=detail_regdate%> : </td>
<td width="50%" align="left"><%=rs("uyelik_tarihi")%></td>
</tr>
<tr align="center" bgcolor="<%=menu_bg%>">
<td width="50%" align="right"><%=detail_lastlogin%> : </td>
<td width="50%" align="left"><%=rs("son_tarih")%></td>
</tr>
<tr align="center" bgcolor="<%=menu_bg%>">
<td width="50%" align="right" valign="top"><%=detail_signature%> : </td>
<td width="50%" align="left"><%=signature%></td>
</tr>
</table>
</td>
</tr>
</table>
<%
If Session("Enter") = "Yes" Then
If Session("Member") <> rs("kul_adi") Then
Set chkF = Connection.Execute("Select * FROM FRIENDS WHERE MEMBER = "&Session("Uye_ID")&" AND FRIEND = "&QS_CLEAR(Request.QueryString("uid"))&"")
IF NOT chkF.EoF Then
addfl_button = " disabled"
End IF
%>
<br>
<table border="0" cellspacing="1" cellpadding="1" align="right">
<tr>
<form method="post" action="msgbox.asp?action=new_msg&who=<%=rs("kul_adi")%>">
<td align="right">
<input type="submit" value="<%=msg_send%>" class="submit">
</td>
</form>
<form method="post" action="?action=add_friendlist&id=<%=rs("uye_id")%>">
<td align="left">
<input type="submit" value="<%=add_fl%>" class="submit" <%=addfl_button%>>
</td>
</form>
</tr>
</table>
<%
End IF
End IF
End IF
End Sub
%>
</td>
</tr>
</table>
<%
call ORTA2
call ALT
%>