<!--#include file="view.asp" -->
<%call UST
call ORTA%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=uyemenu_msgbox%></font></td></tr>
<tr> 
<td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu">
<%IF Session("Enter")="Yes" Then

Set msgs=Connection.Execute("Select Count(*) as count FROM MESSAGES WHERE TYPE=0 AND SEE=True AND TO='"&Session("Member")&"'")
Set s_msgs=Connection.Execute("Select Count(*) as count FROM MESSAGES WHERE TYPE=0 AND FROM='"&Session("Member")&"' AND MSAVE=True")
Set admin=Connection.Execute("Select Count(*) as count FROM MESSAGES WHERE TYPE=1 AND SEE=True")

If msgs("count") >= sett("msg_limit") Then
t_msgs=sett("msg_limit")
Else
t_msgs=msgs("count")
End If
If s_msgs("count") >= sett("msg_limit") Then
s_msgs=sett("msg_limit")
Else
s_msgs=s_msgs("count")
End If
%>
<table border="0" cellspacing="10" cellpadding="10" width="100%" class="td_menu" style="font-weight:bold">
<tr align="center" valign="bottom">
<td><a href="Your_Account.asp?op=Profile"><img src="images/Your_Account/info.gif" border="0"><br /><%=ya_info%></a></td>
<td><a href="Msgbox.asp?action=main"><img src="images/Your_Account/messages.gif" border="0"><br /><%=ya_messages%></a></td>
<td><a href="Your_Account.asp?op=Friends"><img src="images/Your_Account/friends.gif" border="0"><br /><%=msg_friends%></a></td>
<td><a href="Your_Account.asp?op=SelectTheme"><img src="images/Your_Account/themes.gif" border="0"><br /><%=ya_theme%></a></td>
<td><a href="Your_Account.asp?op=Logout"><img src="images/Your_Account/exit.gif" border="0"><br /><%=ya_logout%></a></td>
</tr>
</table>
<hr size="1" color="<%=td_border_color%>">
<table border="0" cellspacing="1" cellpadding="0" width="100%" align="center" height="20" bgcolor="<%=td_border_color%>" class="td_menu" style="font-weight:bold">
<tr>
<td align="center" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><a href="?action=msgbox&id=<%=Session("Uye_ID")%>"><%=msg_box%> (<%=t_msgs%>)</a></td>
<td align="center" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><a href="?action=saved&id=<%=Session("Uye_ID")%>"><%=msg_saved%> (<%=s_msgs%>)</a></td>
<td align="center" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><a href="?action=fromadmin&id=<%=Session("Uye_ID")%>"><%=msg_admin%> (<%=admin("count")%>)</a></td>
<td align="center" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/td_bg.gif"><a href="?action=new_msg"><%=msg_new%></a></td>
</tr>
</table><br />
<%action=Request.QueryString("action"):action=QS_CLEAR(action,"")

if action="main" then
call messages
elseif action="msgbox" then
call messages
elseif action="fromadmin" then
call fromadmin
elseif action="new_msg" then
call new_msg
elseif action="myfriends" then
call friends
elseif action="fact" then
call fact
elseif action="msg_rec" then
call msg_record
elseif action="msg_del" then
call delete_msg
elseif action="read_msg" then
call read
ElseIF action="saved" Then
call saved
end if

Sub fact
f_action=Request.Form("fr_act")
f_member=Request.Form("fr_nm")

IF f_member="" Then
Response.Write ps_member

ElseIF f_action=""&f_delf&"" Then
Set fmi=Connection.Execute("Select * FROM MEMBERS WHERE kul_adi='"&f_member&"'")
Set df=Server.CreateObject("ADODB.RecordSet")
dfSQL="DELETE * FROM FRIENDS WHERE FRIEND="&fmi("uye_id")&""
df.open dfSQL,Connection,3,3
Response.Write f_success_del
fmi.Close
Set fmi=Nothing

ElseIF f_action=""&f_msg&"" Then
Response.Redirect "msgbox.asp?action=new_msg&who="&f_member&""

ElseIF f_action=""&f_info&"" Then
Set f_m=Connection.Execute("Select * FROM MEMBERS WHERE kul_adi='"&f_member&"'")
Response.Redirect "members.asp?action=member_details&uid="&f_m("uye_id")&""
f_m.Close
Set f_m=Nothing

End IF
End Sub
Sub messages
Session("Folder")="Inbox"
Set rs=Server.CreateObject("ADODB.RecordSet")
SQL="Select * FROM MESSAGES WHERE TO='"&Session("Member")&"' AND SEE=True AND TYPE=0 ORDER BY mdate DESC"
rs.open SQL,Connection,3,3
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" align="center" bgcolor="<%=menu_bg%>" class="td_menu">
<tr bgcolor="<%=background%>" style="font-weight:bold"><td width="5%" align="center">&nbsp;</td><td width="20%" align=center><%=msg_sender%></td><td width="40%" align=center><%=msg_subject%></td><td width="40%" align=center><%=the_tarih%></td></tr>
<%
For i=1 TO sett("msg_limit")
If rs.Eof Then Exit For
Set fmem=Connection.Execute("Select * FROM MEMBERS WHERE kul_adi='"&rs("from")&"'")
IF fmem.Eof OR fmem.Bof Then
msg_mid=""
ELSE
msg_mid=fmem("uye_id")
End IF

number=number + 1
select_color=right(number,1) 
Select Case select_color
case "1","3","5","7","9"
color=forum_bg_1
Case Else
color=forum_bg_2
End Select
%>
<tr bgcolor="<%=color%>" class="td_menu2"><td width="5%" align="center"><% If rs("read")=True Then %><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/readed.gif"><% ElseIf rs("read")=False Then %><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/not_readed.gif"><% End If %></td>
<td width="20%" align=center><a href="members.asp?action=member_details&uid=<%=msg_mid%>"><%=duz(rs("from"))%></a></td>
<td width="40%"><a href="?action=read_msg&mid=<%=rs("mid")%>"><%=duz(rs("subject"))%></a></td>
<td width="40%" align=center><%=rs("mdate")%></td>
<td width="2%" align="center"><a href="?action=msg_del&id=<%=rs("mid")%>"><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/msg_del.gif" border="0" alt="<%=msg_delete_want%>"></a></td>
<%
rs.MoveNext
Next
%>
</table>
<%
rs.Close
Set rs=Nothing
End Sub

Sub saved
Session("Folder")="Saved"
Set rs=Server.CreateObject("ADODB.RecordSet")
SQL="Select * FROM MESSAGES WHERE FROM='"&Session("Member")&"' AND TYPE=0 AND MSAVE=True ORDER BY mdate DESC"
rs.open SQL,Connection,3,3
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" align="center" bgcolor="<%=menu_bg%>" class="td_menu">
<tr style="font-weight:bold"><td width="5%" align="center">&nbsp;</td><td width="20%" bgcolor=<%=background%> align=center><%=msg_sender%></td><td width="40%" bgcolor=<%=background%> align=center><%=msg_subject%></td><td width="40%" bgcolor=<%=background%> align=center><%=the_tarih%></td></tr>
<%
Do Until rs.EoF
Set fmem=Connection.Execute("Select * FROM MEMBERS WHERE kul_adi='"&rs("from")&"'")
IF fmem.Eof OR fmem.Bof Then
msg_mid=""
ELSE
msg_mid=fmem("uye_id")
End IF

number=number + 1
select_color=right(number,1) 
Select Case select_color
case "1","3","5","7","9"
color=forum_bg_1
Case Else
color=forum_bg_2
End Select
%>
<tr bgcolor="<%=color%>" class="td_menu2"><td width="5%" align="center" bgcolor="<%=menu_bg%>"></td>
<td width="20%" align=center><a href="members.asp?action=member_details&uid=<%=msg_mid%>"><%=duz(rs("from"))%></a></td>
<td width="40%"><a href="?action=read_msg&mid=<%=rs("mid")%>"><%=duz(rs("subject"))%></a></td>
<td width="40%" align=center><%=rs("mdate")%></td>
<td width="2%" align="center"><a href="?action=msg_del&msg=saved&id=<%=rs("mid")%>"><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/msg_del.gif" border="0" alt="<%=msg_delete_want%>"></a></td>
<%
rs.MoveNext
Loop
%>
</table>
<%
rs.Close
Set rs=Nothing

End Sub
Sub fromadmin

Set rs=Server.CreateObject("ADODB.RecordSet")
SQL="Select * FROM MESSAGES WHERE SEE=True AND TYPE=1 AND MSAVE=False ORDER BY mdate DESC"
rs.open SQL,Connection,3,3
%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" align="center" bgcolor="<%=menu_bg%>" class="td_menu">
<tr style="font-weight:bold"><td width="5%" align="center">&nbsp;</td><td width="20%" bgcolor=<%=background%> align=center><%=msg_sender%></td><td width="40%" bgcolor=<%=background%> align=center><%=msg_subject%></td><td width="40%" bgcolor=<%=background%> align=center><%=the_tarih%></td></tr>
<% Do Until rs.eof %>
<tr class="td_menu2"><td width="5%" align="center"><% If rs("read")=True Then %><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/readed.gif"><% ElseIf rs("read")=False Then %><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/not_readed.gif"><% End If %></td>
<td width="20%" bgcolor=<%=menu_back%> align=center><%=rs("from")%></td>
<td width="40%" bgcolor=<%=menu_back%>><a href="?action=read_msg&mid=<%=rs("mid")%>"><%=rs("subject")%></a></td>
<td width="40%" bgcolor=<%=menu_back%> align=center><%=rs("mdate")%></td>
<%
rs.MoveNext
Loop
%>
</table>
<%
rs.Close
Set rs=Nothing

End Sub
Sub read

id=request.querystring("mid"):id=QS_CLEAR(id,"false")
Set rs=Server.CreateObject("ADODB.RecordSet"):SQL="Select * FROM MESSAGES WHERE mid="&id&"":rs.open SQL,Connection,3,3
IF Session("Member")=rs("TO") Then
protect_msg=True
ElseIF rs("TYPE")="1" Then
protect_msg=True
ELSE
protect_msg=False
End IF

IF protect_msg=True Then
IF Session("Folder")="Inbox" Then
rs("read")=True
rs.Update
End IF
Set frmem=Connection.Execute("Select * FROM MEMBERS WHERE kul_adi='"&rs("from")&"'")

If frmem.Eof Then
msndr="Site Manager"
Else
msndr="<a href=""members.asp?action=member_details&uid="&frmem("uye_id")&""">"&rs("from")&"</a>"
End If
%>
<table border="0" cellspacing="1" cellpadding="2" align="center" width="100%" bgcolor="<%=td_border_color%>" class="td_menu">
<tr><td bgcolor="<%=menu_color%>"><b><%=msg_sender%> : </b><%=msndr%> // <%=rs("mdate")%></td></tr>
<tr><td bgcolor="<%=menu_color%>"><b><%=msg_subject%> : </b><%=rs("subject")%></td></tr>
<tr><td bgcolor="<%=menu_color%>"><%=SmiLey(rs("msg"))%></td></tr>
</table>
<% If rs("type") <> "1" Then %>
<table border="0" cellspacing="0" cellpadding="1" width="100" align="right">
<form method="post" action="?action=new_msg&msgid=<%=rs("mid")%>">
<tr><td><input type="hidden" name="subject" value="<%=rs("subject")%>">
<input type="submit" value="<%=reply_button%>" class="form1"></td>
</form>
<form method="post" action="?action=msg_del&id=<%=rs("mid")%>">
<td>
<input type="submit" value="<%=msg_delete_want%>" class="form1"></td></tr>
</form>
</table>
<%
End IF
End IF
End Sub
Sub new_msg

msg=request.querystring("msgid"):msg=QS_CLEAR(msg,"false")
towho=request.querystring("who"):towho=QS_CLEAR(towho,"")
subje=request.form("subject")
if msg<>"" then
Set mem=Connection.Execute("SELECT * FROM MESSAGES WHERE mid="&msg&"")
msgto=mem("from")
elseif towho<>"" Then
msgto=towho
elseif towho="" OR msg="" Then
msgto=""
end if

If subje<>"" Then subj="Re:"&subje&"" Else subj=""
IF towho<>"" OR msgto<>"" Then msg_opt1="checked" ELSE msg_opt2="checked"
set members=Connection.Execute("select * from MEMBERS ORDER BY kul_adi")%>
<form action="?action=msg_rec" method="post">
<table border="0" cellspacing="1" cellpadding="1" width="100%" align="center" class="td_menu">
<tr><td width="20%" align="right" valign="top"><%=msg_to%> : </td><td width="80%"><input type="radio" name="type" value="manual" <%=msg_opt1%>>&nbsp;<input type="text" name="mem_name" class="form1" value="<%=msgto%>" size="20"><br /><input type="radio" name="type" value="auto" <%=msg_opt2%>>
<select size="1" name="mem_names" class=form1>
<% Do While Not members.EOF
If members("seviye")="1" Then
seviye=" ["&level1&"]"
ElseIf members("seviye")="2" Then
seviye=" ["&level2&"]"
ElseIf members("seviye")="3" Then
seviye=" ["&level3&"]"
ElseIf members("seviye")="4" Then
seviye=" ["&level4&"]"
ElseIf members("seviye")="5" Then
seviye=" ["&level5&"]"
ElseIf members("seviye")="6" Then
seviye=" ["&level6&"]"
ElseIf members("seviye")="8" Then
seviye=" ["&level8&"]"
ELSE
seviye=""
End If
%>
<option value="<%=members("kul_adi")%>"><%=members("kul_adi")%><%=seviye%></option>
<% members.MoveNext
Loop %>
</select>
</td></tr>
<tr>
<td width="30%" align="right"><%=msg_subject%> : </td>
<td width="70%"><input type="text" name="subject" class="form1" value="<%=subj%>" size="40"></td>
</tr>
<tr>
<td width="30%" align="right" valign="top"><%=msg_message%> : </td>
<td width="70%"><textarea name="message" rows="8" cols="40" class="form1"></textarea></td>
</tr>
<tr>
<td width="30%" align="right" valign="top"><%=msg_save%> : </td>
<td width="70%"><input type="checkbox" name="m_save" class="form1" value="ON"></td>
</tr>
<tr>
<td width="30%" align="right">&nbsp;</td>
<td width="70%"><input type="submit" value="<%=msg_ok%>" class="submit"></td>
</tr>
</table>
</form>
<%
End Sub
Sub msg_record

mem_type=Request.Form("type")
subject=Request.Form("subject")
msg=Request.Form("message")

If mem_type="auto" then
mem=Request.Form("mem_names")
ElseIf mem_type="manual" then
mem=Request.Form("mem_name")
Else
Response.Write "<font class=td_menu>"&msg_error2&"</font>"
End If

Set chk_msg_mem=Connection.Execute("Select * FROM MEMBERS WHERE kul_adi='"&mem&"'")
IF chk_msg_mem.Eof Then
Response.Write "<b>" & msg_error3 & "</b>"
ELSE

If subject="" Then subj=""&msg_empty&"" Else subj=subject
IF ucase(Request.Form("m_save"))="ON" Then msg_save=True ELSE msg_save=False

If msg="" Then
Response.Write "<font class=td_menu>"&msg_error1&"</font>"
ELSE

Set enter=Server.CreateObject("ADODB.RecordSet")
entSQL="Select * FROM MESSAGES"
enter.open entSQL,Connection,3,3

enter.AddNew
enter("from")=Session("Member")
enter("to")=mem
enter("msg")=msg
enter("mdate")=Now()
enter("read")=False
enter("subject")=subj
enter("type")=0
enter("see")=True
enter("msave")=msg_save
enter.Update
IF msg_save=True Then success_sent=success_sent2 ELSE success_sent=success_sent
Response.Write "<font class=td_menu><b>"&success_sent&"</b></font>"
END IF
End IF
End Sub
Sub delete_msg

msgid=Request.QueryString("id"):msgid=QS_CLEAR(msgid,"false")
if msgid<>"" Then
Set chk_msg=Connection.Execute("Select * FROM MESSAGES WHERE mid="&msgid&"")
IF Session("Folder")="Saved" Then
Set del_msg=Connection.Execute("UPDATE MESSAGES SET msave=False WHERE mid="&msgid&"")
ELSE
Set del_msg=Connection.Execute("UPDATE MESSAGES SET see=False WHERE mid="&msgid&"")
End IF
else
end if
Response.Redirect Request.ServerVariables("HTTP_REFERER")
del_msg.Close
Set del_msg=Nothing

End Sub
Else
Response.Write WriteMsg(no_entry,100)
End If%>
</td>
</tr>
</table>
<%call ORTA2
call ALT %>