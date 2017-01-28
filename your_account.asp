<!--#include file="view.asp" --><!--#include file="inc/encryption.asp" -->
<%call UST:call ORTA
IF Session("Enter")="Yes" Then%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=ya_topic%></font></td></tr>
<tr> <td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu"><table border="0" cellspacing="10" cellpadding="10" width="100%" class="td_menu" style="font-weight:bold">
<tr align="center" valign="bottom">
<td><a href="?op=Profile"><img src="images/Your_Account/info.gif" border="0"><br /><%=ya_info%></a></td>
<td><a href="Msgbox.asp?action=main"><img src="images/Your_Account/messages.gif" border="0"><br /><%=ya_messages%></a></td>
<td><a href="?op=Friends"><img src="images/Your_Account/friends.gif" border="0"><br /><%=msg_friends%></a></td>
<td><a href="?op=SelectTheme"><img src="images/Your_Account/themes.gif" border="0"><br /><%=ya_theme%></a></td>
<td><a href="?op=Logout"><img src="images/Your_Account/exit.gif" border="0"><br /><%=ya_logout%></a></td>
</tr></table><hr size="1" color="<%=td_border_color%>">

<%IF QS_CLEAR(Request.QueryString("op"),"")="UpdateProfile" Then
email=duz(Request.Form("email"))
isim=duz(Request.Form("isim"))
g_soru=duz(Request.Form("g_soru"))
g_cevap=duz(Request.Form("g_cevap"))
IF g_cevap <> "" Then g_cevap=MN_PP(g_cevap)
icq=duz(Request.Form("icq"))
msn=duz(Request.Form("msn"))
aim=duz(Request.Form("aim"))
sehir=duz(Request.Form("sehir"))
meslek=duz(Request.Form("meslek"))
cinsiyet=duz(Request.Form("cinsiyet"))
yas=duz(Request.Form("yas_1")) & "." & duz(Request.Form("yas_2")) & "." & duz(Request.Form("yas_3"))
url=duz(Request.Form("web"))
signimza=duz(Request.Form("imza"))
mavatar=duz(Request.Form("mavatar"))
IF signimza="" Then signimza="-" ELSE signimza=signimza
If Request.Form("mail_goster")="on" Then mailgoster="True" Else mailgoster="False"
IF email="" then
response.write "<b>" & error3 & "</b>"
ElseIF isim="" then
response.write "<b>" & error4 & "</b>"
ElseIF g_soru="" then
response.write "<b>" & error5 & "</b>"
ElseIF EmailCheck(email)=False Then
Response.Write "<b>" & error13 & "</b>"
ELSE
Connection.Execute("UPDATE MEMBERS SET email='"&email&"' WHERE uye_id="&Session("Uye_ID")&"")
Connection.Execute("UPDATE MEMBERS SET isim='"&isim&"' WHERE uye_id="&Session("Uye_ID")&"")
Connection.Execute("UPDATE MEMBERS SET g_soru='"&g_soru&"' WHERE uye_id="&Session("Uye_ID")&"")
Connection.Execute("UPDATE MEMBERS SET icq='"&icq&"' WHERE uye_id="&Session("Uye_ID")&"")
Connection.Execute("UPDATE MEMBERS SET msn='"&msn&"' WHERE uye_id="&Session("Uye_ID")&"")
Connection.Execute("UPDATE MEMBERS SET aim='"&aim&"' WHERE uye_id="&Session("Uye_ID")&"")
Connection.Execute("UPDATE MEMBERS SET sehir='"&sehir&"' WHERE uye_id="&Session("Uye_ID")&"")
Connection.Execute("UPDATE MEMBERS SET meslek='"&meslek&"' WHERE uye_id="&Session("Uye_ID")&"")
Connection.Execute("UPDATE MEMBERS SET cinsiyet='"&cinsiyet&"' WHERE uye_id="&Session("Uye_ID")&"")
Connection.Execute("UPDATE MEMBERS SET yas='"&yas&"' WHERE uye_id="&Session("Uye_ID")&"")
Connection.Execute("UPDATE MEMBERS SET url='"&url&"' WHERE uye_id="&Session("Uye_ID")&"")
Connection.Execute("UPDATE MEMBERS SET imza='"&signimza&"' WHERE uye_id="&Session("Uye_ID")&"")
Connection.Execute("UPDATE MEMBERS SET mail_goster="&mailgoster&" WHERE uye_id="&Session("Uye_ID")&"")
Connection.Execute("UPDATE MEMBERS SET u_avatar='"&mavatar&"' WHERE uye_id="&Session("Uye_ID")&"")
IF g_cevap<>"" Then Connection.Execute("UPDATE MEMBERS SET g_cevap='"&g_cevap&"' WHERE uye_id="&Session("Uye_ID")&"")
Response.Write cong2
END IF

ElseIF QS_CLEAR(Request.QueryString("op"),"")="Profile" Then
uid=Session("Uye_ID")
Set rs=Server.CreateObject("ADODB.RecordSet"):SQL="Select * FROM MEMBERS WHERE uye_id="&uid&"":rs.open SQL,Connection,3,3
age_check=Split(rs("yas"),".")
if rs("cinsiyet")="b" then section2="selected" else section1="selected"
if rs("mail_goster")=True Then mcheck="checked" else mcheck=""%>
<script language="JavaScript"><!--
function kontrolet(type){
if(type=="email"){var str=document.formUye.email.value;var filtrele=/^(\w+(?:\.\w+)*)@((?:\w+\.)*\w[\w-]{0,66})\.([a-z]{2,6}(?:\.[a-z]{2})?)$/i;if (filtrele.test(str)){document.getElementById('hint_uyelik').innerHTML="";}else{document.getElementById('hint_uyelik').innerHTML="<b><%=error13%></b>";}}
else if(type=="isim"){if(document.formUye.isim.value.length==0){document.getElementById('hint_uyelik').innerHTML="<b><%=error4%></b>";}else{document.getElementById('hint_uyelik').innerHTML="";}}
else if(type=="gsoru"){if(document.formUye.g_soru.value.length==0){document.getElementById('hint_uyelik').innerHTML="<b><%=error5%></b>";}else{document.getElementById('hint_uyelik').innerHTML="";}}
}
function CheckForm () {
var str=document.formUye.email.value;var filtrele=/^(\w+(?:\.\w+)*)@((?:\w+\.)*\w[\w-]{0,66})\.([a-z]{2,6}(?:\.[a-z]{2})?)$/i;
if (document.formUye.email.value==""){document.getElementById('hint_uyelik').innerHTML="<b><%=error3%></b>";return false;}
else if (!filtrele.test(str)){document.getElementById('hint_uyelik').innerHTML="<b><%=error13%></b>";return false;}
else if (document.formUye.isim.value==""){document.getElementById('hint_uyelik').innerHTML="<b><%=error4%></b>";return false;}
else if (document.formUye.g_soru.value==""){document.getElementById('hint_uyelik').innerHTML="<b><%=error5%></b>";return false;}
else {document.formUye.sbmt_uye.disabled=true;return true;}
}// --></script>
<form name="formUye" id="formUye" onSubmit="return CheckForm();" action="?id=<%=uid%>&op=UpdateProfile" method="post"><table width="100%" border="0" cellspacing="2" cellpadding="1" align="center" bgcolor="<%=menu_bg%>" class="td_menu" style="font-weight:bold">
<tr><td align="center"><span id="hint_uyelik"></span><span id="hint_sifre"></span><span id="hint_nick"></span></td></tr>
<tr height="20"><td bgcolor="<%=menu_back%>" width="50%">&nbsp;<%=uye_kuladi%></td><td bgcolor="<%=menu_bg%>" width="50%"><b><%=Session("Member")%></b></td></tr>
<tr height="20"><td bgcolor="<%=menu_back%>" width="50%">&nbsp;<%=uye_password%></td><td bgcolor="<%=menu_bg%>" width="50%"><b>[<a href="?op=ChangePass"><%=pass_change%></a>]</td></tr>
<tr><td bgcolor="<%=menu_back%>" width="50%">&nbsp;<%=register_mail%> *</td><td bgcolor="<%=menu_bg%>" width="50%"><input type="text" name="email" class="form1" value="<%=rs("email")%>" style="width:100%" onBlur="kontrolet('email');"></td></tr>
<tr><td bgcolor="<%=menu_back%>" width="50%">&nbsp;<%=register_name%> *</td><td bgcolor="<%=menu_bg%>" width="50%"><input type="text" name="isim" class="form1" value="<%=rs("isim")%>" style="width:100%" onBlur="kontrolet('isim');"></td></tr>
<tr><td bgcolor="<%=menu_back%>" width="50%">&nbsp;<%=register_question%> *</td><td bgcolor="<%=menu_bg%>" width="50%"><input type="text" name="g_soru" class="form1" value="<%=rs("g_soru")%>" style="width:100%" onBlur="kontrolet('gsoru');"></td></tr>
<tr><td bgcolor="<%=menu_back%>" width="50%">&nbsp;<%=register_answer%><br /><font face="verdana" size="1" color="gray">(<%=gc_information%>)</font></td><td bgcolor="<%=menu_bg%>" width="50%"><input type="text" name="g_cevap" class="form1" style="width:100%"></td></tr>
<tr><td bgcolor="<%=menu_back%>" width="50%">&nbsp;<%=register_icq%></td><td bgcolor="<%=menu_bg%>" width="50%"><input type="text" name="icq" class="form1" value="<%=rs("icq")%>" style="width:100%"></td></tr>
<tr><td bgcolor="<%=menu_back%>" width="50%">&nbsp;<%=register_msn%></td><td bgcolor="<%=menu_bg%>" width="50%"><input type="text" name="msn" class="form1" value="<%=rs("msn")%>" style="width:100%"></td></tr>
<tr><td bgcolor="<%=menu_back%>" width="50%">&nbsp;<%=register_aim%></td><td bgcolor="<%=menu_bg%>" width="50%"><input type="text" name="aim" class="form1" value="<%=rs("aim")%>" style="width:100%"></td></tr>
<tr><td bgcolor="<%=menu_back%>" width="50%">&nbsp;<%=register_city%></td><td bgcolor="<%=menu_bg%>" width="50%"><input type="text" name="sehir" class="form1" value="<%=rs("sehir")%>" style="width:100%"></td></tr>
<tr><td bgcolor="<%=menu_back%>" width="50%">&nbsp;<%=register_job%></td><td bgcolor="<%=menu_bg%>" width="50%"><input type="text" name="meslek" class="form1" value="<%=rs("meslek")%>" style="width:100%"></td></tr>
<tr><td bgcolor="<%=menu_back%>" width="50%">&nbsp;<%=register_sex%></td><td bgcolor="<%=menu_bg%>" width="50%">
<select size="1" class="form1" name="cinsiyet" style="width:100%"><option value="a" <%=section1%>><%=male%></option><option value="b" <%=section2%>><%=female%></option></select></td></tr>
<tr><td bgcolor="<%=menu_back%>" width="50%">&nbsp;Dogum Tarihi</td><td bgcolor="<%=menu_bg%>" width="50%"><select name="yas_1" size="1" class="form1"><%For i_y=1 TO 31:IF Len(i_y)="1" Then i_y="0" & i_y
IF Left(i_y,2)=age_check(0) Then y_opt="selected" ELSE y_opt=""
Response.Write "<option value="""&i_y&""" "&y_opt&">"&i_y&"</option>"
Next%></select>.<select name="yas_2" size="1" class="form1"><%For i_y1=1 TO 12:IF Len(i_y1)="1" Then i_y1="0" & i_y1
IF Left(i_y1,2)=age_check(1) Then y_opt1="selected" ELSE y_opt1=""
Response.Write "<option value="""&i_y1&""" "&y_opt1&">"&i_y1&"</option>"
Next%></select>.<select name="yas_3" size="1" class="form1"><%For i_y2=Year(Now())-100 TO Year(Now())-1
IF Left(i_y2,4)=age_check(2) Then y_opt2="selected" ELSE y_opt2=""
Response.Write "<option value="""&i_y2&""" "&y_opt2&">"&i_y2&"</option>"
Next%></select></td></tr>
<tr><td bgcolor="<%=menu_back%>" width="50%">&nbsp;<%=register_url%></td><td bgcolor="<%=menu_bg%>" width="50%"><input type="text" name="web" class="form1" value="<%=rs("url")%>" style="width:100%"></td></tr>
<tr><td bgcolor="<%=menu_back%>" width="50%" valign="top">&nbsp;<%=register_signature%></td><td bgcolor="<%=menu_bg%>" width="50%"><textarea rows="5" name="imza" class="form1" style="width:100%"><% IF rs("imza") <> "" Then %><%=Oku(rs("imza"))%><% End IF %></textarea></td></tr>
<tr><td width="50%" valign="top" bgcolor="<%=menu_back%>">&nbsp;<%=register_avatar%></td><td width="50%" bgcolor="<%=menu_bg%>"><select name="mavatar" size="8" class="form1" onChange="(avatar.src=mavatar.options[mavatar.selectedIndex].value)"><option value="<%=rs("u_avatar")%>" selected>Seçiniz</option>
<!--#include file="INC/avatar_list.asp" --></select>&nbsp;<img src="<%=rs("u_avatar")%>" name="avatar" align="absbottom" style="border-style: solid; border-width: 1"></td></tr>
<tr><td bgcolor="<%=menu_back%>" width="50%">&nbsp;<%=register_mailgoster%></td><td bgcolor="<%=menu_bg%>" width="50%"><input type="checkbox" name="mail_goster" class="form1" <%=mcheck%>></td></tr>
<tr><td colspan="2" bgcolor="<%=menu_bg%>"><input type="submit" value="<%=uye_guncelle%>" class="submit" style="width:100%" name="sbmt_uye"></td></tr>
</table></form>

<%ElseIF QS_CLEAR(Request.QueryString("op"),"")="UpdatePass" Then
pold=duz(Request.Form("opass"))
pnew=duz(Request.Form("npass"))
pnew=MN_PP(pnew)
pconfirm=duz(Request.Form("cpass"))
Set mInfo=Connection.Execute("Select * FROM MEMBERS WHERE uye_id="&Session("Uye_ID")&"")
IF pold="" OR pnew="" OR pconfirm="" Then
Response.Write "<font class=td_menu><b>"&pass_err1&"</b></font><br /><br /><a href=javascript:history.back();><< "&previous_page&"</a>"
ElseIF MN_PP(pold) <> mInfo("sifre") Then
Response.Write "<font class=td_menu><b>"&pass_err2&"</b></font><br /><br /><a href=javascript:history.back();><< "&previous_page&"</a>"
ElseIF pnew <> MN_PP(pconfirm) Then
Response.Write "<font class=td_menu><b>"&pass_err3&"</b></font><br /><br /><a href=javascript:history.back();><< "&previous_page&"</a>"
ELSE
Connection.Execute("UPDATE MEMBERS SET sifre='"&pnew&"' WHERE uye_id="&Session("Uye_ID")&"")
Response.Write "<b>" & succes_saved & "</b>"
End IF

ElseIF QS_CLEAR(Request.QueryString("op"),"")="ChangePass" Then%>
<table border="0" cellspacing="1" cellpadding="1" width="90%" class="td_menu">
<form action="?op=UpdatePass" method="post">
<tr>
<td width="40%" align="right"><%=old_pass%> : </td>
<td width="60%"><input type="text" name="opass" size="30" class="form1"></td>
</tr>
<tr>
<td width="40%" align="right"><%=new_pass%> : </td>
<td width="60%"><input type="text" name="npass" size="30" class="form1"></td>
</tr>
<tr>
<td width="40%" align="right"><%=confirm_pass%> : </td>
<td width="60%"><input type="text" name="cpass" size="30" class="form1"></td>
</tr>
<tr>
<td width="40%" align="right">&nbsp;</td>
<td width="60%"><input type="submit" value="<%=uye_guncelle%>" class="submit"></td>
</tr>
</form>
</table>
<%
ElseIF QS_CLEAR(Request.QueryString("op"),"")="RegTheme" Then
Connection.Execute("UPDATE MEMBERS SET u_theme='"&lcase(Request.Form("theme"))&"' WHERE uye_id="&Session("Uye_ID")&"")
Session("Theme")=Request.Form("theme")
Response.Redirect "Your_Account.asp?op=SelectTheme"

ElseIF QS_CLEAR(Request.QueryString("op"),"")="SelectTheme" Then%>
<form action="?op=RegTheme" method="post">
<%=ya_selecttheme%> : <select name="theme" size="1" class="form1">
<%Set themes=Server.CreateObject("ADODB.RecordSet"):themes.open "Select * FROM THEMES ORDER BY t_name ASC",Connection,3,3:Do Until themes.EoF
IF Session("SiteTheme")=themes("t_dir") Then opt="selected" ELSE opt=""
Response.Write "<option value="""&themes("t_dir")&""" "&opt&">"&themes("t_name")&"</option>"
themes.MoveNext
Loop%>
</select><br />
<input type="submit" value="<%=uye_guncelle%>" class="submit">
</form>

<%ElseIF QS_CLEAR(Request.QueryString("op"),"")="Logout" Then
Set up=Connection.Execute("UPDATE MEMBERS SET u_online=False WHERE uye_id="&Session("Uye_ID")&"")
Session("Enter")=""
Session("Uye_ID")=""
Session("Member")=""
Session("Name")=""
Session("Email")=""
Session("Level")=""
Session("ISMod")=""
Session("EditPost")=""
Session("Signature")=""
Session("Theme")=""
Session("LastLogged")=""
Response.Cookies("MiniNuke_Gir")("girdi")=""
Response.Cookies("MiniNuke_Gir")("kullanici_id")=""
Response.Cookies("MiniNuke_Gir")("kullanici")=""
Response.Cookies("MiniNuke_Gir")("kullanici_mail")=""
Response.Redirect ("default.asp")
ElseIF QS_CLEAR(Request.QueryString("op"),"")="Friends" Then
Set fRs=Server.CreateObject("ADODB.RecordSet"):fRsSQL="Select * FROM FRIENDS WHERE MEMBER="&Session("Uye_ID")&"":fRs.open frsSQL,Connection,3,3%>
<form action="Msgbox.asp?action=fact" method="post">
<select name="fr_nm" size="10" class="form1" style="width:150">
<%
IF fRs.Eof Then
%>
<option value=""><%=no_friend%></option>
<%
ELSE
Do Until fRs.Eof
Set fr=Connection.Execute("Select * FROM MEMBERS WHERE uye_id="&frs("FRIEND")&"")
IF fr.Eof OR fr.Bof Then
mname="###"
ELSe
mname=fr("kul_adi")
End IF
%>
<option value="<%=mname%>"><%=mname%></option>
<%
frs.MoveNext
Loop
End If
%>
</select>
<br /><br />
<input type="submit" name="fr_act" class="submit" value="<%=f_info%>">
<input type="submit" name="fr_act" class="submit" value="<%=f_msg%>">
<input type="submit" name="fr_act" class="submit" value="<%=f_delf%>">
</form>
<%
	ELSE

Set mem_avatar=Connection.Execute("Select * FROM MEMBERS WHERE uye_id="&Session("Uye_ID")&"")
Set mem_l10n=Server.CreateObject("ADODB.RecordSet")
mem_l10n.open "Select * FROM NEWS WHERE ekleyen='"&Session("Member")&"' AND onay=True ORDER BY tarih DESC",Connection,3,3
Set urMsgs=Connection.Execute("Select Count(*) as count FROM MESSAGES WHERE TYPE=0 AND SEE=True AND READ=False AND TO='"&Session("Member")&"'")
Set TM=Connection.Execute("Select Count(*) as mcount FROM MESSAGES WHERE TYPE=0 AND SEE=True AND TO='"&Session("Member")&"'")

percent=Int((TM("mcount") / sett("msg_limit")) * 100)
If percent >= 100 Then
percent="100"
Else
percent=percent
End If
If urMsgs("count") > sett("msg_limit") Then
urMsgs_COUNT=sett("msg_limit")
Else
urMsgs_COUNT=urMsgs("count")
End If
%>
<table border="0" cellspacing="2" cellpadding="0" align="center" width="90%" class="td_menu">
<tr>
<td valign="top" width="10">
<%
IF mem_avatar("u_avatar")="" Then
Response.Write ""
ELSE
Response.Write "<img src="""&mem_avatar("u_avatar")&""" border=""1"">"
End IF
%>
</td>
<td width="60%" valign="top">
<br /><br /><br />- <%=msg_welcome%> <b><%=Session("Member")%></b><br />
<%
If urMsgs_COUNT="0" Then
Response.Write msg_no_unreaed_msgs
ELSE
%>
   <b><%=urMsgs_COUNT%></b> <%=msg_unreaded_msgs%>
<% End If %>
</td>
<td width="40%" align="left" valign="top">
<br /><br /><%=msg_es%> : <b><%=100-percent%>%</b><br />
<%=msg_us%> : <b><%=percent%>%</b>
<table border="0" cellspacing="1" cellpadding="1" width="115" bgcolor="<%=td_border_color%>">
<tr bgcolor="<%=menu_bg%>">
<td>
<img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/poll_result.gif" width="<%=percent%>" height="14">
</td>
</tr>
</table>
</td>
</tr>
</table>

<%
Response.Write "</div>"
mem_avatar.Close
mem_l10n.Close
Set mem_avatar=Nothing
Set mem_l10n=Nothing
	End IF
%>
</td>
</tr>
</table>
<%
ELSE
Response.Write WriteMsg(no_entry,100)
End IF
call ORTA2
call ALT
%>