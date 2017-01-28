<!--#include file="view.asp" --><!--#include file="inc/encryption.asp" -->
<% call UST:call ORTA
action=Request.QueryString("action"):action=QS_CLEAR(action,"")
if action="new" then
call mem_new
elseif action="register" then
call reg
elseif action="lostpass" then
call lostpassword
end if

Sub reg
kuladi=duz(Request.Form("kuladi"))
password=duz(Request.Form("password")):password=MN_PP(password)
password2=duz(Request.Form("password2")):password2=MN_PP(password2)
email=duz(Request.Form("email"))
isim=duz(Request.Form("isim"))
g_soru=duz(Request.Form("g_soru"))
g_cevap=duz(Request.Form("g_cevap"))
g_cevap=MN_PP(g_cevap)
icq=duz(Request.Form("icq"))
msn=duz(Request.Form("msn"))
aim=duz(Request.Form("aim"))
sehir=duz(Request.Form("sehir"))
meslek=duz(Request.Form("meslek"))
cinsiyet=duz(Request.Form("cinsiyet"))
yas=duz(Request.Form("yas_1")) & "." & duz(Request.Form("yas_2")) & "." & duz(Request.Form("yas_3"))
url=duz(Request.Form("web"))
sc=duz(Request.Form("security_code"))
imza=duz(Request.Form("imza"))
IF imza="" Then imza="-" ELSE imza=imza
If icq="" Then icq=0 Else icq=icq
If yas="" Then yas="0" Else yas=yas
Set nameCheck=Connection.Execute("Select * From MEMBERS where UCase(kul_adi)='"&UCase(kuladi)&"'")
Set mailCheck=Connection.Execute("Select * From MEMBERS where email='"&email&"'")
IF InStr(1,kuladi,"ð",1) OR InStr(1,kuladi,"þ",1) OR InStr(1,kuladi,"ö",1) Then chk_kul_adi="False" ELSE chk_kul_adi="True"
If EmailCheck(email) Then emailC=True ELSE emailC=False
IF ucase(sc) <> ucase(Session("RegSecurityCode")) Then s_code=False ELSE s_code=True
If Request.Form("mail_goster")="on" Then mail_goster=True Else mail_goster=False
if Session("ImagePath")="" then
if Request.Form("mavatar")="" then avatar="images/avatars/58.jpg" else avatar=Request.Form("mavatar")
else
avatar=Session("ImagePath")
end if
if kuladi="" then
response.write WriteMsg(error1,100)
elseif len(password)<4 then
response.write WriteMsg(error2,100)
elseif password<>password2 then
response.write WriteMsg(error2_1,100)
elseif email="" then
response.write WriteMsg(error3,100)
elseif isim="" then
response.write WriteMsg(error4,100)
elseif g_soru="" then
response.write WriteMsg(error5,100)
elseif g_cevap="" then
response.write WriteMsg(error6,100)
elseif sc="" Then
Response.Write WriteMsg(sc_err1,100)
elseif s_code=False Then
Response.Write WriteMsg(sc_err2,100)
elseif NOT mailCheck.eof Then
response.write WriteMsg(error10,100)
elseif NOT nameCheck.eof Then
response.write WriteMsg(error9,100)
elseif emailC=False Then
response.write WriteMsg(error13,100)
elseif Len(kuladi) > 15 Then
response.write WriteMsg(error15,100)
elseif chk_kul_adi="False" Then
response.write WriteMsg(error16,100)
ELSE
Set rec=Server.CreateObjecT("ADODB.RecordSet"):rSQL="Select * FROM MEMBERS":rec.open rSQL,Connection,3,3
Session.LCID=1033
rec.AddNew
rec("kul_adi")=kuladi
rec("sifre")=password
rec("email")=email
rec("isim")=isim
rec("g_soru")=g_soru
rec("g_cevap")=g_cevap
rec("icq")=icq
rec("msn")=msn
rec("aim")=aim
rec("sehir")=sehir
rec("meslek")=meslek
rec("cinsiyet")=cinsiyet
rec("yas")=yas
rec("url")=url
rec("imza")=imza
rec("mail_goster")=mail_goster
rec("login_sayisi")=0
rec("uyelik_tarihi")=Date()
rec("son_tarih")=Now()
rec("seviye")=0
rec("msg_sayisi")=0
rec("u_theme")="Fantasy"
rec("u_avatar")=avatar
rec.Update
Response.Write WriteMsg(cong,100)
END IF
End Sub

Sub mem_new
reg_sec_code=CPASS(4):Session("RegSecurityCode")=reg_sec_code%>
<script language="JavaScript"><!--
var filtrele=/^(\w+(?:\.\w+)*)@((?:\w+\.)*\w[\w-]{0,66})\.([a-z]{2,6}(?:\.[a-z]{2})?)$/i;
function kontrolet(type){var str=document.formUye.email.value;
if(type=="kuladi"){if(document.formUye.kuladi.value.length==0){document.getElementById('hint_nick').innerHTML="<b><%=error1%></b><br />";}else if(document.formUye.kuladi.value.length==0){document.getElementById('hint_nick').innerHTML="<b><%=error15%></b><br />";}}
else if(type=="parola"){if(document.formUye.password.value.length<4){document.getElementById('hint_sifre').innerHTML="<b><%=error2%></b><br />";}}
else if(type=="parola2"){if(document.formUye.password.value!=document.formUye.password2.value){document.getElementById('hint_sifre').innerHTML="<b><%=error2_1%></b><br />";}}
else if(type=="email"){if (filtrele.test(str)){document.getElementById('hint_ml').innerHTML="";}else{document.getElementById('hint_ml').innerHTML="<b><%=error13%></b><br />";}}
else if(type=="isim"){if(document.formUye.isim.value.length==0){document.getElementById('hint_sm').innerHTML="<b><%=error4%></b><br />";}}
else if(type=="gsoru"){if(document.formUye.g_soru.value.length==0){document.getElementById('hint_sr').innerHTML="<b><%=error5%></b><br />";}}
else if(type=="gcevap"){if(document.formUye.g_cevap.value.length==0){document.getElementById('hint_cvp').innerHTML="<b><%=error6%></b><br />";}}
else if(type=="security"){if(document.formUye.security_code.value.length==0){document.getElementById('hint_sc').innerHTML="<b><%=sc_err1%></b><br />";}else if(document.formUye.security_code.value!=<%=reg_sec_code%>){document.getElementById('hint_sc').innerHTML="<b><%=sc_err1%></b><br />";}}
}
function CheckForm(){var str=document.formUye.email.value;
if (document.formUye.kuladi.value==""){document.getElementById('hint_nick').innerHTML="<b><%=error1%></b><br />";return false;}
else if (document.formUye.password.value.length<4){document.getElementById('hint_sifre').innerHTML="<b><%=error2%></b><br />";return false;}
else if (document.formUye.password.value!=document.formUye.password2.value){document.getElementById('hint_sifre').innerHTML="<b><%=error2_1%></b><br />";return false;}
else if (document.formUye.email.value==""){document.getElementById('hint_ml').innerHTML="<b><%=error3%></b><br />";return false;}
else if (!filtrele.test(str)){document.getElementById('hint_ml').innerHTML="<b><%=error13%></b><br />";return false;}
else if (document.formUye.isim.value==""){document.getElementById('hint_sm').innerHTML="<b><%=error4%></b><br />";return false;}
else if (document.formUye.g_soru.value==""){document.getElementById('hint_sr').innerHTML="<b><%=error5%></b><br />";return false;}
else if (document.formUye.g_cevap.value==""){document.getElementById('hint_cvp').innerHTML="<b><%=error6%></b><br />";return false;}
else if (document.formUye.security_code.value==""){document.getElementById('hint_sc').innerHTML="<b><%=sc_err1%></b><br />";return false;}
else if(document.formUye.security_code.value!=<%=reg_sec_code%>){document.getElementById('hint_sc').innerHTML="<b><%=sc_err2%></b><br />";}
else {document.formUye.sbmt_uye.disabled=true;return true;}
}
// --></script><script src="inc/ajax_username.js"></script> 
<form name="formUye" id="formUye" onSubmit="return CheckForm();" action="membership.asp?action=register" method="post"><table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=m_title%></font></td></tr>
<tr><td align="center"><span id="hint_nick"></span><span id="hint_sifre"></span><span id="hint_ml"></span><span id="hint_sm"></span><span id="hint_sr"></span><span id="hint_cvp"></span><span id="hint_sc"></span></td></tr>
<tr><td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu"><table width="100%" border="0" cellspacing="2" cellpadding="2" align="center" class="td_menu">
<tr bgcolor="<%=forum_bg_1%>"><td width="40%">&nbsp;<%=uye_kuladi%> *</td><td><input type="text" name="kuladi" class="form1" style="width:100%" onBlur="kontrolet('kuladi');check_username(this.value)" onKeyPress="document.getElementById('hint_nick').innerHTML=''"></td></tr>
<tr bgcolor="<%=forum_bg_2%>"><td>&nbsp;<%=uye_password%> *</td><td><input type="password" name="password" class="form1"  style="width:100%" onBlur="kontrolet('parola');" onKeyPress="document.getElementById('hint_sifre').innerHTML=''"></td></tr>
<tr bgcolor="<%=forum_bg_2%>"><td>&nbsp;<%=uye_password2%> *</td><td><input type="password" name="password2" class="form1"  style="width:100%" onBlur="kontrolet('parola2');" onKeyPress="document.getElementById('hint_sifre').innerHTML=''"></td></tr>
<tr bgcolor="<%=forum_bg_1%>"><td>&nbsp;<%=register_mail%> *</td><td><input type="text" name="email" class="form1"  style="width:100%" onBlur="kontrolet('email');" onKeyPress="document.getElementById('hint_ml').innerHTML=''"></td></tr>
<tr bgcolor="<%=forum_bg_2%>"><td>&nbsp;<%=register_name%> *</td><td><input type="text" name="isim" class="form1" style="width:100%" onBlur="kontrolet('isim');" onKeyPress="document.getElementById('hint_sm').innerHTML=''"></td></tr>
<tr bgcolor="<%=forum_bg_1%>"><td>&nbsp;<%=register_question%> *</td><td><input type="text" name="g_soru" class="form1"  style="width:100%" onBlur="kontrolet('gsoru');" onKeyPress="document.getElementById('hint_sr').innerHTML=''"></td></tr>
<tr bgcolor="<%=forum_bg_2%>"><td>&nbsp;<%=register_answer%> *</td><td><input type="text" name="g_cevap" class="form1"  style="width:100%" onBlur="kontrolet('gcevap');" onKeyPress="document.getElementById('hint_cvp').innerHTML=''"></td></tr>
<tr bgcolor="<%=forum_bg_1%>"><td>&nbsp;<%=register_icq%></td><td><input type="text" name="icq" class="form1"  style="width:100%"></td></tr>
<tr bgcolor="<%=forum_bg_2%>"><td>&nbsp;<%=register_msn%></td><td><input type="text" name="msn" class="form1"  style="width:100%"></td></tr>
<tr bgcolor="<%=forum_bg_1%>"><td>&nbsp;<%=register_aim%></td><td><input type="text" name="aim" class="form1"  style="width:100%"></td></tr>
<tr bgcolor="<%=forum_bg_2%>"><td>&nbsp;<%=register_city%></td><td><input type="text" name="sehir" class="form1"  style="width:100%"></td></tr>
<tr bgcolor="<%=forum_bg_1%>"><td>&nbsp;<%=register_age%></td><td><select name="yas_1" size="1" class="form1">
<%For i_y=1 TO 31:Response.Write "<option value="""&i_y&""">"&i_y&"</option>":Next%></select>.<select name="yas_2" size="1" class="form1">
<%For i_y1=1 TO 12:Response.Write "<option value="""&i_y1&""">"&i_y1&"</option>":Next%></select>.<select name="yas_3" size="1" class="form1">
<%For i_y2=1920 TO Year(Now())-1:Response.Write "<option value="""&i_y2&""">"&i_y2&"</option>":Next%></select></td></tr>
<tr bgcolor="<%=forum_bg_2%>"><td>&nbsp;<%=register_sex%></td><td><select size="1" class="form1" name="cinsiyet"  style="width:100%"><option value="a"><%=male%></option><option value="b"><%=female%></option></select></td></tr>
<tr bgcolor="<%=forum_bg_1%>"><td>&nbsp;<%=register_job%></td><td><input type="text" name="meslek" class="form1"  style="width:100%"></td></tr>
<tr bgcolor="<%=forum_bg_2%>"><td>&nbsp;<%=register_url%></td><td><input type="text" name="web" class="form1" value="http://"  style="width:100%"></td></tr>
<tr bgcolor="<%=forum_bg_1%>"><td>&nbsp;<%=register_signature%></td><td><textarea name="imza" rows="3"  style="width:100%" class="form1"></textarea></td></tr>
<tr bgcolor="<%=forum_bg_2%>"><td><input type="radio" name="AvatarOption" value="1" onClick="document.getElementById('divAvatar2').style.display='none';document.getElementById('divAvatar1').style.display='';" checked="checked">Avatar Seç<br />
<input type="radio" name="AvatarOption" value="2" onClick="document.getElementById('divAvatar1').style.display='none';document.getElementById('divAvatar2').style.display='';">Avatar Yükle</td>
<td><div id="divAvatar1" style="height:110;"><select name="mavatar" size="7" class="form1" onChange="(avatar.src=mavatar.options[mavatar.selectedIndex].value)"><!--#include file="INC/avatar_list.asp" --></select><img src="IMAGES/avatars/58.jpg" width="86" height="100" name="avatar"></div>
<div id="divAvatar2" style="display:none;height:110;"><iframe src="administrator/news.asp?action=IS&page=new" allowtransparency="true" width="300" height="100" align="center" border="0" MARGINWIDTH=0 MARGINHEIGHT=0 HSPACE=0 VSPACE=0 FRAMEBORDER=0 SCROLLING=NO></iframe></div>
</td></tr>
<tr bgcolor="<%=forum_bg_1%>"><td>&nbsp;<%=register_mailgoster%></td><td><input type="checkbox" name="mail_goster"></td></tr>
<tr bgcolor="<%=forum_bg_2%>"><td>&nbsp;<%=security_code%></td><td><b><%=reg_sec_code%></b></td></tr>
<tr bgcolor="<%=forum_bg_1%>"><td>&nbsp;<%=security_code_type%> *</td><td><input type="text" name="security_code" size="20" class="form1" onBlur="kontrolet('security');" onChange="document.getElementById('hint_sc').innerHTML=''"></td></tr>
<tr bgcolor="<%=forum_bg_2%>"><td colspan="2" align="center"><input type="submit" value="<%=uye_kayit%>" class="submit"  style="width:100%" name="sbmt_uye"></td></tr>
</table></td></tr>
</table></form><script language="javascript">document.formUye.kuladi.focus();</script>
<%End Sub

Sub lostpassword
step=Request.QueryString("step"):step=QS_CLEAR(step,"false")%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=uye_lostpwd%></font></td></tr>
<tr><td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu"><%if step="1" Then %>
<form action="membership.asp?action=lostpass&step=2" method="post"><table border="0" cellspacing="0" cellpadding="0" align="center" width="70%">
<tr><td width="50%" align="right"><font face=verdana size=1><%=lost_memname%>:</font></td><td width="50%"><input type="text" name="memname" size="30" class="form1"></td></tr>
<tr><td width="50%">&nbsp;</font></td><td width="50%"><input type="submit" value="<%=lost_continue%>" class="submit"></td></tr>
</table></form>
<%ElseIf step="2" Then
member=Request.Form("memname")
IF member="" Then Response.Write error1 Else Set uyectrl=Connection.Execute("Select * FROM MEMBERS WHERE kul_adi='"&member&"'")
IF member<>"" Then
IF uyectrl.eof Then
Response.Write error12
ELSE%>
<form action="membership.asp?action=lostpass&step=3&memname=<%=uyectrl("kul_adi")%>" method="post"><table border="0" cellspacing="0" cellpadding="0" align="center" width="70%" class="td_menu">
<tr><td colspan="2" align="center"><%=uyectrl("g_soru")%></td></tr>
<tr><td width="50%" align="right"><font face=verdana size=1><%=register_answer%>:</font></td><td width="50%"><input type="text" name="answer" size="30" class="form1"></td></tr>
<tr><td width="50%">&nbsp;</font></td><td width="50%" align="right"><input type="submit" value="<%=lost_continue%>" class="submit"></td></tr>
</table></form>
<%End If
End If
ElseIf step="3" Then
answer=Request.Form("answer")
member=Request.QueryString("memname"):member=QS_CLEAR(member,"")
If answer="" Then
Response.Write error6
Else
Set answerctrl=Connection.Execute("Select * FROM MEMBERS WHERE kul_adi='"&member&"'")
IF answerctrl("g_cevap") <> MN_PP(answer) Then
response.write error11
Else
IF sett("fpass")="Site" Then
asd=answerctrl("sifre")
mem_new_pass=CPASS(4)
Connection.Execute("UPDATE MEMBERS SET sifre='"&MN_PP(mem_new_pass)&"' WHERE kul_adi='"&member&"'")%>
<table border="0" cellspacing="1" cellpadding="0" align="center" width="75%" class="td_menu">
<tr><td colspan="2" align="center"><%=member%></td></tr>
<tr><td width="50%" align="right"><%=uye_new_password%>:</td><td width="50%"><b><%=mem_new_pass%></b></td></tr>
</table>
<%ElseIF sett("fpass")="Mail" Then
Set objCDO=Server.CreateObject("CDONTS.NewMail")
objCDO.To=answerctrl("email")
objCDO.From=sett("site_mail")
objCDO.Subject=""&sett("site_isim")&" Þifre Hatýrlatma"
objCDO.BodyFormat=0
objCDO.MailFormat=0
govde=govde &" <font face=""Tahoma"" style=""font-size:11px;""> "& chr(10)
govde=govde &" Sayin <b>"&answerctrl("kul_adi")&"</b>; "& chr(10)
govde=govde &" <br /> "& chr(10)
govde=govde &" <b>"&sett("site_isim")&"</b> isimli sitedeki Sifre Hatirlatma servisi kullanilarak sifreniz mailinize iletilmistir.Mailin sizinle bir alakasi olmadigini düsünüyorsaniz bu maili siliniz... "& chr(10)
govde=govde &" <br /><br /> "& chr(10)
govde=govde &" Yeni Sifreniz : <b>"&mem_new_pass&"</b> "& chr(10)
govde=govde &" <br /><br /> "& chr(10)
govde=govde &" <b><a href="""&sett("site_adres")&""">"&sett("site_isim")&"</a></b> (Powered by <a href=""http://www.mini-nuke.com"">Mini-NUKE</a>) "& chr(10)
govde=govde &"</font>"& chr(10)
objCDO.Body=govde
objCDO.Importance=2
objCDO.Send
Set objCDO=Nothing
Response.Write fp_msg
End IF
End If
End If
End If%>
<br /><br /><div align=center class=td_menu2><a href="javascript:history.back();"><%=p_back%></div></td>
</tr>
</table>
<%End Sub

call ORTA2
call ALT %>
