<% Response.Buffer = True %>
<!--#include file="includes.asp" -->
<script language="Javascript1.2"><!-- // load htmlarea
_editor_url = "../htmlarea/";                     // URL to htmlarea files
var win_ie_ver = parseFloat(navigator.appVersion.split("MSIE")[1]);
if (navigator.userAgent.indexOf('Mac')        >= 0) { win_ie_ver = 0; }
if (navigator.userAgent.indexOf('Windows CE') >= 0) { win_ie_ver = 0; }
if (navigator.userAgent.indexOf('Opera')      >= 0) { win_ie_ver = 0; }
if (win_ie_ver >= 5.5) {
 document.write('<scr' + 'ipt src="' +_editor_url+ 'editor.js"');
 document.write(' language="Javascript1.2"></scr' + 'ipt>');  
} else { document.write('<scr'+'ipt>function editor_generate() { return false; }</scr'+'ipt>'); }
// --></script>
<%
If Session("Level") = "1" OR Session("Level") = "7" Then
act = Request.QueryString("action")
If act = "menu" Then
call menu
ElseIf act = "MsgToMembers" Then
call msgmembers
ElseIf act = "SendMsg" Then
call sendmsg
ElseIf act = "members" Then
call members
ElseIf act = "member_details" Then
call mem_details
ElseIf act = "member_update" Then
call mem_update
ElseIf act = "member_delete" Then
call mem_del
ElseIf act = "AdminMsgs" Then
call adminmsgs
ElseIf act = "AdminMsgs_Del" Then
call deleteadminmsgs
ElseIf act = "MemberMsgs" Then
call membermsgs
ElseIf act = "DelMemMsg" Then
call delete_memmsg
ElseIf act = "MemMsgRead" Then
call read_memmsg
ElseIF act = "Mail2Members" Then
call mail_2_members
ElseIF act = "SendMail2Members" Then
call mail_send_2_members
ElseIF act = "DeleteMessages" Then
call del_all_msgs
End If

Sub del_all_msgs
IF Request.QueryString("x") = "OK" Then
Connection.Execute("DELETE FROM MESSAGES")
Response.Write "<div align=center class=td_menu><b>Tebrikler !!!</b><br>Tüm Mesajlar Baþarýyla Silindi...</div>"
ELSE
Response.Write "<div align=center class=td_menu><b>DÝKKAT !!!</b><br><br>Tamam'a bastýðýnýz andan itibaret Toplu Mesajlar ve Üyeler Arasýndaki mesajlarýn hepsi silinecektir...<br><br><input type=""button"" value=""Tamam"" class=""submit"" onClick=""location.href('?action="&act&"&x=OK')""> <input type=""button"" value=""Vazgeç"" class=""submit"" onClick=""location.href('default.asp?pane=Main')"">"
End IF
End Sub

Sub mail_2_members
%>
<script language="JavaScript1.2" defer>
editor_generate('m_msg');
</script>
<table border="0" cellspacing="2" cellpadding="2" align="center" width="90%" align="center" class="td_menu">
<form method="post" action="?action=SendMail2Members">
<tr bgcolor="#EEEEEE">
<td width="20%" align="right"><b>Baþlýk : </b></td>
<td width="80%"><input type="text" name="m_title" class="form1" style="width:100%"></td>
</tr>
<tr bgcolor="#EEEEEE">
<td colspan="2"><textarea name="m_msg" rows="20" class="form1" style="width:100%"></textarea></td>
</tr>
<tr align="center">
<td colspan="2"><input type="submit" value="Gönder" class="submit" style="width:100%"></td>
</tr>
</form>
</table>
<%
End SUB

Sub mail_send_2_members
IF Request.Form("m_msg") = "" Then
Response.Write "<div align=center class=td_menu><b>Lütfen Mesajýnýzý Yazýnýz</b></div>"
ELSE
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM MEMBERS ORDER BY uye_id DESC",Connection,3,3
Set objSettings = Server.CreateObject("ADODB.RecordSet")
objSettings.open "Select * FROM SETTINGS",Connection,3,3

Do Until rs.EoF

Set objCDO = Server.CreateObject("CDONTS.NewMail")
objCDO.To = rs("email")
objCDO.From = objSettings("site_isim")
objCDO.Subject = Request.Form("m_title")
objCDO.BodyFormat=0
objCDO.MailFormat=0

govde = govde &" <font face=""Tahoma"" style=""font-size:11px;""> "& chr(10)
govde = govde &" "&Request.Form("m_msg")&" "& chr(10)
govde = govde &" </font> "& chr(10)

objCDO.Body = govde
objCDO.Importance = 2
objCDO.Send
Set objCDO = Nothing

rs.MoveNext
Loop
Response.Write "<div align=center class=td_menu><b>TEBRÝKLER !</b><br><br>Mesajýnýz Tüm Üyelerin Mail Adreslerine Gönderildi.</div>"
End IF
End SUB

Sub menu
End Sub
Sub members
IF Len(Request.QueryString("y")) < "1" Then
arrange_word = "A"
ELSE
arrange_word = Request.QueryString("y")
End IF
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM MEMBERS WHERE Left(kul_adi,1) LIKE '%"&arrange_word&"%' ORDER BY kul_adi ASC",Connection,3,3
%>
<table border="0" cellspacing="0" cellpadding="1" width="80%" class="td_menu" style="font-weight:bold" style="<%=TableShadow%>" align="center">
<tr bgcolor="#EEEEEE">
<td align="center">
<%
	For b = 65 To 90 Step 1
	IF b <> 65 Then
	Response.Write "&nbsp;-&nbsp;"
	End IF
	Response.Write "<a href='?action=members&y=" & Chr(b) &"'>" & Chr(b) &"</a>"
	Next
%>
</td>
</tr>
</table>
<hr size="1" style="width:80%">
<table width="80%" align="center">
<%
While Not rs.EOF
%>
  <tr>
    <td width="50%" height="15" bgcolor="#F4F4F4" valign="top"><font face="tahoma" style="font-size:11px"><a href="?action=member_details&mid=<%=rs("uye_id")%>"><b><%=rs("kul_adi")%></b></a> [<a href="?action=member_delete&id=<%=rs("uye_id")%>">Sil</a>]</font></td>
    <%
rs.MoveNext
If Not rs.EOF Then
%>
    <td width="50%" height="15" bgcolor="#F4F4F4" valign="top"><font face="tahoma" style="font-size:11px"><a href="?action=member_details&mid=<%=rs("uye_id")%>"><b><%=rs("kul_adi")%></b></a> [<a href="?action=member_delete&id=<%=rs("uye_id")%>">Sil</a>]</font></td>
<%
End If

If Not rs.EOF Then
rs.MoveNext
End If
%>
</tr>
<% Wend %>
</table>
<%
End Sub
Sub mem_details

uid = Request.QueryString("mid")
Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM MEMBERS WHERE uye_id = "&uid&""
rs.open SQL,Connection,1,3

uye_gcevap = rs("g_cevap")
age_check = Split(rs("yas"),".")
%>
<form action="?action=member_update&mid=<%=uid%>" method="post">
<table width="100%" border="0" cellspacing="1" cellpadding="0" align="center" bgcolor="#999999">
<tr>
<td colspan="2" bgcolor="#EEEEEE" align="center"><font face=arial size=2><b><%=rs("kul_adi")%></b></font></td></tr>
  <tr> 
    <td bgcolor="#FFFFFF" width="50%">&nbsp;<font face="verdana" size="1">Seviye : </font></td><td bgcolor="#FFFFFF" width="50%">
<%
If rs("seviye") = "0" Then
sel = "selected"
ElseIf rs("seviye") = "1" Then
sel1 = "selected"
ElseIf rs("seviye") = "2" Then
sel2 = "selected"
ElseIf rs("seviye") = "3" Then
sel3 = "selected"
ElseIf rs("seviye") = "4" Then
sel4 = "selected"
ElseIf rs("seviye") = "5" Then
sel5 = "selected"
ElseIf rs("seviye") = "8" Then
sel8 = "selected"
End If
%>
<select size="1" name="seviye" class="form1">
<option value="0" <%=sel%>>Üye</option>
<option value="1" <%=sel1%>>Yönetici</option>
<option value="2" <%=sel2%>>Download Editörü</option>
<option value="3" <%=sel3%>>Makale Editörü</option>
<option value="4" <%=sel4%>>Haber Editörü</option>
<option value="5" <%=sel5%>>Forum Görevlisi</option>
<option value="8" <%=sel8%>>Galeri Editörü</option>
</select></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" width="50%">&nbsp;<font face="verdana" size="1">Üye Adý : </font></td><td bgcolor="#FFFFFF" width="50%"><input type="text" name="membername" class="form1" value="<%=rs("kul_adi")%>"></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" width="50%">&nbsp;<font face="verdana" size="1">Yeni Þifre (Deðiþtirmek istemiyorsanýz boþ býrakýn) : </font></td><td bgcolor="#FFFFFF" width="50%"><input type="text" name="password" class="form1"></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" width="50%">&nbsp;<font face="verdana" size="1">E-Mail : </font></td><td bgcolor="#FFFFFF" width="50%"><input type="text" name="email" class="form1" value="<%=rs("email")%>"></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" width="50%">&nbsp;<font face="verdana" size="1">Ýsim : </font></td><td bgcolor="#FFFFFF" width="50%"><input type="text" name="isim" class="form1" value="<%=rs("isim")%>"></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" width="50%">&nbsp;<font face="verdana" size="1">Gizli Soru : </font></td><td bgcolor="#FFFFFF" width="50%"><input type="text" name="g_soru" class="form1" value="<%=rs("g_soru")%>"></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" width="50%">&nbsp;<font face="verdana" size="1">Cevap (Deðiþtirmek istemiyorsanýz boþ býrakýn) : </font></td><td bgcolor="#FFFFFF" width="50%"><input type="text" name="g_cevap" class="form1"></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" width="50%">&nbsp;<font face="verdana" size="1">ICQ : </font></td><td bgcolor="#FFFFFF" width="50%"><input type="text" name="icq" class="form1" value="<%=rs("icq")%>"></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" width="50%">&nbsp;<font face="verdana" size="1">MSN : </font></td><td bgcolor="#FFFFFF" width="50%"><input type="text" name="msn" class="form1" value="<%=rs("msn")%>"></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" width="50%">&nbsp;<font face="verdana" size="1">AIM : </font></td><td bgcolor="#FFFFFF" width="50%"><input type="text" name="aim" class="form1" value="<%=rs("aim")%>"></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" width="50%">&nbsp;<font face="verdana" size="1">Þehir : </font></td><td bgcolor="#FFFFFF" width="50%"><input type="text" name="sehir" class="form1" value="<%=rs("sehir")%>"></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" width="50%">&nbsp;<font face="verdana" size="1">Meslek : </font></td><td bgcolor="#FFFFFF" width="50%"><input type="text" name="meslek" class="form1" value="<%=rs("meslek")%>"></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" width="50%">&nbsp;<font face="verdana" size="1">Cinsiyet</font></td><td bgcolor="#FFFFFF" width="50%">
<%
if rs("cinsiyet") = "a" then
section1 = "selected"
elseif rs("cinsiyet") = "b" then
section2 = "selected"
else
section1 = "selected"
end if
%>
<select size="1" class="form1" name="cinsiyet">
<option value="a" <%=section1%>>Bay</option>
<option value="b" <%=section2%>>Bayan</option>
</select>
</td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" width="50%">&nbsp;<font face="verdana" size="1">Yaþ : </font></td><td bgcolor="#FFFFFF" width="50%">
<select name="yas_1" size="1" class="form1">
<%
For i_y = 1 TO 31
IF Len(i_y) = "1" Then
i_y = "0" & i_y
End IF
IF Left(i_y,2) = age_check(0) Then
y_opt = "selected"
ELSE
y_opt = ""
End IF
Response.Write "<option value="""&i_y&""" "&y_opt&">"&i_y&"</option>"
Next
%>
</select>
.
<select name="yas_2" size="1" class="form1">
<%
For i_y1 = 1 TO 12
IF Len(i_y1) = "1" Then
i_y1 = "0" & i_y1
End IF
IF Left(i_y1,2) = age_check(1) Then
y_opt1 = "selected"
ELSE
y_opt1 = ""
End IF
Response.Write "<option value="""&i_y1&""" "&y_opt1&">"&i_y1&"</option>"
Next
%>
</select>
.
<select name="yas_3" size="1" class="form1">
<%
For i_y2 = 1920 TO Year(Now())-1
IF Left(i_y2,4) = age_check(2) Then
y_opt2 = "selected"
ELSE
y_opt2 = ""
End IF
Response.Write "<option value="""&i_y2&""" "&y_opt2&">"&i_y2&"</option>"
Next
%>
</select>
</td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" width="50%">&nbsp;<font face="verdana" size="1">Web Adresi : </font></td><td bgcolor="#FFFFFF" width="50%"><input type="text" name="web" class="form1" value="<%=rs("url")%>"></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" width="50%">&nbsp;<font face="verdana" size="1">Login Sayýsý : </font></td><td bgcolor="#FFFFFF" width="50%"><input type="text" name="login" class="form1" value="<%=rs("login_sayisi")%>"></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" width="50%">&nbsp;<font face="verdana" size="1">Mesaj Sayýsý : </font></td><td bgcolor="#FFFFFF" width="50%"><input type="text" name="msg" class="form1" value="<%=rs("msg_sayisi")%>"></td>
  </tr>
  <tr>
    <td bgcolor="#FFFFFF" width="50%">&nbsp;<font face="verdana" size="1">Diðer üyeler maili görebilsin</font></td><td bgcolor="#FFFFFF" width="50%">
<%
if rs("mail_goster") = True Then
mcheck = "checked"
elseif rs("mail_goster") = False Then
mcheck = ""
else
mcheck = ""
end if
%>
<input type="checkbox" name="mail_goster" class=form1 <%=mcheck%>></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF" width="50%">&nbsp;<font face="verdana" size="1">Ýmza : </font></td><td bgcolor="#FFFFFF" width="50%"><textarea name="imza" rows="2" class="form1"><%=rs("imza")%></textarea></td>
  </tr>
  <tr> 
    <td colspan="2" bgcolor="#FFFFFF"><input type="submit" value="Deðiþiklikleri Kaydet" class="form1" style="width:100%"></td>
  </tr>
</table>
</form>
<%
End Sub
Sub mem_update
uid = Request.QueryString("mid")
memname = duz(Request.Form("membername"))
password = duz(Request.Form("password"))
email = duz(Request.Form("email"))
isim = duz(Request.Form("isim"))
g_soru = duz(Request.Form("g_soru"))
g_cevap = duz(Request.Form("g_cevap"))
icq = duz(Request.Form("icq"))
msn = duz(Request.Form("msn"))
aim = duz(Request.Form("aim"))
sehir = duz(Request.Form("sehir"))
meslek = duz(Request.Form("meslek"))
cinsiyet = duz(Request.Form("cinsiyet"))
yas = Request.Form("yas_1") & "." & Request.Form("yas_2") & "." & Request.Form("yas_3")
url = duz(Request.Form("web"))
logincount = duz(Request.Form("login"))
msgcount = duz(Request.Form("msg"))
level = duz(Request.Form("seviye"))
imza = duz(Request.Form("imza"))

IF Len(imza) <= 0 Then
imza = "-"
ELSE
imza = imza
End IF

If Request.Form("mail_goster") = "on" Then
mailgoster = True
Else
mailgoster = False
End If

if memname="" then
response.write "<center><font face=verdana size=2 color=navy>Lütfen Üye Adý Yazýnýz</font></center>"

elseif email="" then
response.write "<center><font face=verdana size=2 color=navy>Lütfen Mail Adresi Yazýnýz</font></center>"

elseif isim="" then
response.write "<center><font face=verdana size=2 color=navy>Lütfen Ýsim Yazýnýz</font></center>"

elseif g_soru="" then
response.write "<center><font face=verdana size=2 color=navy>Lütfen Gizli Soru Yazýnýz</font></center>"

ELSE

Set rec = Server.CreateObjecT("ADODB.RecordSet")
rSQL = "Select * FROM MEMBERS WHERE uye_id = "&uid&""
rec.open rSQL,Connection,1,3

IF password<>"" Then
password = MN_PP(password)
ELSE
password = rec("sifre")
End IF

IF g_cevap<>"" Then
g_cevap = MN_PP(g_cevap)
ELSE
g_cevap = rec("g_cevap")
End IF

rec("kul_adi") = memname
rec("sifre") = password
rec("email") = email
rec("isim") = isim
rec("g_soru") = g_soru
rec("g_cevap") = g_cevap
rec("icq") = icq
rec("msn") = msn
rec("aim") = aim
rec("sehir") = sehir
rec("meslek") = meslek
rec("cinsiyet") = cinsiyet
rec("yas") = yas
rec("url") = url
rec("mail_goster") = mailgoster
rec("login_sayisi") = logincount
rec("msg_sayisi") = msgcount
rec("seviye") = level
rec("imza") = imza
rec.Update

Response.Write "<center><font face=tahoma size=2>Bilgiler Baþarýyla Güncellendi...</font></center>"
END IF

End Sub
Sub msgmembers
%>
<script language="JavaScript1.2" defer>
editor_generate('msg');
</script>
<form action="?action=SendMsg" method="post">
<table border="0" cellspacing="2" cellpadding="1" width="90%" align="center" class="td_menu">
<tr bgcolor="#E0E0E0">
<td width="20%" align="right"><b>&nbsp;Konu :&nbsp;</b></td><td><input type="text" name="subject" class="form1" style="width:100%"></td>
</tr>
<tr bgcolor="#E0E0E0">
<td colspan="2"><textarea name="msg" rows="20" class="form1" style="width:100%"></textarea></td>
</tr>
<tr bgcolor="#FFFFFF">
<td colspan="2" align="right"><input type="submit" value="   Gönder   " class="submit" style="width:100%"></td>
</tr> 
</table>
</form>
<%
End Sub
Sub sendmsg

subj = Request.Form("subject")
message = Request.Form("msg")
from = "Site Manager"

If subj="" Then
Response.Write "<font face=verdana size=2 color=red>Lütfen Konu Yazýnýz...</font>"
ElseIf message="" Then
Response.Write "<font face=verdana size=2 color=red>Lütfen Mesaj Yazýnýz...</font>"
ELSE

Set enter = Server.CreateObject("ADODB.RecordSet")
eSQL = "Select * FROM MESSAGES"
enter.open eSQL,Connection,1,3

enter.AddNew
enter("from") = from
enter("msg") = message
enter("mdate") = Now()
enter("read") = False
enter("subject") = subj
enter("type") = 1
enter("see") = True
enter.Update
Response.Write "<font face=verdana size=2>Mesaj Tüm Üyelere Ýletildi...</font>"
END IF

End Sub
Sub membermsgs

Set rs = Connection.Execute("Select * FROM MESSAGES WHERE TYPE = 0")
%>
<table border="1" cellspacing="0" cellpadding="0" width="100%">
<tr>
<td width="20%" align="center" bgcolor="#FFFFCC"><font face=arial size=2><b>Kimden</b></font></td>
<td width="20%" align="center" bgcolor="#FFFFCC"><font face=arial size=2><b>Kime</b></font></td>
<td width="40%" align="center" bgcolor="#FFFFCC"><font face=arial size=2><b>Konu</b></font></td>
<td width="20%" align="center" bgcolor="#FFFFCC"><font face=arial size=2><b>Durum</b></font></td>
</tr>
<%
Do While NOT rs.Eof

If rs("see") = True Then
durum = "Gösteriliyor"
ElseIf rs("see") = False Then
durum = "Silinmiþ"
Else
durum = "Belirsiz"
End If
%>
<tr>
<td width="20%" align="center"><font face=arial size=1><%=rs("from")%></font></td>
<td width="20%" align="center"><font face=arial size=1><%=rs("to")%></font></td>
<td width="40%" align="center"><a href="?action=MemMsgRead&mid=<%=rs("mid")%>"><font face=arial size=1><%=rs("subject")%></font></a></td>
<td width="20%" align="center"><font face=arial size=1><%=durum%></font></td>
</tr>
<%
rs.MoveNext
Loop

rs.Close
Set rs = Nothing

End Sub
Sub read_memmsg

id = Request.QueryString("mid")
Set rs = Connection.Execute("Select * FROM MESSAGES WHERE mid = "&id&"")

If rs("type") = "0" Then
tto = rs("to")
ElseIf rs("type") = "1" Then
tto = "Tüm Üyelere"
Else
tto = ""
End If
%>
<table border="1" cellspacing="0" cellpadding="1" width="100%">
<tr>
<td width="20%" align="right" bgcolor="#FFFFCC"><font face=arial size=2><b>Kimden : </b></font></td>
<td width="80%"><font face=verdana size=1><%=rs("from")%></font></td>
</tr>
<tr>
<td width="20%" align="right" bgcolor="#FFFFCC"><font face=arial size=2><b>Kime : </b></font></td>
<td width="80%"><font face=verdana size=1><%=tto%></font></td>
</tr>
<tr>
<td width="20%" align="right" bgcolor="#FFFFCC"><font face=arial size=2><b>Tarih : </b></font></td>
<td width="80%"><font face=verdana size=1><%=rs("mdate")%></font></td>
</tr>
<tr>
<td width="20%" align="right" valign="top" bgcolor="#FFFFCC"><font face=arial size=2><b>Mesaj : </b></font></td>
<td width="80%"><font face=verdana size=1><%=rs("msg")%></font></td>
</tr>
</table>
<br><br>
<center><a href="?action=DelMemMsg&mid=<%=id%>"><font face=tahoma size=2 color=red><b>Bu Mesajý Tamamen Sil</b></font></a></center>
<%
rs.Close
Set rs = Nothing

End Sub
Sub delete_memmsg

id = Request.QueryString("mid")
Set del = Server.CreateObject("ADODB.RecordSet")
dSQL = "DELETE * FROM MESSAGES WHERE mid = "&id&""
del.open dSQL,Connection,1,3
Response.Write "<center><font face=verdana size=2>Mesaj Baþarýyla Silindi...<br><br><a href=?action=MemberMsgs><< Geri</a></font></center>"

End Sub
Sub adminmsgs

Set rs = Connection.Execute("Select * FROM MESSAGES WHERE TYPE = 1")
%>
<table border="1" cellspacing="0" cellpadding="0" width="100%">
<tr>
<td width="10%" align="center" bgcolor="#FFFFCC"><font face=arial size=2><b>Ýþlem</b></font></td>
<td width="60%" align="center" bgcolor="#FFFFCC"><font face=arial size=2><b>Konu</b></font></td>
<td width="30%" align="center" bgcolor="#FFFFCC"><font face=arial size=2><b>Tarih</b></font></td>
</tr>
<% do while not rs.eof %>
<tr>
<td width="10%" align="center"><a href="?action=DelMemMsg&mid=<%=rs("mid")%>"><img src="../themes/default/msg_del.gif" border="0" alt="Bu Mesajý Tamamen Silmek için Týklayýn !!!"></a></td>
<td width="60%" align="center"><a href="?action=MemMsgRead&mid=<%=rs("mid")%>"><font face=arial size=1><%=rs("subject")%></font></a></td>
<td width="30%" align="center"><font face=arial size=1><%=rs("mdate")%></font></td>
</tr>
<%
rs.MoveNext
Loop

rs.Close
Set rs = Nothing
End Sub
Sub mem_del
uid = Request.QueryString("id")
IF uid<>"" Then
Set dm = Server.CreateObject("ADODB.RecordSet")
dmSQL = "DELETE * FROM MEMBERS WHERE uye_id = "&uid&""
dm.open dmSQL,Connection,1,3
End If
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
End If
%>