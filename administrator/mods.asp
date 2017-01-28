<!--#include file="includes.asp" -->
<%
act = Request.QueryString("action")
act = QS_CLEAR(act)

'KATEGORÝ ÝÞLEMLERÝ
IF act = "EditMain" Then
call mainedit
ElseIf act = "UpdateMain" Then
call mainupdate
ElseIf act = "DelMain" Then
call maindel

'FORUM ÝÞLEMLERÝ
ElseIf act = "ShowSettings" Then
call show_settings
ElseIf act = "UpdateSettings" Then
call update_settings
ElseIf act = "NewCat" Then
call cat_new
ElseIf act = "RegCat" Then
call cat_reg
ElseIf act = "NewForum" Then
call fnew
ElseIf act = "RegForum" Then
call freg
ElseIf act = "EditCat" Then
call cat_edit
ElseIf act = "UpdateCat" Then
call upd_cat
ElseIf act = "DeleteCat" Then
call del_cat
ElseIf act = "Lock" Then
call lock
ElseIf act = "UnLock" Then
call unlock
ElseIf act = "SetEntry" Then
call entry
ElseIf act = "SetNoEntry" Then
call noentry
ElseIF act = "NewMod" Then
call mod_new
ElseIF act = "DeleteMod" Then
call mod_del

'MESAJ ÝÞLEMLERÝ
ElseIf act = "EditMsg" Then
call editmsg
ElseIf act = "UpdateMsg" Then
call updmsg
ElseIf act = "DeleteMsg" Then
call delmsg
ElseIf act = "LockMsg" Then
call msglock
ElseIf act = "UnLockMsg" Then
call msgunlock
ElseIF act = "Edit-Rep" Then
call rep_edit
ElseIF act = "Update-Rep" Then
call rep_update
ElseIF act = "Del-Rep" Then
call rep_delete
End If

Sub rep_update
IF Session("Level") = "1" OR Session("Level") = "5" OR Session("ISMod") = "Yes" OR Session("Level") = "7" OR Session("Level") = "9" Then
IF Request.Form("msg") = "" Then
Response.Write "<br><center>Tüm Alanlarý Doldurunuz</center><br>"
ELSE
Set mRs = Connect.Execute("Select * FROM MESSAGES WHERE mid = "&QS_CLEAR(Request.QueryString("id"))&"")
frm_mid = mRs("mhit")
Connect.Execute("UPDATE MESSAGES SET mmsg = '"&Request.Form("msg")&"' WHERE mid = "&QS_CLEAR(Request.QueryString("id"))&"")
Response.Redirect "Forum.asp?action=Topic&id="&frm_mid&""
End IF
End IF
End Sub

Sub rep_edit
call UST
IF Session("Level") = "1" OR Session("Level") = "5" OR Session("ISMod") = "Yes" OR Session("Level") = "7" OR Session("Level") = "9" Then
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM MESSAGES WHERE mid = "&QS_CLEAR(Request.QueryString("id"))&"",Connect,3,3
%>
<table border="0" cellspacing="2" cellpadding="2" width="50%" align="center" class="td_menu">
<form action="?action=Update-Rep&id=<%=QS_CLEAR(Request.QueryString("id"))%>" method="post">
<tr bgcolor="<%=menu_color%>">
<td>
<textarea name="msg" rows="10" class="form1" style="width:100%" cols="20"><%=rs("mmsg")%></textarea>
</td>
</tr>
<tr>
<td>
<input type="submit" value="Güncelle »" class="submit">
</td>
</tr>
</form>
</table>
<%
rs.Close
Set rs = Nothing
End IF
call ALT
End Sub

Sub rep_delete
IF Session("Level") = "1" OR Session("Level") = "5" OR Session("ISMod") = "Yes" OR Session("Level") = "7" OR Session("Level") = "9" Then
Set mRs = Connect.Execute("Select * FROM MESSAGES WHERE mid = "&QS_CLEAR(Request.QueryString("id"))&"")
frm_mid = mRs("mhit")
Connect.Execute("DELETE FROM MESSAGES WHERE mid = "&QS_CLEAR(Request.QueryString("id"))&"")
Response.Redirect "Forum.asp?action=Topic&id="&frm_mid&""
End IF
End Sub

Sub mod_del
IF Session("Level") = "1" OR Session("Level") = "5" OR Session("Level") = "7" OR Session("Level") = "9" Then
IF QS_CLEAR(Request.QueryString("step")) = "2" Then
Set fRs = Server.CreateObject("ADODB.RecordSet")
fRs.open "Select * FROM CATEGORIES WHERE cid = "&Session("SelectedForum")&"",Connect,3,3
IF fRs("mods") = QS_CLEAR(Request.QueryString("sMem")) Then
fRs("mods") = ""
fRs.Update
ELSE
mmmmm = ","
MeMS = Replace(QS_CLEAR(Request.QueryString("sMem")),mmmmm,"k",1,-1,1)
MeMS = Replace(fRs("mods"),QS_CLEAR(Request.QueryString("sMem")),"",1,-1,1)
memStr = Split(MeMS,",")
fRs("mods") = ""
fRs.Update
For I = 1 TO UBound(memStr)
fRs("mods") = fRs("mods") & "," & memStr(""&I&"")
fRs.Update
Next
End IF
Response.Redirect "?action=DeleteMod&forum="&fRs("cid")&""
fRs.Close
Set fRs = Nothing
ELSE
Set fRs = Server.CreateObject("ADODB.RecordSet")
fRs.open "Select * FROM CATEGORIES WHERE cid = "&QS_CLEAR(Request.QueryString("forum"))&"",Connect,3,3
Response.Write "<br><div align=center class=td_menu><b>" & fRs("cname") & "</b> Forumu için Yönetici Silme Sayfasý</div>"
Session("SelectedForum") = QS_CLEAR(Request.QueryString("forum"))
%>
<table border="0" cellspacing="2" cellpadding="2" width="50%" align="center" class="td_menu2">
<tr style="font-weight:bold" bgcolor="<%=menu_color%>">
<td width="70%">Kullanýcý Adý</td>
<td width="30%" align="center">Ýþlem</td>
</tr>
<%
On Error Resume Next
f_member = Split(fRs("mods"),",")
IF Err <> "0" Then
Response.Write "<tr><td colspan=2 align=center><b>Bu Forumun Yöneticisi Yok !</b></td></tr>"
ELSE
For I = 1 TO ubound(f_member)
%>
<tr bgcolor="<%=forum_bg_1%>">
<td width="70%"><%=f_member(""&I&"")%></td>
<td width="30%" align="center"><a href="?action=DeleteMod&step=2&sMem=,<%=f_member(""&I&"")%>">Yöneticiliði Sil</a></td>
</tr>
<%
fRs.MoveNext
Next
End IF
%>
</table>
<br><br>
<div align="center" class="td_menu2">
<a href="javascript:history.back();">« Geri</a>
</div>
<%
fRs.Close
Set fRs = Nothing
End IF
End IF
End Sub

Sub mod_new
IF Session("Level") = "1" OR Session("Level") = "5" OR Session("Level") = "7" OR Session("Level") = "9" Then
IF QS_CLEAR(Request.QueryString("step")) = "3" Then
Session("SelectedForum") = Session("SelectedForum")
Session("SelectedMember") = Request.Form("member")
Set fRs = Server.CreateObject("ADODB.RecordSet")
fRs.open "Select * FROM CATEGORIES WHERE cid = "&Session("SelectedForum")&"",Connect,3,3

fRs("mods") = fRs("mods") & "," & Session("SelectedMember")
fRs.Update
Response.Write "<div align=center class=td_menu2><b>"&Session("SelectedMember")&"</b> nickli üye <b>"&fRs("cname")&"</b> isimli foruma atandý !<br><br>[ <a href=?action=NewMod>« Geri</a> ]</div>"
fRs.Close
Set fRs = Nothing
ElseIF QS_CLEAR(Request.QueryString("step")) = "2" Then
IF Request.Form("option") = "Yönetici Sil" Then
Response.Redirect "?action=DeleteMod&forum="&Request.Form("forum")&""
ELSE
Session("SelectedForum") = Request.Form("forum")
Set mRs = Connection.Execute("Select * FROM MEMBERS WHERE seviye <> 1 ORDER BY kul_adi ASC")
%>
<br>
<div align="center" class="td_menu">
<form method="post" action="?action=NewMod&step=3">
<b>Üye Seçin : </b>
<select name="member" size="1" class="form1">
<% Do Until mRs.EOF %>
<option value="<%=mRs("kul_adi")%>"><%=mRs("kul_adi")%></option>
<%
mRs.MoveNext
Loop
%>
</select>
<input type="submit" value="Devam »" class="submit">
</form>
</div>
<%
mRs.Close
Set mRs = Nothing
End IF
ELSE
Set fRs = Server.CreateObject("ADODB.RecordSet")
fRs.open "Select * FROM CATEGORIES ORDER BY cname ASC",Connect,3,3
%>
<br>
<div align="center" class="td_menu">
<form method="post" action="?action=NewMod&step=2">
<b>Forum Seçin : </b>
<select name="forum" size="1" class="form1">
<% Do Until fRs.EOF %>
<option value="<%=fRs("cid")%>"><%=fRs("cname")%></option>
<%
fRs.MoveNext
Loop
%>
</select>
<br>
<input type="submit" name="option" value="Yeni Yönetici" class="submit">
<input type="submit" name="option" value="Yönetici Sil" class="submit">
</form>
</div>
<%
fRs.Close
Set fRs = Nothing
End IF
End IF
End Sub

Sub update_settings
IF Session("Level") = "1" OR Session("Level") = "5" OR Session("Level") = "7" OR Session("Level") = "9" Then
Connect.Execute("UPDATE SETTINGS SET f_posts = "&Request.Form("topics")&"")
Connect.Execute("UPDATE SETTINGS SET f_replies = "&Request.Form("replies")&"")
Response.Write "<div align=""center"" class=""td_menu""><b>Forum Ayarlarý Güncellendi !</b><br><br><a href=""javascript:window.close();"">[ Kapat ]</a></div>"
End IF
End Sub
Sub show_settings
IF Session("Level") = "1" OR Session("Level") = "5" OR Session("Level") = "7" OR Session("Level") = "9" Then
Set sRs = Server.CreateObject("ADODB.RecordSet")
sRs.open "Select * FROM SETTINGS",Connect,3,3
%>
<form action="?action=UpdateSettings" method="post">
<table border="0" cellspacing="2" cellpadding="2" align="center" width="100%" class="td_menu" style="font-weight:bold">
<tr>
<td width="80%" align="right">Bir Sayfada Gösterilecek Konu Sayýsý : </td>
<td width="20%" align="left">
<select size="1" name="topics" class="form1">
<%
For I = 1 TO 100
IF sRs("f_posts") = I Then
opt = "selected"
Else
opt = ""
End IF
%>
<option value="<%=I%>" <%=opt%>><%=I%></option>
<% Next %>
</select>
</td>
</tr>
<tr>
<td width="80%" align="right">Bir Sayfada Gösterilecek Yanýt Sayýsý : </td>
<td width="20%" align="left">
<select size="1" name="replies" class="form1">
<%
For I = 1 TO 100
IF sRs("f_replies") = I Then
opt = "selected"
Else
opt = ""
End IF
%>
<option value="<%=I%>" <%=opt%>><%=I%></option>
<% Next %>
</select>
</td>
</tr>
<tr>
<td colspan="2" align="center"><input type="submit" value="Tamam »" class="submit" style="width:90%"></td>
</tr>
</table>
</form>
<%
End IF
End Sub
Sub updmsg
IF Session("Level") = "1" OR Session("Level") = "5" OR Session("ISMod") = "Yes" OR Session("Level") = "7" OR Session("Level") = "9" Then
msg_title = Request.Form("m_title")
msg_msg = Request.Form("m_msg")

If msg_title="" OR msg_msg="" Then
Response.Write "<font class=""td_menu""><center>Lütfen Tüm Alanlarý Doldurunuz...</center></font>"
ELSE
Set uMsg = Server.CreateObject("ADODB.RecordSet")
umSQL = "Select * FROM MESSAGES WHERE mid = "&QS_CLEAR(Request.QueryString("id"))&""
uMsg.open umSQL,Connect,3,3

uMsg("mtitle") = msg_title
uMsg("mmsg") = msg_msg
uMsg.Update
Response.Redirect "forum.asp?action=Topics&id="&uMsg("cid")&""
End If
End IF
End Sub
Sub editmsg
IF Session("Level") = "1" OR Session("Level") = "5" OR Session("ISMod") = "Yes" OR Session("Level") = "7" OR Session("Level") = "9" Then
call UST
Set msg = Connect.Execute("Select * FROM MESSAGES WHERE mid = "&QS_CLEAR(Request.QueryString("id"))&"")
%>
<br><br>
<form action="?action=UpdateMsg&id=<%=QS_CLEAR(Request.QueryString("id"))%>" method="post">
<table border="0" cellspacing="3" cellpadding="3" align="center" width="70%" class="td_menu" style="font-weight:bold">
<tr bgcolor="<%=top_menu_color%>">
<td width="40%" align="right">Baþlýk</td>
<td width="60%"><input type="text" name="m_title" size="60" class="form1" value="<%=msg("mtitle")%>"></td>
</tr>
<tr bgcolor="<%=top_menu_color%>">
<td width="40%" valign="top" align="right">Mesaj</td>
<td width="60%"><textarea name="m_msg" rows="10" cols="60" class="form1"><%=msg("mmsg")%></textarea></td>
</tr>
<tr bgcolor="<%=top_menu_color_hover%>">
<td colspan="2" align="right"><input type="submit" value="Deðiþiklikleri Kaydet >>>" class="submit"></td>
</tr>
</table>
</form>
<%
call ALT
End IF
End Sub
Sub mainupdate
IF Session("Level") = "1" OR Session("Level") = "5"  OR Session("Level") = "7" OR Session("Level") = "9" Then
Set updCat = Connect.Execute("UPDATE MAIN_CATS SET mc_name = '"&duz(Request.Form("c_name"))&"' WHERE mcid = "&QS_CLEAR(Request.QueryString("id"))&"")
Response.Redirect ""&sett("site_adres")&"/forum.asp"
End IF
End Sub
Sub mainedit
IF Session("Level") = "1" OR Session("Level") = "5" OR Session("Level") = "7" OR Session("Level") = "9" Then
call UST
Set rs = Connect.Execute("Select * FROM MAIN_CATS WHERE mcid = "&QS_CLEAR(Request.QueryString("id"))&"")
%>
<br><br>
<form action="?action=UpdateMain&id=<%=QS_CLEAR(Request.QueryString("id"))%>" method="post">
<table border="0" cellspacing="2" cellpadding="2" align="center" width="70%" class="td_menu">
<tr>
<td width="50%" align="right"><b>Kategori Adý : </b></td>
<td width="50%" align="left">
<input type="text" name="c_name" class="form1" value="<%=rs("mc_name")%>" size="20"></td>
</tr>
<tr>
<td width="50%" align="right">&nbsp;</td>
<td width="50%" align="left"><input type="submit" value="Güncelle" class="submit"></td>
</tr>
</table>
</form>
<%
call ALT
End IF
End Sub
Sub maindel
IF Session("Level") = "1" OR Session("Level") = "5" OR Session("Level") = "7" OR Session("Level") = "9" Then
mcid = Request.QueryString("id")
mcid = QS_CLEAR(mcid)
Set d1 = Server.CreateObject("ADODB.RecordSet")
d1SQL = "DELETE * FROM MAIN_CATS WHERE mcid = "&mcid&""
d1.open d1SQL,Connect,3,3
Set dsc = Connect.Execute("Select * FROM CATEGORIES WHERE mcid = "&mcid&"")
If NOT dsc.Eof Then
Set d2 = Server.CreateObject("ADODB.RecordSet")
d2SQL = "DELETE * FROM MESSAGES WHERE cid = "&dsc("cid")&""
d2.open d2SQL,Connect,3,3
Set d3 = Server.CreateObject("ADODB.RecordSet")
d3sQL = "DELETE * FROM CATEGORIES WHERE mcid = "&mcid&""
d3.open d3SQL,Connect,3,3
End If
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End IF
End Sub
Sub delmsg
IF Session("Level") = "1" OR Session("Level") = "5" OR Session("ISMod") = "Yes" OR Session("Level") = "7" OR Session("Level") = "9" Then
msid = Request.QueryString("id")
msid = QS_CLEAR(msid)
Set mdel = Server.CreateObject("ADODB.RecordSet")
mdSQL = "DELETE * FROM MESSAGES WHERE mid = "&msid&" AND topic = True"
mdel.open mdSQL,Connect,3,3
Set mdel2 = Server.CreateObject("ADODB.RecordSet")
md2SQL = "DELETE * FROM MESSAGES WHERE mhit = "&msid&" AND topic = False"
mdel2.open md2SQL,Connect,3,3
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End IF
End Sub
Sub noentry
IF Session("Level") = "1" OR Session("Level") = "5" Then
cid = Request.QueryString("id")
cid = QS_CLEAR(cid)
Set cnoentry = Connect.Execute("UPDATE CATEGORIES SET noentry = True WHERE cid="&cid&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End IF
End Sub
Sub entry
IF Session("Level") = "1" OR Session("Level") = "5" OR Session("Level") = "7" OR Session("Level") = "9" Then
cid = Request.QueryString("id")
cid = QS_CLEAR(cid)
Set cnoentry = Connect.Execute("UPDATE CATEGORIES SET noentry = False WHERE cid="&cid&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End IF
End Sub
Sub del_cat
IF Session("Level") = "1" OR Session("Level") = "5" OR Session("Level") = "7" OR Session("Level") = "9" Then
id = Request.QueryString("id")
id = QS_CLEAR(id)
Set del1 = Server.CreateObject("ADODB.RecordSet")
d1SQL = "DELETE FROM CATEGORIES WHERE cid = "&id&""
del1.open d1SQL,Connect,3,3
Set del2 = Server.CreateObject("ADODB.RecordSet")
d2SQL = "DELETE FROM MESSAGES WHERE cid = "&id&""
del2.open d2SQL,Connect,3,3
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End IF
End Sub
Sub cat_edit
IF Session("Level") = "1" OR Session("Level") = "5" OR Session("Level") = "7" OR Session("Level") = "9" Then
call UST
Set rs = Connect.Execute("Select * FROM CATEGORIES WHERE cid = "&QS_CLEAR(Request.QueryString("id"))&"")
Set cats = Connect.Execute("Select * FROM MAIN_CATS")
%>
<form action="?action=UpdateCat&id=<%=QS_CLEAR(Request.QueryString("id"))%>" method="post">
<table border="0" cellspacing="2" cellpadding="2" align="center" width="70%" style="font-family:verdana; font-size:11px;">
<tr>
<td width="30%" align="right">Forum Adý : </td>
<td width="70%" align="left"><input type="text" name="cname" class="form1" size="30" value="<%=rs("cname")%>"></td>
</tr>
<tr>
<td width="30%" align="right">Forum Açýklamasý : </td>
<td width="70%" align="left"><textarea name="cinfo" rows="5" cols="29" class="form1"><%=rs("cinfo")%></textarea></td>
</tr>
<tr>
<td width="30%" align="right">Forum Kategori : </td>
<td width="70%" align="left">
<select name="mcat" size="1" class="form1">
<%
Do While NOT cats.Eof
If rs("mcid") = cats("mcid") Then
sel = "selected"
Else
sel = ""
End If
%>
<option value="<%=cats("mcid")%>" <%=sel%>><%=cats("mc_name")%></option>
<%
cats.MoveNext
Loop
%>
</select>
</td>
</tr>
<tr>
<td width="30%" align="right">&nbsp;</td>
<td width="70%" align="left"><input type="submit" class="submit" value="G ü n c e l l e"></td>
</tr>
</table>
</form>
<%
rs.Close
Set rs = Nothing
call ALT
End IF
End Sub
Sub upd_cat
IF Session("Level") = "1" OR Session("Level") = "5" OR Session("Level") = "7" OR Session("Level") = "9" Then
id = Request.QueryString("id")
id = QS_CLEAR(id)
name = duz(Request.Form("cname"))
info = duz(Request.Form("cinfo"))
cat = duz(Request.Form("mcat"))

Set cupd = Server.CreateObject("ADODB.RecordSet")
cupdSQL = "Select * FROM CATEGORIES WHERE cid = "&id&""
cupd.open cupdSQL,Connect,3,3

If NOT cupd.Eof Then
cupd("cname") = name
cupd("cinfo") = info
cupd("mcid") = cat
cupd.Update
END IF
cupd.Close
Set cupd = Nothing
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End IF
End Sub
Sub unlock
IF Session("Level") = "1" OR Session("Level") = "5" OR Session("Level") = "7" OR Session("Level") = "9" Then
fid = Request.QueryString("id")
fid = QS_CLEAR(fid)
Set upd = Server.CreateObject("ADODB.RecordSet")
updSQL = "SELECT * FROM CATEGORIES WHERE cid = "&fid&""
upd.open updSQL,Connect,3,3

If NOT upd.Eof Then
upd("locked") = False
upd.Update
End IF
upd.Close
Set upd = Nothing
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End IF
End Sub
Sub Lock
IF Session("Level") = "1" OR Session("Level") = "5" OR Session("Level") = "7" OR Session("Level") = "9" Then
fid = Request.QueryString("id")
fid = QS_CLEAR(fid)
Set upd = Server.CreateObject("ADODB.RecordSet")
updSQL = "SELECT * FROM CATEGORIES WHERE cid = "&fid&""
upd.open updSQL,Connect,3,3

If NOT upd.Eof Then
upd("locked") = True
upd.Update
End IF
upd.Close
Set upd = Nothing
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End IF
End Sub
Sub msgunlock
IF Session("Level") = "1" OR Session("Level") = "5" OR Session("ISMod") = "Yes" OR Session("Level") = "7" OR Session("Level") = "9" Then
msgid = Request.QueryString("id")
msgid = QS_CLEAR(msgid)
Set msgul = Connect.Execute("UPDATE MESSAGES SET Locked = False WHERE mid="&msgid&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End IF
End Sub
Sub msglock
IF Session("Level") = "1" OR Session("Level") = "5" OR Session("ISMod") = "Yes" OR Session("Level") = "7" OR Session("Level") = "9" Then
msgid = Request.QueryString("id")
msgid = QS_CLEAR(msgid)
Set msgul = Connect.Execute("UPDATE MESSAGES SET Locked = True WHERE mid="&msgid&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End IF
End Sub
Sub fnew
IF Session("Level") = "1" OR Session("Level") = "5" OR Session("Level") = "7" OR Session("Level") = "9" Then
Set cats = Connect.Execute("Select * FROM MAIN_CATS")
%>
<form action="?action=RegForum" method="post">
<table border=0 cellspacing=1 cellpadding=1 width=100% align=center style="font-family:tahoma; font-size:11px; font-weight:bold">
<tr>
<td width="30%" align="right">Forum Adý : </td>
<td width="70%" align="left"><input type="text" name="fname" size="30" class="form1"></td>
</tr>
<tr>
<td width="30%" align="right">Açýklamasý : </td>
<td width="70%" align="left"><textarea name="finfo" cols="29" rows="5" class="form1"></textarea></td>
</tr>
<tr>
<td width="30%" align="right">Kategori : </td>
<td width="70%" align="left">
<select name="fcat" size="1" class="form1">
<% Do While NOT cats.Eof %>
<option value="<%=cats("mcid")%>"><%=cats("mc_name")%></option>
<%
cats.MoveNext
Loop
%>
</select>
</td>
</tr>
<tr>
<td colspan="2"><input type="submit" value="K A Y D E T" class="submit" style="width: 100%"></td>
</tr>
</table>
</form>
<%
End IF
End Sub
Sub freg
IF Session("Level") = "1" OR Session("Level") = "5" OR Session("Level") = "7" OR Session("Level") = "9" Then
namef = duz(Request.Form("fname"))
infof = duz(Request.Form("finfo"))
catf = duz(Request.Form("fcat"))

If namef="" OR infof="" Then
Response.Write "<center><font face=verdana size=1>Lütfen Tüm Alanlarý Doldurunuz...</font></center>"
ELSE
Set creg = Server.CreateObject("ADODB.RecordSet")
cregSQL = "Select * FROM CATEGORIES"
creg.open cregSQL,Connect,3,3

creg.AddNew
creg("cname") = namef
creg("cinfo") = infof
creg("mcid") = catf
creg("mods") = "-"
creg.Update

creg.Close
Set creg = Nothing
Response.Redirect Request.ServerVariables("HTTP_REFERER")
END IF
End IF
End Sub
Sub cat_new
IF Session("Level") = "1" OR Session("Level") = "5" OR Session("Level") = "7" OR Session("Level") = "9" Then
%>
<form action="?action=RegCat" method="post">
<table border="0" cellspacing="0" cellpadding="1" align="center" width="100%" style="font-family:tahoma; font-size:11px; font-weight:bold">
<tr>
<td width="40%" align="right">Kategori Adý : </td>
<td width="60%" align="right"><input type="text" name="cname" class="form1" size="30"></td>
</tr>
<tr>
<td colspan="2"><input type="submit" value="K A Y D E T" class="submit" style="width:100%"></td>
</tr>
</table>
</form>
<%
End IF
End Sub
Sub cat_reg
IF Session("Level") = "1" OR Session("Level") = "5" OR Session("Level") = "7" OR Session("Level") = "9" Then
name = duz(Request.Form("cname"))
If name="" Then
Response.Write "<center><font face=verdana size=1>Lütfen Kategori Adý Yazýnýz...</font></center>"
ELSE
Set cent = Server.CreateObject("ADODB.RecordSet")
cSQL = "Select * FROM MAIN_CATS"
cent.open cSQL,Connect,3,3

cent.AddNew
cent("mc_name") = name
cent.Update
Response.Redirect Request.ServerVariables("HTTP_REFERER")
END IF
End IF
End Sub
%>