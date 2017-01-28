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
If Session("Level") = "1"  OR Session("Level") = "7" Then
Response.Write "<center><font face=arial size=3 color=navy><b>S A Y F A L A R</b></font></center>"
act = Request.QueryString("action")
If act = "all" Then
call all
ElseIf act = "NewPage" Then
call npage
ElseIf act = "EditPage" Then
call edit
ElseIf act = "UpdatePage" Then
call update
ElseIf act = "DeletePage" Then
call delete
ElseIf act = "RegisterPage" Then
call reg
End If
Sub delete
page_id = Request.QueryString("id")
Set dm = Server.CreateObject("ADODB.RecordSet")
dmSQL = "DELETE * FROM PAGES WHERE p_id = "&page_id&""
dm.open dmSQL,Connection,1,3
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
Sub all
Set rsP = Server.CreateObject("ADODB.RecordSet")
pSQL = "Select * FROM PAGES ORDER BY p_title ASC"
rsP.open pSQL,Connection,1,3
%>
<div align="left" class="td_menu">» <a href="?action=NewPage"><b>Yeni Sayfa Yarat</b></a></div><br>
<table border="1" cellspacing="0" cellpadding="1" align="center" width="80%">
<tr style="font-family:tahoma; font-size:11px; font-weight:bold" bgcolor="#FFFFCC" height="20">
<td width="50%">Sayfa Baþlýðý</td>
<td width="25%" align="center">Durum</td>
<td width="25%" align="center">Ýþlem</td>
</tr>
<%
Do While NOT rsP.Eof
IF rsP("p_cat") <> "" Then
Set cat = Connection.Execute("Select * FROM MENU_CATS WHERE mc_id = "&rsP("p_cat")&"")
p_cat = cat("mc_title")
ELSE
p_cat = "Belirsiz"
End IF
%>
<tr style="font-family:verdana; font-size:11px;">
<td width="50%"><%=rsP("p_title")%> (<%=p_cat%>)</td>
<td width="25%" align="center"><% If rsP("p_members") = True Then %><font color=red>Sadece Üyeler</font><%Else%><font color=red>Herkes</font><% End If %></td>
<td width="25%" align="center"><a href="?action=EditPage&id=<%=rsP("p_id")%>">Düzenle</a> - <a href="?action=DeletePage&id=<%=rsP("p_id")%>">Sil</a></td>
</tr>
<%
rsP.MoveNext
Loop
%>
</table>
<%
End Sub
Sub npage
%>
<script language="JavaScript1.2" defer>
editor_generate('content');
</script>
<form action="?action=RegisterPage" method="post">
<table border="0" cellspacing="2" cellpadding="1" align="center" width="90%" class="td_menu" bgcolor="<%=background%>">
<tr>
<td colspan="2" bgcolor="#CCCCCC" align="center" style="font-size:12px;font-weight:bold">Yeni Sayfa</td>
</tr>
<tr bgcolor="#EEEEEE" style="font-weight:bold">
<td width="40%" align="right">Sayfa Baþlýðý : </td>
<td width="60%"><input type="text" name="title" size="60" class="form1"></td>
</tr>
<tr bgcolor="#EEEEEE" style="font-weight:bold">
<td colspan="2">
<textarea name="content" rows="20" class="form1" style="width:100%" cols="20"></textarea></td>
</tr>
<tr bgcolor="#EEEEEE" style="font-weight:bold">
<td width="40%" align="right">Giriþ : </td>
<td width="60%">
<select name="mems" size="1" class="form1">
<option value="members">Sadece Üyeler</option>
<option value="everyone">Herkes</option>
</select>
</td>
</tr>
<% Set cats = Connection.Execute("Select * FROM MENU_CATS ORDER BY mc_title ASC") %>
<tr bgcolor="#EEEEEE" style="font-weight:bold">
<td width="40%" align="right">Kategori : </td>
<td width="60%">
<select name="cat" size="1" class="form1">
<% Do Until cats.EoF %>
<option value="<%=cats("mc_id")%>"><%=cats("mc_title")%></option>
<%
cats.MoveNext
Loop
%>
</select>
</td>
</tr>
<tr align="center">
<td colspan="2"><input type="submit" value="Sayfayý Kaydet »" class="submit"></td>
</tr>
</table>
</form>
<br><br>
<div align="center" class="td_menu" style="font-weight:bold"><a href="?action=all"><< Geri</a></div>
<%
End Sub
Sub reg
ptitle = Request.Form("title")
pcontent = Request.Form("content")

If Request.Form("mems") = "members" Then
mem = True
Else
mem = False
End If

IF ptitle="" Then
Response.Write "<center><font face=verdana size=2 color=navy>Lütfen Sayfa Baþlýðýný Yazýnýz...</font></center>"
ElseIF pcontent="" Then
Response.Write "<center><font face=verdana size=2 color=navy>Lütfen Sayfa Ýçeriðini Yazýnýz...</font></center>"
ELSE

Set ent = Server.CreateObject("ADODB.RecordSet")
eSQL = "Select * FROM PAGES"
ent.open eSQL,Connection,1,3

ent.AddNew
ent("p_title") = ptitle
ent("p_content") = pcontent
ent("p_members") = mem
ent("p_cat") = Request.Form("cat")
ent.Update
Response.Write "<center><font face=verdana size=2 color=navy>Sayfa Eklendi...<br><br><a href=?action=all><< Geri</a></font></center>"
END IF
End Sub
Sub edit
pid = Request.QueryString("id")
Set pRs = Server.CreateObject("ADODB.RecordSet")
pRs.open "Select * FROM PAGES WHERE p_id = "&pid&"",Connection,1,3
%>
<script language="JavaScript1.2" defer>
editor_generate('content');
</script>
<form action="?action=UpdatePage&id=<%=pid%>" method="post">
<table border="0" cellspacing="2" cellpadding="1" align="center" width="90%" class="td_menu" bgcolor="<%=background%>">
<tr>
<td colspan="2" bgcolor="#CCCCCC" align="center" style="font-size:12px;font-weight:bold"><i>(<%=pRs("p_title")%>)</i> Baþlýklý Sayfayý Düzenliyorsunuz</td>
</tr>
<tr bgcolor="#EEEEEE" style="font-weight:bold">
<td width="40%" align="right">Sayfa Baþlýðý : </td>
<td width="60%"><input type="text" name="title" size="60" class="form1" value="<%=oku(pRs("p_title"))%>"></td>
</tr>
<tr bgcolor="#EEEEEE" style="font-weight:bold">
<td colspan="2">
<textarea name="content" rows="20" class="form1" style="width:100%" cols="20"><%=pRs("p_content")%></textarea></td>
</tr>
<tr bgcolor="#EEEEEE" style="font-weight:bold">
<td width="40%" align="right">Giriþ : </td>
<td width="60%">
<select name="mems" size="1" class="form1">
<%
IF pRs("p_members") = True Then
opt1 = "selected"
ELSE
opt2 = "selected"
End IF
%>
<option value="members" <%=opt1%>>Sadece Üyeler</option>
<option value="everyone" <%=opt2%>>Herkes</option>
</select>
</td>
</tr>
<% Set cats = Connection.Execute("Select * FROM MENU_CATS ORDER BY mc_title ASC") %>
<tr bgcolor="#EEEEEE" style="font-weight:bold">
<td width="40%" align="right">Kategori : </td>
<td width="60%">
<select name="cat" size="1" class="form1">
<%
Do Until cats.EoF
IF pRs("p_cat") = cats("mc_id") Then
opt = "selected"
ELSE
opt = ""
End IF
%>
<option value="<%=cats("mc_id")%>" <%=opt%>><%=cats("mc_title")%></option>
<%
cats.MoveNext
Loop
%>
</select>
</td>
</tr>
<tr align="center">
<td colspan="2"><input type="submit" value="Sayfayý Güncelle »" class="submit"></td>
</tr>
</table>
</form>
<br><br>
<div align="center" class="td_menu" style="font-weight:bold"><a href="?action=all"><< Geri</a></div>
<%
End Sub
Sub update
pid = Request.QueryString("id")
ptitle = Request.Form("title")
pcontent = Request.Form("content")

If Request.Form("mems") = "members" Then
mem = True
Else
mem = False
End If

IF ptitle="" Then
Response.Write "<center><font face=verdana size=2 color=navy>Lütfen Sayfa Baþlýðýný Yazýnýz...</font></center>"
ElseIF pcontent="" Then
Response.Write "<center><font face=verdana size=2 color=navy>Lütfen Sayfa Ýçeriðini Yazýnýz...</font></center>"
ELSE

Set ent = Server.CreateObject("ADODB.RecordSet")
eSQL = "Select * FROM PAGES WHERE p_id = "&pid&""
ent.open eSQL,Connection,1,3

ent("p_title") = ptitle
ent("p_content") = pcontent
ent("p_members") = mem
ent("p_cat") = Request.Form("cat")
ent.Update
Response.Write "<center><font face=verdana size=2 color=navy>Sayfa Bilgileri Baþarýyla Güncellendi...<br><br><a href=?action=all><< Geri</a></font></center>"
END IF
End Sub
End If
%>