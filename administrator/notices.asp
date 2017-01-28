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
IF Session("Level") = "1" OR Session("Level") = "7" Then
action = Request.QueryString("x")

IF action = "New" Then
%>
<script language="JavaScript1.2" defer>
editor_generate('notice_text');
</script>
<table border="0" cellspacing="2" cellpadding="2" width="90%" class="td_menu" align="center">
<form action="?x=Register" method="post">
<tr bgcolor="#EEEEEE">
<td colspan="2">
<textarea name="notice_text" rows="10" class="form1" style="width:100%" cols="20"></textarea></td>
</tr>
<tr>
<td width="50%"></td>
<td width="50%" align="right"><input type="submit" value="     Tamam     " class="submit"></td>
</tr>
</form>
</table>
<%
ElseIF action = "Register" Then
n_text = Request.Form("notice_text")
IF n_text = "" Then
Response.Write "<div align=center>Lütfen Duyuruyu Yazýnýz</div>"
ELSE
Set nRs = Server.CreateObject("ADODB.RecordSet")
nRs.open "Select * FROM NOTICES",Connection,3,3

nRs.AddNew
nRs("n_text") = n_text
nRs("n_date") = Now()
nRs.Update
Response.Redirect "notices.asp"
End IF
ElseIF action = "Edit" Then
Set rs = Connection.Execute("Select * FROM NOTICES WHERE n_id = "&Request.QueryString("id")&"")
%>
<script language="JavaScript1.2" defer>
editor_generate('notice_text');
</script>
<table border="0" cellspacing="2" cellpadding="2" width="90%" class="td_menu" align="center">
<form action="?x=Update&id=<%=Request.QueryString("id")%>" method="post">
<tr bgcolor="#EEEEEE">
<td>
<textarea name="notice_text" rows="10" style="width:100%" class="form1" cols="20"><%=rs("n_text")%></textarea></td>
</tr>
<tr>
<td><input type="submit" value="     Tamam     " class="submit"></td>
</tr>
</form>
</table>
<%
ElseIF action = "Update" Then
n_text = Request.Form("notice_text")
IF n_text = "" Then
Response.Write "<div align=center>Lütfen Duyuruyu Yazýnýz</div>"
ELSE
Set uRs = Server.CreateObject("ADODB.RecordSet")
uRs.open "Select * FROM NOTICES WHERE n_id = "&Request.QueryString("id")&"",Connection,3,3
uRs("n_text") = n_text
uRs.Update
Response.Redirect "notices.asp"
End IF
ElseIF action = "Delete" Then
Set d_n = Connection.Execute("DELETE FROM NOTICES WHERE n_id = "&Request.QueryString("id")&"")
Set d_n = Nothing
Response.Redirect "notices.asp"
ELSE
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM NOTICES",Connection,3,3
%>
<table border="1" cellspacing="0" cellpadding="2" width="90%" class="td_menu" align="center">
<tr>
<td colspan="3" style="font-weight:bold">» <a href="?x=New">Yeni Duyuru Ekle</a></td>
</tr>
<tr style="font-weight:bold" bgcolor="#FFCC77">
<td width="5%" align="center">No</td>
<td width="75%">Duyuru (Ýlk 150 harf)</td>
<td width="20%" align="center">Ýþlemler</td>
</tr>
<% Do Until rs.Eof %>
<tr>
<td width="5%" align="center"><%=rs("n_id")%>&nbsp;</td>
<td width="75%"><%=Left(rs("n_text"),150)%>&nbsp;</td>
<td width="20%" align="center"><a href="?x=Edit&id=<%=rs("n_id")%>">Düzenle</a> - <a href="?x=Delete&id=<%=rs("n_id")%>">Sil</a></td>
</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
End IF

End IF
%>