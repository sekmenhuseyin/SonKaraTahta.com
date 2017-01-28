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

Response.Write "<center><font face=arial size=3 color=navy><b>B L O K L A R</b></font></center>"
act = Request.QueryString("action")
If act = "all" Then
call blocks
ElseIF act = "New" Then
call new_block
ElseIF act = "NewReg" Then
call new_block_reg
ElseIF act = "Delete" Then
call block_delete
ElseIF act = "Edit" Then
call block_edit
ElseIF act = "Update" Then
call block_update
ElseIf act = "Passive" Then
call passive
ElseIf act = "Active" Then
call active
End If

Sub block_update
name = Request.Form("name")
content = Request.Form("content")
include = Request.Form("include")
align = Request.Form("align")
queue = Request.Form("queue")

IF Request.Form("active") = "on" Then
actv = "True"
ELSE
actv = "False"
End IF

IF name="" or queue="" Then
Response.Write "<div align=center class=td_menu>Lütfen Tüm Alanlar Doldurunuz</div>"
ELSE
Set upd = Server.CreateObject("ADODB.RecordSet")
upd.open "Select * FROM BLOCKS WHERE bid = "&Request.QueryString("id")&"",Connection,3,3

upd("b_name") = name
upd("b_content") = content
upd("b_include") = include
upd("b_align") = align
upd("b_active") = actv
upd("b_queue") = queue
upd.Update
Response.Redirect "?action=all"
End IF
End Sub
Sub block_edit
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM BLOCKS WHERE bid = "&Request.QueryString("id")&"",Connection,3,3
%>
<script language="JavaScript1.2" defer>
editor_generate('content');
</script>
<table border="0" cellspacing="1" cellpadding="1" align="center" class="td_menu" bgcolor="#999999">
<form action="?action=Update&id=<%=Request.QueryString("id")%>" method="post">
<tr align="center" bgcolor="#FFFFCC" style="font-size:12px;font-weight:bold">
<td colspan="2"><b>Blok Düzenle (<i><%=rs("b_name")%></i>)</b></td>
</tr>
<tr bgcolor="#F0F0F0" style="font-weight:bold">
<td width="30%">Blok Adý : </td>
<td width="70%"><input type="text" name="name" class="form1" value="<%=rs("b_name")%>" style="width:100%"></td>
</tr>
<tr bgcolor="#F0F0F0">
<td colspan="2"><textarea name="content" rows="15" cols="95" wrap="VIRTUAL" class="form1"><%=rs("b_content")%></textarea></td>
</tr>
<tr bgcolor="#F0F0F0">
<td width="30%">Blok Yeri</td>
<td width="70%">
<%
IF rs("b_align") = "LEFT" Then
opt1 = "selected"
ElseIF rs("b_align") = "RIGHT" Then
opt2 = "selected"
ELSE
opt1 = "selected"
End IF
%>
<select name="align" size="1" class="form1">
<option value="LEFT" <%=opt1%>>Sol Blok</option>
<option value="RIGHT" <%=opt2%>>Sað Blok</option>
</select>
</td>
</tr>
<tr bgcolor="#F0F0F0">
<td width="30%">Blok Sayfasý (Varsa,Yoksa <b>null</b> yazýnýz)</td>
<td width="70%">BLOCKS/
<input type="text" name="include" value="<%=rs("b_include")%>" class="form1" style="width:50%">
</td>
</tr>
<tr bgcolor="#F0F0F0">
<td width="30%">Sýra</td>
<td width="70%">
<select name="queue" size="1" class="form1">
<% Set block_record = Connection.Execute("Select Count(*) as sayi FROM BLOCKS")
For I = 1 TO block_record("sayi")+2
IF I = rs("b_queue") Then
opt = "selected"
ELSE
opt = ""
End IF
%>
<option value="<%=I%>" <%=opt%>><%=I%></option>
<%
Next
%>
</select>
</td>
</tr>
<tr bgcolor="#F0F0F0">
<td width="30%">Aktif</td>
<td width="70%">
<%
IF rs("b_active") = True Then
c_opt = "checked"
End IF
%>
<input type="checkbox" name="active" <%=c_opt%>> <i>( Ýþaretliyse Aktif , Deðilse Pasif )</i></td>
</tr>
<tr bgcolor="#F0F0F0" align="right">
<td colspan="2"><input type="submit" value="Tamam" class="submit" style="width:20%"></td>
</tr>
</form>
</table>
<%
End Sub
Sub block_delete
Connection.Execute("DELETE * FROM BLOCKS WHERE bid = "&Request.QueryString("id")&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
Sub new_block_reg
name = Request.Form("name")
content = Request.Form("content")
include = Request.Form("include")
align = Request.Form("align")
queue = Request.Form("queue")

IF Request.Form("active") = "on" Then
actv = "True"
ELSE
actv = "False"
End IF

IF name="" or queue="" Then
Response.Write "<div align=center class=td_menu>Lütfen Tüm Alanlar Doldurunuz</div>"
ELSE
Connection.Execute("INSERT INTO BLOCKS (b_name, b_content, b_include, b_align, b_active, b_queue) VALUES ('"&name&"', '"&content&"', '"&include&"', '"&align&"', "&actv&", "&queue&")")
Response.Redirect "?action=all"
End IF
End Sub
Sub new_block
%>
<script language="JavaScript1.2" defer>
editor_generate('content');
</script>
<table border="0" cellspacing="1" cellpadding="1" width="90%" align="center" class="td_menu" bgcolor="#999999">
<form action="?action=NewReg" method="post">
<tr align="center" bgcolor="#FFFFCC" style="font-size:12px;font-weight:bold">
<td colspan="2"><b>Yeni Blok</b></td>
</tr>
<tr bgcolor="#F0F0F0" style="font-weight:bold">
<td width="30%">Blok Adý : </td>
<td width="70%"><input type="text" name="name" size="95" class="form1"></td>
</tr>
<tr bgcolor="#F0F0F0">
<td colspan="2"><textarea name="content" rows="15" wrap="VIRTUAL" class="form1" style="width:100%"></textarea></td>
</tr>
<tr bgcolor="#F0F0F0">
<td width="30%">Blok Yeri</td>
<td width="70%">
<select name="align" size="1" class="form1">
<option value="LEFT">Sol Blok</option>
<option value="RIGHT">Sað Blok</option>
</select>
</td>
</tr>
<tr bgcolor="#F0F0F0">
<td width="30%">Blok Sayfasý (Varsa,Yoksa <b>null</b> yazýnýz)</td>
<td width="70%">BLOCKS/
<input type="text" name="include" value="null" class="form1" style="width:50%">
</td>
</tr>
<tr bgcolor="#F0F0F0">
<td width="30%">Sýra</td>
<td width="70%">
<select name="queue" size="1" class="form1">
<% Set block_record = Connection.Execute("Select Count(*) as sayi FROM BLOCKS")
For I = 1 TO block_record("sayi")+2
%>
<option value="<%=I%>"><%=I%></option>
<%
Next
%>
</select>
</td>
</tr>
<tr bgcolor="#F0F0F0">
<td width="30%">Aktif</td>
<td width="70%"><input type="checkbox" name="active"> <i>( Ýþaretliyse Aktif , Deðilse Pasif )</i></td>
</tr>
<tr bgcolor="#F0F0F0" align="right">
<td colspan="2"><input type="submit" value="Tamam" class="submit" style="width:20%"></td>
</tr>
</form>
</table>
<%
End Sub
Sub passive
blid = Request.QueryString("id")
Set activation = Connection.Execute("UPDATE BLOCKS SET b_active = False WHERE bid = "&blid&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
Sub active
blid = Request.QueryString("id")
Set activation = Connection.Execute("UPDATE BLOCKS SET b_active = True WHERE bid = "&blid&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
Sub blocks
Set rs = Connection.Execute("Select * FROM BLOCKS")
%>
<div align="left" class="td_menu"><b><a href="?action=New">» Yeni Blok Ekle</a></b></div>
<table border="1" cellspacing="0" cellpadding="1" align="center" width="90%">
<tr style="font-size:11px; font-family:tahoma; font-weight:bold">
<td width="70%" bgcolor="#FFFFCC">Blok Ýsmi</td>
<td width="15%" align="center" bgcolor="#FFFFCC">Durum</td>
<td width="15%" align="center" bgcolor="#FFFFCC">Ýþlem</td>
</tr>
<% Do While NOT rs.Eof %>
<tr style="font-family:verdana; font-size:11px">
<td width="70%"><%=rs("b_name")%></td>
<td width="15%" align="center"><% If rs("b_active") = True Then %><font color=red>Aktif</font>&nbsp;<a href="?action=Passive&id=<%=rs("bid")%>"><img src="IMAGES/Passive.gif" border="0" alt="Pasifleþtirmek Ýçin Týklayýn" align="absmiddle"></a><%Else%><font color=red>Pasif</font>&nbsp;<a href="?action=Active&id=<%=rs("bid")%>"><img src="IMAGES/Active.gif" border="0" alt="Aktifleþtirmek Ýçin Týklayýn" align="absmiddle"></a><% End If %></td>
<td width="15%" align="center"><a href="?action=Edit&id=<%=rs("bid")%>">Düzenle</a> - <a href="?action=Delete&id=<%=rs("bid")%>">Sil</a></td>
</tr>
<%
rs.MoveNext
Loop
%>
</table>
<br><br>
<font face=verdana size=1>
<font color=red>Pasif Hale Getirmek :</font> Eðer bloðun yanýnda "Aktif" yazýyorsa yanýndaki "çarpý" iþaretine basarak bloðu pasif hale getirebilirsiniz<br>
<font color=red>Aktif Hale Getirmek :</font> Eðer bloðun yanýnda "Pasif" yazýyorsa yanýndaki "tick" iþaretine basarak bloðu aktif hale getirebilirsiniz
</font>
<center>
<br><font face=tahoma color=#000090 style="font-size:11px"><b>Sayfada görünmesini istemediðiniz bloklarý "Pasif" hale getirmeniz yeterli olacaktýr.</b></font></center>
<%
End Sub
End If
%>          