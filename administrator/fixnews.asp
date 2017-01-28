<% Response.Buffer = True %>
<!--#include file="includes.asp" -->
<%
If Session("Level") = "1" OR Session("Level") = "4" OR Session("Level") = "7" Then
act = Request.QueryString("action")
If act = "all" Then
call all_news
ElseIf act = "organize" Then
call organize
ElseIf act = "update" Then
call update
ElseIf act = "new_enter" Then
call enternew
ElseIf act = "register" Then
call reg
ElseIf act = "IS" Then
call ImageSelect
ElseIf act = "IU" Then
call ImageUpload
ElseIf act = "delete" Then
call del
ElseIf act = "change" Then
call change
ElseIF act = "Cats" Then
call cats
ElseIF act = "Cat_New" Then
call cat_new
ElseIF act = "Cat_Register" Then
call cat_newreg
ElseIF act = "Cat_Edit" Then
call cat_edit
ElseIF act = "Cat_Update" Then
call cat_update
ElseIF act = "Cat_Delete" Then
call cat_delete
ElseIF act = "WaitApprove" Then
call wait_approve
ElseIF act = "AllNews" Then
call all_news
End If

Sub all_news
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM FIXED",Connection,3,3
IF rs.EoF OR rs.BoF Then
Response.Write "<div align=center class=td_menu><b>Haber YOK !</b></div>"
ELSE
%>
<table border="1" cellspacing="0" cellpadding="2" align="center" width="90%" class="td_menu">
<%
Do Until rs.EoF
%>
<tr>
<td width="70%" style="font-weight:bold"><%=rs("sf_baslik")%></td>
<td width="30%" align="center" bgcolor="#F4F4F4"><a href="?action=organize&fid=<%=rs("fid")%>">Düzenle</a> - <a href="?action=delete&fid=<%=rs("fid")%>">Sil</a></td>
</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
End IF
%>
<%
End Sub 
Sub enternew 
%>
<script language="JavaScript1.2" defer>
editor_generate('metin');
</script>
<form action="?action=register" method="post">
<table border="0" cellspacing="2" cellpadding="2" align="center" class="td_menu" style="font-weight:bold" bgcolor="#FFFFFF">
<tr bgcolor="#E0E0E0">
<td align="right">Baþlýk : </td>
<td><input type="text" name="baslik" class="form1" size="60"></td>
</tr>
<tr bgcolor="#E0E0E0">
<td colspan="2"><textarea name="metin" rows="10" class="form1" style="width:100%"></textarea>
</td>
</tr>
<tr bgcolor="#E0E0E0">
<td align="right" valign="top">Resim URL : </td>
<td>
Images/
<input type="text" name="image" class="form1" style="width:85%" value="CLIP02.ico"></td>
</tr>
<tr>
<td align="right">&nbsp;</td>
<td align="right"><input type="submit" class="submit" size="40" value="Kaydet »"></td></tr>
</form>
</table>
<br>
<%
End Sub
Sub reg

baslik = Request("baslik")
haber = Request("metin")
image = Request("image")

If baslik="" Then
Response.Write "<center><font face=verdana size=2 color=navy>Lütfen Baþlýk Yazýnýz...</font></center>"
ElseIf haber="" Then
Response.Write "<center><font face=verdana size=2 color=navy>Lütfen Haber Metni Yazýnýz...</font></center>"
ELSE
Set ent = Server.CreateObject("ADODB.RecordSet")
ent.open "Select * FROM FIXED",Connection,3,3

Session.LCID = 2048
ent.AddNew
ent("sf_baslik") = baslik
ent("f_metin") = haber
ent("resim") = image
ent.Update
ent.Close
Set ent = Nothing
Response.Redirect "?action=all"
End IF

End Sub
%>
<%
Sub organize

fid = Request.QueryString("fid")
Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM FIXED WHERE fid = "&fid&""
rs.open SQL,Connection,3,3
%>
<script language="JavaScript1.2" defer>
editor_generate('metin');
</script>
<form action="?action=update&fid=<%=fid%>" method="post">
<table border="0" cellspacing="2" cellpadding="2" class="td_menu" align="center" style="font-weight:bold">
<tr bgcolor="#E0E0E0">
<td align="right">Baþlýk : </td>
<td><input type="text" name="baslik" class="form1" value="<%=Oku(rs("sf_baslik"))%>" style="width:100%"></td>
</tr>
<tr bgcolor="#E0E0E0">
<td colspan="2"><textarea name="metin" rows="10" class="form1" style="width:100%"><%=Oku(rs("f_metin"))%></textarea></td>
</tr>
<tr bgcolor="#E0E0E0">
<td align="right">Resim URL : </td>
<td>
<% IF dbType = "personal" Then
Response.Write "Images/"
ELSE
Response.Write "Images/"
End IF %>
<input type="text" name="image" value="<%=Oku(rs("resim"))%>" class="form1" style="width:75%"></td>
</tr>
<tr bgcolor="#E0E0E0">
<td align="right">&nbsp;</td>
<td><input type="submit" class="submit" size="40" value="Güncelle »"></td>
</tr>
</table>
<%
rs.Close
Set rs = Nothing

End Sub
Sub update

fid = Request.QueryString("fid")
Set ent = Server.CreateObject("ADODB.RecordSet")
entSQL = "Select * FROM FIXED WHERE fid = "&fid&""
ent.open entSQL,Connection,1,3

baslik = Request.Form("baslik")
haber = Request.Form("metin")
image = Request.Form("image")

If baslik="" Then
Response.Write "<center><font face=verdana size=2 color=navy>Lütfen Baþlýk Yazýnýz...</font></center>"
ElseIf haber="" Then
Response.Write "<center><font face=verdana size=2 color=navy>Lütfen Haber Metni Yazýnýz...</font></center>"
ELSE

ent("sf_baslik") = baslik
ent("f_metin") = haber
ent("resim") = image
ent.Update
Response.Redirect "?action=all"
END IF

ent.Close
Set ent = Nothing

End Sub
Sub del

fid = Request.QueryString("fid")
Set delete = Server.CreateObject("ADODB.RecordSet")
delSQL = "DELETE * FROM FIXED WHERE fid = "&fid&""
delete.open delSQL,Connection,1,3

Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
End If
set oFO = Nothing
%>