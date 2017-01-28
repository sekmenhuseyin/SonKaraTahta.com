<% Response.Buffer = True %>
<!--#include file="includes.asp" -->
<%
If Session("Level") = "1" OR Session("Level") = "7" Then

Response.Write "<center><font face=arial size=3 color=navy><b>M O D Ü L L E R</b></font></center>"
act = Request.QueryString("action")
If act = "all" Then
call all
ElseIf act = "EditModule" Then
call edit
ElseIf act = "UpdateModule" Then
call update
ElseIf act = "ChangeModule" Then
call change
ElseIf act = "DeleteModule" Then
call delete
ElseIf act = "Register" Then
call reg
ElseIf act = "Active" Then
call active
ElseIf act = "Passive" Then
call passive
ElseIF act = "New" Then
call new_module
End If

Sub new_module
Set mcats = Connection.Execute("Select * FROM MENU_CATS ORDER BY mc_title ASC")
%>
<form action="?action=Register" method="post">
<table border="0" cellspacing="1" cellpadding="1" width="50%" align="center" bgcolor="#CCCCCC" style="font-family:verdana; font-size:11px;">
<tr>
<td width="40%" align="right" bgcolor="#FFFFCC">Modül Ýsmi : </td>
<td width="60%" bgcolor="#FFFFFF"><input type="text" name="mname" class="form1" size="20"></td>
</tr>
<tr>
<td width="40%" align="right" bgcolor="#FFFFCC">Modül Dosyasý : </td>
<td width="60%" bgcolor="#FFFFFF">
Örnegin: (send_news) <input type="text" name="mdir" size="20" class="form1">
</td>
</tr>
<tr>
<td width="40%" align="right" bgcolor="#FFFFCC">Kategori : </td>
<td width="60%" bgcolor="#FFFFFF">
<select name="mcat" size="1" class="form1">
<% Do Until mcats.EoF %>
<option value="<%=mcats("mc_id")%>"><%=mcats("mc_title")%></option>
<%
mcats.MoveNext
Loop
%>
</select><br><font face="verdana" size="1">( Kategori gözükmüyorsa lütfen "Yeni Kategori" Ekleyin! )</font>
</td>
</tr>
<tr>
<td width="40%" align="right" bgcolor="#FFFFCC">Sadece Üyeler Görebilsin : </td>
<td width="60%" bgcolor="#FFFFFF">
<input type="checkbox" name="mems"></td>
</tr>
<tr>
<td colspan="2" bgcolor="#FFFFFF"><input type="submit" value="Kaydet" class="submit" style="width:100%"></td>
</tr>
</table>
</form>
<%
End Sub

Sub passive
mdlid = Request.QueryString("id")
Set activation = Connection.Execute("UPDATE MODULES SET mactive = False WHERE module_id = "&mdlid&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
Sub active
mdlid = Request.QueryString("id")
Set activation = Connection.Execute("UPDATE MODULES SET mactive = True WHERE module_id = "&mdlid&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
Sub delete
mdlid = Request.QueryString("id")
Set dm = Server.CreateObject("ADODB.RecordSet")
dmSQL = "DELETE * FROM MODULES WHERE module_id = "&mdlid&""
dm.open dmSQL,Connection,1,3
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
Sub all
Set mds = Connection.Execute("Select * FROM MODULES")
%>
<table border="1" cellspacing="0" cellpadding="1" align="center" width="80%">
<tr style="font-family:tahoma; font-size:11px; font-weight:bold" bgcolor="#FFFFCC" height="20">
<td width="40%">Modül Ýsmi</td>
<td width="20%" align="center">Kategori</td>
<td width="15%" align="center">Durum</td>
<td width="25%" align="center">Ýþlem</td>
</tr>
<%
Do While NOT mds.Eof
Set cat = Connection.Execute("Select * FROM MENU_CATS WHERE mc_id = "&mds("mcat")&"")
%>
<tr style="font-family:verdana; font-size:11px;">
<td width="40%"><%=mds("mname")%></td>
<td width="20%" align="center"><%=cat("mc_title")%></td>
<td width="15%" align="center"><% If mds("mactive") = True Then %><font color=red>Aktif</font>&nbsp;<a href="?action=Passive&id=<%=mds("module_id")%>"><img src="IMAGES/Passive.gif" border="0" alt="Pasifleþtirmek Ýçin Týklayýn" align="absmiddle"></a><%Else%><font color=red>Pasif</font>&nbsp;<a href="?action=Active&id=<%=mds("module_id")%>"><img src="IMAGES/Active.gif" border="0" alt="Aktifleþtirmek Ýçin Týklayýn" align="absmiddle"></a><% End If %></td>
<td width="25%" align="center"><a href="?action=EditModule&id=<%=mds("module_id")%>">Düzenle</a> - <a href="?action=DeleteModule&id=<%=mds("module_id")%>">Sil</a></td>
</tr>
<%
mds.MoveNext
Loop
%>
</table>
<%
End Sub
Sub reg
name = duz(Request.Form("mname"))
dir = duz(Request.Form("mdir"))
cat = Request.Form("mcat")

If Request.Form("mems") = "on" Then
mem = True
Else
mem = False
End If

If name="" Then
Response.Write "<center><font face=verdana size=2 color=navy>Lütfen Modül Ýsmini Yazýnýz...</font></center>"
ElseIf dir="" Then
Response.Write "<center><font face=verdana size=2 color=navy>Lütfen Modül Dosya Adresini Yazýnýz...</font></center>"
ELSE

Set ent = Server.CreateObject("ADODB.RecordSet")
eSQL = "Select * FROM MODULES"
ent.open eSQL,Connection,1,3

ent.AddNew
ent("mname") = name
ent("mdir") = dir
ent("mems") = mem
ent("mcat") = cat
ent.Update
Response.Write "<center><font face=verdana size=2 color=navy>Modül Eklendi...<br><br><a href=?action=all><< Geri</a></font></center>"
END IF
End Sub
Sub edit
mdid = Request.QueryString("id")
Set mdl = Connection.Execute("Select * FROM MODULES WHERE module_id = "&mdid&"")
Set cats = Connection.Execute("Select * FROM MENU_CATS ORDER BY mc_title ASC")
%>
<form action="?action=UpdateModule&id=<%=mdid%>" method="post">
<table border="0" cellspacing="1" cellpadding="1" width="50%" align="center" bgcolor="#CCCCCC" style="font-family:verdana; font-size:11px;">
<tr>
<td width="40%" align="right" bgcolor="#FFFFCC">Modül Ýsmi : </td>
<td width="60%" bgcolor="#FFFFFF"><input type="text" name="mname" class="form1" size="20" value="<%=mdl("mname")%>"></td>
</tr>
<tr>
<td width="40%" align="right" bgcolor="#FFFFCC">Modül Dosyasý : </td>
<td width="60%" bgcolor="#FFFFFF">
<input type="text" name="mdir" size="1" class="form1" value="<%=mdl("mdir")%>">
</td>
</tr>
<tr>
<td width="40%" align="right" bgcolor="#FFFFCC">Kategori : </td>
<td width="60%" bgcolor="#FFFFFF">
<select name="mcat" size="1" class="form1">
<%
Do Until cats.EoF
IF cats("mc_id") = mdl("mcat") Then
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
</select><font face="verdana" size="1">( Kategori gözükmüyorsa lütfen "Yeni Kategori" Ekleyin! )</font>
</td>
</tr>
<tr>
<td width="40%" align="right" bgcolor="#FFFFCC">Sadece Üyeler Görebilsin : </td>
<td width="60%" bgcolor="#FFFFFF">
<%
If mdl("mems") = True Then
modS = "checked"
ElseIf mdl("mems") = False Then
modS = ""
End If
%>
<input type="checkbox" name="mems" <%=modS%>></td>
</tr>
<tr>
<td colspan="2" bgcolor="#FFFFFF"><input type="submit" value="Güncelle" class="submit" style="width:100%"></td>
</tr>
</table>
</form>
<%
End Sub
Sub update
mdid = Request.QueryString("id")
name = duz(Request.Form("mname"))
dir = duz(Request.Form("mdir"))

If Request.Form("mems") = "on" Then
mem = True
Else
mem = False
End If

If name="" Then
Response.Write "<center><font face=verdana size=2 color=navy>Lütfen Modül Ýsmini Yazýnýz...</font></center>"
ElseIf dir="" Then
Response.Write "<center><font face=verdana size=2 color=navy>Lütfen Modül Dosya Adresini Yazýnýz...</font></center>"
ELSE

Set ent = Server.CreateObject("ADODB.RecordSet")
eSQL = "Select * FROM MODULES WHERE module_id = "&mdid&""
ent.open eSQL,Connection,1,3

ent("mname") = name
ent("mdir") = dir
ent("mems") = mem
ent.Update
Response.Write "<center><font face=verdana size=2 color=navy>Modül Baþarýyla Güncellendi...<br><br><a href=?action=all><< Geri</a></font></center>"
END IF
End Sub
End If
%>