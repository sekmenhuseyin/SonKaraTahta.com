<% Response.Buffer = True %>
<!--#include file="includes.asp" -->
<%
If Session("Level") = "1" OR Session("Level") = "3" OR Session("Level") = "7"  Then
act = Request.QueryString("action")
If act = "all" Then
call All
ElseIf act = "RegisterNew" Then
call Reg_New
ElseIf act = "DeleteCat" Then
call Del
ElseIf act = "OrganizeCat" Then
call Organize_Cat
ElseIf act = "UpdateCat" Then
call update_cat
ElseIf act = "Articles" Then
call articles
ElseIF act = "WaitApprove" Then
call wait_approve
ElseIF act = "NoCats" Then
call no_cats
ElseIF act = "NewCat" Then
call new_cat
End If

Sub no_cats
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM ARTICLES WHERE a_approved = True",Connection,3,3
%>
<table border="1" cellspacing="0" cellpadding="1" align="center" width="80%" bgcolor="#FFFFFF" class="td_menu2" style="font-size:11px">
<tr>
<td width="70%" align="center" bgcolor="#FFFFCC"><font face=tahoma size=2><b>Baþlýk</b></font></td><td width="30%" align="center" bgcolor="#FFFFCC"><font face=tahoma size=2><b>Ýþlem</b></font></td>
</tr>
<%
Do Until rs.EoF
Set cat = Connection.Execute("Select * FROM ARTICLE_CATS WHERE cat_id = "&rs("cat_id")&"")
%>
<tr>
<td width="70%"><b><%=duz(rs("a_title"))%></b> (<%=cat("cat_name")%>)</td>
<td width="30%" align="center"><a href="?action=Articles&page=Organize&aid=<%=rs("aid")%>">Düzenle</a> - <a href="?action=Articles&page=Delete&aid=<%=rs("aid")%>">Sil</a></td>
</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
End Sub
Sub new_cat
%>
<form action="?action=RegisterNew" method="post">
<table border="0" cellspacing="0" cellpadding="1" width="50%" align="center">
<tr>
<td width="40%" align="right"><font face=verdana size=1>Kategori Adý : </font></td><td width="60%"><input type="text" name="cat_name" class="form1" size="20"></td>
</tr>
<tr>
<td width="40%">&nbsp;</td><td width="60%"><input type="submit" value="K a y d e t" class="submit"></td>
</tr>
</table>
</form>
<%
End Sub
Sub wait_approve
Set o = Server.CreateObject("ADODB.RecordSet")
o.open "Select * FROM ARTICLES WHERE a_approved = False",Connection,3,3
%>
<table border="1" cellspacing="0" cellpadding="1" align="center" width="80%" bgcolor="#FFFFFF">
<tr style="font-family:arial; font-size:12px; color:navy;">
<td width="70%" bgcolor="#FFFFCC"><b>Baþlýk</b></td>
<td width="30%" bgcolor="#FFFFCC" align="center"><b>Ýþlem</b></td>
</tr>
<% Do Until o.EoF %>
<tr style="font-family:verdana; font-size:11px;">
<td width="70%"><%=duz(o("a_title"))%></td>
<td width="30%" align="center"><a href="?action=Articles&page=Approve&aid=<%=o("aid")%>">Onayla</a> - <a href="?action=Articles&page=Delete&aid=<%=o("aid")%>">Reddet</a> - <a href="?action=Articles&page=Organize&aid=<%=o("aid")%>">Düzenle</a></td>
</tr>
<%
o.MoveNext
Loop
%>
</table>
<%
End Sub
Sub All
Set rs = Connection.Execute("Select * FROM ARTICLE_CATS ORDER BY cat_name ASC") %>
<table border="1" cellspacing="0" cellpadding="0" align="center" width="80%" bgcolor="#FFFFFF">
<tr>
<td width="70%" align="center" bgcolor="#FFFFCC"><font face=tahoma size=2><b>Baþlýk</b></font></td><td width="30%" align="center" bgcolor="#FFFFCC"><font face=tahoma size=2><b>Ýþlem</b></font></td>
</tr>
<% do while not rs.eof %>
<tr>
<td width="70%"><font face=verdana size=2><a href="?action=Articles&page=arts&catid=<%=rs("cat_id")%>"><%=rs("cat_name")%></a></font></td>
<td width="30%" align="center"><font face=verdana size=1><a href="?action=OrganizeCat&catid=<%=rs("cat_id")%>">Düzenle</a> - <a href="?action=DeleteCat&catid=<%=rs("cat_id")%>">Sil</a></td>
</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
End Sub
Sub Reg_New
Set ent = Server.CreateObject("ADODB.RecordSet")
eSQL = "Select * FROM ARTICLE_CATS"
ent.open eSQL,Connection,1,3

name = duz(Request.Form("cat_name"))

If name="" Then
Response.Write "<center><font face=verdana size=2>Lütfen Kategori Adýný Yazýnýz...</font></center>"
ELSE

ent.AddNew
ent("cat_name") = name
ent.Update
Response.Redirect "?action=all"
END IF
End Sub
Sub Organize_Cat

id = Request.QueryString("catid")
Set rs = Connection.Execute("Select * FROM ARTICLE_CATS WHERE cat_id = "&id&"")
%>
<form action="?action=UpdateCat&catid=<%=id%>" method="post">
<table border="0" cellspacing="0" cellpadding="1" width="50%" align="center">
<tr>
<td width="40%" align="right"><font face=verdana size=1>Kategori Adý : </font></td><td width="60%"><input type="text" name="cat_name" class="form1" size="20" value="<%=rs("cat_name")%>"></td>
</tr>
<tr>
<td width="40%">&nbsp;</td><td width="60%"><input type="submit" value="G ü n c e l l e" class="submit"></td>
</tr>
</table>
</form>
<%
End Sub
Sub update_cat
id = Request.QueryString("catid")
Set ent = Server.CreateObject("ADODB.RecordSet")
eSQL = "Select * FROM ARTICLE_CATS WHERE cat_id = "&id&""
ent.open eSQL,Connection,1,3

name = Request.Form("cat_name")

If name="" Then
Response.Write "<center><font face=verdana size=2>Lütfen Kategori Adýný Yazýnýz...</font></center>"
ELSE

ent("cat_name") = name
ent.Update
Response.Redirect "?action=all"
END IF
End Sub
Sub Del
id = Request.QueryString("catid")
Set delete = Server.CreateObject("ADODB.RecordSet")
delSQL = "DELETE * FROM ARTICLE_CATS WHERE cat_id = "&id&""
delete.open delSQL,Connection,1,3
Set delete2 = Server.CreateObject("ADODB.RecordSet")
delSQL2 = "DELETE * FROM ARTICLES WHERE cat_id = "&id&""
delete2.open delSQL2,Connection,1,3
Response.Redirect "?action=all"
End Sub
Sub articles
p = Request.QueryString("page")
If p = "arts" Then

id = Request.QueryString("catid")
Set rs = Connection.Execute("Select * FROM ARTICLES WHERE cat_id = "&id&" AND a_approved = True")
%>
<table border="1" cellspacing="0" cellpadding="0" align="center" width="80%" bgcolor="#FFFFFF">
<tr>
<td width="70%" align="center" bgcolor="#FFFFCC"><font face=tahoma size=2><b>Baþlýk</b></font></td><td width="30%" align="center" bgcolor="#FFFFCC"><font face=tahoma size=2><b>Ýþlem</b></font></td>
</tr>
<% do while not rs.eof %>
<tr>
<td width="70%"><font face=verdana size=2><%=duz(rs("a_title"))%></font></td>
<td width="30%" align="center"><font face=verdana size=1><a href="?action=Articles&page=Organize&aid=<%=rs("aid")%>">Düzenle</a> - <a href="?action=Articles&page=Delete&aid=<%=rs("aid")%>">Sil</a></td>
</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
ElseIf p = "New" Then
Set lcats = Connection.Execute("Select * FROM ARTICLE_CATS")
%>
<form action="?action=Articles&page=NewReg" method="post">
<table border="0" cellspacing="0" cellpadding="2" width="90%" class="td_menu2">
<tr>
<td width="40%" align="right">Baþlýk : </td>
<td width="60%"><input type="text" name="atitle" size="60" class="form1"></td>
</tr>
<tr>
<td width="40%" align="right" valign="top">Makale : </td>
<td width="60%"><textarea name="aarticle" rows="10" cols="60" class="form1"></textarea></td>
</tr>
<tr>
<td width="40%" align="right">Kategori : </td>
<td width="60%">
<select name="acat" size="1" class="form1">
<% Do While NOT lcats.Eof %>
<option value="<%=lcats("cat_id")%>"><%=lcats("cat_name")%></option>
<%
lcats.MoveNext
Loop
%>
</td>
</tr>
<tr>
<td width="40%" align="right">&nbsp;</td>
<td width="60%"><input type="submit" class="submit" value="K a y d e t »"></td>
</tr>
</table>
</form>
<%
ElseIf p = "Organize" Then
id = Request.QueryString("aid")
Set lcats = Connection.Execute("Select * FROM ARTICLE_CATS")
Set pr = Connection.Execute("Select * FROM ARTICLES WHERE aid ="&id&"")
%>
<form action="?action=Articles&page=Update&aid=<%=id%>" method="post">
<table border="0" cellspacing="0" cellpadding="2" width="90%" style="font-family:verdana; font-size:10px;">
<tr>
<td width="40%" align="right">Baþlýk : </td>
<td width="60%"><input type="text" name="atitle" size="60" class="form1" value="<%=pr("a_title")%>"></td>
</tr>
<tr>
<td width="40%" align="right" valign="top">Makale : </td>
<td width="60%"><textarea name="aarticle" rows="10" cols="60" class="form1"><%=pr("a_article")%></textarea></td>
</tr>
<tr>
<td width="40%" align="right">Kategori : </td>
<td width="60%">
<select name="acat" size="1" class="form1">
<%
Do While NOT lcats.Eof
If pr("cat_id") = lcats("cat_id") Then
sel = "selected"
Else
sel = ""
End If
%>
<option value="<%=lcats("cat_id")%>" <%=sel%>><%=lcats("cat_name")%></option>
<%
lcats.MoveNext
Loop
%>
</td>
</tr>
<tr>
<td width="40%" align="right">&nbsp;</td>
<td width="60%"><input type="submit" class="submit" value="G ü n c e l l e"></td>
</tr>
</table>
</form>
<%
ElseIf p = "Update" Then
id = Request.QueryString("aid")
title = Request.Form("atitle")
article = Request.Form("aarticle")
cat = Request.Form("acat")

If title="" Then
Response.Write "<font face=verdana size=2 color="&text&">Lütfen Makale Baþlýðýný Yazýnýz</font>"
ElseIf article="" Then
Response.Write "<font face=verdana size=2 color="&text&">Lütfen Makaleyi Yazýnýz</font>"
ELSE

Set pent = Server.CreateObject("ADODB.RecordSet")
peSQL = "Select * FROM ARTICLES where aid = "&id&""
pent.open peSQL,Connection,1,3

pent("a_title") = title
pent("a_article") = article
pent("cat_id") = cat
pent.Update
Response.Redirect "?action=Articles&page=arts&catid="&pent("cat_id")&""
END IF
ElseIF p = "NewReg" Then

title = Request.Form("atitle")
article = Request.Form("aarticle")
cat = Request.Form("acat")

If title="" Then
Response.Write "<font face=verdana size=2 color="&text&">Lütfen Makale Baþlýðýný Yazýnýz</font>"
ElseIf article="" Then
Response.Write "<font face=verdana size=2 color="&text&">Lütfen Makaleyi Yazýnýz</font>"
ElseIf cat="" Then
Response.Write "<font face=verdana size=2 color="&text&">Lütfen Kategori Seçiniz</font>"
ELSE

Set pent = Server.CreateObject("ADODB.RecordSet")
peSQL = "Select * FROM ARTICLES"
pent.open peSQL,Connection,1,3

pent.AddNew
pent("a_title") = title
pent("a_article") = article
pent("hit") = 0
pent("cat_id") = cat
pent("a_date") = Date()
pent("a_writer") = Session("Member")
pent("point") = 0
pent("a_approved") = True
pent.Update
Response.Redirect "?action=all"
END IF
ElseIf p = "Delete" Then
id = Request.QueryString("aid")
Set delPrg = Server.CreateObject("ADODB.RecordSet")
dpSQL = "DELETE * FROM ARTICLES WHERE aid = "&id&""
delPrg.open dpSQL,Connection,1,3
Response.Redirect Request.ServerVariables("HTTP_REFERER")
ElseIf p = "Approve" Then
id = Request.QueryString("aid")
Set upd = Connection.Execute("UPDATE ARTICLES SET a_approved = True WHERE aid = "&id&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End If
Response.Write "<center><font face=verdana size=1><a href=javascript:history.back();><< Geri</a></font></center>"
End Sub
End If
%>