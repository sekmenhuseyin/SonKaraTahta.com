<% Response.Buffer = True %>
<!--#include file="includes.asp" -->
<%
If Session("Level") = "1" OR Session("Level") = "4" OR Session("Level") = "7" Then
klasor = Server.Mappath("../IMAGES/news_cats/")
act = Request.QueryString("action")
If act = "all" Then
call news
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
rs.open "Select * FROM NEWS WHERE onay = True ORDER BY hid DESC",Connection,3,3
IF rs.EoF OR rs.BoF Then
Response.Write "<div align=center class=td_menu><b>Haber YOK !</b></div>"
ELSE
%>
<table border="1" cellspacing="0" cellpadding="2" align="center" width="90%" class="td_menu">
<%
Do Until rs.EoF
Set cat = Connection.Execute("Select * FROM NEWS_CATS WHERE cat_id = "&rs("cat")&"")
%>
<tr>
<td width="70%" style="font-weight:bold"><%=rs("baslik")%> (<%=cat("cat_name")%>)</td>
<td width="30%" align="center" bgcolor="#F4F4F4"><a href="?action=change&op=passive&hid=<%=rs("hid")%>">Pasifleþtir</a> - <a href="?action=organize&hid=<%=rs("hid")%>">Düzenle</a> - <a href="?action=delete&hid=<%=rs("hid")%>">Sil</a></td>
</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
End IF
End Sub
Sub wait_approve
Set rs2 = Server.CreateObject("ADODB.RecordSet")
SQL2 = "Select * FROM NEWS WHERE onay = False ORDER BY hid DESC"
rs2.open SQL2,Connection,1,3
IF rs2.EoF OR rs2.BoF Then
Response.Write "<div align=center class=td_menu><b>Onay Bekleyen Haber Yok !</b></div>"
ELSE
%>
<table border="1" cellspacing="0" cellpadding="2" align="center" width="90%" bgcolor="#FFFFCC" class="td_menu">
<%
Do While Not rs2.Eof
IF rs2("cat") <> "" Then
Set cat = Connection.Execute("Select * FROM NEWS_CATS WHERE cat_id = "&rs2("cat")&"")
cat_name = cat("cat_name")
ELSE
cat_name = "Hata Oluþtu !"
End IF
%>
<tr><td width="70%"><b><%=rs2("baslik")%></b> (<%=cat_name%>)</td><td width="30%" align="center"><a href="?action=change&op=active&hid=<%=rs2("hid")%>">Aktifleþtir</a> - <a href="?action=organize&hid=<%=rs2("hid")%>">Düzenle</a> - <a href="?action=delete&hid=<%=rs2("hid")%>">Sil</a></td></tr>
<%
rs2.MoveNext
Loop
%>
</table>
<%
End IF
rs2.Close
Set rs2 = Nothing
End Sub
Sub cat_delete
Connection.Execute("DELETE FROM NEWS_CATS WHERE cat_id = "&Request.QueryString("id")&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
Sub cat_newreg
f_1 = Request.Form("name")
f_2 = Request.Form("info")
f_3 = Request.Form("n_cat")
f_3 = Replace(f_3, "../", "", 1, -1, 1)
f_3 = "IMAGES/news_cats/" & f_3
IF f_1="" OR f_2="" Then
Response.Write "<div align=center class=td_menu>Lütfen Tüm Alanlarý Doldurunuz</div>"
ElseIF f_1<>"" AND f_2<>"" Then
IF Request.QueryString("x") = "Update" Then
Connection.Execute("UPDATE NEWS_CATS Set cat_name = '"&f_1&"' WHERE cat_id = "&Request.QueryString("id")&"")
Connection.Execute("UPDATE NEWS_CATS Set cat_info = '"&f_2&"' WHERE cat_id = "&Request.QueryString("id")&"")
Connection.Execute("UPDATE NEWS_CATS Set cat_image = '"&f_3&"' WHERE cat_id = "&Request.QueryString("id")&"")
ELSE
Connection.Execute("INSERT INTO NEWS_CATS (cat_name, cat_info, cat_image) VALUES ('"&f_1&"', '"&f_2&"', '"&f_3&"')")
End IF
Response.Redirect "?action=Cats"
End IF
End Sub
Sub cat_new
opt = Request.QueryString("x")
IF opt = "Edit" Then
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM NEWS_CATS WHERE cat_id = "&Request.QueryString("id")&"",Connection,3,3
n_page = "Update"
c_name = rs("cat_name")
c_info = rs("cat_info")
c_img = Replace(rs("cat_image"),"../","",1,-1,1)
c_img = Replace(c_img,"","",1,-1,1)
c_img = Replace(c_img,"news_cats/","",1,-1,1)
btn_submit = "Güncelle"
ELSE
c_img = "blank.gif"
btn_submit = "Ekle +"
End IF
%>
<table border="0" cellspacing="2" cellpadding="2" width="60%" align="center" class="td_menu">
<form method="post" action="?action=Cat_Register&x=<%=n_page%>&id=<%=Request.QueryString("id")%>">
<tr bgcolor="#E0E0E0">
<td width="40%">Kategori Adý : </td>
<td width="60%"><input type="text" name="name" class="form1" style="width:100%" value="<%=c_name%>"></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="40%">Kategori Açýklama : </td>
<td width="60%"><input type="text" name="info" class="form1" style="width:100%" value="<%=c_info%>"></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="40%" valign="top">Kategori Resim : </td>
<td width="60%" align="right">../IMAGES/news_cats/<input type="text" name="n_cat" class="form1" style="width:60%" value="<%=c_img%>"></td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="40%"></td>
<td width="60%"><input type="submit" value="<%=btn_submit%>" class="submit"></td>
</tr>
</form>
</table>
<%
End Sub
Sub cats
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM NEWS_CATS ORDER BY cat_name ASC",Connection,3,3
%>
<table border="1" cellspacing="0" cellpadding="1" align="center" width="90%" class="td_menu">
<% Do While Not rs.Eof %>
<tr>
<td width="70%" style="font-weight:bold"><a href="?action=all&id=<%=rs("cat_id")%>"><%=rs("cat_name")%></a></td>
<td width="30%" align="center" bgcolor="#F4F4F4"><a href="?action=Cat_New&x=Edit&id=<%=rs("cat_id")%>">Düzenle</a> - <a href="?action=Cat_Delete&id=<%=rs("cat_id")%>">Sil</a></td>
</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
rs.Close
Set rs = Nothing
End Sub

Sub news

Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM NEWS WHERE cat = "&Request.QueryString("id")&" AND onay = True ORDER BY hid DESC"
rs.open SQL,Connection,1,3
%>
<table border="1" cellspacing="0" cellpadding="2" align="center" width="90%" class="td_menu">
<% Do While Not rs.Eof %>
<tr>
<td width="70%" style="font-weight:bold"><%=rs("baslik")%></td>
<td width="30%" align="center" bgcolor="#F4F4F4"><a href="?action=change&op=passive&hid=<%=rs("hid")%>">Pasifleþtir</a> - <a href="?action=organize&hid=<%=rs("hid")%>">Düzenle</a> - <a href="?action=delete&hid=<%=rs("hid")%>">Sil</a></td>
</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
rs.Close
Set rs = Nothing

End Sub
Sub enternew
Set n_cats = Connection.Execute("Select * FROM NEWS_CATS ORDER BY cat_name ASC")
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
<td align="right" valign="top">Kategori : </td>
<td>
<select name="cat" size="1" class="form1">
<%
Do Until n_cats.Eof
Response.Write "<option value="""&n_cats("cat_id")&""">"&n_cats("cat_name")&"</option>"
n_cats.MoveNext
Loop
%>
</select>
</td>
</tr>
<tr bgcolor="#E0E0E0">
<td align="right" valign="top">Resim URL : </td>
<td>
<% IF dbType = "personal" Then
Response.Write "Db/"
ELSE
Response.Write ""
End IF
%>
<input name="image" type="text" class="form1" style="width:85%" value="http://"></td>
</tr>
<tr>
<td align="right">&nbsp;</td>
<td align="right"><input type="submit" class="submit" size="40" value="Kaydet »"></td></tr>
</form>
</table>
<br>
<iframe src="?action=IS" width="400" height="80" align="center" border="0" MARGINWIDTH=0 MARGINHEIGHT=0 HSPACE=0 VSPACE=0 FRAMEBORDER=0 SCROLLING=NO></iframe>
<%
End Sub
Sub reg

baslik = Request("baslik")
haber = Request("metin")
cat = Request("cat")
image = Request("image")

If baslik="" Then
Response.Write "<center><font face=verdana size=2 color=navy>Lütfen Baþlýk Yazýnýz...</font></center>"
ElseIf haber="" Then
Response.Write "<center><font face=verdana size=2 color=navy>Lütfen Haber Metni Yazýnýz...</font></center>"
ElseIF cat = "" Then
Response.Write "<center><font face=verdana size=2 color=navy>Lütfen Kategori Seçiniz...</font></center>"
ELSE
Set ent = Server.CreateObject("ADODB.RecordSet")
ent.open "Select * FROM NEWS",Connection,3,3

Session.LCID = 2048
ent.AddNew
ent("baslik") = baslik
ent("metin") = haber
ent("ekleyen") = Session("Member")
ent("resim") = image
ent("tarih") = Now()
ent("onay") = True
ent("puan") = 0
ent("katilimci") = 0
ent("okunma") = 0
ent("cat") = cat
ent.Update
ent.Close
Set ent = Nothing
Response.Redirect "?action=all&id="&cat&""
End IF

End Sub
Sub ImageSelect
%>
<table border="0" cellspacing="0" cellpadding="2" width="100%" class="td_menu">
<tr bgcolor="#000000"><td colspan="2"></td></tr>
<form ENCTYPE="multipart/form-data" ACTION="?action=IU" METHOD="POST">
<tr bgcolor="#E0E0E0">
<td align="right">Resim : </td>
<td><input NAME="File2" class="form1" SIZE="20" TYPE="file"> <input type="submit" value="Gönder »" class="submit"></td>
</tr>
<tr bgcolor="#000000"><td colspan="2"></td></tr>
</form>
</table>
<%
End Sub
Sub ImageUpload
IF dbType = "personal" Then
ImageDir = "../../Db"
ELSE
ImageDir = ""
End IF
	ForWriting = 2
	adLongVarChar = 201
	lngNumberUploaded = 0

	'Get binary data from form		
	noBytes = Request.TotalBytes 
	binData = Request.BinaryRead (noBytes)

	'convery the binary data to a string
	Set RST = CreateObject("ADODB.Recordset")
	LenBinary = LenB(binData)
	
	if LenBinary > 0 Then
	RST.Fields.Append "myBinary", adLongVarChar, LenBinary
	RST.Open
	RST.AddNew
	RST("myBinary").AppendChunk BinData
	RST.Update
	strDataWhole = RST("myBinary")
	End if

	strBoundry = Request.ServerVariables ("HTTP_CONTENT_TYPE")
	lngBoundryPos = instr(1, strBoundry, "boundary=") + 8 
	strBoundry = "--" & right(strBoundry, len(strBoundry) - lngBoundryPos)
	lngCurrentBegin = instr(1, strDataWhole, strBoundry)
	lngCurrentEnd = instr(lngCurrentBegin + 1, strDataWhole, strBoundry) - 1
	Do While lngCurrentEnd > 0
	'Get the data between current boundry and remove it from the whole.
	strData = mid(strDataWhole, lngCurrentBegin, lngCurrentEnd - lngCurrentBegin)
	strDataWhole = replace(strDataWhole, strData,"")
	
	'Get the full path of the current file.
	lngBeginFileName = instr(1, strdata, "filename=") + 10
	lngEndFileName = instr(lngBeginFileName, strData, chr(34)) 
	'Make sure they selected a file.	
	if lngBeginFileName = lngEndFileName and lngNumberUploaded = 0 Then
	Response.Redirect "index.asp?"
	End if
	'There could be an empty file box.	
	if lngBeginFileName <> lngEndFileName Then
	strFilename = mid(strData, lngBeginFileName, lngEndFileName - lngBeginFileName)

	tmpLng = instr(1, strFilename, "\")
	Do While tmpLng > 0
	PrevPos = tmpLng
	tmpLng = instr(PrevPos + 1, strFilename,"\")
	Loop
	
	FileName = right(strFilename, len(strFileName) - PrevPos)
	
	lngCT = instr(1,strData, "Content-Type:")
	
	if lngCT > 0 Then
	lngBeginPos = instr(lngCT, strData, chr(13) & chr(10)) + 4
	Else
	lngBeginPos = lngEndFileName
	End if
	lngEndPos = len(strData) 
	
	'Calculate the file size.	
	lngDataLenth = lngEndPos - lngBeginPos
	'Get the file data	
	strFileData = mid(strData, lngBeginPos, lngDataLenth)

	IF instr(1,FileName,".asp",1) Then
	Response.Write "<div align=center class=td_menu>Hatalý Dosya Türü !</div>"
	ELSE

	'Create the file.	
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set f = fso.OpenTextFile(server.mappath(imagedir) & "/" & FileName, ForWriting, True)
	f.Write strFileData
	Set f = nothing
	Set fso = nothing

	End IF
	
	lngNumberUploaded = lngNumberUploaded + 1
			
	End if
	
	lngCurrentBegin = instr(1, strDataWhole, strBoundry)
	lngCurrentEnd = instr(lngCurrentBegin + 1, strDataWhole, strBoundry) - 1
	loop

	IF instr(1,FileName,".asp",1) Then
	ELSE
Set wSet = Connection.Execute("Select * FROM SETTINGS")
Session("ImagePath") = FileName
wSet.Close
Set wSet = Nothing
With Response
	.Write "<font class=""td_menu"">"
	.Write "<b>Resim Baþarýyla Gönderildi...</b>"
	.Write "<br>NOT : Kýrmýzýyla yazýlmýþ olan resim ismini yukarýdaki Resim URL yerine yazýnýz..."
	.Write "<br><br>"
	.Write "Resmin Adres : <font color=red>" & Session("ImagePath") & "</font>"
	.Write "<br><br>"
	.Write "<a href="&Request.ServerVariables("HTTP_REFERER")&">« Geri</a>"
	.Write "</font>"
End With
	End IF
End Sub
Sub change

operation = Request.QueryString("op")
id = Request.QueryString("hid")

If operation = "active" Then
s = "True"
ElseIf operation = "passive" Then
s = "False"
End If

Set chn = Server.CreateObject("ADODB.RecordSet")
chnSQL = "Select * FROM NEWS WHERE hid = "&id&""
chn.open chnSQL,Connection,1,3

chn("onay") = s
chn.Update
Response.Redirect "?action=all&id="&chn("cat")&""

chn.Close
Set chn = Nothing

End Sub
Sub organize

id = Request.QueryString("hid")
Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM NEWS WHERE hid = "&id&""
rs.open SQL,Connection,1,3

Set s_cats = Connection.Execute("Select * FROM NEWS_CATS ORDER BY cat_name ASC")
%>
<script language="JavaScript1.2" defer>
editor_generate('metin');
</script>
<form action="?action=update&hid=<%=id%>" method="post">
<table border="0" cellspacing="2" cellpadding="2" class="td_menu" align="center" style="font-weight:bold">
<tr bgcolor="#E0E0E0">
<td align="right">Baþlýk : </td>
<td><input type="text" name="baslik" class="form1" value="<%=Oku(rs("baslik"))%>" style="width:100%"></td>
</tr>
<tr bgcolor="#E0E0E0">
<td colspan="2"><textarea name="metin" rows="10" class="form1" style="width:100%"><%=Oku(rs("metin"))%></textarea></td>
</tr>
<tr bgcolor="#E0E0E0">
<td align="right" valign="top">Kategori : </td>
<td>
<select name="cat" size="1" class="form1">
<%
Do Until s_cats.Eof
IF rs("cat") = s_cats("cat_id") Then
opt = "selected"
ELSe
opt = ""
End IF
Response.Write "<option value="""&s_cats("cat_id")&""" "&opt&">"&s_cats("cat_name")&"</option>"
s_cats.MoveNext
Loop
%>
</select>
</td>
</tr>
<tr bgcolor="#E0E0E0">
<td align="right">Resim URL : </td>
<td>
<% IF dbType = "personal" Then
Response.Write "Db/"
ELSE
Response.Write ""
End IF %>
<input type="text" name="image" value="<%=rs("resim")%>" class="form1" style="width:75%"></td>
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

id = Request.QueryString("hid")
Set ent = Server.CreateObject("ADODB.RecordSet")
entSQL = "Select * FROM NEWS WHERE hid = "&id&""
ent.open entSQL,Connection,1,3

baslik = Request.Form("baslik")
haber = Request.Form("metin")
cat = Request.Form("cat")
image = Request.Form("image")

If baslik="" Then
Response.Write "<center><font face=verdana size=2 color=navy>Lütfen Baþlýk Yazýnýz...</font></center>"
ElseIf haber="" Then
Response.Write "<center><font face=verdana size=2 color=navy>Lütfen Haber Metni Yazýnýz...</font></center>"
ELSE

ent("baslik") = baslik
ent("metin") = haber
ent("cat") = cat
ent("resim") = image
ent.Update
Response.Redirect "?action=all&id="&ent("cat")&""
END IF

ent.Close
Set ent = Nothing

End Sub
Sub del

id = Request.QueryString("hid")
Set delete = Server.CreateObject("ADODB.RecordSet")
delSQL = "DELETE * FROM NEWS WHERE hid = "&id&""
delete.open delSQL,Connection,1,3

Set del_comments = Server.CreateObject("ADODB.RecordSet")
delCSQL = "DELETE * FROM COMMENTS WHERE nid = "&id&""
del_comments.open delCSQL,Connection,1,3

Response.Redirect Request.ServerVariables("HTTP_REFERER")
End Sub
End If
set oFO = Nothing
%>