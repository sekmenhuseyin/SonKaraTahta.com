<% Response.Buffer = True %>
<!--#include file="includes.asp" -->
<%
If Session("Level") = "1" OR Session("Level") = "2" OR Session("Level") = "7" Then
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
ElseIf act = "Programs" Then
call programs
ElseIF act = "IU" Then
call ImageUpload
ElseIF act = "New_Cat" Then
call new_cat
ElseIF act = "WaitApprove" Then
call wait_approve
ElseIF act = "NoCats" Then
call no_cats
End If

Sub no_cats
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.open "Select * FROM DOWNLOADS WHERE p_approved = True",Connection,3,3
%>
<table border="1" cellspacing="0" cellpadding="1" align="center" width="80%" bgcolor="#FFFFFF" class="td_menu2" style="font-size:11px">
<tr>
<td width="70%" align="center" bgcolor="#FFFFCC"><font face=tahoma size=2><b>Baþlýk</b></font></td><td width="30%" align="center" bgcolor="#FFFFCC"><font face=tahoma size=2><b>Ýþlem</b></font></td>
</tr>
<%
Do Until rs.EOF
Set cat = Connection.Execute("Select * FROM DOWN_CATS WHERE cid = "&rs("cid")&"")
%>
<tr>
<td width="70%"><b><%=duz(rs("p_name"))%></b> (<%=cat("c_name")%>)</td>
<td width="30%" align="center"><a href="?action=Programs&page=Organize&pid=<%=rs("pid")%>">Düzenle</a> - <a href="?action=Programs&page=Delete&pid=<%=rs("pid")%>">Sil</a></td>
</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
End Sub

Sub wait_approve
Set o = Server.CreateObject("ADODB.RecordSet")
o.open "Select * FROM DOWNLOADS WHERE p_approved = False",Connection,3,3
%>
<table border="1" cellspacing="0" cellpadding="1" align="center" width="80%" bgcolor="#FFFFFF">
<tr style="font-family:arial; font-size:12px; color:navy;">
<td width="70%" bgcolor="#FFFFCC"><b>Programýn Adý</b></td>
<td width="30%" bgcolor="#FFFFCC" align="center"><b>Ýþlem</b></td>
</tr>
<% do while not o.eof %>
<tr style="font-family:verdana; font-size:11px;">
<td width="70%"><%=duz(o("p_name"))%></td>
<td width="30%" align="center"><a href="?action=Programs&page=Approve&pid=<%=o("pid")%>">Onayla</a> - <a href="?action=Programs&page=Delete&pid=<%=o("pid")%>">Reddet - <a href="?action=Programs&page=Organize&pid=<%=o("pid")%>">Düzenle</a></a></td>
</tr>
<%
o.MoveNext
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
Sub ImageUpload
IF dbType = "personal" Then
ImageDir = "../../Db"
ELSE
ImageDir = "../Images"
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
Sub All
Set rs = Connection.Execute("Select * FROM DOWN_CATS ORDER BY c_name DESC") %>
<table border="1" cellspacing="0" cellpadding="0" align="center" width="80%" bgcolor="#FFFFFF">
<tr>
<td width="70%" align="center" bgcolor="#FFFFCC"><font face=tahoma size=2><b>Baþlýk</b></font></td><td width="30%" align="center" bgcolor="#FFFFCC"><font face=tahoma size=2><b>Ýþlem</b></font></td>
</tr>
<% do while not rs.eof %>
<tr>
<td width="70%"><font face=verdana size=2><a href="?action=Programs&page=prgs&catid=<%=rs("cid")%>"><%=rs("c_name")%></a></font></td>
<td width="30%" align="center"><font face=verdana size=1><a href="?action=OrganizeCat&catid=<%=rs("cid")%>">Düzenle</a> - <a href="?action=DeleteCat&catid=<%=rs("cid")%>">Sil</a></td>
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
eSQL = "Select * FROM DOWN_CATS"
ent.open eSQL,Connection,1,3

name = duz(Request.Form("cat_name"))

If name="" Then
Response.Write "<center><font face=verdana size=2>Lütfen Kategori Adýný Yazýnýz...</font></center>"
ELSE

ent.AddNew
ent("c_name") = name
ent.Update
Response.Redirect "?action=all"
END IF
End Sub
Sub Organize_Cat

id = Request.QueryString("catid")
Set rs = Connection.Execute("Select * FROM DOWN_CATS WHERE cid = "&id&"")
%>
<form action="?action=UpdateCat&catid=<%=id%>" method="post">
<table border="0" cellspacing="0" cellpadding="1" width="50%" align="center">
<tr>
<td width="40%" align="right"><font face=verdana size=1>Kategori Adý : </font></td><td width="60%"><input type="text" name="cat_name" class="form1" size="20" value="<%=rs("c_name")%>"></td>
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
eSQL = "Select * FROM DOWN_CATS WHERE cid = "&id&""
ent.open eSQL,Connection,1,3

name = duz(Request.Form("cat_name"))

If name="" Then
Response.Write "<center><font face=verdana size=2>Lütfen Kategori Adýný Yazýnýz...</font></center>"
ELSE

ent("c_name") = name
ent.Update
Response.Redirect "?action=all"
END IF
End Sub
Sub Del
id = Request.QueryString("catid")
Set delete = Server.CreateObject("ADODB.RecordSet")
delSQL = "DELETE * FROM DOWN_CATS WHERE cid = "&id&""
delete.open delSQL,Connection,1,3
Set delete2 = Server.CreateObject("ADODB.RecordSet")
delSQL2 = "DELETE * FROM DOWNLOADS WHERE cid = "&id&""
delete2.open delSQL2,Connection,1,3
Response.Redirect "?action=all"
End Sub
Sub programs
p = Request.QueryString("page")
If p = "prgs" Then

id = Request.QueryString("catid")
Set rs = Connection.Execute("Select * FROM DOWNLOADS WHERE cid = "&id&" AND p_approved = True")
%>
<table border="1" cellspacing="0" cellpadding="0" align="center" width="80%" bgcolor="#FFFFFF">
<tr>
<td width="70%" align="center" bgcolor="#FFFFCC"><font face=tahoma size=2><b>Baþlýk</b></font></td><td width="30%" align="center" bgcolor="#FFFFCC"><font face=tahoma size=2><b>Ýþlem</b></font></td>
</tr>
<% do while not rs.eof %>
<tr>
<td width="70%"><font face=verdana size=2><%=duz(rs("p_name"))%></font></td>
<td width="30%" align="center"><font face=verdana size=1><a href="?action=Programs&page=Organize&pid=<%=rs("pid")%>">Düzenle</a> - <a href="?action=Programs&page=Delete&pid=<%=rs("pid")%>">Sil</a></td>
</tr>
<%
rs.MoveNext
Loop
%>
</table>
<%
ElseIf  p = "New" Then
Set lcats = Connection.Execute("Select * FROM DOWN_CATS")
IF Session("ImagePath") = "" Then
Session("ImagePath") = ""
End IF
%>
<form action="?action=Programs&page=Reg" method="post">
<table border="0" cellspacing="0" cellpadding="2" width="90%" style="font-family:verdana; font-size:10px;">
<tr>
<td width="40%" align="right">Programýn Adý : </td>
<td width="60%"><input type="text" name="pname" size="30" class="form1"></td>
</tr>
<tr>
<td width="40%" align="right" valign="top">Açýklama : </td>
<td width="60%"><textarea name="pinfo" rows="5" cols="29" class="form1"></textarea></td>
</tr>
<tr>
<td width="40%" align="right">Dosya Boyutu : </td>
<td width="60%"><input type="text" name="psize" size="30" class="form1" value="0 KB"></td>
</tr>
<tr>
<td width="40%" align="right">Download Adresi : </td>
<td width="60%"><input type="text" name="purl" size="30" class="form1" value="http://"></td>
</tr>
<tr>
<td width="40%" align="right">Telif : </td>
<td width="60%"><input type="text" name="pauthor" size="30" class="form1"></td>
</tr>
<tr>
<td width="40%" align="right">Kategori : </td>
<td width="60%">
<select name="pcat" size="1" class="form1">
<% do while not lcats.eof %>
<option value=<%=lcats("cid")%>><%=lcats("c_name")%></option>
<%
lcats.MoveNext
Loop
%>
</td>
</tr>
<tr><td width="40%"></td><td width="60%"><b><a href="?action=IS" target="_blank"><b>Resim Ekle</b></a></td></tr>
<tr>
<td width="40%" align="right">&nbsp;</td>
<td width="60%"><input type="submit" class="submit" value="K a y d e t"></td>
</tr>
</form>
<form ENCTYPE="multipart/form-data" ACTION="?action=IU" METHOD="POST">
<tr><td colspan="2">&nbsp;</td></tr>
<tr bgcolor="#F0F0F0">
<td width="40%" align="right">Resim : </td>
<td width="60%"><input NAME="File2" class="form1" SIZE="20" TYPE="file"> <input type="submit" value="Yükle »" class="submit"></td>
</tr>
</table>
<%
ElseIf  p = "Reg" Then
name = Request.Form("pname")
info = Request.Form("pinfo")
size = Request.Form("psize")
url = Request.Form("purl")
author = Request.Form("pauthor")
cat = Request.Form("pcat")

If name="" Then
Response.Write "<font face=verdana size=2 color="&text&">Lütfen Programýn Ýsmini Yazýnýz</font>"
ElseIf info="" Then
Response.Write "<font face=verdana size=2 color="&text&">Lütfen Programýn Açýklamasýný Yazýnýz</font>"
ElseIf url="" Then
Response.Write "<font face=verdana size=2 color="&text&">Lütfen Programýn Download Adresini Yazýnýz</font>"
ElseIf author="" Then
Response.Write "<font face=verdana size=2 color="&text&">Lütfen Programýn Yapýmcýsýný Yazýnýz</font>"
ELSE

Set pent = Server.CreateObject("ADODB.RecordSet")
peSQL = "Select * FROM DOWNLOADS"
pent.open peSQL,Connection,1,3

pent.AddNew
pent("p_name") = name
pent("p_info") = info
pent("p_size") = size
pent("p_hit") = 0
pent("p_url") = url
pent("cid") = cat
pent("p_img") = Session("ImagePath")
pent("p_date") = Date()
pent("p_author") = author
pent("point") = 0
pent("p_approved") = True
pent.Update
Response.Redirect Request.ServerVariables("HTTP_REFERER")
END IF
ElseIf p = "Organize" Then
id = Request.QueryString("pid")
Set lcats = Connection.Execute("Select * FROM DOWN_CATS")
Set pr = Connection.Execute("Select * FROM DOWNLOADS WHERE pid ="&id&"")
%>
<form action="?action=Programs&page=Update&pid=<%=id%>" method="post">
<table border="0" cellspacing="0" cellpadding="2" width="90%" style="font-family:verdana; font-size:10px;">
<tr>
<td width="40%" align="right">Programýn Adý : </td>
<td width="60%"><input type="text" name="pname" size="30" class="form1" value="<%=pr("p_name")%>"></td>
</tr>
<tr>
<td width="40%" align="right" valign="top">Açýklama : </td>
<td width="60%"><textarea name="pinfo" rows="5" cols="29" class="form1"><%=Oku(pr("p_info"))%></textarea></td>
</tr>
<tr>
<td width="40%" align="right">Dosya Boyutu : </td>
<td width="60%"><input type="text" name="psize" size="30" class="form1" value="<%=pr("p_size")%>"></td>
</tr>
<tr>
<td width="40%" align="right">Download Adresi : </td>
<td width="60%"><input type="text" name="purl" size="30" class="form1" value="<%=pr("p_url")%>"></td>
</tr>
<tr>
<td width="40%" align="right">Telif : </td>
<td width="60%"><input type="text" name="pauthor" size="30" class="form1" value="<%=pr("p_author")%>"></td>
</tr>
<tr>
<td width="40%" align="right">Hit : </td>
<td width="60%"><input type="text" name="phit" size="30" class="form1" value="<%=pr("p_hit")%>"></td>
</tr>
<tr>
<td width="40%" align="right">Kategori : </td>
<td width="60%">
<select name="pcat" size="1" class="form1">
<% do while not lcats.eof
If pr("cid") = lcats("cid") Then
sel = "selected"
Else
sel = ""
End If
%>
<option value="<%=lcats("cid")%>" <%=sel%>><%=lcats("c_name")%></option>
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
id = Request.QueryString("pid")
name = Request.Form("pname")
info = Request.Form("pinfo")
size = Request.Form("psize")
url = Request.Form("purl")
author = Request.Form("pauthor")
cat = Request.Form("pcat")
hit = Request.Form("phit")

If name="" Then
Response.Write "<font face=verdana size=2 color="&text&">Lütfen Programýn Ýsmini Yazýnýz</font>"
ElseIf info="" Then
Response.Write "<font face=verdana size=2 color="&text&">Lütfen Programýn Açýklamasýný Yazýnýz</font>"
ElseIf url="" Then
Response.Write "<font face=verdana size=2 color="&text&">Lütfen Programýn Download Adresini Yazýnýz</font>"
ElseIf author="" Then
Response.Write "<font face=verdana size=2 color="&text&">Lütfen Programýn Yapýmcýsýný Yazýnýz</font>"
ELSE

Set pent = Server.CreateObject("ADODB.RecordSet")
peSQL = "Select * FROM DOWNLOADS where pid = "&id&""
pent.open peSQL,Connection,1,3

pent("p_name") = name
pent("p_info") = info
pent("p_size") = size
pent("p_hit") = hit
pent("p_url") = url
pent("cid") = cat
pent("p_author") = author
pent.Update
Response.Redirect Request.ServerVariables("HTTP_REFERER")
END IF
ElseIf p = "Delete" Then
id = Request.QueryString("pid")
Set delPrg = Server.CreateObject("ADODB.RecordSet")
dpSQL = "DELETE * FROM DOWNLOADS WHERE pid = "&id&""
delPrg.open dpSQL,Connection,1,3
Response.Redirect Request.ServerVariables("HTTP_REFERER")
ElseIf p = "Approve" Then
id = Request.QueryString("pid")
Set upd = Connection.Execute("UPDATE DOWNLOADS SET p_approved = True WHERE pid = "&id&"")
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End If
Response.Write "<center><font face=verdana size=1><a href=javascript:history.back();><< Geri</a></font></center>"
End Sub
End If
%>