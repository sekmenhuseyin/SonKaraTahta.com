<% Response.Buffer = True %>
<!--#include file="includes.asp" -->
<% 
IF Session("Level") = "1" OR Session("Level") = "7" Then
x = Request.QueryString("op")

	IF x = "New" Then
Set cats = Server.CreateObject("ADODB.Recordset")
cats.Open "Select * FROM MENU_CATS ORDER BY mc_title ASC",Connection,3,3
%>
<table border="0" cellspacing="2" cellpadding="0" width="50%" align="center" class="td_menu" style="font-weight:bold">
<form action="?op=Enter" method="post">
<tr align="right">
<td width="40%" bgcolor="#FFCC00">Baþlýk : </td>
<td width="60%"><input type="text" name="baslik" class="form1" style="width:100%"></td>
</tr>
<tr align="right">
<td width="40%" bgcolor="#FFCC00">Link : </td>
<td width="60%"><input type="text" name="url" class="form1" style="width:100%"></td>
</tr>
<tr align="right">
<td width="40%" bgcolor="#FFCC00">Açýlýþ Þekli : </td>
<td width="60%">
<select name="acilis" size="1" class="form1" style="width:100%">
<option value="blank">Yeni Sayfa</option>
<option value="normal">Normal Site Düzeni</option>
</select>
</td>
</tr>
<tr align="right">
<td width="40%" bgcolor="#FFCC00">Aktif/Pasif : </td>
<td width="60%" align="left"><input type="checkbox" name="durum"></td>
</tr>
<tr align="right">
<td width="40%" bgcolor="#FFCC00">Kategori : </td>
<td width="60%">
<select name="kategori" size="1" class="form1" style="width:100%">
<% Do Until cats.EoF %>
<option value="<%=cats("mc_id")%>"><%=cats("mc_title")%></option>
<%
cats.MoveNext
Loop
%>
</select>
</td>
</tr>
<tr align="right" bgcolor="#FF9000">
<td colspan="2"><input type="submit" value="Ekle »" class="submit"></td>
</tr>
</form>
</table>
<%
cats.Close
Set cats = Nothing
	ElseIF x = "Enter" Then

baslik = Replace(Request.Form("baslik"),"'","´",1,-1,1)
link = Replace(Request.Form("url"),"'","´",1,-1,1)
acilis = Request.Form("acilis")

IF Request.Form("durum") = "on" Then
durum = "True"
ELSE
durum = "False"
End IF

Set yeni = Server.CreateObject("ADODB.RecordSet")
yeni.Open "Select * FROM MENU_LINKS",Connection,3,3
yeni.AddNew
yeni("m_name") = baslik
yeni("m_url") = link
yeni("m_style") = acilis
yeni("m_status") = durum
yeni("m_cat") = Request.Form("kategori")
yeni.Update
yeni.Close
Set yeni = Nothing

Response.Redirect Request.ServerVariables("URL")

	ElseIF x = "Edit" Then
strName = Replace(Request.QueryString("Name"),"'","´",1,-1,1)
strURL = Replace(Request.QueryString("URL"),"'","´",1,-1,1)
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.Open "Select * FROM MENU_LINKS WHERE m_name = '"&strName&"' AND m_url = '"&strURL&"'",Connection,3,3
IF rs("m_style") = "blank" Then
strOpt10 = "Selected"
ELSe
strOpt11 = "Selected"
End IF
IF rs("m_status") = True Then
strOpt20 = "Checked"
End IF
Set cats = Server.CreateObject("ADODB.Recordset")
cats.Open "Select * FROM MENU_CATS ORDER BY mc_title ASC",Connection,3,3
%>
<table border="0" cellspacing="2" cellpadding="0" width="50%" align="center" class="td_menu" style="font-weight:bold">
<form action="?op=Update&Name=<%=strName%>&URL=<%=strURL%>" method="post">
<tr align="right">
<td width="40%" bgcolor="#FFCC00">Baþlýk : </td>
<td width="60%"><input type="text" name="baslik" class="form1" style="width:100%" value="<%=rs("m_name")%>"></td>
</tr>
<tr align="right">
<td width="40%" bgcolor="#FFCC00">Link : </td>
<td width="60%"><input type="text" name="url" class="form1" style="width:100%" value="<%=rs("m_url")%>"></td>
</tr>
<tr align="right">
<td width="40%" bgcolor="#FFCC00">Açýlýþ Þekli : </td>
<td width="60%">
<select name="acilis" size="1" class="form1" style="width:100%">
<option value="blank" <%=strOpt10%>>Yeni Sayfa</option>
<option value="normal" <%=strOpt11%>>Normal Site Düzeni</option>
</select>
</td>
</tr>
<tr align="right">
<td width="40%" bgcolor="#FFCC00">Aktif/Pasif : </td>
<td width="60%" align="left"><input type="checkbox" name="durum" <%=strOpt20%>></td>
</tr>
<tr align="right">
<td width="40%" bgcolor="#FFCC00">Kategori : </td>
<td width="60%">
<select name="kategori" size="1" class="form1" style="width:100%">
<%
Do Until cats.EoF
IF cats("mc_id") = rs("m_cat") Then
opt = "Selected"
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
<tr align="right" bgcolor="#FF9000">
<td colspan="2"><input type="submit" value="Deðiþiklikleri Kaydet »" class="submit"></td>
</tr>
</form>
</table>
<%
rs.Close
Set rs = Nothing
cats.Close
Set cats = Nothing
	ElseIF x = "Update" Then
strName = Replace(Request.QueryString("Name"),"'","´",1,-1,1)
strURL = Replace(Request.QueryString("URL"),"'","´",1,-1,1)

baslik = Replace(Request.Form("baslik"),"'","´",1,-1,1)
link = Replace(Request.Form("url"),"'","´",1,-1,1)
acilis = Request.Form("acilis")

IF Request.Form("durum") = "on" Then
durum = "True"
ELSE
durum = "False"
End IF

Set yeni = Server.CreateObject("ADODB.RecordSet")
yeni.Open "Select * FROM MENU_LINKS WHERE m_name = '"&strName&"' AND m_url = '"&strURL&"'",Connection,3,3

yeni("m_name") = baslik
yeni("m_url") = link
yeni("m_style") = acilis
yeni("m_status") = durum
yeni("m_cat") = Request.Form("kategori")
yeni.Update
yeni.Close
Set yeni = Nothing

Response.Redirect Request.ServerVariables("URL")
	ElseIF x = "Delete" Then
strName = Replace(Request.QueryString("Name"),"'","´",1,-1,1)
strURL = Replace(Request.QueryString("URL"),"'","´",1,-1,1)
Connection.Execute("DELETE FROM MENU_LINKS WHERE m_name = '"&strName&"' AND m_url = '"&strURL&"'")
Response.Redirect Request.ServerVariables("URL")
	ElseIF x = "Change" then

strName = Replace(Request.QueryString("Name"),"'","´",1,-1,1)
strURL = Replace(Request.QueryString("URL"),"'","´",1,-1,1)
strChange = Replace(Request.QueryString("Option"),"'","´",1,-1,1)
Set upd = Server.CreateObject("ADODB.RecordSet")
upd.Open "Select * FROM MENU_LINKS WHERE m_name = '"&strName&"' AND m_url = '"&strURL&"'",Connection,3,3
upd("m_status") = strChange
upd.Update
upd.Close
Set upd = Nothing
Response.Redirect Request.ServerVariables("URL")

	ELSE
Set rs = Server.CreateObject("ADODB.RecordSet")
rs.Open "Select * FROM MENU_LINKS ORDER BY m_name ASC",Connection,3,3
%>
<table border="0" cellspacing="2" cellpadding="1" width="75%" align="center" class="td_menu">
<tr>
<td colspan="4">» <a href="?op=New">Yeni Menü Linki</a></td>
</tr>
<tr bgcolor="#FFCC00" style="font-weight:bold" align="center">
<td width="20%" align="left">Baþlýk</td>
<td width="25%">Link</td>
<td width="20%">Açýlýþ Þekli</td>
<td width="10%">Durum</td>
<td width="25%">Ýþlemler</td>
</tr>
<%
sayi = 1
Do Until rs.EoF
Select Case Right(sayi,1)
Case 1,3,5,7,9
strBG = "#F0F0F0"
Case Else
strBG = "#E0E0E0"
End Select
IF rs("m_status") = True Then
strDurum = "Aktif"
strYazi = "Pasifleþtir"
strOption = "False"
ELSE
strDurum = "Pasif"
strYazi = "Aktifleþtir"
strOption = "True"
End IF
IF rs("m_style") = "blank" Then
strAcilis = "Yeni Pencere"
ELSE
strAcilis = "Normal"
End IF
strName = rs("m_name")
strURL = rs("m_url")
%>
<tr bgcolor="<%=strBG%>" align="center">
<td width="20%" align="left"><%=strName%></td>
<td width="25%"><%=strURL%></td>
<td width="20%"><%=strAcilis%></td>
<td width="10%"><%=strDurum%></td>
<td width="25%"><a href="?op=Edit&Name=<%=strName%>&URL=<%=strURL%>">Düzenle</a> - <a href="?op=Delete&Name=<%=strName%>&URL=<%=strURL%>">Sil</a> - <a href="?op=Change&Name=<%=strName%>&URL=<%=strURL%>&Option=<%=strOption%>"><%=strYazi%></a></td>
</tr>
<%
sayi = sayi + 1
rs.MoveNext
Loop
%>
</table>
<%
rs.Close
Set rs = Nothing
	End IF

End IF
%>