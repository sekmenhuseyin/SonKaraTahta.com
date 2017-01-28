<% Response.Buffer = True %>
<!--#include file="includes.asp" -->
<%
IF Session("Level") = "1" OR Session("Level") = "7" Then
x = Request.QueryString("action")
%>
<%
	IF x = "Move" Then
xy = Request.Form("type")
IF dbType = "personal" Then
Set NewConnection = Server.CreateObject("ADODB.Connection")
NewConnection.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../../db/"&Request.Form("db_name")&".mdb")
ELSE
Set NewConnection = Server.CreateObject("ADODB.Connection")
NewConnection.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../db/"&Request.Form("db_name")&".mdb")
End IF
	IF xy = "Members" Then
Set rs1 = Server.CreateObject("ADODB.RecordSet")
rs1.open "Select * FROM MEMBERS",Connection,3,3
Set rs2 = Server.CreateObject("ADODB.RecordSet")
rs2.open "Select * FROM MEMBERS",NewConnection,3,3

Do Until rs2.EoF
Set chk_mem = Connection.Execute("Select * FROM MEMBERS WHERE kul_adi = '"&rs2("kul_adi")&"'")
IF chk_mem.EoF Then
day_now = Day(Now())
month_now = Month(Now())
IF Len(day_now) = 1 Then
day_now = "0" & day_now
ELSE
day_now = day_now
End IF
IF Len(month_now) = 1 Then
month_now = "0" & month_now
ELSE
month_now = month_now
End IF
IF rs2("yas") = "" OR Len(rs2("yas")) > 3 Then
ar_age = "0"
ELSE
ar_age = rs2("yas")
End IF
m_age = day_now & "." & month_now & "." & Year(Now())-ar_age

IF Len(rs2("imza")) <= 0 Then
m_signature = "-"
ELSE
m_signature = rs2("imza")
End IF

IF rs2("mail_goster") = 0 Then
m_mgoster =  True
ELSE
m_mgoster = False
End IF

m_pass = MN_PP(rs2("sifre"))
m_gcevap = MN_PP(rs2("g_cevap"))

rs1.AddNEw
rs1("kul_adi") = rs2("kul_adi")
rs1("sifre") = m_pass
rs1("email") = rs2("email")
rs1("isim") = rs2("isim")
rs1("g_soru") = rs2("g_soru")
rs1("g_cevap") = m_gcevap
rs1("icq") = rs2("icq")
rs1("msn") = rs2("msn")
rs1("aim") = rs2("aim")
rs1("sehir") = rs2("sehir")
rs1("meslek") = rs2("meslek")
rs1("cinsiyet") = rs2("cinsiyet")
rs1("yas") = m_age
rs1("url") = rs2("url")
rs1("imza") = m_signature
rs1("mail_goster") = m_mgoster
rs1("login_sayisi") = rs2("login_sayisi")
rs1("uyelik_tarihi") = rs2("uyelik_tarihi")
rs1("son_tarih") = rs2("son_tarih")
rs1("seviye") = rs2("seviye")
rs1("msg_sayisi") = rs2("msg_sayisi")
rs1("last_ip") = rs2("last_ip")
rs1("u_theme") = "default"
rs1("u_avatar") = "IMAGES/avatars/blank.gif"
rs1("u_online") = "False"
rs1("ozellik") = "--"
rs1.Update
End IF
rs2.MoveNext
Loop
Response.Write "<div align=center class=td_menu2>Tüm Üyeler Yeni Database'e Aktarýldý !</div>"
rs1.Close
rs2.Close
Set rs1 = Nothing
Set rs2 = Nothing
	ElseIF xy = "Files" Then

Set cRs = Server.CreateObject("ADODB.RecordSet")
cRs.open "Select * FROM DOWN_CATS",NewConnection,3,3
Set fRs = Server.CreateObject("ADODB.RecordSet")
fRs.open "Select * FROM DOWNLOADS",NewConnection,3,3
Set nRs = Server.CreateObject("ADODB.RecordSet")
nRs.open "Select * FROM DOWNLOADS",Connection,3,3

Do Until cRs.EoF
Connection.Execute("INSERT INTO DOWN_CATS (c_name) VALUES ('"&cRs("c_name")&"')")
cRs.MoveNext
Loop
Response.Write "<div align=center class=td_menu2>Kategoriler Aktarýldý !</div>"

Do Until fRs.EoF
Set cats = NewConnection.Execute("Select * FROM DOWN_CATS WHERE cid = "&fRs("cid")&"")
Set new_cat = Connection.Execute("Select * FROM DOWN_CATS WHERE c_name = '"&cats("c_name")&"'")
nRs.AddNew
nRs("p_name") = fRs("p_name") 
nRs("p_info") = fRs("p_info")
nRs("p_size") = fRs("p_size")
nRs("p_hit") = fRs("p_hit")
nRs("p_url") = fRs("p_url")
nRs("cid") = new_cat("cid") 
nRs("p_date") = fRs("p_date")
nRs("p_author") = fRs("p_author")
nRs("point") = fRs("point")
nRs("p_approved") = fRs("p_approved")
nRs("p_img") = "IMAGES/avatars/blank.gif"
nRs.Update
fRs.MoveNext
Loop
Response.Write "<div align=center class=td_menu2>Dosyalar Aktarýldý !</div>"

	ElseIF xy = "Articles" Then

Set cRs = Server.CreateObject("ADODB.RecordSet")
cRs.open "Select * FROM ARTICLE_CATS",NewConnection,3,3
Set fRs = Server.CreateObject("ADODB.RecordSet")
fRs.open "Select * FROM ARTICLES",NewConnection,3,3
Set aRs = Server.CreateObject("ADODB.RecordSet")
aRs.open "Select * FROM ARTICLES",Connection,3,3

Do Until cRs.EoF
Connection.Execute("INSERT INTO ARTICLE_CATS (cat_name) VALUES ('"&cRs("cat_name")&"')")
cRs.MoveNext
Loop
Response.Write "<div align=center class=td_menu2>Kategoriler Aktarýldý !</div>"

Do Until fRs.EoF
Set cats = NewConnection.Execute("Select * FROM ARTICLE_CATS WHERE cat_id = "&fRs("cat_id")&"")
Set new_cat = Connection.Execute("Select * FROM ARTICLE_CATS WHERE cat_name = '"&cats("cat_name")&"'")

aRs.AddNew
aRs("a_title") = fRs("a_title") 
aRs("a_article") = fRs("a_article")
aRs("a_date") = fRs("a_date")
aRs("a_writer") = fRs("a_writer")
aRs("cat_id") = cats("cat_id")
aRs("hit") = fRs("hit")
aRs("point") = fRs("point")
aRs("a_approved") = fRs("a_approved")
aRs.Update
fRs.MoveNext
Loop
Response.Write "<div align=center class=td_menu2>Makaleler Aktarýldý !</div>"
	ElseIF xy = "Pages" Then
Set opRs = Server.CreateObject("ADODB.RecordSet")
opRs.open "Select * FROM PAGES",NewConnection,3,3
Set npRs = Server.CreateObject("ADODB.RecordSet")
npRs.open "Select * FROM PAGES",Connection,3,3

Do Until opRs.EoF
npRs.AddNew
npRs("p_title") = opRs("p_title")
npRs("p_content") = opRs("p_content")
npRs("p_members") = opRs("p_members")
npRs.Update
opRs.MoveNext
Loop
Response.Write "<div align=center class=td_menu2>Sayfalar Aktarýldý !</div>"

	ElseIF xy = "Counter" Then

Set oC = Server.CreateObject("ADODB.RecordSet")
oC.open "Select * FROM WEBCOUNTER",NewConnection,3,3

Do Until oC.EoF
Connection.Execute("INSERT INTO WEBCOUNTER (today,cdate) VALUES ('"&oC("today")&"','"&oC("cdate")&"')")
oC.MoveNext
Loop
Response.Write "<div align=center class=td_menu2>Sayaç Bilgileri Aktarýldý !</div>"

	ElseIF xy = "Banners" Then

Set oB = Server.CreateObject("ADODB.RecordSet")
oB.open "Select * FROM BANNERS",NewConnection,3,3
Set nB = Server.CreateObject("ADODB.RecordSet")
nB.open "Select * FROM BANNERS",Connection,3,3

Do Until oB.EoF
nB.AddNew
nB("ban_url") = oB("ban_url")
nB("go_url") = oB("go_url")
nB("hit") = oB("hit")
nB("active") = oB("active")
nB("position") = oB("position")
nB("limit") = oB("limit")
nB("show") = oB("show")
nB("password") = oB("password")
nB.Update
oB.MoveNext
Loop
Response.Write "<div align=center class=td_menu2>Bannerlar Aktarýldý !</div>"

	End IF
Set NewConnection = Nothing
	ELSE
%>
<table border="0" cellspacing="2" cellpadding="2" width="60%" align="center" class="td_menu2">
<form action="?action=Move" method="post">
<tr bgcolor="#E0E0E0">
<td width="50%" align="right"><b>Bilgilerin Yerleþtirileceði Database : </b></td>
<td width="50%" align="left">mn<%=dbNum%>.mdb</td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="50%" align="right"><b>Bilgilerin Alýnacaðý Database : </b></td>
<td width="50%" align="left">
<input type="text" name="db_name" class="form1" style="width:80%" size="20">.mdb</td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="50%" align="right"><b>Aktarým Türü : </b></td>
<td width="50%" align="left">
<select name="type" size="1" class="form1">
<option value="Members">Üyeler</option>
<option value="Files">Dosyalar</option>
<option value="Articles">Makaleler</option>
<option value="Pages">Sayfalar</option>
<option value="Counter">Sayaç</option>
<option value="Banners">Bannerlar</option>
</select>
</td>
</tr>
<tr bgcolor="#E0E0E0">
<td width="50%" align="right"></td>
<td width="50%" align="left"><input type="submit" value="Devam »" class="submit"></td>
</tr>
</form>
</table>
<br>
<div align="center" class="td_menu2"><b>NOT : </b>Bilgilerin Alýnacaðý Database,Bilgileri Yerleþtirileceði Database(Þimdi Kullandýðýnýz Database) ile ayný klasörde olmalýdýr.<br><i>Eðer hata alýyorsanýz yazdýðýnýz veritabaný ismini kontrol edin.</i></div>
<%
	End IF

End IF
%>