<!--#include file="INC/functions.asp" --><!--#include file="INC/filter.asp" --><!--#include file="INC/encryption.asp" -->
<title>Mini-NUKE Fixed v1.5 Kurulum Sayfasý</title>
<%Function TextFLTR(Str)
Str = Replace(Str, "[1]", "<", 1, -1, 1)
Str = Replace(Str, "[2]", "%", 1, -1, 1)
Str = Replace(Str, "[3]", ">", 1, -1, 1)
TextFLTR = Str
End Function%>
<body bgcolor="#FF9900" text="#008080">
<style>
.form1{font-family:Tahoma;font-size:8pt;background:#FFBB00;border-width:1px;border-color:#FF3300;page-break-inside:avoid;border-style:solid;border-collapse:collapse;}
.submit{font-family:Tahoma;font-size: 14pt;background:#FFBB00;border-width:1px;border-color:#660000;page-break-inside:avoid;border-style:solid;border-collapse:collapse;}
</style>
<%Set FEsys = Server.CreateObject("Scripting.FileSystemObject")
IF FEsys.FileExists(""&Server.Mappath("DB/data.asp")&"") = True Then
Response.Write "<font class=form1>Site Önceden Kurulmuþ !!!</font>"
ELSE
dbnum = CPASS(10)

If Request.QueryString("action") = "CDBS" Then
name = duz(Request.Form("s_name"))
address = duz(Request.Form("s_address"))
email = duz(Request.Form("s_email"))
site_slogan = duz(Request.Form("s_slogan"))
site_logo = "images/logo.gif"
flood_time = "1"
link = duz(Request.Form("s_link"))
news = duz(Request.Form("s_news"))
pa = duz(Request.Form("s_pacount"))
msgLimit = duz(Request.Form("s_msglimit"))
s_lcid = duz(Request.Form("s_lcid"))
path1 = Server.Mappath("db/MN"&dbnum&".mdb")
path2 = Server.Mappath("db/FORUM"&dbnum&".mdb")

Set objADOX = Server.CreateObject("ADOX.Catalog")
objADOX.Create "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & path1 & "; Jet OLEDB:Engine Type=5;"
objADOX.Create "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & path2 & "; Jet OLEDB:Engine Type=5;"
Set objADOX = Nothing
Set Connection = Server.CreateObject("ADODB.Connection"):Connection.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & path1
Set Connection2 = Server.CreateObject("ADODB.Connection"):Connection2.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & path2

'SMILILER
Dim mn_smiley_n(33)
Dim mn_smiley_i(33)
mn_smiley_n(1)=":)"
mn_smiley_n(2)=":P"
mn_smiley_n(3)=";)"
mn_smiley_n(4)=":*:"
mn_smiley_n(5)=":o"
mn_smiley_n(6)="xx("
mn_smiley_n(7)=":D"
mn_smiley_n(8)="|)"
mn_smiley_n(9)=":s"
mn_smiley_n(10)=":("
mn_smiley_n(11)=":^("
mn_smiley_n(12)=":^:"
mn_smiley_n(13)="8D"
mn_smiley_n(14)=":x"
mn_smiley_n(15)=":o)"
mn_smiley_n(16)="8("
mn_smiley_n(17)=":$"
mn_smiley_n(18)="}:)"
mn_smiley_n(19)=":V:"
mn_smiley_n(20)=":|"
mn_smiley_n(21)=":Y:"
mn_smiley_n(22)=":N:"
mn_smiley_n(23)=":-B"
mn_smiley_n(24)=":["
mn_smiley_n(25)=":?:"
mn_smiley_n(26)="&gt;&lt;"
mn_smiley_n(27)="%("
mn_smiley_n(28)="8-}"
mn_smiley_n(29)=":@)"
mn_smiley_n(30)="&gt;:D&lt;"
mn_smiley_n(31)="=D&gt;"
mn_smiley_n(32)="%"
mn_smiley_n(33)="!"
mn_smiley_i(1) ="smiley001.gif"
mn_smiley_i(2)="smiley017.gif"
mn_smiley_i(3)="smiley002.gif"
mn_smiley_i(4)="smiley010.gif"
mn_smiley_i(5)="smiley003.gif"
mn_smiley_i(6)="smiley011.gif"
mn_smiley_i(7)="smiley004.gif"
mn_smiley_i(8)="smiley012.gif"
mn_smiley_i(9)="smiley005.gif"
mn_smiley_i(10)="smiley006.gif"
mn_smiley_i(11)="smiley019.gif"
mn_smiley_i(12)="smiley014.gif"
mn_smiley_i(13)="smiley016.gif"
mn_smiley_i(14)="smiley007.gif"
mn_smiley_i(15)="smiley008.gif"
mn_smiley_i(16)="smiley018.gif"
mn_smiley_i(17)="smiley009.gif"
mn_smiley_i(18)="smiley015.gif"
mn_smiley_i(19)="smiley013.gif"
mn_smiley_i(20)="smiley022.gif"
mn_smiley_i(21)="smiley020.gif"
mn_smiley_i(22)="smiley021.gif"
mn_smiley_i(23)="smiley023.gif"
mn_smiley_i(24)="smiley024.gif"
mn_smiley_i(25)="smiley025.gif"
mn_smiley_i(26)="smiley026.gif"
mn_smiley_i(27)="smiley028.gif"
mn_smiley_i(28)="smiley029.gif"
mn_smiley_i(29)="smiley030.gif"
mn_smiley_i(30)="smiley031.gif"
mn_smiley_i(31)="smiley032.gif"
mn_smiley_i(32)="smiley033.gif"
mn_smiley_i(33)="smiley034.gif"

'AVATARLAR
Dim mn_avatar_n(39)
Dim mn_avatar_i(39)
mn_avatar_n(1) = "01"
mn_avatar_n(2) = "01"
mn_avatar_n(3) = "02"
mn_avatar_n(4) = "02"
mn_avatar_n(5) = "03"
mn_avatar_n(6) = "03"
mn_avatar_n(7) = "04"
mn_avatar_n(8) = "04"
mn_avatar_n(9) = "05"
mn_avatar_n(10) = "05"
mn_avatar_n(11) = "06"
mn_avatar_n(12) = "06"
mn_avatar_n(13) = "07"
mn_avatar_n(14) = "07"
mn_avatar_n(15) = "08"
mn_avatar_n(16) = "08"
mn_avatar_n(17) = "09"
mn_avatar_n(18) = "09"
mn_avatar_n(19) = "10"
mn_avatar_n(20) = "10"
mn_avatar_n(21) = "11"
mn_avatar_n(22) = "11"
mn_avatar_n(23) = "12"
mn_avatar_n(24) = "12"
mn_avatar_n(25) = "13"
mn_avatar_n(26) = "13"
mn_avatar_n(27) = "14"
mn_avatar_n(28) = "14"
mn_avatar_n(29) = "15"
mn_avatar_n(30) = "15"
mn_avatar_n(31) = "16"
mn_avatar_n(32) = "16"
mn_avatar_n(33) = "17"
mn_avatar_n(34) = "17"
mn_avatar_n(35) = "18"
mn_avatar_n(36) = "18"
mn_avatar_n(37) = "19"
mn_avatar_n(38) = "19"
mn_avatar_n(39) = "20"


mn_avatar_i(1) = "01.gif"
mn_avatar_i(2) = "01.jpg"
mn_avatar_i(3) = "02.gif"
mn_avatar_i(4) = "02.jpg"
mn_avatar_i(5) = "03.gif"
mn_avatar_i(6) = "03.jpg"
mn_avatar_i(7) = "04.gif"
mn_avatar_i(8) = "04.jpg"
mn_avatar_i(9) = "05.gif"
mn_avatar_i(10) = "05.jpg"
mn_avatar_i(11) = "06.gif"
mn_avatar_i(12) = "06.jpg"
mn_avatar_i(13) = "07.gif"
mn_avatar_i(14) = "07.jpg"
mn_avatar_i(15) = "08.gif"
mn_avatar_i(16) = "08.jpg"
mn_avatar_i(17) = "09.gif"
mn_avatar_i(18) = "09.jpg"
mn_avatar_i(19) = "10.gif"
mn_avatar_i(20) = "10.jpg"
mn_avatar_i(21) = "11.gif"
mn_avatar_i(22) = "11.jpg"
mn_avatar_i(23) = "12.gif"
mn_avatar_i(24) = "12.jpg"
mn_avatar_i(25) = "13.gif"
mn_avatar_i(26) = "13.jpg"
mn_avatar_i(27) = "14.gif"
mn_avatar_i(28) = "14.jpg"
mn_avatar_i(29) = "15.gif"
mn_avatar_i(30) = "15.jpg"
mn_avatar_i(31) = "16.gif"
mn_avatar_i(32) = "16.jpg"
mn_avatar_i(33) = "17.gif"
mn_avatar_i(34) = "17.jpg"
mn_avatar_i(35) = "18.gif"
mn_avatar_i(36) = "18.jpg"
mn_avatar_i(37) = "19.gif"
mn_avatar_i(38) = "19.jpg"
mn_avatar_i(39) = "20.gif"

'AYARLAR
Connection.Execute("Create TABLE SETTINGS (a_id AutoIncrement, site_isim MEMO, site_adres MEMO, site_logo MEMO, site_slogan MEMO, flood_time NUMERIC, site_mail MEMO, haber_sayisi NUMERIC, prg_sayisi NUMERIC, msg_limit NUMERIC, theme TEXT, np_msg MEMO, fpass TEXT, s_lcid NUMERIC, mp_news YESNO, mp_forum YESNO, site_locked MEMO, site_locked_msg MEMO, sagtus BIT, merlin BIT, hmenu BIT, site_topbanner MEMO, radyo MEMO, gorunusurl MEMO)")
'TEMALAR
Connection.Execute("Create TABLE THEMES (t_id AutoIncrement,t_name MEMO,t_dir MEMO,t_active YESNO)")
'ILETISIM
Connection.Execute("Create TABLE MESAJ (id AutoIncrement, baslik MEMO, mesaj MEMO, ekleyen MEMO, email MEMO, konu MEMO, tarih DATETIME)")
'SMILILER
Connection.Execute("Create TABLE SMILEYS (s_id AutoIncrement,s_info TEXT,s_img MEMO)")
'AVATARLAR
Connection.Execute("Create TABLE AVATARS (a_id AutoIncrement,a_name MEMO,a_img MEMO)")
'DUYURULAR
Connection.Execute("Create TABLE NOTICES (n_id AutoIncrement,n_text MEMO,n_date DATETIME)")
'BANLANMIÞ IPLER
Connection.Execute("Create TABLE BANNED_IPS (b_id AutoIncrement,b_ip MEMO,b_date DATETIME)")
'MAKALE KATEGORILERI
Connection.Execute("Create TABLE ARTICLE_CATS (cat_id AutoIncrement, cat_name MEMO)")
'MAKALELER
Connection.Execute("Create TABLE ARTICLES (aid AutoIncrement, a_title MEMO, a_article MEMO, a_date TEXT, a_writer TEXT, cat_id NUMERIC, hit NUMERIC, point NUMERIC, pointer NUMERIC, a_approved BIT)")
'MENÜ LINKLERI
Connection.Execute("Create TABLE MENU_LINKS (m_name MEMO, m_url MEMO, m_status YESNO, m_style MEMO, m_cat NUMERIC)")
'REKLAM
Connection.Execute("Create TABLE BANNERS (banner_id AutoIncrement, ban_url MEMO, go_url MEMO, hit NUMERIC, active BIT, position TEXT, limit NUMERIC, show NUMERIC, password MEMO,ban_type TEXT)")
'BLOKLAR
Connection.Execute("Create TABLE BLOCKS (bid AutoIncrement, b_name MEMO, b_content MEMO, b_align TEXT, b_include MEMO, b_active YESNO, b_queue INT)")
'YORUMLAR
Connection.Execute("Create TABLE COMMENTS (cid AutoIncrement, writer TEXT, comment MEMO, cdate DATETIME, nid NUMERIC, page MEMO)")
'SAYAÇ
Connection.Execute("Create TABLE WEBCOUNTER (id AutoIncrement, today NUMERIC, cdate DATETIME)")
'KÜFÜR KORUMASI
Connection.Execute("Create TABLE SWEARWORDS (id AutoIncrement, s_text MEMO, s_value MEMO)")
'DOWNLOAD KATEGORILERI
Connection.Execute("Create TABLE DOWN_CATS (cid AutoIncrement, c_name MEMO)")
'DOWNLOADLAR
Connection.Execute("Create TABLE DOWNLOADS (pid AutoIncrement, p_name TEXT, p_info MEMO, p_size MEMO, p_hit NUMERIC, p_url MEMO, cid NUMERIC, p_date TEXT, p_author TEXT, point NUMERIC, pointer NUMERIC, p_approved BIT, p_img MEMO)")
'ARKADAÞ LiSTESi
Connection.Execute("Create TABLE FRIENDS (member NUMERIC, friend NUMERIC)")
'ZiYARETÇiLER
Connection.Execute("Create TABLE GUESTS (zid AutoIncrement, ip MEMO, zdate DATETIME)")
'üYELER
Connection.Execute("Create TABLE MEMBERS (uye_id AutoIncrement, kul_adi TEXT, sifre MEMO, email MEMO, isim MEMO, g_soru MEMO, g_cevap MEMO, icq NUMERIC, msn MEMO, aim MEMO, sehir MEMO, meslek MEMO, cinsiyet TEXT, yas TEXT, url MEMO, imza MEMO, mail_goster BIT, login_sayisi NUMERIC, uyelik_tarihi TEXT, son_tarih DATETIME, seviye NUMERIC, msg_sayisi NUMERIC, last_ip MEMO, u_theme TEXT, u_avatar TEXT, u_online YESNO, u_browser MEMO, u_ws MEMO, ozellik MEMO)")
'MESAJLAR
Connection.Execute("Create TABLE MESSAGES (mid AutoIncrement, from TEXT, to TEXT, msg MEMO, mdate DATETIME, read BIT, subject MEMO, type NUMERIC, see BIT, msave YESNO)")
'MODüLLER
Connection.Execute("Create TABLE MODULES (module_id AutoIncrement, mname TEXT, mdir MEMO, mactive BIT, mems BIT,mcat NUMERIC)")
'HABER KATEGORILERI
Connection.Execute("Create TABLE NEWS_CATS (cat_id AutoIncrement, cat_name TEXT, cat_info MEMO, cat_image MEMO)")
'HABERLER
Connection.Execute("Create TABLE NEWS (hid AutoIncrement, baslik MEMO, metin MEMO, ekleyen TEXT, tarih DATETIME, onay BIT, resim MEMO, puan NUMERIC, katilimci NUMERIC, okunma NUMERIC, cat NUMERIC)")
'SABIT HABERLER
Connection.Execute("Create TABLE FIXED (fid AutoIncrement, sf_baslik MEMO, f_metin MEMO, resim MEMO)")
'ANKET SEÇENEKLERi
Connection.Execute("Create TABLE POLL_ALTERNATIVES (a_id AutoIncrement, pid NUMERIC, alternative MEMO, a_counter NUMERIC)")
'ANKETLER
Connection.Execute("Create TABLE POLLS (poll_id AutoIncrement, p_question MEMO, p_date DATETIME)")
'SAYFALAR
Connection.Execute("Create TABLE PAGES (p_id AutoIncrement, p_title MEMO, p_content MEMO, p_members BIT,p_cat NUMERIC)")
'MENÜ KATEGORÝLERÝ
Connection.Execute("Create TABLE MENU_CATS (mc_id AutoIncrement,mc_title TEXT,mc_queue NUMERIC)")
'ZIYARETÇI DEFTERI DATABASE
Connection.Execute("Create TABLE GUESTBOOK (ID AutoIncrement, NAME TEXT, EMAIL TEXT,YAZI MEMO)")
'LINKLER DATABASE
Connection.Execute("Create TABLE LINKLER (LID AutoIncrement, ltitle MEMO, lurl MEMO,linfo MEMO, ldate DATETIME, lekleyen MEMO, lhit NUMERIC, cid NUMERIC)")
Connection.Execute("Create TABLE KATEGORILER (CID AutoIncrement, cname MEMO)")
'RESIM GALERISI DATABASE
Connection.Execute("Create TABLE FOTOGRAF (fno AutoIncrement, isim TEXT, aciklama MEMO, tarih TEXT, kucuk MEMO, buyuk MEMO, hit NUMERIC,kategori TEXT)")
Connection.Execute("Create TABLE KATEGORI (kno AutoIncrement, statement MEMO, kategori TEXT)")
Connection.Execute("Create TABLE YORUM (yno AutoIncrement, isim TEXT, yorum MEMO, onay YESNO, tarih DATETIME, hangi NUMERIC)")
'RASTGELE RESIM & DOSTBANNER  & TOPLIST & LINKLIST & GUNUNSOZU DATABASE
Connection.Execute("Create TABLE RASTGELE (id AutoIncrement, resim MEMO, hit MEMO, puan MEMO)")
Connection.Execute("Create TABLE DOSTBAN (banid AutoIncrement, resurl MEMO, giturl MEMO, ryazi TEXT, restik NUMERIC, aktif BIT, pozisyon TEXT, sinir NUMERIC, gor NUMERIC, siffre MEMO)")
Connection.Execute("Create TABLE UFAKBAN (banid AutoIncrement, resurl MEMO, giturl MEMO, ryazi TEXT, restik NUMERIC, aktif BIT, pozisyon TEXT, sinir NUMERIC, gor NUMERIC, siffre MEMO)")
Connection.Execute("Create TABLE LINKLIST (uye_no AutoIncrement, email MEMO, adi MEMO, url MEMO, tekil NUMERIC, cogul NUMERIC, kayit_tarihi DATETIME, onay BIT, sayac_zamani DATETIME)")
Connection.Execute("Create TABLE GUNUNSOZU (id AutoIncrement, sozler MEMO, fotograf MEMO, yazar MEMO)")
'FORUM
Connection2.Execute("Create TABLE SETTINGS (f_posts NUMERIC, f_replies NUMERIC)")
Connection2.Execute("Create TABLE MAIN_CATS (mcid AutoIncrement, mc_name MEMO)")
Connection2.Execute("Create TABLE CATEGORIES (cid AutoIncrement, cname MEMO, cinfo MEMO, locked BIT, noentry BIT, mcid NUMERIC)")
Connection2.Execute("Create TABLE MESSAGES (mid AutoIncrement, mtitle TEXT, mmsg MEMO, mwriter TEXT, mdate DATETIME, mhit NUMERIC, cid NUMERIC, locked BIT, topic BIT, last_post DATETIME, cvp_sayisi NUMERIC html_allow YESNO)")
Connection2.Execute("Create TABLE MODERATORS (modid AutoIncrement, uye_id NUMERIC, cid NUMERIC)")

dataText = "[1][2]"
dataText = dataText + vbCrLf
dataText = dataText + "dbType="""&dbType&""""
dataText = dataText + vbCrLf
dataText = dataText + "dbNum="""&dbnum&""""
dataText = dataText + vbCrLf
dataText = dataText + "S_LCID="""&s_lcid&""""
dataText = dataText + vbCrLf
dataText = dataText + "[2][3]"
dataText = TextFLTR(dataText)

Set Fsys = CreateObject("Scripting.FileSystemObject")
Set CDFile = Fsys.CreateTextFile(""&Server.Mappath("./db/DATA.ASP")&"", 2)
CDFile.WriteLine(""&dataText&"")

a_name = duz(Request.Form("a_name"))
a_pass = duz(Request.Form("a_pass"))
a_pass = MN_PP(a_pass)
a_email = duz(Request.Form("a_email"))
g_soru = duz(Request.Form("g_soru"))
g_cevap = duz(Request.Form("g_cevap"))
g_cevap = MN_PP(g_cevap)

Set objNSettings = Server.CreateObject("ADODB.RecordSet")
nsSQL = "Select * FROM SETTINGS"
objNSettings.open nsSQL,Connection,3,3

objNSettings.AddNew
objNSettings("site_isim") = duz(Request.Form("s_name"))
objNSettings("site_adres") = duz(Request.Form("s_address"))
objNSettings("gorunusurl") = duz(Request.Form("gorunusurl"))
objNSettings("site_slogan") = duz(Request.Form("site_slogan"))
objNSettings("radyo") = duz(Request.Form("radyo"))
objNSettings("site_logo") = "images/logo.gif"
objNSettings("site_topbanner") = duz(Request.Form("site_topbanner"))
objNSettings("flood_time") = "1"
objNSettings("sagtus") = "-1"
objNSettings("merlin") = "-1"
objNSettings("hmenu") = "-1"
objNSettings("site_mail") = duz(Request.Form("s_email"))
objNSettings("haber_sayisi") = duz(Request.Form("s_news"))
objNSettings("prg_sayisi") = duz(Request.Form("s_pacount"))
objNSettings("msg_limit") = duz(Request.Form("s_msglimit"))
objNSettings("theme") = "Mn-Dehset"
objNSettings("np_msg") = "<font class=td_menu><b>Bu Sayfa Sadece Üyeler Açýktýr !!!</b><br><br><div align=left>- <a href=membership.asp?action=new>Üye Ol</a><br>- <a href=membership.asp?action=lostpass&step=1>Þifremi Unuttum !!</a></div></font>"
objNSettings("fpass") = "Site"
objNSettings("mp_news") = Request.Form("ln5")
objNSettings("mp_forum") = Request.Form("lf5")
objNSettings.Update

Connection.Execute("INSERT INTO BLOCKS (b_name, b_content, b_include, b_align, b_active, b_queue) VALUES ('Üyelik', '', 'membership', 'RIGHT', True, 1)")
Connection.Execute("INSERT INTO BLOCKS (b_name, b_content, b_include, b_align, b_active, b_queue) VALUES ('Menü', '', 'pages_modules', 'LEFT', True, 1)")
Connection.Execute("INSERT INTO BLOCKS (b_name, b_content, b_include, b_align, b_active, b_queue) VALUES ('Istatistikler', '', 'stats', 'LEFT', True, 2)")

Connection2.Execute("INSERT INTO SETTINGS (f_posts, f_replies) VALUES(1, 1)")
Connection.Execute("INSERT INTO THEMES (t_name, t_dir, t_active) VALUES('Cehennem', 'Cehennem', True)")
Connection.Execute("INSERT INTO THEMES (t_name, t_dir, t_active) VALUES('default', 'default', True)")
Connection.Execute("INSERT INTO THEMES (t_name, t_dir, t_active) VALUES('Mn-Dehset', 'Mn-Dehset', True)")
Connection.Execute("INSERT INTO THEMES (t_name, t_dir, t_active) VALUES('Fantasy', 'Fantasy', True)")
Connection.Execute("INSERT INTO THEMES (t_name, t_dir, t_active) VALUES('MN-Digi', 'MN-Digi', True)")
Connection.Execute("INSERT INTO BANNERS (ban_url, go_url, hit, active, position, limit, show, password,ban_type) VALUES('http://www.mini-nuke.info/images/sponsor/reklam.gif', 'http://www.epidio.com', 1, True, 'top', 999999999999, 10, 'epidio', 'Normal')")
Connection.Execute("INSERT INTO BANNERS (ban_url, go_url, hit, active, position, limit, show, password,ban_type) VALUES('http://www.mini-nuke.info/images/sponsor/reklam.gif', 'http://www.epidio.com', 1, True, 'bottom', 999999999999, 10, 'epidio', 'Normal')")
Connection.Execute("INSERT INTO UFAKBAN (resurl, giturl, restik, aktif, gor) VALUES ('http://www.mini-nuke.de/images/banner.gif', 'http://www.mini-nuke.de', '0', '-1', 0)")
Connection.Execute("INSERT INTO MENU_CATS (mc_title, mc_queue) VALUES ('GeneL', '1')")
Connection.Execute("INSERT INTO MODULES (mname, mdir, mactive, mems, mcat) VALUES ('Haber Gönder', 'send_news', '-1', '-1', '1')")
Connection.Execute("INSERT INTO MODULES (mname, mdir, mactive, mems, mcat) VALUES ('Resim Galerisi', 'Gallery', '-1', '-1', '1')")


For I = 1 TO 39:Connection.Execute("INSERT INTO AVATARS (a_name,a_img) VALUES ('"&mn_avatar_n(""&I&"")&"','"&mn_avatar_i(""&I&"")&"')"):Next
For I = 1 TO 33:Connection.Execute("INSERT INTO SMILEYS (s_info,s_img) VALUES ('"&mn_smiley_n(""&I&"")&"','"&mn_smiley_i(""&I&"")&"')"):Next
Set entA = Server.CreateObject("ADODB.Recordset"):eaSQL = "Select * FROM MEMBERS":entA.open eaSQL,Connection,3,3
entA.AddNew
entA("kul_adi") = a_name
entA("sifre") = a_pass
entA("email") = a_email
entA("g_soru") = g_soru
entA("g_cevap") = g_cevap
entA("yas") = Request.Form("yas_1") & "." & Request.Form("yas_2") & "." & Request.Form("yas_3")
entA("icq") = 0
entA("login_sayisi") = 0
entA("uyelik_tarihi") = Date()
entA("seviye") = 1
entA("msg_sayisi") = 0
entA("last_ip") = Request.ServerVariables("REMOTE_ADDR")
entA("u_theme") = "Cehennem"
entA("u_avatar") = "IMAGES/avatars/01.gif"
entA("imza") = "-"
entA.Update
objNSettings.Close:Set objNSettings = Nothing
Connection.Close:Set Connection = Nothing
Response.Write "<font face=""Verdana"" style=""font-size:12px"">Siteniz baþarýyla kuruldu... Fixed 1.5'e Hoþ Geldiniz.<br><br>Sitenize gitmek için <a href="""&duz(Request.Form("s_address"))&"""><b>týklayýnýz...</b></a></font>"

ELSE%>
<form action="?action=CDBS" method="post">
<table border="0" cellspacing="1" cellpadding="5" align="center" width="686" style="font-family:tahoma; font-size:11px;" bordercolor="#FF0000">
<tr><td colspan="2" bgcolor="#CC0000" align="center" style="font-size:12px;color:Red" bordercolor="#CC0000" width="674"><b><font class="block_title" color="#00FFFF">Mini-NUKE Fixed v1.5 Kurulum Sayfasý</font></b></td></tr>
<tr><td colspan="2" bgcolor="#CC0000" style="color:#FFCC00" bordercolor="#CC0000" width="674"><font color="#00FFFF"><b>» Genel</b></font></td></tr>
<tr>
<td width="347" bgcolor="#800000" bordercolor="#660000"><font color="#FFFFFF"><b>Sitenizin Ýsmi</b></font></td>
<td width="316" bgcolor="#800000" bordercolor="#660000"><font color="#FFFFFF"><input type="text" name="s_name" size="60" class="form1" value="siteismi"></font></td>
</tr>
<tr>
<td width="347" bgcolor="#800000" bordercolor="#660000"><font color="#FFFFFF"><b>Sitenizin Adresi</b></font></td>
<td width="316" bgcolor="#800000" bordercolor="#660000"><font color="#FFFFFF"><input type="text" name="s_address" size="60" class="form1" value="http://www.siteadresi.com"></font></td>
</tr>
<tr>
<td width="347" bgcolor="#800000" bordercolor="#660000"><font color="#FFFFFF"><b>Görünüþ Adresi</b></font></td>
<td width="316" bgcolor="#800000" bordercolor="#660000"><font color="#FFFFFF"><input type="text" name="gorunusurl" size="60" class="form1" value="görünümadresi"></font></td>
</tr>
<tr>
<td width="347" bgcolor="#800000" bordercolor="#660000"><font color="#FFFFFF"><b>Sitenizin Slogani</b></font></td>
<td width="316" bgcolor="#800000" bordercolor="#660000"><font color="#FFFFFF"><input type="text" name="site_slogan" size="60" class="form1" value="sitesloganý"></font></td>
</tr>
<tr>
<td width="347" bgcolor="#800000" bordercolor="#660000"><font color="#FFFFFF"><b>Site Radyosu</b></font></td>
<td width="316" bgcolor="#800000" bordercolor="#660000"><font color="#FFFFFF"><input type="text" name="radyo" size="60" class="form1" value="siteradyosu"></font></td>
</tr>
<tr>
<td width="347" bgcolor="#800000" bordercolor="#660000"><font color="#FFFFFF"><b>Site TopBanneri</b></font></td>
<td width="316" bgcolor="#800000" bordercolor="#660000"><font color="#FFFFFF"><input type="text" name="site_topbanner" size="60" class="form1" value="topbanner"></font></td>
</tr>
<td width="347" bgcolor="#800000" bordercolor="#660000"><font color="#FFFFFF"><b>Sitenizin E-Mail Adresi</b></font></td>
<td width="316" bgcolor="#800000" bordercolor="#660000"><font color="#FFFFFF"><input type="text" name="s_email" size="60" class="form1" value="email"></font></td>
</tr>
<tr>
<td width="347" bgcolor="#800000" bordercolor="#660000"><font color="#FFFFFF"><b>Anasayfada Listelenecek Haber Sayýsý</b></font></td>
<td width="316" bgcolor="#800000" bordercolor="#660000"><font color="#FFFFFF"><select name="s_news" size="1" class="form1"><% For i = 1 TO 10 %><option value="<%=i%>"><%=i%></option><% Next %></select> </font></td>
</tr>
<tr>
<td width="347" bgcolor="#800000" bordercolor="#660000"><font color="#FFFFFF"><b>Bir Sayfada Listelenecek Makale veya Program Sayýsý</b></font></td>
<td width="316" bgcolor="#800000" bordercolor="#660000"><font color="#FFFFFF"><select name="s_pacount" size="1" class="form1"><% For i = 1 TO 50 %><option value="<%=i%>"><%=i%></option><% Next %></select> </font></td>
</tr>
<tr>
<td width="347" bgcolor="#800000" bordercolor="#660000"><font color="#FFFFFF"><b>Mesajlaþma Limiti</b></font></td>
<td width="316" bgcolor="#800000" bordercolor="#660000"><font color="#FFFFFF"><select name="s_msglimit" size="1" class="form1"><% For i = 1 TO 1000 %><option value="<%=i%>"><%=i%></option><% Next %></select> Mesaj</font></td>
</tr>
<tr>
<td width="347" bgcolor="#800000" bordercolor="#660000"><font color="#FFFFFF"><b>Site LCID (Saat/Tarih Formatý)</b></font></td>
<td width="316" bgcolor="#800000" bordercolor="#660000"><font color="#FFFFFF"><select name="s_lcid" size="1" class="form1"><option value="1055">1055 (Türkiye)</option><option value="1033">1033 (United States)</option><option value="2057">2057 (United Kingdom)</option><option value="2048">2048</option></select> </font></td>
</tr>
<tr>
<td width="347" bgcolor="#800000" bordercolor="#660000"><font color="#FFFFFF"><b>Anasayfada Son 5 Haber</b></font></td>
<td width="316" bgcolor="#800000" bordercolor="#660000"><font color="#FFFFFF"><select name="ln5" size="1" class="form1"><option value="True">Olsun</option><option value="False">Olmasýn</option></select> </font></td>
</tr>
<tr>
<td width="347" bgcolor="#800000" bordercolor="#660000"><font color="#FFFFFF"><b>Anasayfada Son 5 Forum</b></font></td>
<td width="316" bgcolor="#800000" bordercolor="#660000"><font color="#FFFFFF"><select name="lf5" size="1" class="form1"><option value="True">Olsun</option><option value="False">Olmasýn</option></select> </font></td>
</tr>
<tr><td colspan="2" bgcolor="#CC0000" style="color:#FFCC00" bordercolor="#800000" width="674"><font color="#00FFFF"><b>» Yönetim</b></font></td></tr>
<tr>
<td width="347" bgcolor="#800000" bordercolor="#800000"><font color="#FFFFFF"><b>Yönetici Kullanýcý Adý : </b></font></td>
<td width="316" bgcolor="#800000" bordercolor="#800000"><font color="#FFFFFF"><input type="text" name="a_name" size="60" class="form1" value="admin"></font></td>
</tr>
<tr>
<td width="347" bgcolor="#800000" bordercolor="#800000"><font color="#FFFFFF"><b>Yönetici Þifresi : </b></font></td>
<td width="316" bgcolor="#800000" bordercolor="#800000"><font color="#FFFFFF"><input type="text" name="a_pass" size="60" class="form1" value="admin"></font></td>
</tr>
<tr>
<td width="347" bgcolor="#800000" bordercolor="#800000"><font color="#FFFFFF"><b>Yönetici E-Maili : </b></font></td>
<td width="316" bgcolor="#800000" bordercolor="#800000"><font color="#FFFFFF"><input type="text" name="a_email" size="60" class="form1" value="admin@admin.admin"></font></td>
</tr>
<tr>
<td width="347" bgcolor="#800000" bordercolor="#800000"><font color="#FFFFFF"><b>Yönetici Gizli Sorusu : </b></font></td>
<td width="316" bgcolor="#800000" bordercolor="#800000"><font color="#FFFFFF"><input type="text" name="g_soru" size="60" class="form1" value="admin"></font></td>
</tr>
<tr>
<td width="347" bgcolor="#800000" bordercolor="#800000"><font color="#FFFFFF"><b>Yönetici Gizli Sorusu Cevabi : </b></font></td>
<td width="316" bgcolor="#800000" bordercolor="#800000"><font color="#FFFFFF"><input type="text" name="g_cevap" size="60" class="form1" value="admin"></font></td>
</tr>
<tr> 
<td bgcolor="#800000" width="347" bordercolor="#800000"><font color="#FFFFFF"><b>Dogum Tarihi : </b></font></td>
<td bgcolor="#800000" width="316" bordercolor="#800000"><font color="#FFFFFF">
<select name="yas_1" size="1" class="form1"><% For i = 1 TO 31 %><option value="<%=i%>"><%=i%></option>"<% Next %></select>.
<select name="yas_2" size="1" class="form1"><% For i = 1 TO 12 %><option value="<%=i%>"><%=i%></option>"<% Next %></select>.
<select name="yas_3" size="1" class="form1"><% For i = 1910 TO 2010 %><option value="<%=i%>"><%=i%></option>"<% Next %></select>
</font></td>
</tr>
<tr>
<td colspan="2" bgcolor="#CC0000" align="center" width="674"><input type="submit" value="Tamam »" class="submit" style="width:500px"></td>
</tr>
</table>
</form>
<%End IF
End IF%>