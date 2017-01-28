<%
Set Sur = Server.CreateObject("ADODB.Connection")
Sur.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../../db/resimdb.mdb")
%>
<%
Set abiniz = Server.CreateObject("ADODB.Recordset")
SQL = "Select * from abiniz ORDER BY id desc"
abiniz.Open SQL,Sur,1,3
Randomize
id = Int((toplam * Rnd)+ 0) 
abiniz.Move(id)
%>
<html>

<head>
<meta http-equiv="Content-Type"
content="text/html; charset=windows-1254">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta name="GENERATOR" content="Microsoft FrontPage Express 2.0">
<title>Rastgele Resim by Abiniz</title>
</head>

<body bgcolor="#FFFFFF" text="#EFAB3E" link="#414343"
vlink="#414343" alink="#414343" topmargin="0" leftmargin="0">
<%
Set abiniz = Server.CreateObject("ADODB.Recordset")
SQL = "Select * from abiniz ORDER BY id desc"
abiniz.Open SQL,Sur,1,3
toplam=abiniz.recordcount
Randomize
id = Int((toplam * Rnd)+ 0) 
abiniz.Move(id)
%>
<p align="center"> <img src="<%=abiniz("resim")%>" width="150" height="150" border="1"></p>
</body>
</html>
<hr style="border: 1 dashed #0059B3" color="#81BAE5" size="0" width="550">
<p align="center"><font face="Tahoma" size="1" color="#525454">Coded by Abiniz for Mini-Nuke Fixed Ver.<br />
</font><a href="http://www.mini-nuke.info"><font face="Tahoma" size="1" color="#525454">http://www.mini-nuke.info</font></a><font face="Tahoma" size="1" color="#525454">
</font><br />
