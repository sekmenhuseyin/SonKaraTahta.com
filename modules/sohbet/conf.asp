<%
'DATABASE

IF dbType="brinkster" Then
dataPath=""&Server.MapPath("../db/shoutdata.txt")&""
ELSE
dataPath=""&Server.MapPath("db/shoutdata.txt")&""
END IF

'FORM
	msg_submit="Gönder »"

'MESAJLAR

	err0="Lütfen Mesajýnýzý Yazýnýz..."
	guest="Ziyaretçi"
%>