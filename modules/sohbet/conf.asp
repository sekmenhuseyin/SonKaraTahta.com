<%
'DATABASE

IF dbType="brinkster" Then
dataPath=""&Server.MapPath("../db/shoutdata.txt")&""
ELSE
dataPath=""&Server.MapPath("db/shoutdata.txt")&""
END IF

'FORM
	msg_submit="G�nder �"

'MESAJLAR

	err0="L�tfen Mesaj�n�z� Yaz�n�z..."
	guest="Ziyaret�i"
%>