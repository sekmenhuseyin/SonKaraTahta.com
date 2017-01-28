<!--#include file="../../DB/data.asp" -->
<% 
If dbType = "personal" Then
Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../../db/mn"&dbNum&".mdb")
Set Connect = Server.CreateObject("ADODB.Connection")
Connect.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../../db/forum"&dbNum&".mdb")
ELSE
Set Connection = Server.CreateObject("ADODB.Connection")
Connection.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../db/mn"&dbNum&".mdb")
Set Connect = Server.CreateObject("ADODB.Connection")
Connect.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../db/forum"&dbNum&".mdb")
End IF
%>