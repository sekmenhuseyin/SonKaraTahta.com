<!--#include file="../db/data.asp" -->
<%Set Connection= Server.CreateObject("ADODB.Connection"):Connection.Open"DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("db/mn"&dbNum&".mdb")
Set Connect=Server.CreateObject("ADODB.Connection"):Connect.Open"DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("db/forum"&dbNum&".mdb")%>