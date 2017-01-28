<!--#include file="../../db/data.asp" -->

<% 
IF dbType="personal" Then
Set Connection=Server.CreateObject("ADODB.Connection")
Connection.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../../db/mn"&dbNum&".mdb")
ELSE
Set Connection=Server.CreateObject("ADODB.Connection")
Connection.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../../db/mn"&dbNum&".mdb")
END IF

yorid=request.querystring("silyorum")
Set dc=Server.CreateObject("ADODB.RecordSet")
dcSQL="DELETE * FROM yorum WHERE yno="&yorid&""
dc.open dcSQL,Connection,3,3
Response.Redirect Request.ServerVariables("HTTP_REFERER")
%>