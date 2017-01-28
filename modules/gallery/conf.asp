<!--#include file="../../db/data.asp" -->
<%
'DATABASE

IF dbType="personal" Then
Set Connection=Server.CreateObject("ADODB.Connection")
Connection.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("../db/mn"&dbNum&".mdb")
ELSE
Set Connection=Server.CreateObject("ADODB.Connection")
Connection.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("db/mn"&dbNum&".mdb")
END IF

ksay=10
%>