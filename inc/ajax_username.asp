<!--#include file="../db/data.asp" --><!--#include file="filter.asp" --><!--#include file="functions.asp" --><!--#include file="language_file.asp" -->
<%Response.ContentType = "text/html":Response.Charset = "iso-8859-9"
Set Connection=Server.CreateObject("ADODB.Connection"):Connection.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ="&Server.MapPath("../db/mn"&dbNum&".mdb")
psid=duz(Request.QueryString("psid")):kuladi=duz(Request.QueryString("q"))
if psid<>"" and kuladi<>"" then
Set nameCheck=Connection.Execute("Select * From MEMBERS where UCase(kul_adi)='"&UCase(kuladi)&"'")
IF InStr(1,kuladi,"ð",1) OR InStr(1,kuladi,"þ",1) OR InStr(1,kuladi,"ö",1) Then chk_kul_adi="False" ELSE chk_kul_adi="True"
if NOT nameCheck.eof Then
response.write ("<b>"&error9&"</b><br />")
elseif chk_kul_adi="False" Then
response.write (error16)
end if
end if
Connection.Close:Set Connection=Nothing%>
