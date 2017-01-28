<!--#include file="../../INC/b_include.asp" --><%
Set top_5_mems = Server.CreateObject("ADODB.RecordSet")
top_5_mems.open "Select * FROM MEMBERS ORDER BY msg_sayisi DESC",Connection,3,3
For I = 1 TO 10
IF top_5_mems.Eof Then Exit For
Response.Write "- <a href=""members.asp?action=member_details&uid="&top_5_mems("uye_id")&""">"&top_5_mems("kul_adi")&"<br /></a>"
top_5_mems.MoveNext
Next%>