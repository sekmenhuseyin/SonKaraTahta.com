<!--#include file="../../INC/b_include.asp" -->
<marquee style="cursor:  url(../../images/site/mavi.cur);" onmouseover="this.stop();" onmouseout="this.start();" direction="up" height="90" SCROLLAMOUNT="1"><%
Set todaymems = Server.CreateObject("ADODB.RecordSet")
todaymems.open "Select * FROM MEMBERS WHERE uyelik_tarihi= Date()-0",Connection,3,3
For I = 1 TO 20
IF todaymems.Eof Then Exit For
With Response
	.Write "- "
	.Write "<a href=""members.asp?action=member_details&uid="&todaymems("uye_id")&""">"
	.Write "<b>"&todaymems("kul_adi") &"</b>"& "<br />"
	.Write "</a>"
End With
todaymems.MoveNext
Next
%></td>