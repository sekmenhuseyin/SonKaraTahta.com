<!--#include file="../../INC/b_include.asp" -->
<%
Set admin_mems = Server.CreateObject("ADODB.RecordSet")
admin_mems.open "Select * FROM MEMBERS WHERE seviye=1",Connection,3,3
%>
<p align="center">
<b><img border="0" src="Images/bestmod.gif">Yöneticiler:</b><br />
<%Do Until admin_mems.Eof %> 
	<a href="members.asp?action=member_details&uid=<%=admin_mems("uye_id")%>" alt="Yönetici">
	<%=admin_mems("kul_adi")%> <br />
</a>
<%
admin_mems.MoveNext
Loop
%>
<%
Set hdm_mems = Server.CreateObject("ADODB.RecordSet")
hdm_mems.open "Select * FROM MEMBERS WHERE seviye=7",Connection,3,3
%>
<%
Set haber_mems = Server.CreateObject("ADODB.RecordSet")
haber_mems.open "Select * FROM MEMBERS WHERE seviye=4",Connection,3,3
%>
<b><img border="0" src="Images/Haber.gif">Haber Editörleri:</b><br />
<% If haber_mems.Eof OR haber_mems.Bof Then %> -------<br />
<%end if%>
<%Do Until haber_mems.Eof
%>
	<a href="members.asp?action=member_details&uid=<%=haber_mems("uye_id")%>">
	<%=haber_mems("kul_adi")%> <br />
</a>
<%
haber_mems.MoveNext
Loop
%>
<%
Set haber_mems = Server.CreateObject("ADODB.RecordSet")
haber_mems.open "Select * FROM MEMBERS WHERE seviye=8",Connection,3,3
%>
<b><img border="0" src="Images/linkler.gif">GaLeri Editörleri:</b><br />
<% If haber_mems.Eof OR haber_mems.Bof Then %> -------<br />
<%end if%>
<%Do Until haber_mems.Eof
%>
	<a href="members.asp?action=member_details&uid=<%=haber_mems("uye_id")%>">
	<%=haber_mems("kul_adi")%> <br />
</a>
<%
haber_mems.MoveNext
Loop
%>

<%
Set down_mems = Server.CreateObject("ADODB.RecordSet")
down_mems.open "Select * FROM MEMBERS WHERE seviye=3",Connection,3,3
%>
<b><img border="0" src="Images/yazi.gif">Makale Editörleri:</b><br />
<% If down_mems.Eof OR down_mems.Bof Then %>&nbsp; -------<br />
<%end if%>
<%Do Until down_mems.Eof
%>
	<a href="members.asp?action=member_details&uid=<%=down_mems("uye_id")%>">
	<%=down_mems("kul_adi")%> <br />
</a>
<%
down_mems.MoveNext
Loop
%>

<%
Set mak_mems = Server.CreateObject("ADODB.RecordSet")
mak_mems.open "Select * FROM MEMBERS WHERE seviye=2",Connection,3,3
%>
<b><img border="0" src="Images/download.gif">Dosya Editörleri:</b><br />
<% If mak_mems.Eof OR mak_mems.Bof Then %>&nbsp; -------<br />
<%end if%>
<%Do Until mak_mems.Eof
%>
	<a href="members.asp?action=member_details&uid=<%=mak_mems("uye_id")%>">
	<%=mak_mems("kul_adi")%> <br />
</a>
<%
mak_mems.MoveNext
Loop
%>

<%
Set forum_mems = Server.CreateObject("ADODB.RecordSet")
forum_mems.open "Select * FROM MEMBERS WHERE seviye=5",Connection,3,3
%>
<b><img border="0" src="Images/forum.gif">Forum Görevlileri:</b><br />
<% If forum_mems.Eof OR forum_mems.Bof Then %>&nbsp; -------<br />
<%end if%>
<%Do Until forum_mems.Eof
%>
	<a href="members.asp?action=member_details&uid=<%=forum_mems("uye_id")%>">
	<%=forum_mems("kul_adi")%> <br />
</a>
<%
forum_mems.MoveNext
Loop
%>