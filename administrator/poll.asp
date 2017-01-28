<% Response.Buffer = True %>
<!--#include file="includes.asp" -->
<%
If Session("Level") = "1" OR Session("Level") = "7" Then
act = Request.QueryString("action")
If act = "all" Then
Call polls
ElseIf act = "poll_details" Then
Call details
ElseIf act = "poll_delete" Then
Call delete
ElseIf act = "new_poll" Then
Call newpoll
ElseIf act = "poll_register" Then
Call regpoll
ElseIF act = "New" Then
call create_poll
End If

Sub create_poll
%>
<div class=td_menu><b>» Yeni Anket Yarat</b></div>
<form action="?action=new_poll" method="post">
<font face=verdana size=1>Soru : </font><input type="text" name="question" class="form1" size="30">&nbsp;<input type="submit" value="Yarat »" class="submit">
</form>
<%
End Sub

Sub polls

Set rs = Server.CreateObject("ADODB.RecordSet")
SQL = "Select * FROM POLLS ORDER BY poll_id DESC"
rs.open SQL,Connection,1,3

do while not rs.eof %>
<font face="tahoma" size="2"><b>» <a href="?action=poll_details&pid=<%=rs("poll_id")%>"><%=rs("p_question")%></a></b>&nbsp;&nbsp;<a href="?action=poll_delete&pid=<%=rs("poll_id")%>">[Sil]</a></font></font><br>
<%
rs.MoveNext
Loop

rs.Close
Set rs = Nothing

End Sub
Sub newpoll

qstn = duz(Request.Form("question"))

If qstn="" Then
Response.Write "<center><font face=verdana size=2 color=navy>Lütfen Anket Sorusunu Yazýnýz...</font></center>"
ELSE

Set entp = Server.CreateObject("ADODB.RecordSet")
entpSQL = "Select * FROM POLLS"
entp.open entpSQL,Connection,1,3

entp.AddNew
entp("p_question") = qstn
entp("p_date") = Date()
entp.Update
%>
<font face=tahoma size=2><b><%=qstn%></b></font><br><br>
<form action="?action=poll_register&pid=<%=entp("poll_id")%>" method="post">
<font face=verdana size=2>Seçenek 1 : </font><input type="text" name="alternative1" class="form1" size="40"><br>
<font face=verdana size=2>Seçenek 2 : </font><input type="text" name="alternative2" class="form1" size="40"><br>
<font face=verdana size=2>Seçenek 3 : </font><input type="text" name="alternative3" class="form1" size="40"><br>
<font face=verdana size=2>Seçenek 4 : </font><input type="text" name="alternative4" class="form1" size="40"><br>
<font face=verdana size=2>Seçenek 5 : </font><input type="text" name="alternative5" class="form1" size="40"><br>
<input type="submit" value="Seçenekleri Ekle ve Anketi Yarat" class="submit">
</form>
<%
entp.Close
Set entp = Nothing

End If
End Sub
Sub regpoll

pollid = Request.QueryString("pid")
alt1 = duz(Request.Form("alternative1"))
alt2 = duz(Request.Form("alternative2"))
alt3 = duz(Request.Form("alternative3"))
alt4 = duz(Request.Form("alternative4"))
alt5 = duz(Request.Form("alternative5"))

If alt1="" AND alt2="" Then
Response.Write "<center><font face=verdana size=2 color=navy>Anketin Doðru Çalýþmasý için EN AZ Ýki Seçenek Doldurmalýsýnýz...</font></center>"
ELSE

Set pent1 = Server.CreateObject("ADODB.RecordSet")
p1SQL = "Select * FROM POLL_ALTERNATIVES"
pent1.open p1SQL,Connection,1,3
Set pent2 = Server.CreateObject("ADODB.RecordSet")
p2SQL = "Select * FROM POLL_ALTERNATIVES"
pent2.open p2SQL,Connection,1,3
Set pent3 = Server.CreateObject("ADODB.RecordSet")
p3SQL = "Select * FROM POLL_ALTERNATIVES"
pent3.open p3SQL,Connection,1,3
Set pent4 = Server.CreateObject("ADODB.RecordSet")
p4SQL = "Select * FROM POLL_ALTERNATIVES"
pent4.open p4SQL,Connection,1,3
Set pent5 = Server.CreateObject("ADODB.RecordSet")
p5SQL = "Select * FROM POLL_ALTERNATIVES"
pent5.open p5SQL,Connection,1,3

pent1.AddNew
pent1("alternative") = alt1
pent1("pid") = pollid
pent1("a_counter") = 0
pent1.Update
pent2.AddNew
pent2("alternative") = alt2
pent2("pid") = pollid
pent2("a_counter") = 0
pent2.Update
If alt3<>"" Then
pent3.AddNew
pent3("alternative") = alt3
pent3("pid") = pollid
pent3("a_counter") = 0
pent3.Update
End If
If alt4<>"" Then
pent4.AddNew
pent4("alternative") = alt4
pent4("pid") = pollid
pent4("a_counter") = 0
pent4.Update
End If
If alt5<>"" Then
pent5.AddNew
pent5("alternative") = alt5
pent5("pid") = pollid
pent5("a_counter") = 0
pent5.Update
End If
Response.Redirect "?action=all"
END IF

End Sub
Sub delete

id = Request.QueryString("pid")
Set del = Server.CreateObject("ADODB.RecordSet")
delSQL = "DELETE * FROM POLLS WHERE poll_id = "&id&""
del.open delSQL,Connection,1,3
Set del_alts = Server.CreateObject("ADODB.RecordSet")
altsdelSQL = "DELETE * FROM POLL_ALTERNATIVES WHERE pid = "&id&""
del_alts.open altsdelSQL,Connection,1,3
Response.Redirect "?action=all"
End Sub

Sub details

pollid = Request.QueryString("pid")
Set prs = Server.CreateObject("ADODB.RecordSet")
pSQL = "Select * FROM POLLS WHERE poll_id = "&pollid&""
prs.open pSQL,Connection,3,3

Set prs2 = Server.CreateObject("ADODB.RecordSet")
pSQL2 = "Select * FROM POLL_ALTERNATIVES WHERE pid = "&pollid&""
prs2.open pSQL2,Connection,3,3
Set prss = Connection.Execute("Select * FROM POLL_ALTERNATIVES WHERE pid = "&pollid&"")
%>
<div align="center" class="td_menu" style="font-size:12px;font-weight:bold"><%=prs("p_question")%></div>
<br>
<table border="0" cellspacing="3" cellpadding="0" width="100%" align="center" class="td_menu">
<%
do while not prs2.eof
Set p_tpl = Connection.Execute("SELECT SUM(a_counter) AS count FROM POLL_ALTERNATIVES WHERE pid = "&prs2("pid")&"")
If p_tpl("count") = 0 Then
percent = "0"
Else
percent = Int((prs2("a_counter") / p_tpl("count")) * 100)
End If
%>
<tr><td width="40%" align="right"><b><%=prs2("alternative")%></b></td><td width="60%"><img src="../THEMES/default/poll_result.gif" width="<%=percent%>" height="12">&nbsp;&nbsp;( <%=prs2("a_counter")%> - <%=percent%>% )</td></tr>
<%
prs2.MoveNext
Loop %>
</table><br>
<center>
<font class="td_menu" size="1"><b>Katýlýmcý : </b><%=p_tpl("count")%></font>
</center>
<%
prs.Close
set prs = Nothing
prs2.Close
set prs2 = Nothing

End Sub
End If
%>