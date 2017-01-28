<!--#include file="../../INC/b_include.asp" -->
<marquee style="cursor:  url(../../images/site/mavi.cur);" onmouseover="this.stop();" onmouseout="this.start();" direction="up" height="90" SCROLLAMOUNT="1"><%
Set l_forum = Connect.Execute("Select * FROM MESSAGES WHERE topic = True ORDER BY last_post DESC")


For last_forum = 1 TO 15
if l_forum.eof then exit for

Set reps = Connect.Execute("Select Count(*) AS rCount FROM MESSAGES WHERE mhit = "&l_forum("mid")&" AND topic = False")
%>
<img src="images/forum.gif" align="bottom">&nbsp<a href="Forum.asp?action=Topic&id=<%=l_forum("mid")%>"><%=duz(SmiLey(Left(l_forum("mtitle"),35)))%></a> (<%=reps("rCount")%>&nbsp;<%=replies_count%>)<br /></marquee><%
l_forum.MoveNext
Next
%></td>