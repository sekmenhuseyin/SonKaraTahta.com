<!--#include file="../../INC/b_include.asp" -->
<marquee style="cursor:  url(../../images/site/mavi.cur);" onmouseover="this.stop();" onmouseout="this.start();" direction="up" height="90" SCROLLAMOUNT="1"><%
Set l5 = Connection.Execute("Select * FROM ARTICLES WHERE a_approved = True ORDER BY hit DESC")

For p = 1 To 20
if l5.eof Then exit For
%>
 <img border="0" src="images/yazi.gif"> <%=duz(l5("a_title"))%> <b>[<a href="Articles.asp?action=Read&aid=<%=l5("aid")%>"><%=a_read%></a>]</b><br /><%
l5.MoveNext
Next
%></td>