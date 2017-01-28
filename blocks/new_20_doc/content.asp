<!--#include file="../../INC/b_include.asp" -->
<marquee style="cursor:  url(../../images/site/mavi.cur);" onmouseover="this.stop();" onmouseout="this.start();" direction="up" height="90" SCROLLAMOUNT="1"><%
Set n5 = Connection.Execute("Select * FROM ARTICLES WHERE a_approved = True ORDER BY a_date DESC")

For p = 1 To 20
if n5.eof Then exit For
%>
 <img border="0" src="images/yazi.gif"> <%=duz(n5("a_title"))%> <b>[<a href="Articles.asp?action=Read&aid=<%=n5("aid")%>"><%=a_read%></a>]</b><br /></marquee><%
n5.MoveNext
Next
%></td>