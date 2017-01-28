<!--#include file="../../INC/b_include.asp" -->
<marquee style="cursor:  url(../../images/site/mavi.cur);" onmouseover="this.stop();" onmouseout="this.start();" direction="up" height="90" SCROLLAMOUNT="1"><%
Set n5 = Connection.Execute("Select * FROM DOWNLOADS WHERE p_approved = True ORDER BY p_date DESC")

For p = 1 To 20
if n5.eof Then exit For
%>
 <img border="0" src="images/download.gif"> <%=n5("p_name")%> <b>[<a href="Programs.asp?action=p_details&pid=<%=n5("pid")%>"><%=p_more%></a>]</b><br /><%
n5.MoveNext
Next
%></td>