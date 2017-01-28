<!--#include file="../../INC/b_include.asp" -->
<marquee style="cursor:  url(../../images/site/mavi.cur);" onmouseover="this.stop();" onmouseout="this.start();" direction="up" height="90" SCROLLAMOUNT="1"><%
Set l5 = Connection.Execute("Select * FROM DOWNLOADS WHERE p_approved = True ORDER BY p_hit DESC")

For p = 1 To 20
if l5.eof Then exit For
%>
 <img border="0" src="images/download.gif"> <%=l5("p_name")%> <b>[<a href="Programs.asp?action=p_details&pid=<%=l5("pid")%>"><%=p_more%></a>]</b><br /><%
l5.MoveNext
Next
%></td>