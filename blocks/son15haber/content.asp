<!--#include file="../../INC/b_include.asp" -->
<marquee style="cursor:  url(../../images/site/mavi.cur);" onmouseover="this.stop();" onmouseout="this.start();" direction="up" height="90" SCROLLAMOUNT="1"><%
Set hbr5 = Connection.Execute("Select * FROM NEWS WHERE onay = True ORDER BY hid DESC")
For last_news = 1 TO 15
IF hbr5.eof then exit for
%>
<img src="images/haber.gif" align="absmiddle">&nbsp<a href="News.asp?action=Read&hid=<%=hbr5("hid")%>"><%=duz(Left(hbr5("baslik"),30))%></a> &nbsp(<%=hbr5("ekleyen")%>)<br /></marquee><%
hbr5.MoveNext
Next%></td>