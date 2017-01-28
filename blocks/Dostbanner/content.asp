<!--#include file="../../INC/includes.asp" -->
<p align="center">
<MARQUEE style="cursor:  url(../../images/site/mavi.cur);" onmouseover=this.stop() onmouseout=this.start() scrollAmount=2 
      scrollDelay=90 direction=up height="170" width="95">
<%
Set topban = Connection.Execute(" Select * FROM DOSTBAN ")
Do Until topban.Eof
response.write "<a href="""&topban("giturl")&""" target=""_blank""><img src="""&topban("resurl")&""" width=""88"" height=""31""></a> <br />"
topban.MoveNext
Loop%></marquee></p>