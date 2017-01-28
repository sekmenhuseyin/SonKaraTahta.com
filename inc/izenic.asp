<!--#include file="c.inc" -->
<table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="<%=td_border_color%>" class="td_menu" style="border-collapse: collapse" bordercolor="#111111">
<tr><td bgcolor="<%=menu_back%>" align="center"><b>© <%=sett("site_isim")%></b></td></tr>
<tr><td bgcolor="<%=menu_back%>" align="center"><MARQUEE style="cursor:  url(images/site/mavi.cur);" height="51" scrollamount="7" scrolldelay="1" onmouseover='this.stop()' onmouseout='this.start()' width="100%">
<%Set uban=Connection.Execute("Select * FROM UFAKBAN WHERE aktif=TRUE"):Do Until uban.Eof
response.write "<a href="""&uban("giturl")&""" target=""_blank""><img src="""&uban("resurl")&""" width=""88"" height=""31"" border=""0""></a>&nbsp;&nbsp;&nbsp;"
uban.MoveNext
Loop%></marquee></td></tr>
</table>
<%sett.Close:Set sett=Nothing:LB.Close:Set LB=Nothing:RB.Close:Set RB=Nothing
Connection.Close:Set Connection=Nothing:Connect.Close:Set Connect=Nothing%>
<b><font style="font-size: 7pt;" face="Tahoma">(Powered by <a target="_blank" href="http://www.mini-nuke.de">Mini-NUKE Fixed Versiyon)</a> &amp; (Coders <a target="_blank" href="http://www.mini-nuke.info">Mini-Nuke TeaM</a>)</font></b>
