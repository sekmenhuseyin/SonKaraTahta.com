<%Set sRs=Server.CreateObject("ADODB.RecordSet"):sRs.open "Select * FROM SMILEYS ORDER BY s_info ASC",Connection,3,3%>
<table border="0" cellspacing="2" cellpadding="2" width="100%" class="td_menu2"><% Do Until sRs.EoF %>
<tr><td width="50%"><%=sRs("s_info")%></td><td width="50%" align="center"><img src="../IMAGES/smileys/<%=sRs("s_info")%>.gif"></td></tr>
<%sRs.MoveNext
Loop%></table>
<%sRs.Close:Set sRs=Nothing%>