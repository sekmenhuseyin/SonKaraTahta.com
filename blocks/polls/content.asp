<!--#include file="../../INC/b_include.asp" -->
<div align="center"><%
Set poll = Connection.Execute("Select * FROM POLLS ORDER BY p_date DESC")
IF poll.Eof OR poll.Bof Then
	Response.Write error14
ELSE
Set p_res = Connection.Execute("Select * FROM POLL_ALTERNATIVES WHERE pid = "&poll("poll_id")&"")
Set per = Connection.Execute("Select * FROM POLL_ALTERNATIVES WHERE pid = "&poll("poll_id")&"")

vote = 0
Do While Not per.EOF
	vote = vote +  per("a_counter")
	per.MoveNext
Loop%>
<font class="td_menu"><%=poll("p_question")%><br /></font><hr />
<form action="polls.asp?action=poll_process&poll_id=<%=poll("poll_id")%>" method="post" name="form">
<table border="0" cellspacing="1" cellpadding="0" width="100%" class="td_menu">
<%do while not p_res.eof
	Set p_tpl = Connection.Execute("SELECT SUM(a_counter) AS count FROM POLL_ALTERNATIVES WHERE pid = "&p_res("pid")&"")
	If p_res("a_counter") = "0" Then
		percent = "0"
	Else
		percent = Int((p_res("a_counter") / vote) * 100)
	End If
	IF Request.Cookies("MiniNUKE_POLL")("VOTE") = ""&poll("poll_id")&"" Then%>
		<tr bgcolor="<%=menu_back%>">
		  <td><b><%=p_res("alternative")%></b><br /><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/poll_result.gif" width="<%=percent/2%>" height="12">&nbsp;(<%=p_res("a_counter")%> - <%=percent%>%)</td>
		</tr>
	<% Else %>
		  <tr>
			<td width="1%"><input type="radio" name="alternative" value="<%=p_res("a_id")%>"></td>
			<td width="99%" bgcolor="<%=background%>"><%=p_res("alternative")%><br />
			  (<%=p_res("a_counter")%> - <%=percent%>%)</td>
		  </tr>
	<%End If
	p_res.MoveNext
Loop%>
</table>
<%IF Request.Cookies("MiniNUKE_POLL")("VOTE") <> ""&poll("poll_id")&"" Then%>
  <br />
  <input type="submit" value="<%=poll_button%>" style="width:100%" class="submit" onClick="this.form.submit();this.disabled=true; return true;">
</form>
<% End IF %>
<font class="td_menu" style="font-size:10px"><b><%=poll_katilan%> : </b><%=p_tpl("count")%></font><br />
<br />
<font class="td_menu"> <b>[</b> <a href="polls.asp?action=polls"><%=poll_others%></a> <b>|</b> <a href="polls.asp?action=poll_details&pid=<%=poll("poll_id")%>"><%=poll_results%></a> <b>]</b> </font>
<% End If %>
</div>
