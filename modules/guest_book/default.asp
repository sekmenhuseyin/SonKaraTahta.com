<!--#include file="includes.asp" -->
<% IF Request.QueryString("op")="New" Then %>
<table border="0" cellspacing="0" cellpadding="2" align="center" width="100%" class="td_menu">
  <form action="?name=<%=Request.QueryString("name")%>&op=Reg" method="post">
    <tr style="font-weight:bold">
      <td width="20%" align="right"><%=gb_name%> : </td>
      <td align="left"><%
		IF Session("Enter")="Yes" Then
			Response.Write Session("Member")
		ELSE%>
			<input type="text" name="gb_name" size="50" class="form1">
        <% End IF %>
      </td>
    </tr>
    <tr style="font-weight:bold">
      <td align="right"><%=gb_email%> : </td>
      <td><%
		IF Session("Enter")="Yes" Then
			Response.Write Session("Email")
		ELSE%>
			<input type="text" name="gb_email" size="50" class="form1">
        <% End IF %>
      </td>
    </tr>
    <tr style="font-weight:bold">
      <td colspan="2" bgcolor=buttonface>
		<script src="htmlarea/editor.js" language="Javascript1.2"></script>
		<script language="JavaScript1.2" defer>editor_generate('gb_yazi');</script>
		<textarea name="gb_yazi" style="width:100%" rows="7" class="form1"></textarea>
	  </td>
    </tr>
    <tr style="font-weight:bold">
      <td colspan="2" align="center"><input style="width:80%;" type="submit" class="submit" value="<%=gb_submit%>"></td>
    </tr>
  </form>
</table>


<%ElseIF Request.QueryString("op")="Reg" Then
	Set objSet=Connection.Execute("Select * FROM SETTINGS")
	IF DateDiff("N" ,Session("FloodTimer_GB"),Now()) > objSet("flood_time") Then
		IF Session("Enter")="Yes" Then
			name=Session("Member")
			email=Session("Email")
		ELSE
			name=duz(Request.Form("gb_name"))
			email=duz(Request.Form("gb_email"))
		End IF
		yazi=Request.Form("gb_yazi")
		IF name="" OR email="" OR yazi="" Then
			Response.Write err1
		ELSE
			SET gbEnt=Server.CreateObject("ADODB.RecordSet")
			gbSQL="Select * FROM GUESTBOOK"
			gbEnt.open gbSQL,Connection,1,3
			gbEnt.AddNew
			gbEnt("NAME")=name
			gbEnt("EMAIL")=email
			gbEnt("YAZI")=yazi
			gbEnt.Update
			Response.Write success
			Session("FloodTimerStart_GB")="Yes"
			Set Cmsg=Connection.Execute("Select * FROM MEMBERS WHERE kul_adi='"&Session("Member")&"'")
			Session("FloodTimer_GB")=Cmsg("son_tarih")
			Set Cmsg=Nothing
		END IF
	ELSE
		Response.Write "<div align=center class=td_menu2><b>"&err3&"</b></div>"
	End IF
	Set objSet=Nothing


ElseIF Request.QueryString("op")="del" Then
	Set rsGB=Connection.Execute("Select * FROM GUESTBOOK WHERE ID="&Request.QueryString("id")&"")
	IF rsGB.Eof Then
		Response.Write err2
	ELSE
		Set delGB=Server.CreateObject("ADODB.RecordSet")
		dgbSQL="DELETE * FROM GUESTBOOK WHERE ID="&Request.QueryString("id")&""
		delGB.open dgbSQL,Connection,1,3
		Response.Redirect Request.ServerVariables("HTTP_REFERER")
	End IF
ELSE

	Set rs=Server.CreateObject("ADODB.RecordSet")
	SQL="Select * FROM GUESTBOOK ORDER BY ID DESC"
	rs.open SQL,Connection,1,3
	IF rs.Eof OR rs.Bof Then
		Response.Write "<b>"&err0&"</b>"
	ELSE
		Do Until rs.Eof%>
			<table border="0" cellspacing="1" cellpadding="1" width="90%" class="td_menu" bgcolor="<%=table_border_color%>">
			  <tr bgcolor="<%=background%>">
				<td><%=rs("NAME")%> - <a href="mailto:<%=rs("EMAIL")%>"><%=rs("EMAIL")%></a>
				  <%IF Session("Level")="1" Then
						Response.Write " / <a href=""?name="&Request.QueryString("name")&"&op=del&id="&rs("ID")&"""><b>Bu Yorumu Sil</b></a>"
					End IF%>
				</td>
			  </tr>
			  <tr bgcolor="<%=background%>">
				<td><%=rs("YAZI")%></td>
			  </tr>
			</table>
			<br />
			<%rs.MoveNext
		Loop
	End IF%>
	<br />
	<b>[</b> <a href="?name=<%=Request.QueryString("name")%>&op=New">Deftere Yaz</a> <b>]</b>
<% End IF %>
