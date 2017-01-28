<!--#include file="includes.asp" -->
<%
IF Request.QueryString("op")="editCat" Then
IF Session("Level")="1" Then
Set Cat=Connection.Execute("Select * FROM KATEGORILER WHERE CID="&Request.QueryString("id")&"")
%>
<form action="?name=<%=Request.QueryString("name")%>&op=updateCat&id=<%=Request.QueryString("id")%>" method="post">
Kategori Adý : <input type="text" name="c_name" size="30" class="form1" value="<%=Cat("CNAME")%>">&nbsp;<input type="submit" value="<%=s_submit%>" class="submit">
</form>
<%
End IF
ElseIF Request.QueryString("op")="updateCat" Then
IF Session("Level")="1" Then
cname=duz(Request.Form("c_name"))
IF cname="" Then
Response.Write  err3
ELSE
Connection.Execute("UPDATE KATEGORILER SET CNAME='"&cname&"' WHERE CID="&Request.QueryString("id")&"")
Response.Redirect "?name="&Request.QueryString("name")&""
End IF
End IF
ElseIF Request.QueryString("op")="delCat" Then
IF Session("Level")="1" Then
Set sC=Connection.Execute("Select * FROM KATEGORILER WHERE CID="&Request.QueryString("id")&"")
IF sC.Eof Then
Response.Write err4
ELSE
Set dC=Server.CreateObject("ADODB.RecordSet")
dcSQL="DELETE * FROM KATEGORILER WHERE CID="&Request.QueryString("id")&""
dC.open dcSQL,Connection,1,3

Set dL=Server.CreateObject("ADODB.RecordSet")
dlSQL="DELETE * FROM LINKLER WHERE CID="&Request.QueryString("id")&""
dL.open dlSQL,Connection,1,3
Response.Redirect Request.ServerVariables("HTTP_REFERER")
End IF
End IF
ElseIF Request.QueryString("op")="redirect" Then
Set RedL=Server.CreateObject("ADODB.RecordSet")
RLSQL="Select * FROM LINKLER WHERE LID="&Request.QueryString("id")&""
RedL.open RLSQL,Connection,1,3

IF RedL.Eof OR RedL.Bof Then
Response.Write err4
ELSE
RedL("LHIT")=RedL("LHIT") + 1
RedL.Update
Response.Redirect RedL("LURL")
End IF
ElseIF Request.QueryString("op")="links" Then
Set rs=Server.CreateObject("ADODB.RecordSet")
SQL="Select * FROM LINKLER WHERE CID="&Request.QueryString("cat")&" ORDER BY LID DESC"
rs.open SQL,Connection,1,3
IF rs.Eof Then
Response.Write err1
End IF
Do Until rs.Eof
%>
<table border="0" cellspacing="1" cellpadding="1" align="center" width="90%" class="td_menu" bgcolor="<%=td_border_color%>">
<tr bgcolor="<%=menu_back%>">
<td><img src="IMAGES/site/site.gif" align="absmiddle">&nbsp;<b><a href="?name=<%=Request.QueryString("name")%>&op=redirect&id=<%=rs("LID")%>" target="_blank"><%=rs("LTITLE")%></a></b> | <%=rs("LEKLEYEN")%> - <%=rs("LDATE")%> - <%=s_hit%> : <b><%=rs("LHIT")%></b>
<%
IF Session("Level")="1" Then
Response.Write " /// <b><a href=""?name="&Request.QueryString("name")&"&op=delLink&id="&rs("LID")&""">Sil</a></b>"
End IF
%>
</td>
</tr>
<tr bgcolor="<%=menu_back%>">
<td><%=rs("LINFO")%></td>
</tr>
</table><br />
<%
rs.MoveNext
Loop
%>
<br />
<%=t_links%> : <b><%=rs.recordcount%></b>
<br /><br />
<a href="?name=<%=Request.QueryString("name")%>"><%=p_back%></a>
<%
ElseIF Request.QueryString("op")="New" Then %>
<table border="0" cellspacing="0" cellpadding="2" align="center" width="100%" class="td_menu">
<form action="?name=<%=Request.QueryString("name")%>&op=Reg" method="post">
<tr style="font-weight:bold">
<td width="30%" align="right"><%=s_name%> : </td>
<td width="70%"><input type="text" name="s_name" size="30" class="form1"></td>
</tr>
<tr style="font-weight:bold">
<td width="30%" align="right"><%=s_url%> : </td>
<td width="70%"><input type="text" name="s_url" size="30" class="form1" value="http://"></td>
</tr>
<tr style="font-weight:bold">
<td width="30%" align="right"><%=s_add%> : </td>
<td width="70%">
<% IF Session("Enter")="Yes" Then
Response.Write "<b>"&Session("Member")&"</b>"
ELSE
%>
<input type="text" name="s_add" size="30" class="form1">
<% End IF %>
</td>
</tr>
<tr style="font-weight:bold">
<td width="30%" align="right" valign="top"><%=s_info%> : </td>
<td width="70%"><textarea name="s_info" cols="30" rows="5" class="form1"></textarea></td>
</tr>
<tr style="font-weight:bold">
<td width="30%" align="right" valign="top"><%=s_kat%> : </td>
<td width="70%">
<select name="s_kat" size="1" class="form1">
<%
Set LCats=Server.CreateObject("ADODB.RecordSet")
LCSQL="Select * FROM KATEGORILER ORDER BY CNAME ASC"
LCats.open LCSQL,Connection,1,3
Do Until LCats.Eof
%>
<option value="<%=LCats("CID")%>"><%=LCats("CNAME")%></option>
<%
LCats.MoveNext
Loop
%>
</select>
</td>
</tr>
<tr style="font-weight:bold">
<td width="30%"></td>
<td width="70%"><input type="submit" class="submit" value="<%=s_submit%>"></td>
</tr>
</form>
</table>
<%
ElseIF Request.QueryString("op")="Reg" Then
IF Session("Enter")="Yes" Then
add=Session("Member")
ELSE
add=duz(Request.Form("s_add"))
End IF
name=duz(Request.Form("s_name"))
info=duz(Request.Form("s_info"))
kat=duz(Request.Form("s_kat"))
url=duz(Request.Form("s_url"))

IF name="" OR add="" OR info="" OR url="" Then
Response.Write err3
ELSE
SET LEnt=Server.CreateObject("ADODB.RecordSet")
LSQL="Select * FROM LINKLER"
LEnt.open LSQL,Connection,1,3

LEnt.AddNew
LEnt("LTITLE")=name
LEnt("LURL")=url
LEnt("LINFO")=info
LEnt("LDATE")= Date()
LEnt("LHIT")=0
LEnt("LEKLEYEN")=add
LEnt("CID")=kat
LEnt.Update
Response.Write success
End IF
ElseIF Request.QueryString("op")="delLink" Then
IF Session("Level")="1" Then
Set rsL=Connection.Execute("Select * FROM LINKLER WHERE LID="&Request.QueryString("id")&"")
IF rsL.Eof Then
Response.Write err4
ELSE
Set delLink=Server.CreateObject("ADODB.RecordSet")
dLSQL="DELETE * FROM LINKLER WHERE LID="&Request.QueryString("id")&""
delLink.open dLSQL,Connection,1,3

Response.Redirect Request.ServerVariables("HTTP_REFERER")
End IF
End IF
ElseIf Request.QueryString("op")="kat_ekle" Then
IF Session("Level")="1" Then
k_name=duz(Request.Form("kat_name"))
IF k_name="" Then
Response.Write err2
ELSE
LinkSQL="INSERT INTO KATEGORILER(CNAME) VALUES ('"&k_name&"');"
Connection.Execute(LinkSQL)
Response.Redirect Request.ServerVariables("HTTP_REFERER")	
End IF
End IF
ELSE
Set rs=Server.CreateObject("ADODB.RecordSet")
SQL="Select * FROM KATEGORILER ORDER BY CNAME ASC"
rs.open SQL,Connection,1,3
IF rs.Eof OR rs.Bof Then
Response.Write "<b>"&err0&"</b><br /><br />"
ELSE
%>
<table border="0" cellspacing="2" cellpadding="2" width="90%" class="td_menu" bgcolor="<%=menu_bg%>">
<%
While Not rs.EOF

Set links=Connection.Execute("Select Count(*) AS s_count FROM LINKLER WHERE CID="&rs("CID")&"")
%>
  <tr>
    <td width="50%" height="15" bgcolor="<%=menu_back%>"><img src="<%=sett("site_adres")%>/IMAGES/site/folder.gif" align="absmiddle">&nbsp;<a href="?name=<%=Request.QueryString("name")%>&op=links&cat=<%=rs("CID")%>"><b><%=rs("CNAME")%></b></a> (<%=links("s_count")%>)
<% IF Session("Level")="1" Then %>
 - [<a href="?name=<%=Request.QueryString("name")%>&op=editCat&id=<%=rs("CID")%>">Düzenle</a>-<a href="?name=<%=Request.QueryString("name")%>&op=delCat&id=<%=rs("CID")%>">Sil</a>]
<% End IF %>
</td>
    <%
rs.MoveNext
Set links=Nothing
If Not rs.EOF Then

Set links=Connection.Execute("Select Count(*) AS s_count FROM LINKLER WHERE CID="&rs("CID")&"")
%>
    <td width="50%" height="15" bgcolor="<%=menu_back%>"><img src="<%=sett("site_adres")%>/IMAGES/site/folder.gif" align="absmiddle">&nbsp;<a href="?name=<%=Request.QueryString("name")%>&op=links&cat=<%=rs("CID")%>"><b><%=rs("CNAME")%></b></a> (<%=links("s_count")%>)
<% IF Session("Level")="1" Then %>
 - [<a href="?name=<%=Request.QueryString("name")%>&op=editCat&id=<%=rs("CID")%>">Düzenle</a>-<a href="?name=<%=Request.QueryString("name")%>&op=delCat&id=<%=rs("CID")%>">Sil</a>]
<% End IF %>
</td>
<% End If

If Not rs.EOF Then
rs.MoveNext
End If
%>
</tr>
<% Wend %>
</table>
<br />
<%
End IF
Set tCats=Connection.Execute("Select Count(*) AS cats_count FROM KATEGORILER")
Set tLinks=Connection.Execute("Select Count(*) AS links_count FROM LINKLER")
Set tHits=Connection.Execute("Select SUM(LHIT) AS hits_count FROM LINKLER")
%>
[ <%=t_cat%> : <b><%=tCats("cats_count")%></b> | <%=t_links%> : <b><%=tLinks("links_count")%></b> | <%=t_hits%> : <b><%=tHits("hits_count")%></b> ]
<br /><br />
<b>[</b> <a href="?name=<%=Request.QueryString("name")%>&op=New"><%=new_link%></a> <b>]</b>
<% IF Session("Level")="1" Then %>
<br />
<form action="?name=<%=Request.QueryString("name")%>&op=kat_ekle" method="post">
Kategori Adý : <input type="text" name="kat_name" size="20" class="form1">&nbsp;<input type="submit" class="submit" value="Kategori Ekle">
</form>
<%
End IF
End IF
%>