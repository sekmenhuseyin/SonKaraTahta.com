<!--#include file="../../INC/b_include.asp" -->
<%
Set cats = Connection.Execute("Select * FROM MENU_CATS ORDER BY mc_queue")
Do Until cats.EoF
Response.Write "<b>» " & cats("mc_title") & "</b><br />"
Set module = Connection.Execute("SELECT * FROM MODULES WHERE mactive = True AND mcat = "&cats("mc_id")&" ORDER BY mname ASC")
Set links = Connection.Execute("Select * FROM MENU_LINKS WHERE m_status = True AND m_cat = "&cats("mc_id")&"")
Set pages = Connection.Execute("SELECT * FROM PAGES WHERE p_cat = "&cats("mc_id")&" ORDER BY p_title ASC")
Do Until links.EoF
IF links("m_style") = "normal" Then
strStyle = ""
ELSE
strStyle = "target="""&links("m_style")&""""
End IF
Response.Write "<img src=""IMAGES/Site/menu_select.gif"" align=""absmiddle""> <a href="""&links("m_url")&""" "&strStyle&">" & links("m_name") & "</a><br />"
links.MoveNext
Loop
Do Until module.Eof
name = module("mname")
%>
<img src="IMAGES/Site/menu_select.gif" align="absmiddle"> <a href="modules.asp?name=<%=name%>"><%=name%></a><br />
<%
module.MoveNext
Loop
Do Until pages.Eof
Response.Write "<img src=""IMAGES/Site/menu_select.gif"" align=""absmiddle""> <a href=""pages.asp?id="&pages("p_id")&""">"&pages("p_title")&"</a><br />"
pages.MoveNext
Loop
cats.MoveNext
Loop

Set cats = Nothing
Set module = Nothing
Set links = Nothing
Set pages = Nothing
%>