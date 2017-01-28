<!--#include file="view.asp" -->
<%call UST:call ORTA
IF Request.QueryString("action")="show" Then
m=Request.QueryString("month"):m=QS_CLEAR(m,"false")
d=Request.QueryString("day"):d=QS_CLEAR(d,"false")
IF Len(d) < 2 AND S_LCID="1033" Then d="0"&d&"" ELSE d=d
IF Len(m) < 2 AND S_LCID="1033" Then m="0"&m&"" ELSE m=m
IF S_LCID="1033" then thisDate=""&d&"."&m&"."&Year(Now())&"" ELSE thisDate=""&m&"/"&d&"/"&Year(Now())&""

Set rs1=Server.CreateObject("ADODB.RecordSet"):SQL1="Select * FROM MEMBERS WHERE uyelik_tarihi='"&thisDate&"'":rs1.open SQL1,Connection,3,3
Set rs2=Server.CreateObject("ADODB.RecordSet"):SQL2="Select * FROM DOWNLOADS WHERE p_date='"&thisDate&"' AND p_approved=True":rs2.open SQL2,Connection,3,3
Set rs3=Server.CreateObject("ADODB.RecordSet"):SQL3="Select * FROM ARTICLES WHERE a_date='"&thisDate&"' AND a_approved=True":rs3.open SQL3,Connection,3,3
Set rs_b=Server.CreateObject("ADODB.RecordSet"):rs_b.open "Select * FROM MEMBERS WHERE Left(yas,5)='"&Left(thisDate,5)&"'",Connection,3,3%>
              <table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
                <tr> 
                  <td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=menu5%></font></td>
                </tr>
                <tr> 
<td align="center" valign="top" bgcolor="<%=menu_bg%>" class="td_menu">
<%
Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""1"" width=""100%"" align=""center"" class=""td_menu"" bgcolor="&td_border_color&">"
Response.Write "<tr style=""font-weight:bold"" height=""20""><td colspan=""3"" background="&sett("site_adres")&"/THEMES/"&Session("SiteTheme")&"/td_bg.gif>"&clndr_birthday&"</td></tr>"
Response.Write "<tr class=""td_menu2"" style=""font-weight:bold"" bgcolor="&menu_back&"><td width=""15%"">"&detail_memname&"</td><td width=""15%"">"&detail_name&"</td><td width=""15%"">"&detail_mail&"</td></tr>"
Do Until rs_b.Eof
If rs_b("mail_goster")=True Then
email=rs_b("email")
Else
email=""&detail_invisiblemail&""
End If
Response.Write "<tr bgcolor="&menu_bg&"><td width=""15%""><a href=""members.asp?action=member_details&uid="&rs_b("uye_id")&""">"&rs_b("kul_adi")&"</a></td><td width=""15%"">"&rs_b("isim")&"</td><td width=""15%"">"&email&"</td></tr>"
rs_b.MoveNext
Loop
Response.Write "</table>"
Response.Write "<hr size=1 color="&menu_back&">"
Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""1"" width=""100%"" align=""center"" class=""td_menu"" bgcolor="&td_border_color&">"
Response.Write "<tr style=""font-weight:bold"" height=""20""><td colspan=""3"" background="&sett("site_adres")&"/THEMES/"&Session("SiteTheme")&"/td_bg.gif>"&clndr_members&"</td></tr>"
Response.Write "<tr class=""td_menu2"" style=""font-weight:bold"" bgcolor="&menu_back&"><td width=""15%"">"&detail_memname&"</td><td width=""15%"">"&detail_name&"</td><td width=""15%"">"&detail_mail&"</td></tr>"
Do Until rs1.Eof
If rs1("mail_goster")=True Then
email=rs1("email")
Else
email=""&detail_invisiblemail&""
End If
Response.Write "<tr bgcolor="&menu_bg&"><td width=""15%""><a href=""members.asp?action=member_details&uid="&rs1("uye_id")&""">"&rs1("kul_adi")&"</a></td><td width=""15%"">"&rs1("isim")&"</td><td width=""15%"">"&email&"</td></tr>"
rs1.MoveNext
Loop
Response.Write "</table>"
Response.Write "<hr size=1 color="&menu_back&">"
Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""1"" width=""100%"" align=""center"" class=""td_menu"" bgcolor="&td_border_color&">"
Response.Write "<tr style=""font-weight:bold"" height=""20""><td background="&sett("site_adres")&"/THEMES/"&Session("SiteTheme")&"/td_bg.gif>"&clndr_files&"</td></tr>"
Do Until rs2.Eof
Response.Write "<tr bgcolor="&menu_bg&"><td><a href=""programs.asp?action=p_details&pid="&rs2("pid")&""">"&rs2("p_name")&"</a> ("&Left(rs2("p_info"),50)&")</td></tr>"
rs2.MoveNext
Loop
Response.Write "</table>"
Response.Write "<hr size=1 color="&menu_back&">"
Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""1"" width=""100%"" align=""center"" class=""td_menu"" bgcolor="&td_border_color&">"
Response.Write "<tr style=""font-weight:bold"" height=""20""><td background="&sett("site_adres")&"/THEMES/"&Session("SiteTheme")&"/td_bg.gif>"&clndr_articles&"</td></tr>"
Do Until rs3.Eof
Response.Write "<tr bgcolor="&menu_bg&"><td><a href=""articles.asp?action=read&aid="&rs3("aid")&""">"&rs3("a_title")&"</a> ("&rs3("a_writer")&")</td></tr>"
rs3.MoveNext
Loop
Response.Write "</table>"%>
</td>
</tr>
</table>
<%End IF
call ORTA2:call ALT%>