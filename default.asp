<!--#include file="view.asp" --><meta name="verify-v1" content="z/Fbs/SmTMuBFDapOFpPx7v7s9uqBLdKofgzVHa9SZI=" /><%call UST:call ORTA
Set rs=Server.CreateObject("ADODB.RecordSet"):sorgu="Select * FROM NEWS WHERE onay=True ORDER BY hid DESC":rs.open sorgu,Connection,3,3
Page=Request.QueryString("Page"):Page=QS_CLEAR(Page,"false"):if Page="" then Page=1
''son 5 haber ve son 5 forum
mp_news=sett("mp_news")
mp_forum=sett("mp_forum")
IF mp_news=True Then Set hbr5=Connection.Execute("Select * FROM NEWS WHERE onay=True ORDER BY hid DESC")
IF mp_forum=True Then Set l_forum=Connect.Execute("Select * FROM MESSAGES WHERE topic=True and locked=false ORDER BY last_post DESC")
IF mp_news=True OR mp_forum=True Then%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" align="center"><tr>
<%End IF
IF mp_news=True and mp_forum=True Then
tblWidth=" width='98%'":tbl1Position="":tbl2Position=" align=right"
ELSe
tblWidth=" width='100%'":tbl1Position=" align=center":tbl2Position=" align=center"
End IF
''son 5 haber
IF mp_news=True Then%>
<td valign="top"<%=tbl1Position%>><table border="0" cellpadding="2" cellspacing="0" style="<%=TableBorder%>"<%=tblWidth%>>
<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=top_haber%></font></td></tr>
<tr><td align="left" valign="top" bgcolor="<%=menu_bg%>" class="td_menu">
<marquee style="cursor:  url(images/site/mavi.cur);" onMouseOver="this.stop();" onMouseOut="this.start();" direction="up" height="90" SCROLLAMOUNT="1">
<% If hbr5.Eof OR hbr5.Bof Then %>
<font class=td_menu><%=no_news%></font>
<% Else
For last_news=1 TO 5
IF hbr5.eof then exit for%>
<font face=webdings color="<%=text%>" size=1>4</font><a href="News.asp?action=Read&hid=<%=hbr5("hid")%>"><%=duz(Left(hbr5("baslik"),35))%></a><br />
<%hbr5.MoveNext
Next
End If%></marquee>
</td></tr>
</table></td>
<%End IF
''son 5 forum
IF mp_forum=True Then%>
<td valign="top"<%=tbl2Position%>><table border="0" cellpadding="2" cellspacing="0" style="<%=TableBorder%>"<%=tblWidth%>>
<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=top_forum%></font></td></tr>
<tr><td align="left" valign="top" bgcolor="<%=menu_bg%>" class="td_menu">
<marquee style="cursor:  url(images/site/mavi.cur);" onMouseOver="this.stop();" onMouseOut="this.start();" direction="up" height="90" SCROLLAMOUNT="1">
<% If l_forum.Eof OR l_forum.Bof Then %>
<font class=td_menu><%=no_forum%></font>
<% Else
For last_forum=1 TO 5
if l_forum.eof then exit for%>
<font face=webdings color="<%=text%>" size=1>4</font><a href="Forum.asp?action=Topic&id=<%=l_forum("mid")%>"><%=duz(SmiLey(Left(l_forum("mtitle"),35)))%></a> (<%=l_forum("cvp_sayisi")%>&nbsp;<%=replies_count%>)<br />
<%l_forum.MoveNext
Next
End If %></marquee>
</td></tr>
</table></td>
<%End IF
IF sett("mp_news")=True OR sett("mp_forum")=True Then%>
</tr></table><br />
<%End IF
''sabit haberler
Set rs2=Server.CreateObject("ADODB.RecordSet"):sorgu2="Select * FROM FIXED order by fid desc":rs2.open sorgu2,Connection,3,3
If rs2.Eof OR rs2.Bof Then%>
<font class=td_menu></font>
<%Else
CharToWrite=1000
FOR I=1 TO UBound(saryHTMLtags)%>
<table border="0" cellpadding="0" cellspacing="0" width="100%"  align="center" style="<%=TableBorder%>">
<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><center><font class="block_title"><b><%=rs2("sf_baslik")%></b></font><center></td></tr>
<tr><td align="center" valign="top" bgcolor="<%=menu_back%>" class="td_menu">
<table border="0" cellspacing="5" cellpadding="0" align="center" width="100%" class="td_menu">
<tr><td valign="top">
<%IF rs2("resim")="" Then Response.Write " " Else%><img src="Images/<%=rs2("resim")%>" align="left">
<%If Len(rs2("f_metin")) <= CharToWrite Then Response.Write MakeURL(DelHTML(rs2("f_metin"))) Else Response.Write ""&DelHTML(Left(rs2("f_metin"),CharToWrite))&" ..."%>
</tr><tr><td><font size="1" class="td_menu"><b>[<a href="news.asp?Action=Read_sbt&fid=<%=rs2("fid")%>"><%=a_read%></a>]</b></font></td></tr>
</table>
</td></tr>
</table><br />
<%rs2.MoveNext
Next
End If
''haberler
If rs.Eof OR rs.Bof Then %>
<font class=td_menu><%=no_news%></font><br />
<%eLse
CharToWrite=500
rs.pagesize=sett("haber_sayisi"):rs.absolutepage=Page:sayfa=rs.pagecount
For J=1 TO rs.pagesize
IF rs.EoF OR rs.BoF Then Exit For
Set comments=Connection.Execute("Select Count(*) AS count FROM COMMENTS WHERE nid="&rs("hid")&" AND page='news'")
Set cats=Connection.Execute("Select * FROM NEWS_CATS WHERE cat_id="&rs("cat")&"")%>
<table border="0" cellpadding="0" cellspacing="0" width="100%"  align="center" style="<%=TableBorder%>">
<tr> <td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=duz(rs("baslik"))%></font></td></tr>
<tr> <td align="center" valign="top" bgcolor="<%=menu_back%>" class="td_menu">
<table border="0" cellspacing="5" cellpadding="0" align="center" width="100%" class="td_menu">
<tr><td valign="top">
<a href="News.asp?action=News&id=<%=cats("cat_id")%>"><img src="<%=cats("cat_image")%>" border="0" align="left" hspace="5" vspace="0" alt="<%=cats("cat_info")%>"></a>
<%IF Len(rs("resim")) > 0 Then%><img src="<%=rs("resim")%>" align="right" height=150><%End IF
If Len(rs("metin")) <= CharToWrite Then Response.Write MakeURL(DelHTML(rs("metin"))) Else Response.Write ""&DelHTML(Left(rs("metin"),CharToWrite))&"..."%>
</td></tr><tr><td><font size="1" class="td_menu"><b>[<a href="news.asp?Action=Read&hid=<%=rs("hid")%>"><%=a_read%></a>]</b></font></td></tr>
</table>
<hr size="1" color="<%=td_border_color%>"><div align="center" class="td_menu2">
<%Set nwmem=Connection.Execute("Select * FROM MEMBERS WHERE kul_adi='"&rs("ekleyen")&"'")
If NOT nwmem.Eof then %><a href="members.asp?action=member_details&uid=<%=nwmem("uye_id")%>"><%End IF %><%=duz(rs("ekleyen"))%><% IF NOT nwmem.Eof then %></a><% End If %>
 | <%=rs("tarih")%> (<%=rs("okunma")%>&nbsp;<%=h_readed%>)<br />( <% If Len(rs("metin")) >= CharToWrite Then %><a href="news.asp?Action=Read&hid=<%=rs("hid")%>"><b><%=news_read%></b></a> | <% end if %><a href="comments.asp?action=comments&page=news&id=<%=rs("hid")%>"><%=news_comments%>(<%=comments("count")%>)</a> | <a href="news.asp?Action=Print&hid=<%=rs("hid")%>" target="_blank"><img src="images/site/print.gif" border="0" align="absmiddle" alt="<%=h_print%>"></a> <a href="news.asp?Action=SendFriend&ST=1&hid=<%=rs("hid")%>"><img src="images/site/sendfriend.gif" border="0" align="absmiddle" alt="<%=h_sendfriend%>"></a> )
</div>
</td></tr>
</table><br />
<%rs.MoveNext
Next
End If
''sayfa numaralarý%>
<br /><table border="0" cellspacing="1" cellpadding="0" align="center" width="70%" bgcolor="<%=td_border_color%>" class="td_menu">
<tr><td bgcolor="<%=menu_back%>" align="center"><b>
<%if sayfa>1 then response.Write(top_sayfa2) else response.Write(top_sayfa1)%> : 
<%If cint(Page)>1 Then Response.Write "&nbsp;<a href=""?Page="&(cint(Page)-1)&""" style=""font-size:10px"">« "&previous_page&"</a> - "
For y=1 to sayfa
IF cint(Page)=cint(y) then Response.Write " <font class=""td_menu"" style=""font-size:10px"">["&y&"]</font> " ELSE Response.Write " <a href=""?Page="&y&""" style=""font-size:10px"">["&y&"]</a>"
Next
If NOT rs.Eof Then Response.Write " - <a href=""?Page="&(cint(Page)+1)&""" style=""font-size:10px"">"&next_page&" »</a>"%>
<br /><%=top_haberler%> :</font><font face="verdana" size="1"> <b><%=rs.recordcount%></b> 
</b></td></tr>
</table>
<%IF sett("mp_forum")=True Then l_forum.Close:Set l_forum=Nothing
IF sett("mp_news")=True Then hbr5.Close:Set hbr5=Nothing
rs.Close:Set rs=Nothing
call ORTA2:call ALT%>