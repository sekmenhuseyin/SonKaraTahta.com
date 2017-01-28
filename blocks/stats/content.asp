<!--#include file="../../INC/b_include.asp" -->
<%IF Request.Cookies("MiniNuke_Counter")("C_ENTER") <> "Yes" AND Request.Cookies("MiniNuke_Counter")("C_DATE") <> Date() Then
Set count = Connection.Execute("Select * FROM WEBCOUNTER WHERE cdate = Date()-0")
Set countAdd = Connection.Execute("UPDATE WEBCOUNTER SET today = "&count("today")&" + 1 WHERE cdate = Date()-0")
Response.Cookies("MiniNuke_Counter")("C_ENTER") = "Yes"
Response.Cookies("MiniNuke_Counter")("C_DATE") = Date()
End IF
Session.LCID = S_LCID
Set mem = Connection.Execute("Select * FROM MEMBERS ORDER BY uye_id DESC")
Set memc = Connection.Execute("Select Count(*) AS count FROM MEMBERS")
Set counter = Connection.Execute("Select * FROM WEBCOUNTER WHERE cdate = date()-0")
Set counter2 = Connection.Execute("Select * FROM WEBCOUNTER WHERE cdate = date()-1")
Set counter_t = Connection.Execute("SELECT SUM(today) AS count FROM WEBCOUNTER")
Set today = Connection.Execute("Select Count(*) AS m_count FROM MEMBERS WHERE uyelik_tarihi=date()-0")
Set yesterday = Connection.Execute("Select Count(*) AS m_count FROM MEMBERS WHERE uyelik_tarihi=date()-1")
IF counter2.EoF OR counter2.BoF Then counter2_yesterday = "0" ELSE counter2_yesterday = counter2("today")%>
<table border="0" cellspacing="0" cellpadding="0" align="center" width="100%" class="td_menu2">
<tr><td><img src="IMAGES/stats/mem.gif"> <b><%=menu4%></b></td></tr>
<tr><td><img src="IMAGES/stats/arrow.gif"> <%=last_member%>:<a href="members.asp?action=member_details&uid=<%=mem("uye_id")%>"><b><%=mem("kul_adi")%></b></a></td></tr>
<tr><td><img src="IMAGES/stats/arrow.gif"> <%=c_today%>:<b><%=today("m_count")%></b></td></tr>
<tr><td><img src="IMAGES/stats/arrow.gif"> <%=c_yesterday%>:<b><%=yesterday("m_count")%></b></td></tr>
<tr><td><img src="IMAGES/stats/arrow.gif"> <%=total_member%>:<b><%=memc("count")%></b></td></tr>
<tr><td align="center"><hr size="1" color="<%=td_border_color%>"></td></tr>
<tr><td><img src="IMAGES/stats/visitors.gif"> <b><a href="Members.asp?action=OnlineUsers"><%=sitedekiler%></a></b></td></tr>
<tr><td><img src="IMAGES/stats/arrow.gif"> <%=bottom_members%>:<b><%=mem_count%></b></td></tr>
<tr><td><img src="IMAGES/stats/arrow.gif"> <%=bottom_ziyaretciler%>:<b><%=guests("sayi")%></b></td></tr>
<tr><td><img src="IMAGES/stats/arrow.gif"> <%=bottom_total%>:<b><%=mem_count+guests("sayi")%></b></td></tr>
<tr><td align="center"><hr size="1" color="<%=td_border_color%>"></td></tr>
<tr><td><img src="IMAGES/stats/online.gif"> <b><%=online_mems%></b></td></tr>
<tr><td>
<%num = "0"
Do Until om.Eof
IF om("seviye") = "1" Then
om_title = ""&level1&""
ElseIf om("seviye") = "2" Then
om_title = ""&level2&""
ElseIf om("seviye") = "3" Then
om_title = ""&level3&""
ElseIf om("seviye") = "4" Then
om_title = ""&level4&""
ElseIf om("seviye") = "5" Then
om_title = ""&level5&""
ElseIf om("seviye") = "0" Then
om_title = ""&normal&""
End If
IF om("seviye") = "1" Then
om_t_c = online_admin:om_t_style = "bold"
ElseIF om("seviye") = "2" OR om("seviye") = "3" OR om("seviye") = "4" OR om("seviye") = "5" OR om("seviye") = "6"  OR om("seviye") = "8"  Then
om_t_c = online_editor:om_t_style = "bold"
ELSE
om_t_c = text:om_t_style = "regular"
End IF
num = num+1
Response.Write ""&num&": <a href=""members.asp?action=member_details&uid="&om("uye_id")&""" title="""&om_title&"""><font color="&om_t_c&" style=""font-weight:"&om_t_style&""">"&om("kul_adi")&"</font></a><br />"
om.MoveNext
Loop
IF num = "0" Then Response.Write no_online%>
</td></tr>
<tr><td align="center"><hr size="1" color="<%=td_border_color%>"></td></tr>
<tr><td><img src="IMAGES/stats/graph.gif"> <b><%=site_hits%></b></td></tr>
<tr><td><img src="IMAGES/stats/arrow.gif"> <%=c_today%>&nbsp;:&nbsp;<b><%=counter("today")%></b></td></tr>
<tr><td><img src="IMAGES/stats/arrow.gif"> <%=c_yesterday%>&nbsp;:&nbsp;<b><%=counter2_yesterday%></b></td></tr>
<tr><td><img src="IMAGES/stats/arrow.gif"> Sizin :<b>
<script language="Javascript"><!--
function getCookieVal (offset) {
var endstr=document.cookie.indexOf (";", offset);
if(endstr == -1)endstr=document.cookie.length;
return unescape(document.cookie.substring(offset, endstr));
}function GetCookie(name){var arg=name + "=";var alen=arg.length;
var clen=document.cookie.length;var i=0;
while(i < clen){var j=i + alen;if(document.cookie.substring(i, j)==arg)return getCookieVal(j);i=document.cookie.indexOf(" ", i) + 1;if (i == 0)break;}
return null;
}function SetCookie (name, value) {
var argv=SetCookie.arguments;var argc=SetCookie.arguments.length;var expires=(argc > 2) ? argv[2] : null;
var path=(argc > 3) ? argv[3] : null;var domain=(argc > 4) ? argv[4] : null;var secure=(argc > 5) ? argv[5] : false;
document.cookie=name + "=" + escape (value) + ((expires == null) ? "" : ("; expires=" + expires.toGMTString())) +
((path == null) ? "" : ("; path=" + path)) + ((domain == null) ? "" : ("; domain=" + domain)) +	((secure == true) ? "; secure" : "");
}function DeleteCookie(name) {var exp=new Date();FixCookieDate(exp); // Correct for Mac bug
exp.setTime (exp.getTime() - 1);var cval=GetCookie (name);if(cval != null)document.cookie=name + "=" + cval + "; expires=" + exp.toGMTString();}
var expdate=new Date();var num_visits;expdate.setTime(expdate.getTime() + (5*24*60*60*1000));
if (!(num_visits=GetCookie("num_visits")))  num_visits=0;num_visits++;
SetCookie("num_visits",num_visits,expdate);//--></script>
<script language="JavaScript"><!--
document.write(""+num_visits+"");//-->
</script>
</b></td></tr>
<tr><td><img src="IMAGES/stats/arrow.gif"> <%=c_total%>&nbsp;:&nbsp;<b><%=counter_t("count")%></b></td></tr>
</table>
<%mem.Close:Set mem = Nothing
memc.Close:Set memc = Nothing
counter.Close:Set counter = Nothing
counter2.Close:Set counter2 = Nothing
counter_t.Close:Set counter_t = Nothing
today.Close:Set today = Nothing
yesterday.Close:Set yesterday = Nothing%>
</td>