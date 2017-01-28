<!--#include file="../../INC/b_include.asp" -->
<%
d_month=Request.QueryString("month")
d_year=Year(Now())
if d_month="" then d_month=month(date)
if d_year="" then d_year=year(date)
if d_month>12 then
d_month=1
d_year=d_year+1
end if
if d_month<1 then
d_month=12
d_year=d_year-1
end if
d_day=28
while isdate(d_day&","&d_month&","&d_year)=true
d_day=d_day+1
wend
d_day=d_day-1
Session.LCID = S_LCID
%>
<form method="GET" action="default.asp" id="form1" name="form1">
<table border="0" cellspacing="1" cellpadding="0" width="100%" class="td_menu" align="center">
<tr>
	<td style="font-size:12px"><a href="default.asp?month=<%=d_month-1%>&year=<%=d_year%>">«</a></td>
	<td align=center colspan=5><b><%=Monthname(d_month)%></b></td>
	<td style="font-size:12px"><a href="default.asp?month=<%=d_month+1%>&year=<%=d_year%>">»</a></td>
</tr>
<tr>
<td><font face="tahoma" size="1"><b><%=day1%></b></font></td>
<td><font face="tahoma" size="1"><b><%=day2%></b></font></td>
<td><font face="tahoma" size="1"><b><%=day3%></b></font></td>
<td><font face="tahoma" size="1"><b><%=day4%></b></font></td>
<td><font face="tahoma" size="1"><b><%=day5%></b></font></td>
<td><font face="tahoma" size="1"><b><%=day6%></b></font></td>
<td><font face="tahoma" size="1"><b><%=day7%></b></font></td>
</tr>
<tr>
<%
sayac=weekday("01,"&d_month&","&d_year)
if sayac=1 then 
	sayac=6
else
	sayac=sayac-2
end if	

for i=1 to sayac
Response.Write("<td> </td>")
next

for i=1 to d_day
if i = Day(Now()) then
a = i
else
a = i
end if
if sayac mod 7=0 then Response.Write("</tr> <tr>")%>
<td <% IF a = Day(Now()) Then%>style="BACKGROUND-REPEAT: repeat-x; BORDER-RIGHT: <%=td_border_color%> 1px solid;BORDER-LEFT: <%=td_border_color%> 1px solid;BORDER-TOP: <%=td_border_color%> 1px solid;BORDER-BOTTOM: <%=td_border_color%> 1px solid"<% End IF %>><font face="verdana" size="1"><a href="calender.asp?action=show&day=<%=a%>&month=<%=d_month%>"><% IF a = Day(Now()) Then%><b><%=a%></b><%Else%><%=a%><% End IF %></a></font></td>
<%
sayac=sayac+1
next
%>
</tr>
</table>
<div align="center">
<select name="month" class="form1">
	<option value="01"><%=m1%></option>
	<option value="02"><%=m2%></option>
	<option value="03"><%=m3%></option>
	<option value="04"><%=m4%></option>
	<option value="05"><%=m5%></option>
	<option value="06"><%=m6%></option>
	<option value="07"><%=m7%></option>
	<option value="08"><%=m8%></option>
	<option value="09"><%=m9%></option>
	<option value="10"><%=m10%></option>
	<option value="11"><%=m11%></option>
	<option value="12"><%=m12%></option>
</select>	
<input type="submit" name="button" value="<%=calender_go%>" class="submit">
</div>
</form>