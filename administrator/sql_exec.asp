<%
Response.Buffer = True
function buda(yen,max,budandi)
	str = yen
	max_karakter = max
	istif = "..."
	If len(str)>max_karakter Then 
		yenistr = mid(str,1,max_karakter - len(istif))
		yenistr = yenistr + istif
		budandi = true
	else
		yenistr = str
	End If
	buda = yenistr
' © function by mucit 
end function ' buda
%>
<!--#include file="includes.asp" -->
<% IF Session("Level") = "1" OR Session("Level") = "7" Then %>
<table border="0" cellspacing="2" cellpadding="2" width="90%" align="center" class="td_menu2">
<tr>
<td><b>Dikkat ! : </b>Bu sistemi bilmeyen kiþilerin kullanmasý tehlikeli sonuçlar doðurabilir.</td>
</tr>
<% IF Request.QueryString("op") = "Exec" Then %>
<tr bgcolor="#00EE99">
<td>'<b><%=Request.Form("sql_code")%></b>' çalýþtýrýldý...</td>
</tr>
<% End IF %>
<tr>
<td>
<%
IF Request.QueryString("op") = "Exec" Then
IF InStr(1,Request.Form("sql_code"),"select",1) Then

IF Request.Form("db_name") = "FORUM_DB" Then
Set ks = Server.CreateObject("ADODB.RecordSet")
ks.open ""&Request.Form("sql_code")&"",Connect,3,3
ELSE
Set ks = Server.CreateObject("ADODB.RecordSet")
ks.open ""&Request.Form("sql_code")&"",Connection,3,3
End IF

	sayif = ks.fields.count
		Response.Write "<table border=0 cellpadding=2 cellspacing=2 class=td_menu2 width=300><tr>"
		For i = 0 to sayif-1
			Response.Write "<th bgcolor=""#FFCC33"">" & ks.fields(i).name & "</th>"
		next
		s = 1
		Do While not ks.eof
			If s mod 2 = 0 then 	
			bg = "#E4E4E4"
				else 
				bg = "#CDCDCD"
				end if
			Response.Write "<tr>"
			For i = 0 to sayif-1 
				ksf = buda(ks.fields(i), 30, budandi)
				Response.Write "<td bgcolor="""& bg &""">" & ksf& "</td>"
			next
			Response.Write "</tr>"
		ks.movenext
		s = s +1
		Loop
Response.Write "</table>"
ELSE
Connection.Execute(""&Request.Form("sql_code")&"")
End IF
%>
</td>
</tr>
<% End IF %>
<form method="post" action="?op=Exec">
<tr>
<td>
<textarea name="sql_code" rows="10" class="form1" style="width:100%" cols="20"><%=Request.Form("sql_code")%></textarea>
</td>
</tr>
<tr>
<td>Database Seçin : 
<select name="db_name" size="1" class="form1">
<option value="MAIN_DB">Genel Database</option>
<option value="FORUM_DB">Forum Database</option>
</select>
</td>
</tr>
<tr>
<td><input type="submit" value="Çalýþtýr" class="submit"></td>
</tr>
</form>
</table>
<% End IF %>