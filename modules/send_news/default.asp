<!--#include file="includes.asp" -->
<%
If Request.QueryString("oper")="EnterNew" Then
title=duz(Request.Form("ntitle"))
news=duz(Request.Form("ntext"))
cat=Request.Form("ncat")

IF title="" Then
Response.Write "<center><font class=td_menu>"&err1&"</font></center>"
ElseIF news="" Then
Response.Write "<center><font class=td_menu>"&err2&"</font></center>"
ElseIF cat="" Then
Response.Write "<center><font class=td_menu>"&err3&"</font></center>"
ELSE

Set ent=Server.CreateObject("ADODB.RecordSet")
entSQL="Select * FROM NEWS"
ent.open entSQL,Connection,1,3

ent.AddNew
ent("baslik")=title
ent("metin")=news
ent("ekleyen")=Session("Member")
ent("tarih")=Now()
ent("onay")=False
ent("puan")=0
ent("katilimci")=0
ent("okunma")=0
ent("cat")=cat
ent.Update
Response.Write "Göndermiþ olduðunuz haber yöneticiler tarafýndan incelenip sitedeki yerini alacaktýr.Teþekkür ederiz..."
ent.Close
Set ent=Nothing
END IF
ELSE
Set cats=Connection.Execute("Select * FROM NEWS_CATS ORDER BY cat_name ASC")
%>
<form action="?name=<%=Request.QueryString("name")%>&oper=EnterNew" method="post">
<table border="0" cellspacing="1" cellpadding="1" width="100%" align="center" style="font-family:tahoma; font-size:11px; font-weight:bold">
<tr>
<td width="40%" align="right"><%=f_title%> : </td>
<td width="60%"><input type="text" name="ntitle" size="50" class="form1"></td>
</tr>
<tr>
<td width="40%" align="right" valign="top"><%=f_text%> : </td>
<td width="60%"><textarea name="ntext" rows="8" cols="50" class="form1"></textarea></td>
</tr>
<tr>
<td width="40%" align="right" valign="top"><%=f_cat%> : </td>
<td width="60%">
<select name="ncat" size="1" class="form1">
<%
Do Until cats.Eof
Response.Write "<option value="""&cats("cat_id")&""">"&cats("cat_name")&"</option>"
cats.MoveNext
Loop
%>
</select>
</td>
</tr>
<tr>
<td width="40%">&nbsp;</td>
<td width="60%" align="right"><input type="submit" value="<%=f_submit%>" class="submit"></td>
</tr>
</table>
</form>
<%
cats.Close
Set cats=Nothing
End IF
%>