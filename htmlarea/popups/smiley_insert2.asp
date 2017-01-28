<!--#include file="../../inc/database.asp" --><!--#include file="../../inc/language_file.asp" -->
<!--#include file="../../inc/filter.asp" --><!--#include file="../../inc/functions.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD W3 HTML 3.2//EN">
<HTML id=dlgImage STYLE="width: 460px; height: 435px"><HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta http-equiv="MSThemeCompatible" content="Yes">
<LINK REL="stylesheet" type="text/css" href="../../THEMES/default/style.css">
<TITLE>Smiley Ekle</TITLE>
<style>html, body, button, div, input, select, fieldset { font-family: MS Shell Dlg; font-size: 8pt; position: absolute; };</style>
<script language='JavaScript' type='text/javascript'>
function insertSmiley(){var img=window.event.srcElement;if(img){var src=img.src.replace(/^[a-z]*:[/][/][^/]*/, "");window.returnValue='<IMG border=0 align=absmiddle src=' + src + '>';window.close();}}
function cancel(){window.returnValue=null;window.close();}
</script>
</HEAD><BODY style="background: threedface; color: windowtext;" scroll=no>
<FIELDSET id=fldSmiley style="left: 10px; top: 10px; width: 400px; height: 440px;">
<LEGEND id=lgdSmiley>Smiley Listesi</LEGEND>
</FIELDSET><BUTTON ID=btnCancel style="left: 30em; top: 42em; height: 2.2em; " type=reset tabIndex=45 onClick="cancel();">Kapat/Ç&#305;k</BUTTON><br />
<table class="td_menu" border="0" align="left" cellpadding="0" cellspacing="5" style="border-collapse: collapse" width="380" bordercolor="#000000"><tr>
<%Set rs=Server.CreateObject("ADODB.RecordSet"):sorgu="Select * FROM SMILEYS ORDER BY s_img DESC":rs.open sorgu,Connection,3,3
Page=Request.QueryString("Page"):Page=QS_CLEAR(Page,"false"):if Page="" then Page=1
totalpics=rs.recordcount:rs.pagesize=35:rs.absolutepage=Page:sayfa=rs.pagecount
For i=1 TO rs.pagesize
IF rs.EoF Then Exit For%>
<td align='center' width='20%' height="50"><img onclick='insertSmiley()' border=0 src='../../images/smileys/<%=rs("s_img")%>' style="cursor:hand;" /></td>
<%if right(i,1)=0 or right(i,1)=5 then response.write("</tr><tr>")
rs.movenext
Next%>
</tr><tr><td colspan="5"><b>
<%if sayfa>1 then response.Write(top_sayfa2) else response.Write(top_sayfa1)%> : 
<%If cint(Page)>1 Then Response.Write "&nbsp;<a href=""?Page="&(cint(Page)-1)&""" style=""font-size:10px"">« "&previous_page&"</a> - "
For y=1 to sayfa
IF cint(Page)=cint(y) then Response.Write " <font class=""td_menu"" style=""font-size:10px"">["&y&"]</font> " ELSE Response.Write " <a href=""?Page="&y&""" style=""font-size:10px"">["&y&"]</a>"
Next
If NOT rs.Eof Then Response.Write " - <a href=""?Page="&(cint(Page)+1)&""" style=""font-size:10px"">"&next_page&" »</a>"%>
</b></td></tr>
</tr></table>
</BODY></HTML>
