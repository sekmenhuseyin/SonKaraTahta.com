<!--#include file="../../inc/filter.asp" --><!--#include file="../../inc/functions.asp" -->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD W3 HTML 3.2//EN">
<HTML id=dlgImage STYLE="width: 460px; height: 435px"><HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta http-equiv="MSThemeCompatible" content="Yes">
<LINK REL="stylesheet" type="text/css" href="../../THEMES/default/style.css">
<TITLE>Smiley Ekle</TITLE>
<style>html, body, button, div, input, select, fieldset{font-family: MS Shell Dlg; font-size: 8pt; position: absolute; };</style>
<script language='JavaScript' type='text/javascript'>
function insertSmiley(){var img=window.event.srcElement;if(img){var src=img.src.replace(/^[a-z]*:[/][/][^/]*/, "");window.returnValue='<IMG border=0 align=absmiddle src=' + src + '>';window.close();}}
function cancel(){window.returnValue=null;window.close();}
</script>
</HEAD><BODY style="background:threedface;color:windowtext;" scroll="no" style="cursor:url(../../images/site/mavi.cur);">
<table class="td_menu" border="0" align="left" cellpadding="0" cellspacing="5" style="border-collapse: collapse" width="380" bordercolor="#000000" id="AutoNumber1"><tr>
<%totalpics=99:pageno=totalpics/35+1
nextpage=QS_CLEAR(request.querystring("nextpage"),"false"):if nextpage="" then nextpage=1
number=nextpage+34:if number>totalpics then number=totalpics
For i=nextpage to number
if len(i)=1 then
j="00"&i
elseif len(i)=2 then
j="0"&i
end if%><td align='center' width='20%' height="50"><img onclick='insertSmiley()' border=0 src='../../images/smileys/smiley<%=j%>.gif' style="cursor:hand;" /></td>
<%if right(i,1)=0 or right(i,1)=5 then response.write("</tr><tr>")
Next:for i=i to nextpage+34
response.write("<td align='center' width='20%' height='50'></td>"):if right(i,1)=0 or right(i,1)=5 then response.write("</tr><tr>")
next%>
<tr><td colspan="5"><b><%for i=1 to pageno
if i=(nextpage-1)/35+1 then response.write(i&" ") else response.write("<a href='?nextpage="&(i-1)*35+1&"'>"&i&"</a> ")
next%></b></td></tr>
</tr></table>
<BUTTON ID="btnCancel" style="left:28.5em;top:38.6em;height:2.2em;" type="reset" tabIndex="45" onClick="cancel();">Kapat/Ç&#305;k</BUTTON><br />
</BODY></HTML>
