<%@ Language=VBScript %>
<%Response.Buffer=TRUE%>
<% Response.CacheControl="no-cache" %>
<% Response.AddHeader "Pragma", "no-cache" %>
<% Response.Expires=-1 %>
<%
fromid="from" & request.cookies("myid")
%>
<HTML>
<HEAD>
<meta http-equiv=Content-Type content="text/html; charset=iso-8859-1">
<TITLE>mc_im_server</TITLE>
<script LANGUAGE="JavaScript">
function popupwindow(){
window.open("display_message.asp","","height=200,width=250,menubar=0,resizabl e=0,scrollbars=0,status=0,titlebar=0,toolbar=0,left=0,top=0")
}
</script>
<%
if request("Action")="Display" then
application.lock
session("message")=application("" & request.cookies("boxid") & "")
response.cookies("fromid")=application("" & fromid & "")
session("flag")="yes"
application("" & request.cookies("boxid") & "")=""
application("" & request.cookies("flagid") & "")="no"
application.unlock
end if
%>
</HEAD>
<BODY bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">

<%if request("Action")="Display" then%>
<%
if session("flag")="yes" then
%>
<%
if session("message")=request.cookies("messagelast") then
else
%>

<script>
popupwindow()
</script>

<%end if%>
<%end if%>
<%end if%>

<div align="center">
  <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="200" height="40">
    <param name="movie" value="im_server.swf">
    <param name="quality" value="high">
    <embed src="im_server.swf" quality="high" PLUGINSPAGE="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="200" height="40"></embed></object>
</div>
</BODY>
</HTML>
<%
session("flag")=""
%>
