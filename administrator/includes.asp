<%
If session("Member") = "" then
Response.Write "<center><font class=""tabloalt_font"">Lütfen Giriþ Yapýnýz.</font></center>"
Else
%>
<!--#include file="../db/data.asp" -->
<!--#include file="INC/database.asp" -->
<!--#include file="INC/date.asp" -->
<!--#include file="../INC/functions.asp" -->
<!--#include file="../INC/filter.asp" -->
<!--#include file="../INC/encryption.asp" -->
<title>Yönetim PANELI</title>

<link rel="stylesheet" type="text/css" href="tema/style.css">
<script><!-- // load htmlarea

_editor_url = "../htmlarea/";                     // URL to htmlarea files
var win_ie_ver = parseFloat(navigator.appVersion.split("MSIE")[1]);
if (navigator.userAgent.indexOf('Mac')        >= 0) { win_ie_ver = 0; }
if (navigator.userAgent.indexOf('Windows CE') >= 0) { win_ie_ver = 0; }
if (navigator.userAgent.indexOf('Opera')      >= 0) { win_ie_ver = 0; }
if (win_ie_ver >= 5.5) {
 document.write('<scr' + 'ipt src="' +_editor_url+ 'editor.js"');
 document.write(' language="Javascript1.2"></scr' + 'ipt>');  
} else { document.write('<scr'+'ipt>function editor_generate() { return false; }</scr'+'ipt>'); }
// --></script> 
<%
Set upddIP = Connection.Execute("UPDATE MEMBERS SET son_tarih = now() WHERE uye_id = "&Session("Uye_ID")&"")
%>
<% END IF %>