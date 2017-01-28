<!--#include file="INC/includes.asp" -->
        
 <td bgcolor="<%=background%>">



<html>
<body bgcolor="<%=background%>" text="<%=text%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"<%=body_have_new_msg%>>
<link rel="stylesheet" type="text/css" href="THEMES/<%=Session("SiteTheme")%>/style.css">
<head>
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<title>Dinle</title>
<title>Dinle</title><title>Dinle</title><title>Dinle</title><title>Dinle</title><title>Dinle</title><title>Dinle</title><title>Dinle</title><title>Dinle</title><title>Dinle</title><title>Dinle</title><title>Dinle</title><title>Dinle</title>
</head>
<style>
<!--
BODY{
cursor:  url(mavi.cur);
}
-->
</style>

<script language="JavaScript1.2">

//Disable select-text script (IE4+, NS6+)- By Andy Scott
//Exclusive permission granted to Dynamic Drive to feature script
//Visit http://www.dynamicdrive.com for this script

function disableselect(e){
return false
}

function reEnable(){
return true
}

//if IE4+
document.onselectstart=new Function ("return false")

//if NS6
if (window.sidebar){
document.onmousedown=disableselect
document.onclick=reEnable
}
</script>
<SCRIPT language=JavaScript> 
<!-- //Disable right click script 
var message=""; 
/////////////////////////////////// 
function clickIE() {if (document.all) {(message);return false;}} 
function clickNS(e) {if 
(document.layers||(document.getElementById&&!document.all)) { 
if (e.which==2||e.which==3) {(message);return false;}}} 
if (document.layers) 
{document.captureEvents(Event.MOUSEDOWN);document.onmousedown=clickNS;} 
else{document.onmouseup=clickNS;document.oncontextmenu=clickIE;} 
// --> 
</SCRIPT>
<body>

<p align="center">
<img border="0" src="modules/radyo/resimler/top-left.jpg" align="center"></p>
<p align="center">
<object id="NSPlay0" codeBase="http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701" type="application/x-oleobject" height="50" standby="Loading Microsoft Windows Media Player components..." width="291" align="texttop" classid="clsid:22D6F312-B0F6-11D0-94AB-0080C74C7E95">
  <param NAME="AudioStream" VALUE="-1">
  <param NAME="AutoSize" VALUE="0">
  <param NAME="AutoStart" VALUE="-1">
  <param NAME="AnimationAtStart" VALUE="0">
  <param NAME="AllowScan" VALUE="-1">
  <param NAME="AllowChangeDisplaySize" VALUE="0">
  <param NAME="AutoRewind" VALUE="0">
  <param NAME="Balance" VALUE="0">
  <param NAME="BaseURL" VALUE>
  <param NAME="BufferingTime" VALUE="5">
  <param NAME="CaptioningID" VALUE>
  <param NAME="ClickToPlay" VALUE="-1">
  <param NAME="CursorType" VALUE="0">
  <param NAME="CurrentPosition" VALUE="-1">
  <param NAME="CurrentMarker" VALUE="0">
  <param NAME="DefaultFrame" VALUE>
  <param NAME="DisplayBackColor" VALUE="0">
  <param NAME="DisplayForeColor" VALUE="16777215">
  <param NAME="DisplayMode" VALUE="0">
  <param NAME="DisplaySize" VALUE="4">
  <param NAME="Enabled" VALUE="-1">
  <param NAME="EnableContextMenu" VALUE="-1">
  <param NAME="EnablePositionControls" VALUE="-1">
  <param NAME="EnableFullScreenControls" VALUE="0">
  <param NAME="EnableTracker" VALUE="0">
  <param NAME="Filename" VALUE="http://72.21.62.146:8000" ref>
  <param NAME="InvokeURLs" VALUE="0">
  <param NAME="Language" VALUE="0">
  <param NAME="Mute" VALUE="0">
  <param NAME="PlayCount" VALUE="0">
  <param NAME="PreviewMode" VALUE="0">
  <param NAME="Rate" VALUE="1">
  <param NAME="SAMILang" VALUE>
  <param NAME="SAMIStyle" VALUE>
  <param NAME="SAMIFileName" VALUE>
  <param NAME="SelectionStart" VALUE="-1">
  <param NAME="SelectionEnd" VALUE="-1">
  <param NAME="SendOpenStateChangeEvents" VALUE="-1">
  <param NAME="SendWarningEvents" VALUE="-1">
  <param NAME="SendErrorEvents" VALUE="-1">
  <param NAME="SendKeyboardEvents" VALUE="-1">
  <param NAME="SendMouseClickEvents" VALUE="-1">
  <param NAME="SendMouseMoveEvents" VALUE="-1">
  <param NAME="SendPlayStateChangeEvents" VALUE="-1">
  <param NAME="ShowCaptioning" VALUE="-1">
  <param NAME="ShowControls" VALUE="-1">
  <param NAME="ShowAudioControls" VALUE="-1">
  <param NAME="ShowDisplay" VALUE="0">
  <param NAME="ShowGotoBar" VALUE="0">
  <param NAME="ShowPositionControls" VALUE="-1">
  <param NAME="ShowStatusBar" VALUE="-1">
  <param NAME="ShowTracker" VALUE="0">
  <param NAME="TransparentAtStart" VALUE="0">
  <param NAME="VideoBorderWidth" VALUE="0">
  <param NAME="VideoBorderColor" VALUE="0">
  <param NAME="VideoBorder3D" VALUE="0">
  <param NAME="Volume" VALUE="-1">
  <param NAME="WindowlessVideo" VALUE="-1">
  <embed type="application/x-mplayer2" 
pluginspage="http://www.microsoft.com/Windows/Downloads/Contents/Products/MediaPlayer/" 
filename="http://72.21.62.146:8000/listen.pls" 
src="http://72.21.62.146:8000/listen.pls" showcontrols="-1" showdisplay="-1" 
showstatusbar="-1" width="278" height="74" align="texttop">  <br />        
</embed></object>
<%
private Function GETHTTP(adres)   
Set StrHTTP=Server.CreateObject("Microsoft.XMLHTTP" )   
    StrHTTP.Open "GET" , adres, false   
    StrHTTP.sEnd   
    GETHTTP=StrHTTP.Responsetext   
Set StrHTTP=Nothing   
End Function   
URL="http://radyo.mini-nuke.info:8000/admin.cgi?pass=radyomininuke2005news&mode=viewxml"   
bilgicek=GETHTTP(URL)   
ds=InStr(1,bilgicek,"<LISTEN>" )  
de=InStr(1,bilgicek,"</LISTEN>" )  
list=Mid(bilgicek,ds,de-ds) 
ss=InStr(1,bilgicek,"<SONGTITLE>" )  
se=InStr(1,bilgicek,"</SONGTITLE>" )  
song=Mid(bilgicek,ss,se-ss)  
st=InStr(1,bilgicek,"<SERVERTITLE>" )  
ts=InStr(1,bilgicek,"</SERVERTITLE>" )  
radyo=Mid(bilgicek,st,ts-st)  
%>
<center><font face="Trebuchet MS" size="2">Online dinleyici : <%=list%> <br /> Su Anki Parca : <%=song%> <br /></center>

</body>

</html>