<%@ LANGUAGE = VBScript.Encode %><!--#include file="INC/includes.asp" --><% Response.Buffer = True %>
<html><head>
<META http-equiv="Content-Type" content="text/html; charset=windows-1254">
<META http-equiv="imagetoolbar" content="no">
<META http-equiv="Reply-to" content="">
<META http-equiv="Copyright" content="Copyright © 2007">
<META http-equiv="Resource-type" content="document">
<META http-equiv="Distribution" content="global">
<META http-equiv="Content-Language" content="tr">
<META http-equiv="Window-target" content="_top">
<META http-equiv="Pragma" content="no-cache">
<META NAME="Copyright" content="Copyright © 2007">
<META NAME="Content-Language" content="tr">
<META NAME="verify-v1" content="z/Fbs/SmTMuBFDapOFpPx7v7s9uqBLdKofgzVHa9SZI=" />
<META NAME="description" content="Türkiyenin En Gelişmiş Web Portalı">
<META NAME="keywords" content="bilgisayar, programlama, asp, php, msn, hack, forum, ödül, yarışma, edebiyat, felsefe, bilim, sağlık, döküman, arşiv, sohbet, chat, bayan, arkadaş">
<META NAME="author" content="İsmail & Hüseyin SEKMENOĞLU">
<META NAME="ROBOTS" content="INCLUDE, FOLLOW">
<META NAME="revisit-after" CONTENT="1">
<LINK REV="made" href="mailto:sekmenismail@sonkaratahta.com">
<LINK REL="stylesheet" type="text/css" href="THEMES/<%=Session("SiteTheme")%>/style.css">
<style type="text/css">.topstrip {background-image: url("Themes/<%=Session("SiteTheme")%>/topstrip.gif");background-repeat: repeat-x;}</style>
<SCRIPT LANGUAGE='JAVASCRIPT' TYPE='TEXT/JAVASCRIPT'><!--
var win=null;
function NewWindow(mypage,myname,w,h,pos,infocus){
if(pos=="random"){myleft=(screen.width)?Math.floor(Math.random()*(screen.width-w)):100;mytop=(screen.height)?Math.floor(Math.random()*((screen.height-h)-75)):100;}
if(pos=="center"){myleft=(screen.width)?(screen.width-w)/2:100;mytop=(screen.height)?(screen.height-h)/2:100;}
else if((pos!='center' && pos!="random") || pos==null){myleft=0;mytop=20}
settings="width=" + w + ",height=" + h + ",top=" + mytop + ",left=" + myleft + ",scrollbars=no,location=no,directories=no,status=no,menubar=no,toolbar=no,resizable=no";win=window.open(mypage,myname,settings);
win.focus();
}// --></script>
<script language="Javascript1.2"><!-- // load htmlarea
_editor_url = "htmlarea/";var win_ie_ver = parseFloat(navigator.appVersion.split("MSIE")[1]);
if(navigator.userAgent.indexOf('Mac')>= 0){win_ie_ver = 0;}
if(navigator.userAgent.indexOf('Windows CE')>= 0){win_ie_ver = 0;}
if(navigator.userAgent.indexOf('Opera')>= 0){win_ie_ver = 0;}
if(win_ie_ver >= 5.5){document.write('<scr'+'ipt src="'+_editor_url+'editor.js"');document.write(' language="Javascript1.2"></scr'+'ipt>');  }
else{document.write('<scr'+'ipt>function editor_generate(){return false;}</scr'+'ipt>'); }// --></script>
</head><body bgcolor="<%=background%>" text="<%=text%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"<%=body_have_new_msg%>>
<%Sub UST%>
<table width="100%" align="center" cellpadding="0" cellspacing="0"><tr>
<td width="23%" valign="top" class="topstrip"><table width="164" cellpadding="0" cellspacing="0"><tr><td width="160" valign="top"><img src="Themes/<%=Session("SiteTheme")%>/logo.gif" border="0" /></td></tr></table></td>
<td width="77%" align="right" valign="top" class="topstrip"><table width="660" border="0" cellspacing="0" cellpadding="0">
</table></td>
</tr></table>
<table width="100%" height="25" border="0" cellspacing="0" cellpadding="0" align="center" class="td_menu" style="BACKGROUND-REPEAT: repeat-x; BORDER-RIGHT: <%=top_menu_border%> 1px solid;BACKGROUND-REPEAT: repeat-x; BORDER-BOTTOM: <%=top_menu_border%> 1px solid">
<tr style="font-weight:bold">
<td width="10%" align="center" background="Themes/<%=Session("SiteTheme")%>/menu_bg.gif" onMouseOver="this.style.cursor='hand'" onClick="window.location.href='<%=sett("site_adres")%>/<%=top_menu_link1%>';"><font class="block_title"><%=top_menu1%></font> </td><td valign="top"><img src="Images/Site/top_menu_blank.gif" align="absmiddle"></td>
<td width="10%" align="center" background="Themes/<%=Session("SiteTheme")%>/menu_bg.gif" onMouseOver="this.style.cursor='hand'" onClick="window.location.href='<%=sett("site_adres")%>/<%=top_menu_link2%>';"><font class="block_title"><%=top_menu2%></font> </td><td valign="top"><img src="Images/Site/top_menu_blank.gif" align="absmiddle"></td>
<td width="10%" align="center" background="Themes/<%=Session("SiteTheme")%>/menu_bg.gif" onMouseOver="this.style.cursor='hand'" onClick="window.location.href='<%=sett("site_adres")%>/<%=top_menu_link3%>';"><font class="block_title"><%=top_menu3%></font> </td><td valign="top"><img src="Images/Site/top_menu_blank.gif" align="absmiddle"></td>
<td width="10%" align="center" background="Themes/<%=Session("SiteTheme")%>/menu_bg.gif" onMouseOver="this.style.cursor='hand'" onClick="window.location.href='<%=sett("site_adres")%>/<%=top_menu_link4%>';"><font class="block_title"><%=top_menu4%></font> </td><td valign="top"><img src="Images/Site/top_menu_blank.gif" align="absmiddle"></td>
<td width="10%" align="center" background="Themes/<%=Session("SiteTheme")%>/menu_bg.gif" onMouseOver="this.style.cursor='hand'" onClick="window.location.href='<%=sett("site_adres")%>/<%=top_menu_link5%>';"><font class="block_title"><%=top_menu5%></font> </td><td valign="top"><img src="Images/Site/top_menu_blank.gif" align="absmiddle"></td>
<td width="10%" align="center" background="Themes/<%=Session("SiteTheme")%>/menu_bg.gif" onMouseOver="this.style.cursor='hand'" onClick="window.location.href='<%=sett("site_adres")%>/<%=top_menu_link6%>';"><font class="block_title"><%=top_menu6%></font> </td><td valign="top"><img src="Images/Site/top_menu_blank.gif" align="absmiddle"></td>
<td width="10%" align="center" background="Themes/<%=Session("SiteTheme")%>/menu_bg.gif" onMouseOver="this.style.cursor='hand'" onClick="window.location.href='<%=sett("site_adres")%>/<%=top_menu_link7%>';"><font class="block_title"><%=top_menu7%></font> </td><td valign="top"><img src="Images/Site/top_menu_blank.gif" align="absmiddle"></td>
<form action="<%=sett("site_adres")%>/search.asp" method="get"><td width="30%" align="right" background="Themes/<%=Session("SiteTheme")%>/menu_bg.gif"><font class="td_menu"><b><%=arama%></b> : <input type="text" name="search" class="form1" value="<%=Request("search")%>" size="20"> <input type="image" src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/ara.gif" border="0" align="absmiddle">	&nbsp; </font></td></form></tr>
<%Set notices = Connection.Execute("Select * FROM NOTICES ORDER BY n_date DESC"):IF NOT notices.EoF Then
notice_msg = notices("n_text"):FOR I = 1 TO UBound(saryHTMLtags):notice_msg = Replace(notice_msg,"<"&saryHTMLtags(""&I&"")&">","&lt;"&saryHTMLtags(""&I&"")&"&gt;",1,-1,1):notice_msg = Replace(notice_msg,"<p>","",1,-1,1):notice_msg = Replace(notice_msg,"</p>","",1,-1,1):Next%>
<tr>
<td colspan="15" class="td_menu2" bgcolor="<%=background%>"><table border="0" colspan="0" cellpadding="1" width="100%" class="td_menu2">
<tr><td><%Response.Write "<img src=""IMAGES/stats/notice.gif"" align=""absmiddle""> <b>"&notice&"<i>("&notices("n_date")&")</i></b> " & notice_msg%></td></tr>
</table></td>
</tr>
<%End IF:notices.Close:Set notices = Nothing%>
</table>
<font class="block_title">&nbsp;Merhaba, <%If Session("Member")="" or "0" Then Response.Write "Ziyaretçi" Else Response.Write Session("Name")%></font><br>
<%End Sub 'UST
Sub ORTA
%><table width="100%" border="0" cellspacing="0" cellpadding="0"><tr><%
IF LB.recordcount>0 Then%><td width="20%" align="left" valign="top"><%Do Until LB.Eof%><table border="0" cellpadding="0" cellspacing="0" width="150"><tr><td>
<table border="0" cellpadding="0" cellspacing="0"><tr><td align="left" valign="top" width="32" height="33" rowspan="2"><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/sidebox-title-left.gif" width="32" height="33"></td>
	<td align="left" valign="middle" width="100%" class="td_menu" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/sidebox-title-bg.gif" height="21"><b>&nbsp;</b><b><font class="block_title"><%=LB("b_name")%></font></b></td>
	<td align="left" valign="top" width="23" height="33" rowspan="2"><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/sidebox-title-right.gif" width="23" height="33"></td>
	</tr><tr><td align="left" valign="middle" bgcolor="<%=content_bg%>" class="td_menu" height="10"></td></tr>
</table><table border="0" cellpadding="0" cellspacing="0"><tr>
	<td align="left" valign="top" width="12" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/sidebox-content-left.gif"><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/sidebox-content-left.gif" width="12"></td>
	<td align="left" valign="top" bgcolor="<%=content_bg%>" width="100%" class="td_menu"><%IF LB("b_include") <> "null" Then Server.Execute ("BLOCKS/"&LB("b_include")&"/content.asp")
	Response.Write LB("b_content")%></td>
	<td align="right" valign="top" width="13" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/sidebox-content-right.gif"><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/sidebox-content-right.gif" width="13"></td>
</tr></table>
<table border="0" cellpadding="0" cellspacing="0"><tr><td width="32"><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/sidebox-bottom-left.gif" width="32"></td><td width="100%" class="td_menu" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/sidebox-bottom-bg.gif"></td><td width="23"><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/sidebox-bottom-right.gif" width="23"></td></tr></table>
</td></tr></table><%LB.MoveNext
Loop%></td><%else:response.write"<td width='5%'>&nbsp;</td>":End IF
%><td align="center" valign="top">
<%End Sub 'ORTA
Sub ORTA2 
%></td><%
IF RB.recordcount>0 Then%><td width="20%" align="right" valign="top"><%Do Until RB.Eof%><table border="0" cellpadding="0" cellspacing="0" width="150"><tr><td>
<table border="0" cellpadding="0" cellspacing="0"><tr>
	<td align="left" valign="top" width="32" height="33" rowspan="2"><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/sidebox-title-left.gif" width="32" height="33"></td>
	<td align="left" valign="middle" width="100%" class="td_menu" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/sidebox-title-bg.gif" height="21"><b>&nbsp;</b><b><font class="block_title"><%=RB("b_name")%></font></b></td>
	<td align="left" valign="top" width="23" height="33" rowspan="2"><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/sidebox-title-right.gif" width="23" height="33"></td>
	</tr><tr><td align="left" valign="middle" bgcolor="<%=content_bg%>" class="td_menu" height="10"></td></tr>
</table><table border="0" cellpadding="0" cellspacing="0"><tr>
	<td align="left" valign="top" width="12" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/sidebox-content-left.gif"><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/sidebox-content-left.gif" width="12"></td>
	<td align="left" valign="top" bgcolor="<%=content_bg%>" width="100%" class="td_menu"><%IF RB("b_include") <> "null" Then Server.Execute ("BLOCKS/"&RB("b_include")&"/content.asp")
	Response.Write RB("b_content")%></td>
	<td align="right" valign="top" width="13" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/sidebox-content-right.gif"><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/sidebox-content-right.gif" width="13"></td>
</tr></table>
<table border="0" cellpadding="0" cellspacing="0"><tr><td width="32"><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/sidebox-bottom-left.gif" width="32"></td><td width="100%" class="td_menu" background="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/sidebox-bottom-bg.gif"></td><td width="23"><img src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/sidebox-bottom-right.gif" width="23"></td></tr></table>
</td></tr></table><%RB.MoveNext
Loop
%></td><%else:response.write"<td width='5%'>&nbsp;</td>":End IF
%></tr></table><%
End Sub 'ORTA2
Sub ALT%>
<br><%Set urMsgs = Connection.Execute("Select Count(*) as count FROM MESSAGES WHERE TYPE = 0 AND SEE = True AND READ = False AND TO = '"&Session("Member")&"'")
Set TM = Connection.Execute("Select Count(*) as mcount FROM MESSAGES WHERE TYPE = 0 AND SEE = True AND TO = '"&Session("Member")&"'")
Set urMsgs = Connection.Execute("Select Count(*) as count FROM MESSAGES WHERE TYPE = 0 AND SEE = True AND READ = False AND TO = '"&Session("Member")&"'")
Set TM = Connection.Execute("Select Count(*) as mcount FROM MESSAGES WHERE TYPE = 0 AND SEE = True AND TO = '"&Session("Member")&"'")
percent = Int((TM("mcount") / sett("msg_limit")) * 100)
If percent >= 100 Then percent = "100" Else percent = percent
If urMsgs("count") > sett("msg_limit") Then urMsgs_COUNT = sett("msg_limit") Else urMsgs_COUNT = urMsgs("count")
Set msgs = Connection.Execute("Select Count(*) as count FROM MESSAGES WHERE TYPE = 0 AND SEE = True AND TO = '"&Session("Member")&"'")
Set s_msgs = Connection.Execute("Select Count(*) as count FROM MESSAGES WHERE TYPE = 0 AND FROM = '"&Session("Member")&"' AND MSAVE = True")
Set admin = Connection.Execute("Select Count(*) as count FROM MESSAGES WHERE TYPE = 1 AND SEE = True")
If msgs("count") >= sett("msg_limit") Then t_msgs = sett("msg_limit") Else t_msgs = msgs("count")
If admin("count") >= sett("msg_limit") Then admin = sett("msg_limit") Else admin = admin("count")
If s_msgs("count") >= sett("msg_limit") Then s_msgs = sett("msg_limit") Else s_msgs = s_msgs("count")
%><!--#include file="inc/izenic.asp" --><%
IF Session("Enter")="Yes" and urMsgs_COUNT>0 Then %>
<SCRIPT type=text/javascript>//<![CDATA[
var popupWinoldonloadHndlr=window.onload, popupWinpopupHgt, popupWinactualHgt, popupWintmrId=-1, popupWinresetTimer;
var popupWintitHgt, popupWincntDelta, popupWintmrHide=-1, popupWinhideAfter=-1, popupWinhideAlpha, popupWinhasFilters=true;
var popupWinnWin, popupWinshowBy=null, popupWindxTimer=-1, popupWinpopupBottom;
var popupWinnText,popupWinnMsg,popupWinnTitle,popupWinbChangeTexts=false;
var popupWinmousemoveBack,popupWinmouseupBack;
var popupWinofsX,popupWinofsY;
function popupWinespopup_dxTimer(){clearInterval(popupWindxTimer); popupWindxTimer=-1;}
function popupWinespopup_Close(){if (popupWintmrId==-1){el=document.getElementById('popupWin');el.style.filter='';el.style.display='none';if (popupWintmrHide!=-1) clearInterval(popupWintmrHide); popupWintmrHide=-1;}}
function popupWinespopup_DragDropStop(){document.body.onmousemove=popupWinmousemoveBack;document.body.onmouseup=popupWinmouseupBack;}
function popupWinespopup_DragDrop(e){popupWinmousemoveBack=document.body.onmousemove;popupWinmouseupBack=document.body.onmouseup;ox=(e.offsetX==null)?e.layerX:e.offsetX;oy=(e.offsetY==null)?e.layerY:e.offsetY;popupWinofsX=ox; popupWinofsY=oy;document.body.onmousemove=popupWinespopup_DragDropMove;document.body.onmouseup=popupWinespopup_DragDropStop;if(popupWintmrHide!=-1) clearInterval(popupWintmrHide);}
window.onload=popupWinespopup_winLoad;
function popupWinespopup_ShowPopup(show){
if(popupWindxTimer!=-1){el.filters.blendTrans.stop();}
if((popupWintmrHide!=-1) && ((show!=null) && (show==popupWinshowBy))){clearInterval(popupWintmrHide);popupWintmrHide=setInterval(popupWinespopup_tmrHideTimer,popupWinhideAfter);return;}
if (popupWintmrId!=-1) return;
popupWinshowBy=show;
elCnt=document.getElementById('popupWin_content')
elTit=document.getElementById('popupWin_header');
el=document.getElementById('popupWin');
el.style.left='';
el.style.top='';
el.style.filter='';
if (popupWintmrHide!=-1)clearInterval(popupWintmrHide); popupWintmrHide=-1;
document.getElementById('popupWin_header').style.display='none';
document.getElementById('popupWin_content').style.display='none';
if (navigator.userAgent.indexOf('Opera')!=-1)el.style.bottom=(document.body.scrollHeight*1-document.body.scrollTop*1-document.body.offsetHeight*1+1*popupWinpopupBottom)+'px';
if (popupWinbChangeTexts){popupWinbChangeTexts=false;document.getElementById('popupWinaCnt').innerHTML=popupWinnMsg;document.getElementById('popupWintitleEl').innerHTML=popupWinnTitle;}
popupWinactualHgt=0; el.style.height=popupWinactualHgt+'px';
el.style.visibility='';
if (!popupWinresetTimer) el.style.display='';
popupWintmrId=setInterval(popupWinespopup_tmrTimer,(popupWinresetTimer?1000:20));
}
function popupWinespopup_winLoad(){
if (popupWinoldonloadHndlr!=null) popupWinoldonloadHndlr();
elCnt=document.getElementById('popupWin_content')
elTit=document.getElementById('popupWin_header');
el=document.getElementById('popupWin');
popupWinpopupBottom=el.style.bottom.substr(0,el.style.bottom.length-2);
popupWintitHgt=elTit.style.height.substr(0,elTit.style.height.length-2);
popupWinpopupHgt=el.style.height;
popupWinpopupHgt=popupWinpopupHgt.substr(0,popupWinpopupHgt.length-2); popupWinactualHgt=0;
popupWincntDelta=popupWinpopupHgt-(elCnt.style.height.substr(0,elCnt.style.height.length-2));
if (true){popupWinresetTimer=true;popupWinespopup_ShowPopup(null);}
}
function popupWinespopup_tmrTimer(){
el=document.getElementById('popupWin');
if (popupWinresetTimer){el.style.display='';clearInterval(popupWintmrId); popupWinresetTimer=false;popupWintmrId=setInterval(popupWinespopup_tmrTimer,20);}
popupWinactualHgt+=5;
if(popupWinactualHgt>=popupWinpopupHgt){popupWinactualHgt=popupWinpopupHgt; clearInterval(popupWintmrId); popupWintmrId=-1;document.getElementById('popupWin_content').style.display='';if (popupWinhideAfter!=-1)popupWintmrHide=setInterval(popupWinespopup_tmrHideTimer,popupWinhideAfter);}
if(popupWintitHgt<popupWinactualHgt-6)
document.getElementById('popupWin_header').style.display='';
if((popupWinactualHgt-popupWincntDelta)>0){elCnt=document.getElementById('popupWin_content');elCnt.style.display='';elCnt.style.height=(popupWinactualHgt-popupWincntDelta)+'px';}
el.style.height=popupWinactualHgt+'px';
}
function popupWinespopup_tmrHideTimer(){
clearInterval(popupWintmrHide); popupWintmrHide=-1;
el=document.getElementById('popupWin');
if (popupWinhasFilters){
backCnt=document.getElementById('popupWin_content').innerHTML;
backTit=document.getElementById('popupWin_header').innerHTML;
document.getElementById('popupWin_content').innerHTML='';
document.getElementById('popupWin_header').innerHTML='';
el.style.filter='blendTrans(duration=1)';
el.filters.blendTrans.apply();
el.style.visibility='hidden';
el.filters.blendTrans.play();
document.getElementById('popupWin_content').innerHTML=backCnt;
document.getElementById('popupWin_header').innerHTML=backTit;
popupWindxTimer=setInterval(popupWinespopup_dxTimer,1000);
}
else el.style.visibility='hidden';
}
function popupWinespopup_DragDropMove(e){
el=document.getElementById('popupWin');          
if (e==null&&event!=null){
el.style.left=(event.clientX*1+document.body.scrollLeft-popupWinofsX)+'px';
el.style.top=(event.clientY*1+document.body.scrollTop-popupWinofsY)+'px';
event.cancelBubble=true;
}
else{
el.style.left=(e.pageX*1-popupWinofsX)+'px';
el.style.top=(e.pageY*1-popupWinofsY)+'px';
e.cancelBubble=true;
}
if ((event.button&1)==0) popupWinespopup_DragDropStop();
}
//]]></SCRIPT>
<DIV onselectstart="return false;" onMouseDown="return popupWinespopup_DragDrop(event);" id=popupWin style="BORDER-RIGHT: 1px solid #455690; BORDER-TOP: 1px solid #b9c9ef; DISPLAY: none; Z-INDEX: 9999; RIGHT: 8; BACKGROUND: #e0e9f8; BORDER-LEFT: 1px solid #b9c9ef; WIDTH: 216; BOTTOM: 8; BORDER-BOTTOM: 1px solid #455690; POSITION: absolute; HEIGHT: 93">
<DIV id=popupWin_header style="DISPLAY: none; FILTER: progid:DXImageTransform.Microsoft.Gradient(GradientType=0,StartColorStr='#FFE0E9F8', EndColorStr='#FFFFFFFF'); LEFT: 3; WIDTH: 210; CURSOR: default; COLOR: #1f336b; POSITION: absolute; TOP: 2; HEIGHT: 14; TEXT-DECORATION: none; font-style:normal; font-variant:normal; font-weight:normal; font-size:11px; font-family:arial, sans-serif">
<p align="center"><SPAN id=popupWintitleEl0> <IMG style="LEFT: 1px; POSITION: absolute; TOP: 1px" src="images/popup_attention.gif"><b>Sayın <%=Session("Name")%></b></SPAN></p>
<SPAN onmousedown=event.cancelBubble=true; onMouseOver="style.color='#455690';" style="RIGHT: 3px; FONT: bold 12px arial,sans-serif; CURSOR: pointer; COLOR: #728eb8; POSITION: absolute; TOP: 2px" onclick=popupWinespopup_Close() onMouseOut="style.color='#728EB8';"><IMG alt="" src="images/popup_close.gif" border=0></SPAN> 
</DIV>
<DIV onmousedown=event.cancelBubble=true; id=popupWin_content style="padding:2px; border-right:1px solid #b9c9ef; border-top:1px solid #728eb8; DISPLAY: none; BACKGROUND: #e0e9f8; FILTER: progid:DXImageTransform.Microsoft.Gradient(GradientType=0,StartColorStr='#FFE0E9F8', EndColorStr='#FFFFFFFF'); LEFT: 6; OVERFLOW: hidden; BORDER-LEFT: 1px solid #728eb8; WIDTH: 200; BORDER-BOTTOM: 1px solid #b9c9ef; POSITION: absolute; TOP: 18; HEIGHT: 65; TEXT-ALIGN: center">
<A onMouseOver="style.textDecoration='underline';" onMouseOut="style.textDecoration='none';" href="msgbox.asp?action=main&id=<%=Session("Uye_ID")%>" style="COLOR: #1f336b; TEXT-DECORATION: none; font-style:normal; font-variant:normal; font-weight:700; line-height:15px; font-size:12px; font-family:arial, sans-serif"> <b> <%=urMsgs_COUNT%> </b> <%=msg_unreaded_msgs%><BR>Toplam <%=t_msgs%> Tane Mesajınız Var.</A>
<br /><br /><a onMouseOver="style.textDecoration='underline';" onMouseOut="style.textDecoration='none';" href="msgbox.asp?action=fromadmin&id=<%=Session("Uye_ID")%>" style="COLOR: #3f538b; TEXT-DECORATION: none; font-size:11px; font-family:arial, sans-serif"><%=msg_admin%> (<%=admin%>)</a>
</DIV>
</DIV>
<% END IF%>
<% End Sub 'ALT %>
<script>document.all.yukleme.style.visibility = 'hidden';</script>
</body>
</html>
