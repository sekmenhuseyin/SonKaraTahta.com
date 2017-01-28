<!-- Üst Bar -->
<title><%=sett("site_isim")%> - <%=sett("site_slogan")%></title>
<link rel="stylesheet" href="images/site/ssmitems.css" type="text/css" />
<!-- Mouse Simgesi -->
<style><!--
BODY{cursor:url(images/site/mavi.cur);}--></style>
<!-- Alt Bar -->
<SCRIPT language=JavaScript>  <!--
var current=0
var x=0
var y=0
var speed=100
var speed2=2000
function initArray(n) {this.length=n;for (var i =1; i <= n; i++) {this[i]=' '}}
typ=new initArray(4)
typ[0]="<%=sett("site_isim")%>"
typ[1]=" <%=sett("site_slogan")%>"
typ[2]=" <%=sett("gorunusurl")%>"
typ[3]=" <%=sett("site_mail")%>"
function typnslide() {
	var m=typ[current]
	window.status=m.substring(0, x++)
	if (x == m.length + 1) {x=0;setTimeout("typnslide2()", speed2);}
	else {setTimeout("typnslide()", speed)}
}
function typnslide2() {
	var m=typ[current]
	window.status=m.substring(m.length, y++)
	if (y == m.length) {
		y=0
		current++
		if (current > typ.length - 1) {current=0}
		setTimeout("typnslide()", speed)
	}
	else{setTimeout("typnslide2()", speed)}
}
typnslide();
var bookmarkurl="<%=sett("gorunusurl")%>"
var bookmarktitle="<%=sett("site_isim")%>"
function addbookmark(){if(document.all)window.external.AddFavorite(bookmarkurl,bookmarktitle)}
//--></SCRIPT>
<!----- Açýlýr Menu ------>
<% Set rightc=Connection.Execute("Select * FROM SETTINGS")
IF rightc("hmenu")=TRUE then%>
<script src="images/site/ssm.js" type="text/javascript"></script>
<script type="text/javascript">
YOffset=145;XAlign=2;XOffset=0;staticYOffset=20;waitTime=100;slideX=1;slideXSpeed=0;slideY=1;slideYSpeed=0;slideOnYOverflow=1;
autoHideXOverflow=1;targetFrame="";targetDomain="";operaFix=0;menuOpacity=90;menuPosition=1;menuBGColor="Black";menuWidth=168;
hdrBGColor="white";hdrPadding=1;hdrAlign="full";hdrVAlign="center";linkBGColor="Black";linkOverBGColor="Red";linkAlign="left";
linkVAlign="center";linkPadding=0;barWidth=10;barBGColor="White";barAlign="left";barVAlign="center";barType=0;barText="images/site/sidebar.gif";
//The Menu's Items
addHdr("<img src='images/site/blank1.gif' align='left' border='0'> MENU");
addItem("<img src='images/site/blank.gif' align='left' border='0'> Anasayfa", "Default.asp?MN-Fixed=anasayfa&", "");
addItem("<img src='images/site/blank.gif' align='left' border='0'> Haberler", "News.asp", "");
addItem("<img src='images/site/blank.gif' align='left' border='0'> Dosyalar", "Programs.asp?action=cats", "");
addItem("<img src='images/site/blank.gif' align='left' border='0'> Yazýlar", "Articles.asp?action=cats", "");
addItem("<img src='images/site/blank.gif' align='left' border='0'> Üyeler", "Members.asp?action=members", "");
addItem("<img src='images/site/blank.gif' align='left' border='0'> Anketler", "Polls.asp?action=polls", "");
addItem("<img src='images/site/blank.gif' align='left' border='0'> Forum", "Forum.asp", "");
<%if Session("Member")="" then%>
addHdr("<img src='images/site/blank1.gif' align='left' border='0'> ZÝYARETÇÝLER ÝÇÝN");
addItem("<img src='images/site/blank.gif' align='left' border='0'> Üye Ol", "Membership.asp?action=new", "");
addItem("<img src='images/site/blank.gif' align='left' border='0'> Þifremi Unuttum", "membership.asp?action=lostpass&step=1", "");
<%else%>
addHdr("<img src='images/site/blank1.gif' align='left' border='0'> ÜYELER ÝÇÝN");
addItem("<img src='images/site/blank.gif' align='left' border='0'> Bilgilerimi Deðiþtir", "Your_Account.asp?op=Profile", "");
addItem("<img src='images/site/blank.gif' align='left' border='0'> Mesajlarým", "Msgbox.asp?action=main", "");
addItem("<img src='images/site/blank.gif' align='left' border='0'> Üye Arama", "search.asp?action=Member", "");
addItem("<img src='images/site/blank.gif' align='left' border='0'> Arkadaþlarým", "Your_Account.asp?op=Friends", "");
addItem("<img src='images/site/blank.gif' align='left' border='0'> Güvenli Çýkýþ", "Your_Account.asp?op=Logout", "");
<% End IF %>
buildMenu();
</script>
<% End IF %>
<div id="yukleme" style="POSITION: absolute; left:40%; top:40%; width:300; height:100">
<link rel="stylesheet" type="text/css" href="THEMES/<%=Session("SiteTheme")%>/style.css">
<table cellspacing="1" cellpadding="0" bgcolor="<%=ArkaBorder%>" bordercolorlight="<%=bordercolorlight%>" bordercolordark="<%=bordercolordark%>" border="1" width="50%">
<tr><td width="100%" nowrap background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" align="center"><font face="Tahoma" size="2"><b>Yükleniyor...</b></font></td></tr>
<tr><td align="center"><img border="0" src="images/site/yukleniyor.gif" ></td></tr>
<tr><td align="center" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif"><b><font face="Tahoma" size="2">Bekleyiniz...</font></b></td></tr>
</table>
</div>
<div align="center">
<center>
<!----- Sað Tuþ Kilit Menu ------>
<% Set rightc=Connection.Execute("Select * FROM SETTINGS")
IF rightc("sagtus")=TRUE then%>
<style><!--
#ie5menu     { position: absolute; width: 210px; background-color: menu; font-family: Tahoma; font-size: 12px; line-height: 20px; cursor: default;  visibility: hidden; border: 2px outset default }
.menuitems   { padding-left: 15px; padding-right: 15px }
//--></style>
<script>
var display_url=0
function showmenuie5(){
ie5menu.style.left=document.body.scrollLeft+event.clientX
ie5menu.style.top=document.body.scrollTop+event.clientY
ie5menu.style.visibility="visible"
return false
}
function hidemenuie5(){
ie5menu.style.visibility="hidden"
}
function highlightie5(){
if (event.srcElement.className=="menuitems"){
event.srcElement.style.backgroundColor="highlight"
event.srcElement.style.color="white"
if (display_url==1)window.status=event.srcElement.url
}
}
function lowlightie5(){
if (event.srcElement.className=="menuitems"){
event.srcElement.style.backgroundColor=""
event.srcElement.style.color="black"
window.status=''
}
}
function jumptoie5(){if (event.srcElement.className=="menuitems")window.location=event.srcElement.url}
function correct(){if (finished){setTimeout("begin()",10000)}return true}
window.onerror=correct
function begin(){
if (!document.all)
return
if (maxheight==null)
maxheight=temp.offsetHeight
whatsnew.style.height=maxheight
temp.style.display="none"
c=1
finished=true
change()
}
//--></script>
<!-- EDIT LINKS BELOW HERE -->
<!--[if IE]>
<div id="ie5menu" onMouseover="highlightie5()" onMouseout="lowlightie5()" onClick="jumptoie5()" style="width: 56; height: 45">
  <div class="menuitems" url="default.asp" style="width: 194; height: 20">
    Anasayfa<font color="#008000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font></div><div class="menuitems" url="news.asp">
    Haberler<font color="#008000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font></div><div class="menuitems" url="programs.asp?action=cats">
    Dosyalar<font color="#008000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font></div><div class="menuitems" url="articles.asp?action=cats">
    Yazýlar<font color="#008000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font></div><div class="menuitems" url="members.asp?action=members">
    Üüyeler<font color="#008000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font></div><div class="menuitems" url="polls.asp?action=polls">
    Anketler<font color="#008000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font></div><div class="menuitems" url="forum.asp">
    Forum<font color="#008000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font></div><div class="menuitems" url="search.asp?action=detailed">
    Arama<font color="#008000">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </font></div></div>
<![endif]-->
<script>
document.oncontextmenu=showmenuie5
if (document.all&&window.print)document.body.onclick=hidemenuie5
</script>
<Script language="JavaScript"><!-- Force all the previous frames off
if (self.parent.frames.length != 0) self.parent.location=document.location.href;
// --></script>
<%end if%>
