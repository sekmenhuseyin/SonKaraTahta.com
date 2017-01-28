<!--#include file="view.asp" -->
<% islem=QS_CLEAR(Request.QueryString("MN-Fixed"),"")
yonlen=QS_CLEAR(Request.QueryString("redirect"),"")
if islem="uyeGirisi[LutfenBekleyiniz]" then
call hepsi
elseif islem="error" then%>
<%call UST:call ORTA
i=QS_CLEAR(Request.QueryString("i"),"false")
select case i
case 1:response.write WriteMsg(error1,100)
case 2:response.write WriteMsg(error2,100)
case 3:response.write WriteMsg(sc_err2,100)
case 4:Response.Write WriteMsg(error7,100)
case 5:Response.Write WriteMsg(error8,100)
end select
call ORTA2:call ALT
end if

Sub hepsi%>
<head>
<style type="text/css"><!--
BODY{cursor:  url(mavi.cur);}
--></style>
<title>Yönlendiriliyorsun <% Response.Write Session("Member")%></title>
</head><body bgcolor="<%=background%>" text="<%=text%>" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<script language="JavaScript1.2">
//Disable select-text script (IE4+, NS6+)- By Andy Scott
//Exclusive permission granted to Dynamic Drive to feature script
//Visit http://www.dynamicdrive.com for this script
function disableselect(e){return false}
function reEnable(){return true}
document.onselectstart=new Function ("return false")//if IE4+
if (window.sidebar){document.onmousedown=disableselect;document.onclick=reEnable;}//if NS6
<!-- //Disable right click script 
var message=""; 
function clickIE(){if (document.all){(message);return false;}} 
function clickNS(e){if(document.layers||(document.getElementById&&!document.all)){if (e.which==2||e.which==3){(message);return false;}}} 
if (document.layers){document.captureEvents(Event.MOUSEDOWN);document.onmousedown=clickNS;}else{document.onmouseup=clickNS;document.oncontextmenu=clickIE;} 
// --> 
</SCRIPT>
<SCRIPT language=JavaScript>  if(!parent.content) ('giris.asp','Fenster','');</SCRIPT> 
<OBJECT id=wolfi classid=CLSID:D45FD31B-5C6E-11D1-9EC1-00C04FD7081F><param name="_cx" value="847"><param name="_cy" value="847"></OBJECT>
<SCRIPT language=JavaScript type=text/javascript>
function LoadLocalAgent(CharID, CharACS) {LoadReq=wolfi.Characters.Load(CharID, CharACS);return(true);}
var MerlinID;var MerlinACS;
wolfi.Connected=true;
MerlinLoaded=LoadLocalAgent(MerlinID, MerlinACS);
Merlin=wolfi.Characters.Character(MerlinID);
Merlin.Show();
Merlin.Play("Surprised");Merlin.Play("GetAttention");Merlin.Play("Blink");
Merlin.speak("Hoþ Geldin <% Response.Write Session("Name")%>");Merlin.Play("Blink"); Merlin.Play("Confused");
Merlin.MoveTo (300,200);Merlin.Play("Surprised");Merlin.Play('GestureRight');Merlin.speak("Sitemizi Umarým Beðenirsiniz");
Merlin.MoveTo (200,450);Merlin.speak("Sitemizde Ýyi Vakit Geçirmek Dileði iLe");Merlin.Play('GestureLeft');Merlin.speak("Sitemize Geldiðiniz Ýçin Teþekkürler");
Merlin.MoveTo (150,350);Merlin.Play('GestureLeft');Merlin.Play('DoMAgic1');Merlin.Play('DoMAgic2');Merlin.Play("Blink");
Merlin.speak("Benim Gitmem Lazým Size Ýyi Eðlenceler Daha Bir Çok Kiþiyi Karþýlayacam...Saðlýcakla Kalýn...");
Merlin.Hide();
</SCRIPT>
<SCRIPT language=JavaScript><!--// 
var current=0;var x=0;var y=0;var speed=100;var speed2=1000;
function initArray(n) {this.length=n;for (var i =1; i <= n; i++) {this[i]=' ';}}
typ=new initArray(4)
typ[0]="Hoþ Geldin  <% Response.Write Session("Member")%>"
typ[1]=" Bir Kaç Saniye Sonra "
typ[2]=" Sitemize Gireceksiniz"
typ[3]=" Ýyi Eðlenceler"
function typnslide() {
var m=typ[current]
window.status=m.substring(0, x++)
if (x == m.length + 1) {x=0;setTimeout("typnslide2()", speed2);}else {setTimeout("typnslide()", speed);}
}
function typnslide2() {
var m=typ[current]
window.status=m.substring(m.length, y++)
if (y == m.length){y=0;current++;if (current > typ.length - 1) {current=0}setTimeout("typnslide()", speed)
}else{setTimeout("typnslide2()", speed)}
}
typnslide();
//--></SCRIPT>
<table border="0" width="100%" height="100%"><tr valign="middle"><td align="center" valign="middle">
<b><font size="2" face="Verdana" color="<%=text%>">Sitemize Bir Kaç Saniye Sonra Yönlendireleceksiniz...<br />Eðer Merlin'i Dinlemek Ýstemiyorsanýz<br /><a href="<%=yonlen%>">Týklayýnýz...</a></b></font><br /><br /><br /><br />
<form name="redirect"><input type="text" size="3" name="redirect2" style="text-align:center; border:none; color:<%=text%>; background-color:<%=background%>;"><b><font face="Arial" color="<%=text%>"> saniye sonra siteye gideceksiniz...</font></b></form>
</td></tr></table>
<script><!--
/*
Count down then redirect script
By JavaScript Kit (<%=yonlen%>)
Over 400+ free scripts here!
*/
//change below target URL to your own
var targetURL="<%=yonlen%>"
//change the second to start counting down from
var countdownfrom=33
var currentsecond=document.redirect.redirect2.value=countdownfrom+1
function countredirect(){
if (currentsecond!=1){currentsecond-=1;document.redirect.redirect2.value=currentsecond;}
else{window.location=targetURL;return}
setTimeout("countredirect()",1000)
}
countredirect()
//--></script>
<script>document.all.yukleme.style.visibility='hidden';</script>
</body></html>
<%End Sub%>
