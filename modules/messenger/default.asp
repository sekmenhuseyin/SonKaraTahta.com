<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Mini Messenger</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript">
function newWindow(mypage,myname,w,h,features) {
  if(screen.width){
  var winl=(screen.width-w)/2;
  var wint=(screen.height-h)/2;
  }else{winl=0;wint =0;}
  if (winl < 0) winl=0;
  if (wint < 0) wint=0;
  var settings='height=' + h + ',';
  settings += 'width=' + w + ',';
  settings += 'top=' + wint + ',';
  settings += 'left=' + winl + ',';
  settings += features;
  win=window.open(mypage,myname,settings);
  win.window.focus();
}
</script>
<script language="JavaScript">
<!-- Hide the script from old browsers --
// Michael P. Scholtis (mpscho@planetx.bloomu.edu)
// All rights reserved.  December 22, 1995
// You may use this JavaScript example as you see fit, as long as the
// information within this comment above is included in your script.
function browsertest () 
        {document.write("Bu Hizmet için Java GerekLi")}
// --End Hiding Here -->
</script>
<script src="plugins.js">

//Detect Plugin (Flash, Java, RealPlayer etc) script- By Frederic (fw4@tvd.be)
//Visit http://javascriptkit.com for this script and more

//SAMPLE USAGE- detect "Flash"
//if (pluginlist.indexOf("Flash")!=-1)
//document.write("Flash Yükleyin")
</script>
</head>

<body>
<p><br />
  <br />
  <br />
</p>
<CENTER>
<SCRIPT LANGUAGE="JavaScript">
<!--
   {browsertest();} 
//-->
</SCRIPT>
</CENTER>
<CENTER>
<script>
if (pluginlist.indexOf("Flash")!=-1)
document.write("You have Flash installed.")
</script>
</CENTER>
<p align="center"><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>
<a href="JavaScript:newWindow('modules/messenger/login.asp','popup',310,440,'')">
<img src="modules/messenger/bullet.gif" border="0">Baglan  
  Messenger</a></strong></font><br />
  &nbsp;</p>
</body>
</html>