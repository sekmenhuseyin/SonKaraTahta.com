<%@ Language=VBScript %>
<%
newid="user" & request("id")
boxid="box" & request("id")
flagid="flag" & request("id")
fromid="from" & request.cookies("myid")
%>
<html>
<head>
<title>Mesaj Gönder - Tüm Kullanıcılara</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<script language="JavaScript">
<!--
function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->
</script>
<SCRIPT LANGUAGE="JavaScript">
<!--
<!-- Begin
function showLines(max, text) {
max--;
text="" + text;
var temp="";
var chcount=0; 
for (var i=0; i < text.length; i++) 
{   
var ch=text.substring(i, i+1); 
var ch2=text.substring(i+1, i+2); 
if (ch == '\n') 
{  
temp += ch;
chcount=1;
}
else
{
if (chcount == max) 
{
temp += '\n' + ch; 
chcount=1;
}
else  
{
temp += ch;
chcount++; 
      }
   }
}
return (temp); 
}
//  End -->
</script>
<script language="javascript">

function MM_jumpMenu(targ,selObj,restore){ //v3.0
  eval(targ+".location='"+selObj.options[selObj.selectedIndex].value+"'");
  if (restore) selObj.selectedIndex=0;
}
//-->
</script>
<SCRIPT LANGUAGE="JavaScript">
function toForm() {
document.form1.message.focus();
}
</script>

</head>
<body bgcolor="C1D8EE" link="#0000FF" vlink="#0000FF" alink="#0000FF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="toForm()">
  
<form name="form1" method="post" action="addmessage_all.asp" >
  <center>
    <strong><img src="sendmessage.gif" width="150" height="35"></strong> <br />
    <b><font size="2" face="Arial, Helvetica, sans-serif">Kime</font></b><br />
    <table width="220" border="0" bgcolor="#5996D2">
      <tr>
        <td>
          <div align="center"><font color="white"><b><font face="Arial, Helvetica, sans-serif">Tüm 
            Kullanıcılara </font></b></font></div>
        </td>
      </tr>
    </table>
    <b><font size="2" face="Arial, Helvetica, sans-serif">Mesaj</font></b><br />
      
    <textarea rows=4 cols=35 wrap=virtual name="message" style="background-color: #FFFFFF; color: #000000; font-size: 10px; font-family: Verdana; border: 1 solid #D2D2D2"></textarea>
    <br />
    <table width="150" border="0" cellspacing="0" cellpadding="0" height="45">
      <tr> 
        <td height="10" valign="bottom">
          <div align="center"> 
            <input type="submit" name="Submit" value="Gonder" style="background-color: #FFFFFF; color: #000000; font-size: 10px; font-family: Verdana; border: 1 solid #D2D2D2">
          </div>
        </td>
      </tr>
    </table>
</center></form>




</body>
