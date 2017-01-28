<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Mini Messenger Giriş</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
</head>

<body bgcolor="C1D8EE">
<form name="form1" method="post" action="login.asp">
<div align="center">
  <table width="280" height="400" border="0" bgcolor="#FFFFFF">
    <tr>
      <td width="288" valign="top"><table width="267" height="295" border="0" align="center">
          <tr> 
            <td height="37">
              <div align="center"></div>
            </td>
          </tr>
          <tr> 
            <td width="261" height="248" valign="top" background="login.gif"><div align="center">
                  <table width="227" height="100%" border="0" cellpadding="0" cellspacing="0">
                    <tr> 
                      <td width="84" height="167">&nbsp;</td>
                      <td width="143">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td height="56"><div align="right"><font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Kullanıcı 
                          Adı </strong></font></div></td>
                      <td><font color="#FFFFFF"> 
                        <input type="text" onFocus="this.value='';" name="kuladi" style="background-color: #FFFFFF; color: #000000; font-size: 10px; font-family: Verdana;  border: 1 solid #D2D2D2" size="15" value="<% If Session("Member")="" or "0" Then Response.Write "Ziyaretçi" Else Response.Write Session("Member") End If%>">
                        </font></td>
                    </tr>
                    <tr> 
                      <td>&nbsp;</td>
                      <td>&nbsp;</td>
                    </tr>
                  </table>
              </div></td>
          </tr>
        </table>
        <div align="center">
                    <font class="td_menu">
            <input class="submit" onclick="this.form.submit();this.disabled=true; return true;" type="submit" value="Giriş"></font>
        </div></td>
    </tr>
  </table>
</div>
</form>
</body>
</html>