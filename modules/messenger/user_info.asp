<%@ Language=VBScript %>
<%
newid="user" & request.cookies("myid")
name=application("" & newid & "")
%>
<html>
<head>
<title>Kullanıcı Bilgisi</title>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1254">

</head>
<body bgcolor="C1D8EE" link="#0000FF" vlink="#0000FF" alink="#0000FF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" onLoad="toForm()">
  <center>
    <strong><font size="3" face="Verdana, Arial, Helvetica, sans-serif"><br />
  Kullanıcı Bilgisi </font></strong> <br />
  <hr />
  <b><font size="2" face="Arial, Helvetica, sans-serif">Kullanıcı Adı</font></b><br />
    <table width="220" border="0" bgcolor="#5996D2">
      <tr>
        <td>
          <div align="center"><font color="white"><b><font face="Arial, Helvetica, sans-serif"><%=name%></font></b></font></div>
        </td>
      </tr>
    </table>
    
  <br />
  <b><font size="2" face="Arial, Helvetica, sans-serif">Yetkisi</font></b><br />
    <table width="220" border="0" bgcolor="#5996D2">
      <tr> 
        <td> <div align="center"><font color="white"><b><font face="Arial, Helvetica, sans-serif">Normal 
            Kullanıcı </font></b></font></div></td>
      </tr>
    </table>
    
  <hr />
  <a href="javascript:window.close()"><font size="2">(kapat)</font></a> 
</center>




</body>
