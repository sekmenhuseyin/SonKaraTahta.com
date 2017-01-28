<!--#include file="../../INC/b_include.asp" -->
<body leftmargin="0">
<div align="center">
<center>
  <p></p>
  </center>
</div>
<div align="center"> 
  <center>
    <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1" height="58">
      <tr> 
        <%
yilinayi = month(date)
%>
        <td width="35%" height="58" rowspan="3"> <table width="70" height="80" border="0" align="center" cellpadding="0" cellspacing="0">
            <tr> 
              <td background="Blocks/Tarih/Tarih.gif"><div align="center"><font size="9" face="Verdana, Arial, Helvetica, sans-serif"><%=day(date)%></font><font color="#000000" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                  <strong> 
                  <%if yilinayi="1" then%>
                  Ocak 
                  <%elseif yilinayi="2" then%>
                  Þubat 
                  <%elseif yilinayi="3" then%>
                  Mart 
                  <%elseif yilinayi="4" then%>
                  Nisan 
                  <%elseif yilinayi="5" then%>
                  Mayýs 
                  <%elseif yilinayi="6" then%>
                  Haziran 
                  <%elseif yilinayi="7" then%>
                  Temmuz 
                  <%elseif yilinayi="8" then%>
                  Aðustos 
                  <%elseif yilinayi="9" then%>
                  Eylül 
                  <%elseif yilinayi="10" then%>
                  Ekim 
                  <%elseif yilinayi="11" then%>
                  Kasým 
                  <%elseif yilinayi="12" then%>
                  Aralýk 
                  <%end if%>
                  <br />
                  </strong> </font></div></td>
            </tr>
          </table>
          <div align="center"><font color="#000000" size="3" face="Verdana, Arial, Helvetica, sans-serif"><strong><%=year(date)%></strong></font> 
            <font size="1" face="Verdana, Arial, Helvetica, sans-serif"><br />
            <font color="#000000"><%=date%></font></font> </div></td>
      </tr>
      <tr> 
        
      </tr>
    </table>
</center>
</div>
<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,29,0" width="115" height="55">
  <param name="movie" value="blocks/tarih/clock.swf">
  <param name="quality" value="High">
  <param name="_cx" value="3043">
  <param name="_cy" value="1455">
  <param name="FlashVars" value>
  <param name="Src" value="blocks/tarih/clock.swf">
  <param name="WMode" value="Window">
  <param name="Play" value="-1">
  <param name="Loop" value="-1">
  <param name="SAlign" value>
  <param name="Menu" value="-1">
  <param name="Base" value>
  <param name="AllowScriptAccess" value="always">
  <param name="Scale" value="ShowAll">
  <param name="DeviceFont" value="0">
  <param name="EmbedMovie" value="0">
  <param name="BGColor" value>
  <param name="SWRemote" value>
  <param name="MovieData" value>
  <param name="SeamlessTabbing" value="1">
  <embed src="blocks/calendar/clock1.swf" quality="high" pluginspage="http://www.macromedia.com/go/getflashplayer" type="application/x-shockwave-flash" width="115" height="55"></embed></object>