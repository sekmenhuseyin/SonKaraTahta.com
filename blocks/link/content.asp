<!--#include file="../../INC/b_include.asp" -->
<p align="center"><font face="Verdana" size="2">
<MARQUEE style="cursor:  url(../../images/site/mavi.cur);" onmouseover=this.stop() onmouseout=this.start() scrollAmount=1 
      scrollDelay=100 direction=up height="160" width="124">
<%Set baglanti = Server.CreateObject("ADODB.Connection"):baglanti.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("db/mn"&dbNum&".mdb")%>
<body text="#000000">
<%set liste = Server.CreateObject("ADODB.RecordSet"):SQL_L = "Select * from linklist WHERE onay=1 ORDER by tekil DESC":liste.open SQL_L,baglanti,1,3%>
<br />
<table width="1" border="0" cellspacing="0" cellpadding="0" align="center" >
  <tr>
    <td width="200">
      <div align="center">
        <center>
      <table width="200" border="0" cellspacing="0" cellpadding="3" height="21" style="border-collapse: collapse">
        <%
		for i = 1 to liste.recordcount
		if liste.eof then exit for
		if i mod 2 = 0 then
		end if
		%>
        <tr bgcolor="<%=bg%>"> 
          </td>
          <td height="18"><font size="1"><b><a href="<%=liste("url")%>" target="_blank"><%=liste("adi")%></a><b></font></td>
        </tr>
		<%
		liste.movenext
		Next
		liste.close
		set liste = nothing
		%>
      </table>
        </center>
      </div>
    </td>
  </tr>
</table>
<div align="center"><br />
   </div>
</body>
</html></marquee></font><br />
<br />
<FONT 
face=Verdana size=1 > 

<a onClick="window.name='listecom'; window.open('Link_kayit.asp','Kayit','toolbar=no,status=no,menubar=no,scrollbars=no,resizable=no,width=325,height=140'); return false;" style="text-decoration: none; font-weight: 700" href="default.aspx?ekle=LinkEklePleace);">&lt; Siteni Ekle &gt;</a></FONT>