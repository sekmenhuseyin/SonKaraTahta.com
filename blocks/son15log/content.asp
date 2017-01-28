<!--#include file="../../INC/b_include.asp" -->
<font size=tahoma size=1>
          <%set llo = connection.execute("select * from members order by son_tarih desc")%> </font>
    <%for sttr = 1 to 15
if llo.eof Then exit For

son = llo("son_tarih")

fark = datediff("n",son,now) & " Dk."
IF fark="0 Dk." then 
fark="Online"
End IF
%> <a href="members.asp?action=member_details&uid=<%=llo("uye_id")%>"><b> [<%=llo("kul_adi")%>]</b></a> <%=fark%><br /><%llo.movenext
next%><b>önce giriþ yaptýlar..</b>
<br />
<br />