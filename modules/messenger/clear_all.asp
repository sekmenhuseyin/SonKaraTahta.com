<%
application.lock
users=application("totusers")
application("totusers")=""
for x=1 to users
user="user" & x
box="box" & x
from="from" & x
flag="flag" & x
auth="auth" & x
application("" & user & "")=""
application("" & box & "")=""
application("" & from & "")=""
application("" & flag & "")=""
application("" & auth & "")=""
next
application.unlock
%>
<center><br /><br /><b>Mesajlar Silindi</b></center>