<%Response.Buffer=TRUE%>
<%
if request("kuladi")="" then
response.redirect("login_default.asp")
else
%>
<%
application.lock
newusers=application("totusers")
if application("totusers")="" then
application("totusers")=1
newid=1
else
newid=application("totusers") + 1
application("totusers")=newid
end if
userid="user" & newid
boxid="box" & newid
flagid="flag" & newid
authid="auth" & newid
updateid="update" & newid

myid=newid
application("" & userid & "")=request("kuladi")
application("" & boxid & "")=""
application("" & flagid & "")="no"
application("" & authid & "")=""
application("" & updateid & "")=now
response.cookies("newid")=newid
response.cookies("boxid")=boxid
response.cookies("flagid")=flagid
response.cookies("authid")=authid
response.cookies("imlogdin")="yes"
response.cookies("myid")=myid
response.redirect("messenger.asp")
application.unlock
%>
<%
end if
%>