<%
'Mix Masta Mikee's Flash/ASP IM Server

if request.form("Action")="Check" then
boxid=request.cookies("boxid")
updateid="update" & request.cookies("newid")
application("" & updateid & "")=now
if application("" & request.cookies("flagid") & "")="no" then
else
if application("" & request.cookies("flagid") & "")="yes" then
response.write "found=leroy"
end if
end if
end if
%>

<%
if request.form("Action")="Display" then
%>
<html>
<head>
<script LANGUAGE="JavaScript">
function popupwindow(){
window.open("display_message.asp","","height=130,width=220")
}
</script>
</head>
<body onload="javascript:popupwindow()">
</body>
</html>

<%
end if
%>