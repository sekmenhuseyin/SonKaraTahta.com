<%Session.LCID=S_LCID
Set sett=Server.CreateObject("ADODB.RecordSet"):settSQL="Select * FROM SETTINGS":sett.open settSQL,Connection,1,3
IF sett("site_locked")="1" Then
Response.Write "<html><head><title>"&sett("site_isim")&"</title></head><body>"
response.Write "<table width='100%' height='100%'><tr><td align='center' valign='middle' style='font-family:tahoma;font-size:12px'><b>"
Response.Write sett("site_locked_msg")&"</b></td></tr></table></body></html>"
Response.End
End IF
Set nCount=Server.CreateObject("ADODB.RecordSet"):nSQL="Select * FROM WEBCOUNTER":nCount.open nSQL,Connection,1,3
Set c_check=Connection.Execute("Select * FROM WEBCOUNTER WHERE cdate=date()-0")
If c_check.Eof OR c_check.Bof Then
If Hour(Now())=0 OR Hour(Now()) > 0 Then
nCount.AddNew
nCount("today")="0"
nCount("cdate")=date()
nCount.Update
End If
End If
ip=Request.ServerVariables("REMOTE_ADDR")
Set checkIP=Connection.Execute("SELECT * FROM GUESTS WHERE ip='"&ip&"'")
If checkIP.EOF Then
Set addIP=Server.CreateObject("ADODB.RecordSet"):addIP.Open "SELECT * FROM GUESTS", Connection, 1, 3
addIP.AddNew
addIP("ip")=ip
addIP("zdate")=Now()
addIP.Update
addIP.close
Else
IF Session("Enter")="Yes" Then
Set updIP=Connection.Execute("DELETE * FROM GUESTS WHERE ip='"&ip&"'")
Else
Set updIP=Server.CreateObject("ADODB.RecordSet"):updIP.Open "SELECT * FROM GUESTS WHERE ip='"&ip&"'", Connection, 1, 3
updIP("ip")=ip
updIP("zdate")=Now()
updIP.Update
updIP.close
End If
End If
Set delIP=Connection.Execute("DELETE * FROM GUESTS WHERE zdate < now()")
Set LB=Server.CreateObject("ADODB.RecordSet"):LB.open "Select * FROM BLOCKS WHERE b_align='LEFT' and b_active=true ORDER BY b_queue",Connection,3,3
Set RB=Server.CreateObject("ADODB.RecordSet"):RB.open "Select * FROM BLOCKS WHERE b_align='RIGHT' and b_active=true ORDER BY b_queue",Connection,3,3
IF Session("Enter")="Yes" Then Session("SiteTheme")=Session("Theme") ELSE Session("SiteTheme")=sett("theme")
Session.LCID=1033
IF Session("Enter")="Yes" Then Set updtIP=Connection.Execute("UPDATE MEMBERS SET last_ip='"&ip&"',son_tarih=now() WHERE uye_id="&Session("Uye_ID")&"")%>