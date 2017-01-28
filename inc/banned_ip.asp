<%Set bipRs=Server.CreateObject("ADODB.RecordSet"):bipRs.open "Select * FROM BANNED_IPS",Connection,3,3
bannedIP=Request.ServerVariables("REMOTE_ADDR")
Do Until bipRs.EoF
IF bipRs("b_ip")=bannedIP OR bipRs("b_ip")=""&Session("Uye_ID")&"" Then
Response.Write "<html><head><title>"&sett("site_isim")&"</title></head><body>"
response.Write "<table width='100%' height='100%'><tr><td align='center' valign='middle' style='font-family:tahoma;font-size:12px'><b>"
Response.Write banned_ip_text&"</b></td></tr></table></body></html>"
Response.End
End IF
bipRs.MoveNext
Loop
bipRs.Close
Set bipRs=Nothing%>