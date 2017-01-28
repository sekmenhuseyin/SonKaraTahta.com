<%IF Session("Enter")="Yes" Then
Set chkML=Server.CreateObject("ADODB.RecordSet"):chkML.Open"Select * FROM MEMBERS WHERE uye_id="&Session("Uye_ID")&"",Connection,3,3
Do Until chkML.Eof
IF chkML("seviye")="1" Then
Session("Level")="1"
ElseIF chkML("seviye")="2" Then
Session("Level")="2"
ElseIF chkML("seviye")="3" Then
Session("Level") ="3"
ElseIF chkML("seviye")="4" Then
Session("Level") ="4"
ElseIF chkML("seviye")="5" Then
Session("Level") ="5"
ElseIF chkML("seviye")="6" Then
Session("Level") ="6":Session("ISMod")="Yes"
ElseIF chkML("seviye")="7" Then
Session("Level") ="7"
ElseIF chkML("seviye")="8" Then
Session("Level") ="8"
ELSE
Session("Level")=""
End IF
chkML.MoveNext
Loop
ELSE
Session("Level")="":Session("ISMod")=""
End IF%>