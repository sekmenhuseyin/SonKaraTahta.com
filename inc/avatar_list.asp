<%Set aRs=Server.CreateObject("ADODB.RecordSet"):aRs.open "Select * FROM AVATARS ORDER BY a_name ASC",Connection,3,3
Do Until aRs.EoF%><option value="IMAGES/avatars/<%=aRs("a_img")%>"><%=aRs("a_name")%></option><%aRs.MoveNext:Loop
aRs.Close:Set aRs=Nothing%>