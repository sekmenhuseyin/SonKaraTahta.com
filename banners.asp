<%
Set ban1=Server.CreateObject("ADODB.Recordset" ):b1SQL="Select * FROM BANNERS WHERE active=True AND position='top'":ban1.Open b1SQL,Connection,3,3
IF ban1.Eof Then
Response.Write ""
Set banner1=Nothing
ELSE
TotalRecord1=ban1.RecordCount

Randomize
MoveEntry1=Int((Rnd*TotalRecord1))

ban1.Move(MoveEntry1)

ban1("show")=ban1("show") + 1
ban1.Update
Set banner1=Server.CreateObject("ADODB.RecordSet")
bn1SQL="Select * FROM BANNERS WHERE banner_id="&ban1("banner_id")&""
banner1.open bn1SQL,Connection,3,3
END IF

Set ban2=Server.CreateObject("ADODB.Recordset" )
b2SQL="Select * FROM BANNERS WHERE active=True AND position='bottom'"
ban2.Open b2SQL,Connection,3,3
IF ban2.Eof Then
Response.Write ""
ELSE
TotalRecord2=ban2.RecordCount

Randomize
MoveEntry2=Int((Rnd*TotalRecord2))

ban2.Move(MoveEntry2)

ban2("show")=ban2("show") + 1
ban2.Update
Set banner2=Server.CreateObject("ADODB.RecordSet")
bn2SQL="Select * FROM BANNERS WHERE banner_id="&ban2("banner_id")&""
banner2.open bn2SQL,Connection,3,3
End If
%>