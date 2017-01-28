<%
Set Connection= Server.CreateObject ("ADODB.connection")
connection.mode=3
Connection.Provider= "Microsoft.Jet.OLEDB.4.0"
Connection.Properties ("Data Source")= Server.MapPath ("DB\snake.mdb")
Connection.Open

Set RSsnakescore=Server.CreateObject("ADODB.recordset")
    RSsnakescore.Open "SELECT * FROM tblsnake order by score,ID", Connection, 3

lowscore=0
If RSsnakescore.eof=false then
lowscore=RSsnakescore.fields("score")
lownr=RSsnakescore.fields("ID")
end if

aantalrecords=RSsnakescore.recordcount


score=request.form("score")
score=cint(score)
name=request.form("username")
if (name="") then
 name="Unkown"
end if
insert_date= date()
msg="nok"
if score > lowscore or aantalrecords < 10 then
msg="ok"
connection.execute "insert into tblsnake (Name,Score,insert_date) values ('"&Replace(name,"'","''")&"',"&score&",'"&Insert_date&"');"

	if aantalrecords >= 10 then
		
		connection.execute "delete * from tblsnake where id="&lownr&";"

	end if

end if

RSsnakescore.close
Set RSsnakescore=nothing
Connection.close
Set Connection=nothing
response.redirect "modules\snake\snake.asp?msg="&msg
%>