<%Session.LCID = 1033
liste = DateAdd("n", -1 * 5, Now())
Set guests = Connection.Execute("SELECT COUNT(*) AS sayi FROM GUESTS")
Set om = Connection.Execute("SELECT * FROM MEMBERS where u_online=true ORDER BY son_tarih DESC")
Set w_offline = Connection.Execute("UPDATE MEMBERS SET u_online = False where son_tarih <= #"&liste&"#")
Set mem_c = Connection.Execute("Select COUNT(*) AS say FROM MEMBERS where son_tarih >= #"&liste&"#")
IF om.eof Then mem_count = "0" Else mem_count = mem_c("say")%>