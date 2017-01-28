<html><head><meta http-equiv="Content-Type" content="text/html; charset=windows-1254">
<style>
.form { font-family:verdana; font-size: 11px; border-width:1px; page-break-inside: avoid; border-style:solid; border-collapse:collapse; }
.buton { font-family:verdana; font-size: 11px; }
.tabloalt_font { font-family:tahoma; font-size:10px; color:#000000; }
</style>
</head><body leftmargin=0 topmargin=0 marginwidth=0 marginheight=0>
<%If session("Member") = "" then
Response.Write "<center><font class=""tabloalt_font"">Giriş yapmadınız..</font></center>"
Else%>
<FORM METHOD="post" ACTION="msg_yaz.asp" name="frmMSG"><center><INPUT NAME="text" size="12" maxlength="90" class="form">&nbsp;<INPUT type="submit" value="»»" class="buton"></center></FORM>
<script language="javascript">document.frmMSG.text.focus();</script>
<% End If
mesaj = Request.Form("text") 
mesaj = server.HTMLencode(mesaj)
If NOT Replace(mesaj," ","")="" then
APPLICATION.LOCK
Application("msg20") = Application("msg19")
Application("msg19") = Application("msg18")
Application("msg18") = Application("msg17")
Application("msg17") = Application("msg16")
Application("msg16") = Application("msg15")
Application("msg15") = Application("msg14")
Application("msg14") = Application("msg13")
Application("msg13") = Application("msg12")
Application("msg12") = Application("msg11")
Application("msg11") = Application("msg10")
Application("msg10") = Application("msg9")
Application("msg9") = Application("msg8")
Application("msg8") = Application("msg7")
Application("msg7") = Application("msg6")
Application("msg6") = Application("msg5")
Application("msg5") = Application("msg4")
Application("msg4") = Application("msg3")
Application("msg3") = Application("msg2")
Application("msg2") = Application("msg1")
Application("msg1") = "<B>"&session("Member")&" :</B> "&mesaj
APPLICATION.UNLOCK
END IF%>
</body></html>