<!--#include file="inc/database.asp" --><!--#include file="inc/language_file.asp" --><!--#include file="inc/encryption.asp" --><!--#include file="inc/filter.asp" --><!--#include file="inc/functions.asp" -->
<%kuladi=duz(Request.Form("kuladi"))
password=duz(Request.Form("password"))
guvenlik=duz(Request.Form("guvenlik"))
gguvenlik=duz(Request.Form("gguvenlik"))
action=duz(Request.Form("action"))
If kuladi="" Then
if action="admin" then response.Write error1 else response.Redirect("giris.asp?MN-Fixed=error&i=1")
ElseIf password="" Then
if action="admin" then response.Write error2 else response.Redirect("giris.asp?MN-Fixed=error&i=2")
ElseIf not ucase(gguvenlik)=ucase(guvenlik) Then
if action="admin" then response.Write sc_err2 else response.Redirect("giris.asp?MN-Fixed=error&i=3")
ElseIf not ucase(session("login_gguvenlik"))=ucase(guvenlik) Then
if action="admin" then response.Write sc_err2 else response.Redirect("giris.asp?MN-Fixed=error&i=3")
Else
if action="admin" then set rs=Connection.Execute("SELECT * FROM MEMBERS where seviye>0") else set rs=Connection.Execute("SELECT * FROM MEMBERS")
Do Until rs.Eof
IF ucase(kuladi)=ucase(rs("kul_adi")) Then
kul_adi=True
IF MN_PP(password)=rs("sifre") Then 
password=True
Session("Enter")="Yes"
Session("Uye_ID")=rs("uye_id")
Session("Member")=rs("kul_adi")
Session("Name")=rs("isim")
Session("Email")=rs("email")
Session("Level")=rs("seviye")
Session("Signature")=rs("imza")
Session("Theme")=rs("u_theme")
Session("LastLogged")=""&rs("son_tarih")&""
Session.LCID=1033
Set up=Connection.Execute("UPDATE MEMBERS SET son_tarih=now(),u_online=true,login_sayisi="&rs("login_sayisi")&"+1 WHERE uye_id="&Session("Uye_ID")&"")
Set delete_user=Connection.Execute("DELETE * FROM GUESTS WHERE ip='"&Request.ServerVariables("REMOTE_ADDR")&"'")
Set rightc=Connection.Execute("Select * FROM SETTINGS")
if action="admin" then Response.Redirect ("administrator/")
IF rightc("merlin")=false then Response.Redirect Request.ServerVariables("HTTP_REFERER") ELSE Response.Redirect "Giris.asp?Mn-Fixed=uyeGirisi[LutfenBekleyiniz]&redirect="&Request.ServerVariables("HTTP_REFERER")
Set rightc=Nothing
End If
End If
rs.MoveNext
Loop
If kul_adi <> True then
if action="admin" then response.Write error7 else response.Redirect("giris.asp?MN-Fixed=error&i=4")
ElseIf password <> True then
if action="admin" then response.Write error8 else response.Redirect("giris.asp?MN-Fixed=error&i=5")
End If
rs.Close:Set rs=Nothing
End IF%>
