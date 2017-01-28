<!--#include file="view.asp" -->
<%go=Request.QueryString("go"):go=QS_CLEAR(go,"")
If go="ads" Then
id=Request.QueryString("bid"):id=QS_CLEAR(id,"false")
Set ads=Server.CreateObject("ADODB.RecordSet"):adsSQL="Select * FROM BANNERS WHERE banner_id="&id&"":ads.open adsSQL,Connection,3,3
If ads.Eof Then
Response.Redirect ""&sett("site_adres")&""
ELSE
ads("hit")=ads("hit") + 1
ads.Update
Response.Redirect ""&ads("go_url")&""
End If
ads.Close:Set ads=Nothing
End If %>