<!--#include file="database.asp" --><!--#include file="themes.asp" --><!--#include file="filter.asp" -->
<!--#include file="functions.asp" --><!--#include file="online.asp" --><!--#include file="Language_File.asp" -->
<%Session.LCID = 1033:Set sett = Server.CreateObject("ADODB.RecordSet"):settSQL = "Select * FROM SETTINGS":sett.open settSQL,Connection,1,3%>

