<!--#include file="view.asp" -->
<%call UST:call ORTA
mdle=Request.QueryString("name"):mdle=QS_CLEAR(mdle,"")
Set mdl=Connection.Execute("SELECT * FROM MODULES WHERE mname='"&mdle&"'")
IF mdl.Eof OR mdl.Bof OR mdl.recordcount=0 or isnumeric(mdle)=false Then
Response.Write "<table border=""0"" cellspacing=""1"" cellpadding=""0"" align=""center"" width=""100%"" class=""td_menu"" bgcolor="""&td_border_color&""" height=""20""><tr><td align=""center"" bgcolor="""&menu_back&"""><b>"&module_eof&"</b></td></tr></table>"
ELSE%>
<table border="0" cellpadding="2" cellspacing="0" width="100%" style="<%=TableBorder%>">
<tr><td class="td_menu" background="THEMES/<%=Session("SiteTheme")%>/menu_bg.gif" height="25"><font class="block_title">&nbsp;<img src="THEMES/<%=Session("SiteTheme")%>/menu_bullet.gif" align="absmiddle">&nbsp;<%=mdl("mname")%></font></td></tr>
<tr><td align="left" valign="top" bgcolor="<%=menu_bg%>" class="td_menu">
<%filen="default.asp"
If mdl("mems")=True Then
If Session("Enter")="Yes" Then Server.Execute ("MODULES\"&mdl("mdir")&"\"&filen&"") Else Response.Write sett("np_msg")
ELSE
Server.Execute ("MODULES\"&mdl("mdir")&"\"&filen&"")
End IF%>
</td></tr>
</table>
<%End IF
call ORTA2:call ALT %>