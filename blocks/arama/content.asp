<!--#include file="../../INC/includes.asp" -->
   <form action="<%=sett("site_adres")%>/search.asp" method="get">

      <font class="td_menu"><b><%=arama%></b> : 
        <input type="text" name="search" class="form1" value="<%=Request("search")%>" size="14">
        <input type="image" src="<%=sett("site_adres")%>/THEMES/<%=Session("SiteTheme")%>/ara.gif" border="0" align="absmiddle"></font>
</form>