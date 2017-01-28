<%'programlayan oilcf oilcf@yahoo.com http://www.oilcf.vze.com %>
<%BakCol="#FFCC00":NibCol="#000000":LinCol="#FF0000"%>
<HTML>
<HEAD>
<META NAME="GENERATOR" Content="Microsoft FrontPage 5.0">
<TITLE>aYHaNWeB :: Yýlan Oyunu</TITLE>
<Style Type="text/css">
<!--
    .normalB{font-family:verdana,helvetica,arial,sans serif;color:#CCFFCC;font-size:10pt;font-weight:bold;}
    .normalL{font-family:verdana,helvetica,arial,sans serif;color:#0033CC;font-size:12pt;font-weight:bold;Cursor:hand;}
    .normalR{font-family:verdana,helvetica,arial,sans serif;color:#FFFFFF;font-size:14pt;font-weight:bold;}
    .normal{font-family:verdana,helvetica,arial,sans serif;color:#FFCC66;font-size:12pt;font-weight:bold;}
    .chica{font-family:verdana,helvetica,arial,sans serif;color:#FFCC66;font-size:8pt;font-weight:bold;}
	.boton1{font-family:verdana,helvetica,arial,sans serif;font-weight:bold;background-Color:#FFCC00;Cursor:hand;}
	.boton2{font-family:verdana,helvetica,arial,sans serif;font-weight:bold;background-Color:#FFCC00;Cursor:hand;}
    a:link{color:#0033CC;}
    a:active{color:#990000;}
    a:visited{color:#0033CC;}
    a:hover{color:#ff3300;}
-->
</Style>
<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
'programlayan oilcf oilcf@yahoo.com http://www.oilcf.vze.com
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Dim py,px, Cola,Direc,Timer1,Lev,price,jumplev,Pist,BlkLev,Pausa,Score
jumplev=0
redim cola(200,2)
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
function Iniciar(index)
	blklev=document.all("blocksLev").value
	select case document.all("NoLev").value
		case "1":velo=500
		case "2":velo=300
		case "3":velo=200
		case "4":velo=100
		case "5":velo=80
		case "6":velo=40
		case "7":velo=10
	end select
	window.clearInterval Timer1
	if index=1 then
		if not lev="" then
			document.all(price).innertext=""
			for A=1 to lev
				document.all(cola(a,1) & "x" & cola(a,2)).bgcolor="<%=BakCol%>"
			next
		end if	
	else
		if not lev="" then
			document.all(price).innertext=""
			for A=1 to lev-1
				document.all(cola(a,1) & "x" & cola(a,2)).bgcolor="<%=BakCol%>"
			next
		end if	
	end if
	if index=1 then 
		pist= document.all("StartLev").value
		score=0:document.all("tablero").innertext=score
		for a=1 to 7
			pista1 a,"<%=BakCol%>"
		next	
	else
		pist=pist + 1	
		Pista1 pist-1,"<%=BakCol%>"
	end if
	Pista1 pist,"<%=LinCol%>"
	Direc=40:py=3:px=31:lev=1
	document.all("Nivel").innertext=1
	document.all("Pista").innertext=Pist
	cambio 1
	if not index=1 then msgbox "Bravo Bölüm Bitti " & pist
	Timer1=window.setInterval("Rolar()",velo)
end function
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
function cambio(index)
	Randomize
	price=(round(rnd*22,0)+1) & "x" & (round(rnd*31,0)+1)
	if ucase(document.all(price).bgcolor)="<%=NibCol%>" or ucase(document.all(price).bgcolor)="<%=LinCol%>" then
		Randomize
		price=(round(rnd*22,0)+1) & "x" & (round(rnd*31,0)+1)
		if ucase(document.all(price).bgcolor)="<%=NibCol%>" or ucase(document.all(price).bgcolor)="<%=LinCol%>" then price=(round(rnd*22,0)+1) & "x" & (round(rnd*31,0)+1)	
		if ucase(document.all(price).bgcolor)="<%=NibCol%>" or ucase(document.all(price).bgcolor)="<%=LinCol%>" then price=(round(rnd*22,0)+1) & "x" & (round(rnd*31,0)+1)
		document.all(price).innertext="¤"
	else
		document.all(price).innertext="¤"
	end if
	jumplev=1
	if index=0 then 
		lev=lev+1
		document.all("Nivel").innertext=Lev
		score=score+lev
		document.all("tablero").innertext=score
	end if
	if lev=cint(blklev) then
		Iniciar 0
		score=score+1
		document.all("tablero").innertext=score
	end if
end function
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
function Rolar()
Dim x,y
	select case Direc
		case 39:px=px+1
		case 37:px=px-1
		case 38:py=py-1
		case 40:py=py+1
	end select
	if  py=0 or px=0 or py=24 or px=33 then
		msgbox "Üzgünüm Oyun Bitti",vbinformation,"aYHaNWeB"
		detener
	else
		if not ucase(document.all(py & "x" & px).bgcolor)="<%=BakCol%>" then
			msgbox "Üzgünüm Oyun Bitti",vbinformation,"aYHaNWeB"
			detener
		else
			if document.all(py & "x" & px).innertext="¤" then
				document.all(py & "x" & px).innertext=""
				cambio 0
			end if
			document.all(py & "x" & px).bgcolor="<%=NibCol%>"
			if jumplev=1 then
				jumplev=0
			else			
				document.all(cola(lev,1) & "x" & cola(lev,2)).bgcolor="<%=BakCol%>"
			end if
			nulev="-" & lev
			for a=nulev to -2
				cola(CInt(Right(CStr(A), Len(CStr(A)) - InStr(1, CStr(A), "-"))),1)=cola(CInt(Right(CStr(A), Len(CStr(A)) - InStr(1, CStr(A), "-")))-1,1)
				cola(CInt(Right(CStr(A), Len(CStr(A)) - InStr(1, CStr(A), "-"))),2)=cola(CInt(Right(CStr(A), Len(CStr(A)) - InStr(1, CStr(A), "-")))-1,2)				
			next
			cola(1,1)=py
			cola(1,2)=px			
		end if
	end if
end function
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
function Pista1(pindex,pcolor)
	select case pindex
	case 1
	case 2
		for a=7 to 26
			document.all("11x" & a).bgcolor=pcolor
		next
	case 3
		for a=1 to 16
			document.all("5x" & a).bgcolor=pcolor
			document.all("18x" & a + 16).bgcolor=pcolor
		next
		for a=1 to 14
			document.all(a & "x27").bgcolor=pcolor	
			document.all(a + 9 & "x5").bgcolor=pcolor
		next
	case 4
		for a=6 to 18
			document.all(a & "x7").bgcolor=pcolor
			document.all(a & "x26").bgcolor=pcolor
		next
		for a=9 to 24
			document.all("4x" & a).bgcolor=pcolor
			document.all("20x" & a).bgcolor=pcolor
		next
	case 5
		for a=1 to 32
			if b=1 then
				document.all("11x" & a).bgcolor=pcolor
				b=0
			else
				b=1
			end if
		next
	case 6
		for a=1 to 20
			document.all(a & "x4").bgcolor=pcolor
			document.all(a+3 & "x8").bgcolor=pcolor
			document.all(a & "x12").bgcolor=pcolor
			document.all(a+3 & "x16").bgcolor=pcolor
			document.all(a & "x20").bgcolor=pcolor
			document.all(a+3 & "x24").bgcolor=pcolor
			document.all(a & "x28").bgcolor=pcolor
		next		
	case 7
		for a=3 to 21
			document.all(a & "x" & a).bgcolor=pcolor
			document.all(a & "x" & a+9).bgcolor=pcolor
		next
	case else
		for c=4 to 28 step 8
			for a=1 to 23 step 2
				document.all(a & "x" & c).bgcolor=pcolor
			next
		next
		for c=8 to 28 step 8
			for a=2 to 23 step 2
				document.all(a & "x" & c).bgcolor=pcolor
			next
		next
	end select
end function
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
function Detener()
	window.clearInterval Timer1
end function
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
function Pauss(index)
	msgbox "Oyun Durduruldu.Devam Ýçin Tamama Basýn",vbinformation,"aYHaNWeB"
end function
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Sub document_onkeydown
	select case window.event.keyCode
		case 37 : if not direc=39 then direc=window.event.keyCode
		case 39 : if not direc=37 then direc=window.event.keyCode
		case 40 : if not direc=38 then direc=window.event.keyCode
		case 38 : if not direc=40 then direc=window.event.keyCode
		case 27 : pauss pausa
	end select
End Sub
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
Sub document_onmousedown
	if window.event.button="2" then msgbox "Sað Tuþla Hiç Bir Özellik Kullanamazsýnýz",vbinformation,"aYHaNWeB"
End Sub
'¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤¤
</SCRIPT>
</HEAD>
<BODY leftmargin=3 topmargin=3 bgcolor="#0099FF">
<div align=center>
<table border="2" bordercolor="#0033FF" bgcolor="#0099FF" cellspacing="0" cellpadding="0" style="border-collapse: collapse"><tr>
  <td width="175" valign="top" bordercolor="#0033FF">
<p align="center">
<input type=button onclick="Iniciar(1)" value="    Oyuna Baþla    " class=boton2 id=button1 name=button1></p>
<center>
<p align="center" class=normalB>Puan<br /><a class=normalR id=tablero name=tablero>&nbsp;</a></p>
<a class=normalB>Hýz</a><br />
<select size=1 id="NoLev" name="NoLev" class=boton1>
<%For a=1 to 5%><option value=<%=a%>><%=a%></option><%next%>
<option selected value=6>6</option>
<option value=7>7</option>
</select>
<br /><br />
<a class=normalB>Bölüm</a><br />
<a ID="Pista" Name="Pista" class=normal>&nbsp;</a><br /><br />
<a class=normalB>Toplanan Blok</a><br />
<a ID="Nivel" Name="Nivel" class=normal>&nbsp;</a><br />
<br />
<a class=normalB>Baþlangýç Bölümü</a><br />
<select size=1 id="StartLev" name="StartLev" class=boton1>
<%For a=1 to 8%><option value=<%=a%>><%=a%></option><%next%>
</select>
<br /><br />
<a class=normalB>Bölümdeki Blok Sayýsý</a><br />
<select size=1 id="blocksLev" name="blocksLev" class=boton1>
<option value=5>5</option>
<option value=10>10</option>
<%For a=20 to 180 step 20%><option value=<%=a%>><%=a%></option><%next%>
<option selected value=200>200</option>
</select>
<br /><br />
<a class=chica>&quot;escape&quot; tuþu<br />oyunu durdurur</a>
</center>
</td><td>
<table border="5" bordercolor="#0000FF" cellspacing="0" cellpadding="0"><tr><td>
<table cellspacing="0" cellpadding="0">
<%For b=1 to 23%><tr height="18"><%For a=1 to 32%><td ID="<%=b%>x<%=a%>" width="18" align=center bgcolor="<%=BakCol%>"><%next%></tr><%next%>
</table>
</td>
</td></tr></table></table>
</div>
</BODY>