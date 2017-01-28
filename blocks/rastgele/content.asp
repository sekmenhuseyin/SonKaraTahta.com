<!--#include file="../../INC/b_include.asp" -->
<!--#include file="../../db/data.asp" -->
<%
Set Sur = Server.CreateObject("ADODB.Connection")
Sur.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("db/mn"&dbNum&".mdb")
%>
<%
Set liste = Server.CreateObject("ADODB.Recordset")
SQL = "Select * from RASTGELE ORDER BY id desc"
liste.Open SQL,Sur,1,3
Randomize
id = Int((toplam * Rnd)+ 0) 
liste.Move(id)
%>

<%
Set liste = Server.CreateObject("ADODB.Recordset")
SQL = "Select * from Rastgele ORDER BY id desc"
liste.Open SQL,Sur,1,3
toplam=liste.recordcount
Randomize
id = Int((toplam * Rnd)+ 0) 
liste.Move(id)
%>
<p align="center"><a href="javascript:void window.open('blocks/rastgele/goster.asp?menu=show&imgid=<%=liste("resim")%>','','width=750,height=600,scrollbars=2')"> <script language="JavaScript1.1">
//Picture Cube slideshow - By Tony Foster III
//Modifications by JK
var specifyimage=new Array() //Your images / resimler
specifyimage[0]="<%=liste("resim")%>"
specifyimage[1]="<%=liste("resim")%>"
specifyimage[2]="<%=liste("resim")%>"

var delay=3000 //3 seconds / 3 saniye bekleme

//Counter for array / sýralamadaki ikinci resim 
var count =0;

var cubeimage=new Array()
for (i=0;i<specifyimage.length;i++){
cubeimage[i]=new Image()
cubeimage[i].src=specifyimage[i]
}
function movecube(){
if (window.createPopup)
cube.filters[0].apply()
document.images.cube.src=cubeimage[count].src;
if (window.createPopup)
cube.filters[0].play()
count++;
if (count==cubeimage.length)
count=0;
setTimeout("movecube()",delay)
}
window.onload=new Function("setTimeout('movecube()',delay)")
</script>
<img src="<%=liste("resim")%>" name="cube"  border=0 style="filter:progid:DXImageTransform.Microsoft.Stretch(stretchStyle='PUSH')" width="118" height="118">