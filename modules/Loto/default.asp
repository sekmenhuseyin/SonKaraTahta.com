<html>
<HEAD>

<title>z-loto-3</title>


<style type="text/css">
.a1{
position:relative;
font-family:Verdana;
font-size:20px;
color:#888888;
}
</style>

<script language="JavaScript">
<!-- UK National Lottery Number Picker by kurt.grigg@virgin.net
function lotto(){
B=' ';	
LottoNumbers=new Array();	
 for (i=1; i <= 6; i++)
 {
 RandomNumber=Math.round(1+Math.random()*48);
  for (j=1; j <= 6; j)
  {
  if (RandomNumber == LottoNumbers[j])
    {
     RandomNumber=Math.round(1+Math.random()*48);
     j=0;
    }
  j++;
  }
 LottoNumbers[i]=RandomNumber;
 }
LottoNumbers=LottoNumbers.toString();
X=LottoNumbers.split(',');
 for (i=0; i < X.length; i++)
 {
 X[i]=X[i]+' ';
 if (X[i].length==2) 
 X[i]='0'+X[i];
 } 
X=X.sort();
 for (i=0; i < X.length; i++)
 {
 OutPut=B+=X[i];
 }
if (document.all)document.all.layer1.innerHTML=OutPut;
if (document.getElementById)document.getElementById("layer1").innerHTML=OutPut;
if (document.layers){
  document.layers.layer1.document.open();
  document.layers.layer1.document.write("<span style='position:absolute;top:0px;left:0px;font-family:Verdana;font-size:20px;color:#888888;text-align:center'>&nbsp;"+OutPut+"</span>");
  document.layers.layer1.document.close();
  }
  T=setTimeout('lotto()',50);
window.status=OutPut;
}
function StOp(){
setTimeout('clearTimeout(T)',3000);
}
//-->
</script>


</HEAD>



<BODY>
<table border='0' width="100%" align="left" style="border-collapse: collapse" bordercolor="#111111" cellpadding="0" cellspacing="0">

<tr valign='middle'>

<td align='center' width="100%">

<form name=form>

<input type=button value='Týklayýn Loto Tahminlerimizi Alýn' onClick="lotto();StOp()">

</form>

<span id=layer1 class=a1>Sonuçlar</span>

</td>

</tr>

</table>
</BODY>
</html>