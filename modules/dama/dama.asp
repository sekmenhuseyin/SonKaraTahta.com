
<HTML><HEAD><TITLE>Dama Oyunu</TITLE>
<META content="text/html; charset=windows-1254" http-equiv=Content-Type>
<script>
<!--
/* 
zeus internet  
*/
isNN=document.layers ? 1 : 0; 
function noContext(){return false;}
function noContextKey(e) {
    if(isNN){
        if (e.keyCode == 96){ return (false);}
    } else {
        if (event.keyCode == 96){ return (false);}
    }
}
function noClick(e){
    if(isNN){
        if(e.which > 1) {return false;}
    } else { 
        if(event.button > 1){return false;}
    }
}
if(isNN){ 
    document.captureEvents(Event.MOUSEDOWN);
}
document.oncontextmenu=noContext;
document.onkeypress   =noContextKey;
document.onmousedown  =noClick;
document.onmouseup    =noClick;
// -->
</script>
<SCRIPT language=JavaScript>
<!-- Begin

var pos=new Array(49);
var jumps=new Array();
var boardType="Solitaire";
var numMoves=0;
var finished=false;
var selectnum=false;
var autosolve=false;
var running=false;
var basenum=0;
var destnum=0;
var destnum1=0;
var destnum2=0;
var destnum3=0;
var destnum4=0;
var delaynum=500;

if (document.images)
{
blank=new Image(19,19);
blank.src="modules/dama/blank.gif";
empty=new Image(19,19);
empty.src="modules/dama/empty.gif";
emptysel=new Image(19,19);
emptysel.src="modules/dama/emptysel.gif";
peg=new Image(19,19);
peg.src="modules/dama/peg.gif";
pegact=new Image(19,19);
pegact.src="modules/dama/pegact.gif";
}

function display(pos,basenum,destnum)
{
selectnum=false;
if (!basenum && !destnum)
{
for (var i=0; i<pos.length; i++)
{
if (pos[i]==-1) document.images["img"+i].src=blank.src;
else if (pos[i]==1) document.images["img"+i].src=peg.src;
else document.images["img"+i].src=empty.src;
}
}
else
{
document.images["img"+basenum].src=empty.src;
document.images["img"+(basenum+destnum/2)].src=empty.src;
document.images["img"+(basenum+destnum)].src=peg.src;
for (var i=0; i<pos.length; i++)
{
if (document.images["img"+i].src==emptysel.src)
document.images["img"+i].src=empty.src;
}
}
if (numMoves>1) win();
}

function move(num)
{
var curNumMoves=numMoves;
if (!document.images)
alert("Your browser does not support the 'document.images' property.  You\n" +
"should upgrade to at least Netscape 3.0 or Internet explorer 4.0.");
else if (autosolve && running) {}
else if (autosolve && !finished)
{
if (confirm('You interrupted the \'Solve\' function. Want to try it yourself?'))
newGame();
}
else if (selectnum)
{
if (num!=basenum && num!=basenum+destnum1 && num!=basenum+destnum2 &&
num!=basenum+destnum3 && num!=basenum+destnum4)
alert("Select a destination or click on the original peg again!");
else if (num==basenum)
{
document.images["img"+basenum].src=peg.src;
if (destnum1!=0)
document.images["img"+(basenum+destnum1)].src=empty.src;
if (destnum2!=0)
document.images["img"+(basenum+destnum2)].src=empty.src;
if (destnum3!=0)
document.images["img"+(basenum+destnum3)].src=empty.src;
if (destnum4!=0)
document.images["img"+(basenum+destnum4)].src=empty.src;
selectnum=false;
}
else if (num==basenum+destnum1) movePeg(basenum,destnum1)
else if (num==basenum+destnum2) movePeg(basenum,destnum2)
else if (num==basenum+destnum3) movePeg(basenum,destnum3)
else if (num==basenum+destnum4) movePeg(basenum,destnum4)
}
else if (pos[num]==0)
{
}
else if ((num==3 || num==10) && pos[num+7]==1 && pos[num+14]==0) movePeg(num,14);
else if ((num==45 || num==38) && pos[num-7]==1 && pos[num-14]==0) movePeg(num,-14);
else if ((num==21 || num==22) && pos[num+1]==1 && pos[num+2]==0) movePeg(num,2);
else if ((num==26 || num==27) && pos[num-1]==1 && pos[num-2]==0) movePeg(num,-2);
else if (num==4 || num==11 || num==19 || num==20)
{
if (pos[num-1]==1 && pos[num-2]==0 && pos[num+7]==1 && pos[num+14]==0)
selPeg(num,-2,14);
else if (pos[num-1]==1 && pos[num-2]==0) movePeg(num,-2);
else if (pos[num+7]==1 && pos[num+14]==0) movePeg(num,14);
}
else if (num==2 || num==9 || num==14 || num==15)
{
if (pos[num+1]==1 && pos[num+2]==0 && pos[num+7]==1 && pos[num+14]==0)
selPeg(num,2,14);
else if (pos[num+1]==1 && pos[num+2]==0) movePeg(num,2);
else if (pos[num+7]==1 && pos[num+14]==0) movePeg(num,14);
}
else if (num==28 || num==29 || num==37 || num==44)
{
if (pos[num+1]==1 && pos[num+2]==0 && pos[num-7]==1 && pos[num-14]==0)
selPeg(num,2,-14);
else if (pos[num+1]==1 && pos[num+2]==0) movePeg(num,2);
else if (pos[num-7]==1 && pos[num-14]==0) movePeg(num,-14);
}
else if (num==33 || num==34 || num==39 || num==46)
{
if (pos[num-1]==1 && pos[num-2]==0 && pos[num-7]==1 && pos[num-14]==0)
selPeg(num,-2,-14);
else if (pos[num-1]==1 && pos[num-2]==0) movePeg(num,-2);
else if (pos[num-7]==1 && pos[num-14]==0) movePeg(num,-14);
}
else if (num==16 || num==17 || num==18 || num==23 || num==24 || num==25 ||
num==30 || num==31 || num==32)
{
var cond1=(pos[num-1]==1 && pos[num-2]==0);
var cond2=(pos[num-7]==1 && pos[num-14]==0);
var cond3=(pos[num+1]==1 && pos[num+2]==0);
var cond4=(pos[num+7]==1 && pos[num+14]==0);
if ((cond1 && (cond2 || cond3 || cond4)) ||
(cond2 && (cond1 || cond3 || cond4)) ||
(cond3 && (cond1 || cond2 || cond4)))
{
basenum=num;
destnum1=destnum2=destnum3=destnum4=0;
document.images["img"+basenum].src=pegact.src;
if (cond1)
{
destnum1=-2;
document.images["img"+(basenum+destnum1)].src=emptysel.src;
}
if (cond2)
{
destnum2=-14;
document.images["img"+(basenum+destnum2)].src=emptysel.src;
}
if (cond3)
{
destnum3=2;
document.images["img"+(basenum+destnum3)].src=emptysel.src;
}
if (cond4)
{
destnum4=14;
document.images["img"+(basenum+destnum4)].src=emptysel.src;
}
selectnum=true;
}
else if (cond1) movePeg(num,-2);
else if (cond2) movePeg(num,-14);
else if (cond3) movePeg(num,2);
else if (cond4) movePeg(num,14);
}
if (curNumMoves!=numMoves) display(pos,basenum,destnum);
else if (finished) win();
}

function selPeg(num,ofset1,ofset2)
{
basenum=num;
destnum1=ofset1;
destnum2=ofset2;
destnum3=destnum4=0;
document.images["img"+basenum].src=pegact.src;
document.images["img"+(basenum+destnum1)].src=emptysel.src;
document.images["img"+(basenum+destnum2)].src=emptysel.src;
selectnum=true;
}

function movePeg(num,ofset)
{
pos[num+ofset]=1;
pos[num+ofset/2]=pos[num]=0
basenum=num;
destnum=ofset;
numMoves++;
}

function win()
{
var cnt=0;
for(var i=0; i<pos.length; i++)
{
if (pos[i]!=-1) cnt+=pos[i];
}
if (cnt==1 && autosolve)
{
if (confirm(' �imdi S�ra Sende....'))
newGame();
}
else if (cnt==1 && pos[24]==1)
{
finished=true;
if (confirm('oyun bitti tekrar oynayacakm�s�n?')) newGame();
}
else if (cnt==1)
{
finished=true;
if (confirm('You did it! Do you want to restart?')) newGame();
}
else
{
var legalMoves=false;
var num=0;
while (num<pos.length && !legalMoves)
{
if (pos[num]==1 &&
(((num==2 || num==9 || num==14 || num==15 || num==16 || num==17 ||
num==18 || num==23 || num==24 || num==25 || num==30 || num==31 ||
num==32 || num==21 || num==22 || num==28 || num==29 || num==37 ||
num==44) && pos[num+1]==1 && pos[num+2]==0) ||
((num==4 || num==11 || num==19 || num==20 || num==16 || num==17 ||
num==18 || num==23 || num==24 || num==25 || num==30 || num==31 ||
num==32 || num==26 || num==27 || num==33 || num==34 || num==39 ||
num==46) && pos[num-1]==1 && pos[num-2]==0) ||
((num==2 || num==9 || num==14 || num==15 || num==16 || num==17 ||
num==18 || num==23 || num==24 || num==25 || num==30 || num==31 ||
num==32 || num==4 || num==11 || num==19 || num==20 || num==3 ||
num==10) && pos[num+7]==1 && pos[num+14]==0) ||
((num==33 || num==34 || num==39 || num==46 || num==16 || num==17 ||
num==18 || num==23 || num==24 || num==25 || num==30 || num==31 ||
num==32 || num==45 || num==38 || num==28 || num==29 || num==37 ||
num==44) && pos[num-7]==1 && pos[num-14]==0)))
legalMoves=true;
num++;
}
if (!legalMoves)
{
finished=true;
if (confirm('Oyun bitti tekrar oynayacakm�s�n?')) newGame();
}
}
}

function newGame()
{
if (autosolve && running) {}
else if (document.images)
{
autosolve=false;
finished=false;
if (boardType=="Cross")
{
for (var i=0; i<pos.length; i++) pos[i]=0;
pos[0]=pos[1]=pos[5]=pos[6]=-1;
pos[7]=pos[8]=pos[12]=pos[13]=-1;
pos[10]=pos[16]=pos[17]=pos[18]=pos[24]=pos[31]=1;
pos[35]=pos[36]=pos[40]=pos[41]=-1;
pos[42]=pos[43]=pos[47]=pos[48]=-1;
}
else if (boardType=="Plus")
{
for (var i=0; i<pos.length; i++) pos[i]=0;
pos[0]=pos[1]=pos[5]=pos[6]=-1;
pos[7]=pos[8]=pos[12]=pos[13]=-1;
pos[10]=pos[17]=pos[22]=pos[23]=pos[24]=1;
pos[25]=pos[26]=pos[31]=pos[38]=1;
pos[35]=pos[36]=pos[40]=pos[41]=-1;
pos[42]=pos[43]=pos[47]=pos[48]=-1;
}
else if (boardType=="Fireplace")
{
for (var i=0; i<pos.length; i++) pos[i]=0;
pos[0]=pos[1]=pos[5]=pos[6]=-1;
pos[7]=pos[8]=pos[12]=pos[13]=-1;
pos[2]=pos[3]=pos[4]=pos[9]=pos[10]=1;
pos[11]=pos[16]=pos[17]=pos[18]=1;
pos[23]=pos[25]=1;
pos[35]=pos[36]=pos[40]=pos[41]=-1;
pos[42]=pos[43]=pos[47]=pos[48]=-1;
}
else if (boardType=="Up Arrow")
{
for (var i=0; i<pos.length; i++) pos[i]=0;
pos[0]=pos[1]=pos[5]=pos[6]=-1;
pos[7]=pos[8]=pos[12]=pos[13]=-1;
pos[3]=pos[9]=pos[10]=pos[11]=pos[15]=1;
pos[16]=pos[17]=pos[18]=pos[19]=1;
pos[24]=pos[31]=pos[37]=pos[38]=1;
pos[39]=pos[44]=pos[45]=pos[46]=1;
pos[35]=pos[36]=pos[40]=pos[41]=-1;
pos[42]=pos[43]=pos[47]=pos[48]=-1;
}
else if (boardType=="Pyramid")
{
for (var i=0; i<pos.length; i++) pos[i]=0;
pos[0]=pos[1]=pos[5]=pos[6]=-1;
pos[7]=pos[8]=pos[12]=pos[13]=-1;
pos[10]=pos[16]=pos[17]=pos[18]=pos[22]=1;
pos[23]=pos[24]=pos[25]=pos[26]=1;
pos[28]=pos[29]=pos[30]=pos[31]=1;
pos[32]=pos[33]=pos[34]=1;
pos[35]=pos[36]=pos[40]=pos[41]=-1;
pos[42]=pos[43]=pos[47]=pos[48]=-1;
}
else if (boardType=="Diamond")
{
for (var i=0; i<pos.length; i++) pos[i]=1;
pos[0]=pos[1]=pos[5]=pos[6]=-1;
pos[7]=pos[8]=pos[12]=pos[13]=-1;
pos[2]=pos[4]=pos[14]=pos[20]=pos[24]=0;
pos[28]=pos[34]=pos[44]=pos[46]=0;
pos[35]=pos[36]=pos[40]=pos[41]=-1;
pos[42]=pos[43]=pos[47]=pos[48]=-1;
}
else if (boardType=="Solitaire")
{
for (var i=0; i<pos.length; i++) pos[i]=1;
pos[0]=pos[1]=pos[5]=pos[6]=-1;
pos[7]=pos[8]=pos[12]=pos[13]=-1;
pos[24]=0;
pos[35]=pos[36]=pos[40]=pos[41]=-1;
pos[42]=pos[43]=pos[47]=pos[48]=-1;
}
numMoves=0;
running=true;
changeBoard();
running=false;
solveArray();
display(pos);
}
else
alert("Your browser does not support the 'document.images' property.  You\n" +
"should upgrade to at least Netscape 3.0 or Internet explorer 4.0.");
}

function initArray()
{
this.length=initArray.arguments.length;
for (var i=0; i<this.length; i++)
{
this[i]=initArray.arguments[i]
}
}

function drawPreview(start,end)
{
i=start; j=end;
baseref=jumps[start];
offset=jumps[start+1];
pos[baseref]=pos[baseref+offset/2]=0;
pos[baseref+offset]=1;
document.images["img"+baseref].src=pegact.src;
document.images["img"+(baseref+offset)].src=emptysel.src;
solveRunning=setTimeout('drawJump(i,j)',delaynum);
}

function drawJump(start,end)
{
i=start; j=end;
baseref=jumps[start];
offset=jumps[start+1];
document.images["img"+baseref].src=empty.src;
document.images["img"+(baseref+offset/2)].src=empty.src;
document.images["img"+(baseref+offset)].src=peg.src;
if (start+2==end)
{
document.buttonsform.solve.value="    Tamam    ";
running=false;
finished=true;
setTimeout('win()',delaynum);
}
else solveRunning=setTimeout('drawPreview(i+2,j)',delaynum);
}

function solve()
{
if (!document.images)
alert("Your browser does not support the 'document.images' property.  You\n" +
"should upgrade to at least Netscape 3.0 or Internet explorer 4.0.");
else if (autosolve && running)
{
clearTimeout(solveRunning);
document.buttonsform.solve.value="  Otomatik oyna  ";
running=false;
}
else
{
document.buttonsform.solve.value="     Dur    ";
newGame();
autosolve=true;
running=true;
solveRunning=setTimeout('drawPreview(0,jumps.length)',delaynum);
}
}

function changeBoard()
{
formName=document.buttonsform;
if (!running)
{
boardType=formName.options[formName.options.selectedIndex].value;
newGame();
}
else
{
optlength=formName.options.length;
for (var m=0; m<optlength; m++)
{
if (formName.options[m].value==boardType)
{
formName.options.selectedIndex=m;
break;
}
}
}
}

function solveArray()
{
if (boardType=="Cross")
{
jumps=new initArray(17,-2,31,-14,18,-2,15,2,10,14);
}
else if (boardType=="Plus")
{
jumps=new initArray(23,-2,25,-2,10,14,24,-2,21,2,
38,-14,23,2,26,-2);
}
else if (boardType=="Fireplace")
{
jumps=new initArray(17,2,4,14,25,-14,2,2,4,14,
19,-2,10,14,24,-2,9,14,22,2);
}
else if (boardType=="Up Arrow")
{
jumps=new initArray(46,-14,31,2,45,-14,44,-14,30,2,33,-2,
18,-14,4,-2,16,2,2,14,15,2,18,-2,31,
-14,16,2,19,-2,10,14);
}
else if (boardType=="Pyramid")
{
jumps=new initArray(23,14,25,14,28,2,34,-2,37,-14,39,-14,
16,14,18,-2,31,-2,29,-14,15,2,17,14,
26,-2,31,-14,10,14);
}
else if (boardType=="Diamond")
{
jumps=new initArray(30,14,44,2,32,2,34,-14,18,-14,4,-2,
16,-2,14,14,46,-14,20,-2,2,14,28,2,
38,-14,17,-2,15,14,29,2,31,2,33,-14,
19,-2,24,-2,10,14,25,-2,22,2);
}
else if (boardType=="Solitaire")
{
jumps=new initArray(38,-14,33,-2,46,-14,25,14,44,2,46,-14,
11,14,20,-2,17,2,34,-14,20,-2,
15,2,2,14,23,-14,4,-2,2,14,
37,-14,28,2,31,-2,14,14,28,2,
17,-2,15,14,29,2,31,2,33,-14,19,-2,
24,-2,10,14,25,-2,22,2);
}
}

// End -->
</SCRIPT>
<!------------------------------------------########## End of Script Part 1 ############-------------------------------------------><!------------------------------------------########## Script Part 2 ###################Set the OnLoad event to "window.newGame()" as shown below------------------------------------------->
<META content="Microsoft FrontPage 5.0" name=GENERATOR></HEAD>
<BODY bgColor=#0099FF link=#800000 onload=window.newGame() text=#800000 vLink=#800000 alink="#800080">
<SCRIPT language=JavaScript>
<!-- Begin

document.write(
"<CENTER>\n"+
"<TABLE WIDTH=\"80%\" HEIGHT=\"80%\" CELLPADDING=\"0\" CELLSPACING=\"0\" BORDER=\"0\">\n"+
"<TR><TD VALIGN=\"MIDDLE\" ALIGN=\"CENTER\">\n"+
"<TABLE BGCOLOR=\"#FFCC66\" BORDER=\"1\">\n"+
"<TR><TD ALIGN=\"CENTER\">\n"+
"<H2 ALIGN=\"CENTER\"><FONT FACE=\"Verdana, Arial, Helvetica\" COLOR=\"#000080\">\n"+
"Dama Oyunu\n"+
"</H2></FONT>\n"+
"<P>\n"+
"<TABLE BORDER=\"1\" BGCOLOR=\"#007777\" CELLPADDING=\"15\" CELLSPACING=\"0\">\n"+
"<TR><TD ALIGN=\"CENTER\">");

if (navigator.appName != "Microsoft Internet Explorer")
{
document.write(
"<IMG SRC=\"modules\dama\blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img0\">\n"+
"<IMG SRC=\"modules\dama\blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img1\">\n"+
"<A HREF=\"#\" onClick=\"window.move(2);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"modules\dama\peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img2\"></A>\n"+
"<A HREF=\"#\" onClick=\"window.move(3);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"modules\dama\peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img3\"></A>\n"+
"<A HREF=\"#\" onClick=\"window.move(4);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"modules\dama\peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img4\"></A>\n"+
"<IMG SRC=\"modules\dama\blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img5\">\n"+
"<IMG SRC=\"modules\dama\blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img6\"><br />\n"+
"<IMG SRC=\"modules\dama\blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img7\">\n"+
"<IMG SRC=\"modules\dama\blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img8\">\n"+
"<A HREF=\"#\" onClick=\"window.move(9);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"modules\dama\peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img9\"></A>\n"+
"<A HREF=\"#\" onClick=\"window.move(10);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"modules\dama\peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img10\"></A>\n"+
"<A HREF=\"#\" onClick=\"window.move(11);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"modules\dama\peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img11\"></A>\n"+
"<IMG SRC=\"modules\dama\blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img12\">\n"+
"<IMG SRC=\"modules\dama\blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img13\"><br />\n"+
"<A HREF=\"#\" onClick=\"window.move(14);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"modules\dama\peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img14\"></A>\n"+
"<A HREF=\"#\" onClick=\"window.move(15);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"modules\dama\peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img15\"></A>\n"+
"<A HREF=\"#\" onClick=\"window.move(16);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"modules\dama\peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img16\"></A>\n"+
"<A HREF=\"#\" onClick=\"window.move(17);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"modules\dama\peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img17\"></A>\n"+
"<A HREF=\"#\" onClick=\"window.move(18);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"modules\dama\peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img18\"></A>\n"+
"<A HREF=\"#\" onClick=\"window.move(19);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img19\"></A>\n"+
"<A HREF=\"#\" onClick=\"window.move(20);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img20\"></A><br />\n"+
"<A HREF=\"#\" onClick=\"window.move(21);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img21\"></A>\n"+
"<A HREF=\"#\" onClick=\"window.move(22);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img22\"></A>\n"+
"<A HREF=\"#\" onClick=\"window.move(23);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img23\"></A>\n"+
"<A HREF=\"#\" onClick=\"window.move(24);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"empty.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img24\"></A>\n"+
"<A HREF=\"#\" onClick=\"window.move(25);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img25\"></A>\n"+
"<A HREF=\"#\" onClick=\"window.move(26);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img26\"></A>\n"+
"<A HREF=\"#\" onClick=\"window.move(27);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img27\"></A><br />\n"+
"<A HREF=\"#\" onClick=\"window.move(28);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img28\"></A>\n"+
"<A HREF=\"#\" onClick=\"window.move(29);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img29\"></A>\n"+
"<A HREF=\"#\" onClick=\"window.move(30);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img30\"></A>\n"+
"<A HREF=\"#\" onClick=\"window.move(31);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img31\"></A>\n"+
"<A HREF=\"#\" onClick=\"window.move(32);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img32\"></A>\n"+
"<A HREF=\"#\" onClick=\"window.move(33);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img33\"></A>\n"+
"<A HREF=\"#\" onClick=\"window.move(34);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img34\"></A><br />\n"+
"<IMG SRC=\"blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img35\">\n"+
"<IMG SRC=\"blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img36\">\n"+
"<A HREF=\"#\" onClick=\"window.move(37);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img37\"></A>\n"+
"<A HREF=\"#\" onClick=\"window.move(38);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img38\"></A>\n"+
"<A HREF=\"#\" onClick=\"window.move(39);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img39\"></A>\n"+
"<IMG SRC=\"blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img40\">\n"+
"<IMG SRC=\"blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img41\"><br />\n"+
"<IMG SRC=\"blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img42\">\n"+
"<IMG SRC=\"blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img43\">\n"+
"<A HREF=\"#\" onClick=\"window.move(44);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img44\"></A>\n"+
"<A HREF=\"#\" onClick=\"window.move(45);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img45\"></A>\n"+
"<A HREF=\"#\" onClick=\"window.move(46);return false\" onMouseOver=\"window.status='';\n"+
"return true\"><IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img46\"></A>\n"+
"<IMG SRC=\"blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img47\">\n"+
"<IMG SRC=\"blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img48\"><br />")
}
else
{
document.write(
"<IMG SRC=\"blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img0\">\n"+
"<IMG SRC=\"blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img1\">\n"+
"<IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img2\" \n"+
"onClick=\"window.move(2);return false\">\n"+
"<IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img3\" \n"+
"onClick=\"window.move(3);return false\">\n"+
"<IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img4\" \n"+
"onClick=\"window.move(4);return false\">\n"+
"<IMG SRC=\"blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img5\">\n"+
"<IMG SRC=\"blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img6\"><br />\n"+
"<IMG SRC=\"blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img7\">\n"+
"<IMG SRC=\"blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img8\">\n"+
"<IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img9\" \n"+
"onClick=\"window.move(9);return false\">\n"+
"<IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img10\" \n"+
"onClick=\"window.move(10);return false\">\n"+
"<IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img11\" \n"+
"onClick=\"window.move(11);return false\">\n"+
"<IMG SRC=\"blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img12\">\n"+
"<IMG SRC=\"blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img13\"><br />\n"+
"<IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img14\" \n"+
"onClick=\"window.move(14);return false\">\n"+
"<IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img15\" \n"+
"onClick=\"window.move(15);return false\">\n"+
"<IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img16\" \n"+
"onClick=\"window.move(16);return false\">\n"+
"<IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img17\" \n"+
"onClick=\"window.move(17);return false\">\n"+
"<IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img18\" \n"+
"onClick=\"window.move(18);return false\">\n"+
"<IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img19\" \n"+
"onClick=\"window.move(19);return false\">\n"+
"<IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img20\" \n"+
"onClick=\"window.move(20);return false\"><br />\n"+
"<IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img21\" \n"+
"onClick=\"window.move(21);return false\">\n"+
"<IMG SRC=\"peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img22\" \n"+
"onClick=\"window.move(22);return false\">\n"+
"<IMG SRC=\"\modules\dama\peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img23\" \n"+
"onClick=\"window.move(23);return false\">\n"+
"<IMG SRC=\"empty.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img24\" \n"+
"onClick=\"window.move(24);return false\">\n"+
"<IMG SRC=\"\modules\dama\peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img25\" \n"+
"onClick=\"\modules\dama\window.move(25);return false\">\n"+
"<IMG SRC=\"\modules\dama\peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img26\" \n"+
"onClick=\"\modules\dama\window.move(26);return false\">\n"+
"<IMG SRC=\"\modules\dama\peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img27\" \n"+
"onClick=\"\modules\dama\window.move(27);return false\"><br />\n"+
"<IMG SRC=\"\modules\dama\peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img28\" \n"+
"onClick=\"window.move(28);return false\">\n"+
"<IMG SRC=\"\modules\dama\peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img29\" \n"+
"onClick=\"window.move(29);return false\">\n"+
"<IMG SRC=\"\modules\dama\peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img30\" \n"+
"onClick=\"window.move(30);return false\">\n"+
"<IMG SRC=\"\modules\dama\peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img31\" \n"+
"onClick=\"window.move(31);return false\">\n"+
"<IMG SRC=\"\modules\dama\peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img32\" \n"+
"onClick=\"window.move(32);return false\">\n"+
"<IMG SRC=\"\modules\dama\peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img33\" \n"+
"onClick=\"window.move(33);return false\">\n"+
"<IMG SRC=\"\modules\dama\peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img34\" \n"+
"onClick=\"window.move(34);return false\"><br />\n"+
"<IMG SRC=\"\modules\dama\blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img35\">\n"+
"<IMG SRC=\"\modules\dama\blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img36\">\n"+
"<IMG SRC=\"\modules\dama\peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img37\" \n"+
"onClick=\"window.move(37);return false\">\n"+
"<IMG SRC=\"\modules\dama\peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img38\" \n"+
"onClick=\"window.move(38);return false\">\n"+
"<IMG SRC=\"\modules\dama\peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img39\" \n"+
"onClick=\"window.move(39);return false\">\n"+
"<IMG SRC=\"\modules\dama\blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img40\">\n"+
"<IMG SRC=\"\modules\dama\blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img41\"><br />\n"+
"<IMG SRC=\"\modules\dama\blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img42\">\n"+
"<IMG SRC=\"\modules\dama\blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img43\">\n"+
"<IMG SRC=\"\modules\dama\peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img44\" \n"+
"onClick=\"window.move(44);return false\">\n"+
"<IMG SRC=\"\modules\dama\peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img45\" \n"+
"onClick=\"window.move(45);return false\">\n"+
"<IMG SRC=\"\modules\dama\peg.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img46\" \n"+
"onClick=\"window.move(46);return false\">\n"+
"<IMG SRC=\"\modules\dama\blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img47\">\n"+
"<IMG SRC=\"\modules\dama\blank.gif\" BORDER=\"0\" WIDTH=\"19\" HEIGHT=\"19\" NAME=\"img48\"><br />\n");
}

document.write(
"</TD></TR>\n"+
"</TABLE>\n"+
"<P>\n"+
"<FORM NAME=\"buttonsform\">\n"+
"<INPUT TYPE=\"button\" NAME=\"new\" VALUE=\"Yeni Oyun\" onClick=\"window.newGame()\">\n"+
"<INPUT TYPE=\"button\" NAME=\"solve\" VALUE=\"  Otomatik Oyna  \" onClick=\"window.solve()\">\n"+
"<SELECT NAME=\"options\" onChange=\"(window.changeBoard())\">\n"+
"<OPTION VALUE=\"Cross\">Ha&#231;\n"+
"<OPTION VALUE=\"Plus\">Plus\n"+
"<OPTION VALUE=\"Fireplace\">��mine\n"+
"<OPTION VALUE=\"Up Arrow\">Yukar� Ok\n"+
"<OPTION VALUE=\"Pyramid\">Piramit\n"+
"<OPTION VALUE=\"Diamond\">Baklava\n"+
"<OPTION SELECTED VALUE=\"Solitaire\">Tek Bo�luklu\n"+
"</SELECT>\n"+
"</FORM>\n"+
"<P>\n"+
"<TABLE WIDTH=\"600\" BORDER=\"0\">\n"+
"<TR><TD>\n"+
"<FONT FACE=\"Verdana, Arial, Helvetica\" SIZE=\"-1\" COLOR=\"#000080\">\n"+
"<P ALIGN=\"JUSTIFY\">\n"+

"</FONT>\n"+
"</TD></TR>\n"+
"</TABLE>\n"+
"</TD></TR>\n"+
"</TABLE>\n"+
"</TD></TR>\n"+
"</TABLE>\n"+
"</CENTER>");

// End -->
</SCRIPT>
<NOSCRIPT>
<TABLE border=0 cellPadding=0 cellSpacing=0 height="100%" width="100%">
  <TBODY>
  <TR>
    <TD align=middle vAlign=center>
      <TABLE bgColor=#ffffbb border=1 cellPadding=15 cellSpacing=0>
        <TBODY>
        <TR>
          <TD align=middle><FONT color=#000080 
            face="Verdana, Arial, Helvetica" 
      size=-1><B></B></FONT></TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></NOSCRIPT>
<CENTER>
<p></p>
</CENTER></BODY></HTML>