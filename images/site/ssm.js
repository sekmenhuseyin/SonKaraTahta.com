<!--

/*
Copyright © MaXimuS 2002, All Rights Reserved.
Site: http://maximus.ravecore.com
E-mail: maximusforever@hotmail.com
Script: Static Slide Menu
Version: 6.6 Build 34
*/

NS6=(document.getElementById&&!document.all)
IE=(document.all);IE4=(document.all&&!document.getElementById)
NS=(navigator.appName=="Netscape" && navigator.appVersion.charAt(0)=="4")
OP=(navigator.userAgent.indexOf('Opera')>-1)

tempBar='';barBuilt=0;lastY=0;lastX=0;sI=new Array();moving=setTimeout('null',1);

function moveOut() {
	if(parseInt(ssm.left)<0&&mPos||parseInt(ssm.left)>0&&!mPos){
		clearTimeout(moving);
		moving=setTimeout('moveOut()', slideXSpeed);
		slideMenu((!mPos)?"out":"out");
		}
	else {
		clearTimeout(moving);
		moving=setTimeout('null',1);
	}
}
function moveBack() {
	clearTimeout(moving);
	moving=setTimeout('moveBack1()',waitTime);
}
function moveBack1() {
	if(parseInt(ssm.left)>-(menuWidth+1)&&mPos||parseInt(ssm.left)<menuWidth+1&&!mPos) {
		clearTimeout(moving);
		moving=setTimeout('moveBack1()',slideXSpeed);
		slideMenu((!mPos)?"in":"in");
	}
	else{
		clearTimeout(moving);
		moving=setTimeout('null',1);
	}
}
function slideMenu(way){
	fHow=(NS6)?0.4:0.2;
	if(way=="out")flow=fHow*-(parseInt(ssm.left));
	if(way=="in"&&!mPos)flow=fHow* (menuWidth+1-parseInt(ssm.left));
	else if(way=="in")flow=fHow*-(menuWidth+1+parseInt(ssm.left));
	if(flow>0)flow=Math.ceil(flow);
	else flow=Math.floor(flow);
	if(IE||NS6){
		lastX+=flow;
		bssm.clip="rect(0 "+((!mPos)?(barWidth+menuWidth+3):(barWidth+2+lastX))+" "+(((IE4)?document.body.clientHeight:0)+tssm.offsetHeight)+" "+((!mPos)?(lastX+1):0)+")";
		}
	ssm.left=parseInt(ssm.left)+flow;
	if(NS){
		if(!mPos){
			bssm.clip.left+=flow;
			bssm2.clip.left+=flow;
		}
		else{
			bssm.clip.right+=flow;
			bssm2.clip.right+=flow;
		}
		if(bssm.left+bssm.clip.right>document.width)document.width+=flow;
	}
}

function makeStatic() {
	winY=(IE)?document.body.scrollTop:window.pageYOffset;
	sHow=(NS6)?0.4:0.2;
	if(winY!=lastY&&winY>YOffset-staticYOffset)smooth=sHow*(winY-lastY-YOffset+staticYOffset);
	else if(YOffset-staticYOffset+lastY>YOffset-staticYOffset&&winY<=YOffset-staticYOffset)smooth=sHow*(winY-lastY-(YOffset-(YOffset-winY)));
	else smooth=0;
	if(smooth>0)smooth=Math.ceil(smooth);
	else smooth=Math.floor(smooth);
	bssm.top=parseInt(bssm.top)+smooth;
	lastY=lastY+smooth;
	setTimeout('makeStatic()',slideYSpeed);
}

function menuClick(id) { 
	obj=(document.all)?document.all(id):document.getElementById(id);
	with(obj){
		if(event.srcElement.id!=id){
			if(target=="_top")top.location=href;
			else if(target=="_parent")parent.location=href;
			else if(target=="_blank")window.open(href);
			else if(target>""&&top.frames[target])top.frames[target].location=href;
			else if(target>"")eval('window.open("'+href+'","'+target+'")');
			else location=href;
		}
	}
}

function buildBar() {
	if(!barType)tempBar='<IMG SRC="'+barText+'" BORDER="0">';
	else{
		for(b=0;b<barText.length;b++)tempBar+=barText.charAt(b)+"<br />"
	}
	ssmHTML+='<td align="center" rowspan="100" width="'+barWidth+'" bgcolor="'+barBGColor+'" valign="'+barVAlign+'" align="'+barAlign+'" class="ssmBar" NOWRAP>'+tempBar+'</td>';
}

function initSlide() {
	if (!mPos)lastX=menuWidth
	if ((NS6||IE)&&!OP||(operaFix!=2&&OP)){
		ssm=(NS6)?document.getElementById("thessm").style:document.all("thessm").style;
		tssm=(NS6)?document.getElementById("thessm"):document.all("thessm");
		bssm=(NS6)?document.getElementById("basessm").style:document.all("basessm").style;
		bssm.clip="rect(0 "+(barWidth+2+((!mPos)?menuWidth+1:0))+" "+(((IE4)?document.body.clientHeight:0)+tssm.offsetHeight)+" "+((!mPos)?(menuWidth+1):0)+")";
		if (OP&&operaFix==1)XOff=(!mPos)?document.body.clientWidth-barWidth-3:0;
		bssm.left=(!mPos)?XOff-menuWidth:XOff;
		if(OP)ssm.left=ssm.left;
		bssm.visibility="visible";
		if(NS6&&!OP){
			bssm.top=YOffset;
			if(menuOpacity!=100)ssm.MozOpacity=menuOpacity/100;
			slideIsGo=window.innerHeight>tssm.offsetHeight+staticYOffset;
		}
		else{
			if(menuOpacity!=100)ssm.filter="alpha(opacity="+menuOpacity+")";
			slideIsGo=((OP)?window.innerHeight:document.body.clientHeight)>parseInt(tssm.offsetHeight)+staticYOffset;
			}
		if(autoHideXOverflow&&((IE?document.body.clientWidth:window.innerWidth-16)<parseInt(bssm.left)+parseInt(ssm.left)+menuWidth+barWidth+3)){
			document.body.style.overflowX="hidden";
			document.body.style.overflowY="scroll";
		}
	}
	else if(NS){
		bssm=document.layers["basessm1"];
		bssm2=bssm.document.layers["basessm2"];
		ssm=bssm2.document.layers["thessm"];
		bssm.clip.left=(!mPos)?menuWidth+1:0;
		bssm.clip.right=(!mPos)?(menuWidth+barWidth+3):barWidth+2;
		bssm.left=(!mPos)?XOff-menuWidth:XOff;
		ssm.visibility="show";
		slideIsGo=window.innerHeight>ssm.clip.bottom+staticYOffset;
	}
	if(slideY&&(slideOnYOverflow||(!slideOnYOverflow&&slideIsGo)))makeStatic();
	if(!slideX)moveOut();
}

function getXOff() {
	return (((!XAlign)?((IE||OP)?document.body.clientWidth-barWidth-3:window.innerWidth-barWidth-3-17):(XAlign==1)?Math.floor(((IE||OP)?document.body.clientWidth/2-barWidth/2-1.5:window.innerWidth/2-barWidth/2-1.5)):0)+XOffset);
}

function buildMenu() {
	mPos=menuPosition;
	ssmHTML="";
	XOff=getXOff();
	if(IE||NS6)ssmHTML+='<DIV ID="basessm" style="visibility:hidden;Position : Absolute ;Top : '+YOffset+' ;Z-Index : 20;width:'+(barWidth+2)+';"><DIV ID="thessm" style="Position : Absolute ;Left : '+((!mPos)?menuWidth+1:-menuWidth-1)+' ;Top : 0px ;Z-Index : 30;'+((IE)?"width:1px":"")+'" '+((slideX)?'onmouseover="moveOut()" onmouseout="moveBack()")':'')+'>';
	if(NS)ssmHTML+='<LAYER name="basessm1" top="'+YOffset+'" visibility="show" onload="initSlide()"><ILAYER name="basessm2"><LAYER visibility="hide" name="thessm" bgcolor="'+menuBGColor+'" left="'+((!mPos)?menuWidth+1:-menuWidth-1)+'" '+((slideX)?'onmouseover="moveOut()" onmouseout="moveBack()")':'')+'>';
	if(NS6)ssmHTML+='<table border="0" cellpadding="0" cellspacing="0" width="'+(menuWidth+barWidth+3)+'"><TR><TD>';
	ssmHTML+='<table border="0" cellpadding="0" cellspacing="1" width="'+(menuWidth+barWidth+3)+'" bgcolor="'+((!NS)?menuBGColor:"")+'">';
	for(i=0;i<sI.length;i++){
		ssmHTML+='<TR>';
		if(barBuilt==0&&!mPos){
			buildBar();
			barBuilt=1
		}
		if(sI[i][3]>1)ssmHTML+='<TD BGCOLOR="'+hdrBGColor+'" ALIGN="'+hdrAlign+'" WIDTH="'+menuWidth+'"'+((NS6)?' style="padding:'+hdrPadding+'px"':'><TABLE CELLPADDING="'+hdrPadding+'" CELLSPACING="0" BORDER="0"><TR><TD')+' CLASS="ssmHdr" VALIGN="'+hdrVAlign+'">'+((sI[i][3]==3)?'<a HREF="'+((sI[i][1].indexOf("://")==-1&&sI[i][1].indexOf("../")==-1)?targetDomain:'')+sI[i][1]+'" target="'+sI[i][2]+'" class="ssmHdr">':'')+sI[i][0]+((sI[i][3]==3)?'</a>':'')+((NS6)?'':'</TD></TR></TABLE>')+'</TD>';
		else{
			if(!sI[i][2])sI[i][2]=targetFrame;
			ssmHTML+='<TD WIDTH="'+menuWidth+'"'+(NS&&!sI[i][3]?'':'BGCOLOR="'+linkBGColor+'"')+' '+((NS6)?'CLASS="ssmItem" style="padding:'+linkPadding+'px;" ALIGN="'+linkAlign+'"':'')+' '+((sI[i][3])?'>':'onmouseover="style.backgroundColor=\''+linkOverBGColor+'\'" onmouseout="style.backgroundColor=\''+linkBGColor+'\'" onclick="menuClick(\'item_'+i+'\');"'+(IE?' style="cursor:hand;"':'')+'><ILAYER><LAYER onmouseover="bgColor=\''+linkOverBGColor+'\'" onmouseout="bgColor=\''+linkBGColor+'\'" WIDTH="100%" ALIGN="'+linkAlign+'" bgcolor="'+linkBGColor+'">')+((NS6)?'':'<DIV ALIGN="'+linkAlign+'" CLASS="ssmItem"><TABLE CELLPADDING="'+linkPadding+'" CELLSPACING="0" BORDER="0"><TR><TD VALIGN="'+linkVAlign+'" CLASS="ssmItem">')+((sI[i][3])?'':'<A HREF="'+((sI[i][1].indexOf("://")==-1&&sI[i][1].indexOf("../")==-1)?targetDomain:'')+sI[i][1]+'" target="'+sI[i][2]+'" CLASS="ssmItem" id="item_'+i+'">')+sI[i][0]+''+((sI[i][3])?'':'</A>')+((NS6)?'':'</TD></TR></TABLE></DIV>')+((sI[i][3])?'':'</LAYER></ILAYER>')+'</TD>';
		}
		if(barBuilt==0&&mPos){
			buildBar();
			barBuilt=1;
		}
		ssmHTML+='</TR>';
	}
	ssmHTML+='</table>';
	if(NS6)ssmHTML+='</TD></TR></TABLE>';
	if(IE||NS6){
		ssmHTML+='</DIV></DIV>';
		setTimeout('initSlide();',1);
	}
	if(NS)ssmHTML+='</LAYER></ILAYER></LAYER>';
	document.write(ssmHTML);
}

function addHdr(text){sI[sI.length]=[text, '', '', 2]}

function addLink(text, link, target){if(!link)link="javascript://";sI[sI.length]=[text, link, target, 3]}

function addItem(text, link, target){if(!link)link="javascript://";sI[sI.length]=[text, link, target, 0]}

function addText(text){sI[sI.length]=[text, '', '', 1]}

//window.onresize=function(){setTimeout('alert(getOff());XOff=getXOff();bssm.left=(!mPos)?XOff-menuWidth:XOff;');}

//-->