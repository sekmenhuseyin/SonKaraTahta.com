var xmlHttp
function check_username(str){
if (str.length==0){return}
xmlHttp=GetXmlHttpObject()
if (xmlHttp==null){return} 
var url="inc/ajax_username.asp?q="+str
url=url+"&psid="+Math.random()
xmlHttp.onreadystatechange=stateChangedUserName
xmlHttp.open("GET",url,true)
xmlHttp.setRequestHeader("Content-type","application/x-www-form-urlencoded;charset=iso-8859-9");
xmlHttp.setRequestHeader("Content-type","application/x-www-form-urlencoded;language=tr");
xmlHttp.send(null)
} 
function stateChangedUserName(){
if(xmlHttp.readyState==4||xmlHttp.readyState=="complete"){document.getElementById("hint_nick").innerHTML=xmlHttp.responseText }
} 
function GetXmlHttpObject(){ 
var objXMLHttp=null
if(window.XMLHttpRequest){objXMLHttp=new XMLHttpRequest()}
else if(window.ActiveXObject){objXMLHttp=new ActiveXObject("Microsoft.XMLHTTP")}
return objXMLHttp
} 