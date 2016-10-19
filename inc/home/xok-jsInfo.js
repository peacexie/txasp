
// ******** 更多 >>> jsPlugs.js ******** ******** ******** ******** ******** ****

// ******************************************************************************
// 判断浏览器，操作系统相关；
// ******************************************************************************
var brsUAgent = navigator.userAgent.toLowerCase();
brsCheck = function(r){
  return r.test(brsUAgent);
};
isOpera = brsCheck(/opera/);
isChrome = brsCheck(/chrome/);
isWebKit = brsCheck(/webkit/);
isSafari = !isChrome && brsCheck(/safari/);
isIE = !isOpera && brsCheck(/msie/);
isIE7 = isIE && brsCheck(/msie 7/);
isIE8 = isIE && brsCheck(/msie 8/);
isIE9 = isIE && brsCheck(/msie 9/);
isIE6 = isIE && !isIE7 && !isIE8 && !isIE9;
isGecko = !isWebKit && brsCheck(/gecko/); //isFirefox
isFirefox = brsCheck(/firefox/);



// ******************************************************************************
// 基本函数，有无DOCTYPE声明等 
// ******************************************************************************
var getElmObj = document.documentElement;
function isDBody(){
  return (document.compatMode && document.compatMode!="BackCompat")? document.documentElement : document.body;
}
function getElmID(xID){
  return document.getElementById(xID);
}

//得到跟目录 /asp, /u/demo
function getBasePath(xDir,xUrl){
  var sPath = xUrl.replace('http://','');
  var nPos1 = sPath.indexOf('/');
  var nPos2 = sPath.indexOf(xDir);
  var nLen = sPath.length; 
  if(nPos2>nPos1){
	sPath = sPath.substring(nPos1,nPos2); // /asp	   
  }else{
	sPath = ''; //
  }
  return sPath;
}
//得到后缀名 // .asp
function getFileExt(){
  var sHref = document.location.href;
  var nPos3 = sHref.indexOf('?'); 
  if(nPos3>0) { sHref = sHref.substring(0,nPos3); }
  var myExt = '.htm'; //默认
  var nPos4 = sHref.lastIndexOf('.');
  if(nPos4>0) { myExt = sHref.substring(nPos4); }
	myExt = myExt.toLowerCase();
  return myExt;
}

// 加入收藏夹
function setFavorite(url, title){　 
	try {
		window.external.addFavorite(url, title);
	} catch (e){
		try {
			window.sidebar.addPanel(title, url, '');
        	} catch (e) {
			alert("请按 Ctrl+D 键添加到收藏夹", 'notice');
		}
	}
}
// 设置首页
function setHomepage(xUrl){　
  if (document.all){
    document.body.style.behavior = 'url(#default#homepage)';
    document.body.setHomePage(xUrl);
  }else if (window.sidebar){
	if (window.netscape){
	  try {
		netscape.security.PrivilegeManager.enablePrivilege("UniversalXPConnect");
	  }catch (e) {
		alert("该操作被浏览器拒绝，如果想启用该功能，请在地址栏内输入 about:config ,然后将项为signed.applets.codebase_principal_support的值改为true");
	  }
	}
	var prefs = Components.classes['@mozilla.org/preferences-service;1'].getService(Components.interfaces.nsIPrefBranch);
	prefs.setCharPref('browser.startup.homepage', xUrl);
  }
}
//获取地址栏参数
function getUrlPara(key){
  return (new RegExp("([^(&|\?)]*)" + key + "=([^(&|#)]*)").test(location.href+"#")) ? RegExp.$2 : null;
}

function PicReLoad(xPath,xType,xSess){
  if(xType==undefined) { xType=""; } if(xSess==undefined) { xSess=""; }
  var dImg = getElmID("ChkCImg"); 
  var dTyp = ""; if(xType!="") { dTyp="&Config_PCode="+xType; }
  var dSes = ""; if(xSess!="") { dSes="&Config_PSess="+xSess; }
  dImg.setAttribute("src",xPath+"sadm/pcode/img_frnd.asp?xRnd="+Math.random()+dTyp+dSes+"");
}
function goOptUrl(e,f){
  var url=e.options[e.selectedIndex].value;
  if(url!=''){ 
    if(f=="_self") { location.href=url; } 
	else           { window.open(url);  }
  }
}
function InnSearch(xThis){
  var KW = xThis.KW.value;
  var TP = xThis.TP.value; if(TP.length==0) { alert("请选择一个类别"); return false; }
  var Url = TP+'&KeyWD='+KW; 
  xThis.action = Url;
  xThis.submit(); //return false;
  return true;
}
function GstLogin(xTimID){
  var url = "../smod/info/guest_login.asp?act=time&TimID="+xTimID;
  getOutScript(url,"utf-8");
  //window.open(url);
}
function ShowDiv(xID){
 nDiv = getElmID(xID);
 if(nDiv.style.display=='none') { nDiv.style.display = ''; }
 else { nDiv.style.display = 'none'; } 
}   



// ******************************************************************************
// 1. typeof 给firefox定义contains()方法，ie下不起作用
// 2. getEvent 通过循环对比来判断是不是obj的父元素 解决js中onMouseOut事件冒泡的问题；
// 3. setEvent 给每个td增加onmouseover,onclick属性... 4. Demo... 
// ******************************************************************************
if(typeof(HTMLElement)!="undefined")    {   
  HTMLElement.prototype.contains=function(obj)   
  {   
	 while(obj!=null&&typeof(obj.tagName)!="undefind"){ 
　　　  if(obj==this) return true;   
　　　　obj=obj.parentNode;
　　}   
	return false;   
  };   
}  
function getEvent(Event,ID){  //Event用来传入事件，Firefox的方式
  if (Event){
    var topID = document.getElementById(ID);
	if (isFirefox){ //如果是Firefox
      if (topID.contains(Event.relatedTarget)) { //如果是子元素
        return false;   //结束函式
      } 
    }else{
      if (topID.contains(event.toElement)) { //如果是子元素
        return false; //结束函式
      }
	}
  }
  return true;
}
function setEvent(onAct,doAct,ID,tag) {
  var oItems = document.getElementById(ID).getElementsByTagName(tag);
  for(var i = 0;i<oItems.length;i++)
  {
	if(isFirefox||isIE8||isIE9){
	  oItems[i].setAttribute(onAct,doAct+"(this)");
	}else{
	  oItems[i].setAttribute(onAct,function(){eval(doAct+"(this)")}); 
	}
  }
}
//* Demo-js: var MD = "menu01About"; // default
//*          setEvent("onmouseover","muOver","menu01Tags","td");
//* Demo-html: onmouseout="muOutTop(event,'menu01Top')"
function muOver(e) { 
	var mID = e.id.toString();
	//getElmID("idTest").innerHTML = mID;
	if(mID.length>6){ menu01Show(mID); }
	//else { menu01Show("Home"); }
}
function muOut(Event,ID) { 
  if(getEvent(Event,ID)){
    try{setTimeout("menu01Show('"+MD+"')",100);}catch(objError){ menu01Show("Home"); };
  }
}
function muOutTop(Event,ID){ 
  if(getEvent(Event,ID)){
	try{setTimeout("menu01Show('"+MD+"')",300);}catch(objError){ menu01Show("Home"); };
  }
}



// ******************************************************************************
// 新插入一个 外部动态js文件；
// 图片缩放 onload="javascript:ReSizePic(this,150,150);"
// 折叠目录树 
// ******************************************************************************
function getOutScript(xUrl,xSet){
  if(xUrl.indexOf("?")<=0) { xUrl += "?"; }
  else { xUrl += "&"; }
  if(xSet=="") { xSet="utf-8"; }
  try{
    var ds = document.createElement("script");
    ds.type = "text/javascript";
    ds.charset = xSet;
    ds.src = xUrl+"rOut___Random="+Math.random();
    document.getElementsByTagName('head')[0].appendChild(ds);
	//alert(ds.src); 
  }catch(objError){ 
    //getID("vmMsg").textContent = "提交错误！";
	alert("Error(getOutScript)!"+xUrl); 
  }
}

function setImgSize(obj,w,h){
  img = new Image(); img.src = obj.src; 
  zw = img.width; zh = img.height; 
  zr = zw / zh;
  if(w){ fixw = w; }
  else { fixw = obj.getAttribute('width'); }
  if(h){ fixh = h; }
  else { fixh = obj.getAttribute('height'); }
  if(zw > fixw) {
	zw = fixw; zh = zw/zr;
  }
  if(zh > fixh) {
	zh = fixh; zw = zh*zr;
  }
  obj.width = zw; obj.height = zh;
}

function ChkTree(xID){
  var aLays,j; aLays=sLays.split(";"); // sLays
  for(j=0;j<aLays.length;j++){ 
    if ( (aLays[j]!="") && (xID.indexOf(aLays[j])<0) )
	  { getElmID("PeaceM" + aLays[j]).style.display = 'none';  }
  }
  var aID,i; aID = xID.split(";"); 
  for(i=0;i<aID.length;i++){
    if(aID[i]!=""){
	  try{ ////////////////////////////////////
		e = getElmID("PeaceM" + aID[i]); //eval("PeaceM" + aID[i]); 
		if ( (aID.length<=2) || (aID.length>2)&&(i==aID.length-2) ) {
		   if (e.style.display == "none"){ e.style.display = '';    }
		   else                          { e.style.display = 'none'; }
		}else {
		   e.style.display = '';
		}
	  }catch(o){ } //////////////////
    }
  }
}



// ******************************************************************************
// Ajax 相关；
// ******************************************************************************
function getXmlHttp() {
  var xmlHttp = false;
  try {
     xmlHttp = new XMLHttpRequest();
  } catch (trymicrosoft) {
    try {
       xmlHttp = new ActiveXObject("Msxml2.XMLHTTP");
    } catch (othermicrosoft) {
      try {
        xmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
      } catch (failed) {
        xmlHttp = false;
      }
    }
  }
  if (!xmlHttp){
    alert("无法创建 XMLHttpRequest 对象！");
  }
  return xmlHttp;
}
//var xmlHttp = getXmlHttp();
function getXml_Send(xMod) { //getForm
  var url = "pub_ajax.php?yAct=GetForm&tMod="+xMod; 
  xmlHttp.open("GET", url, true); //这里的true代表是异步请求
  xmlHttp.onreadystatechange = getXml_Update;
  xmlHttp.send(null);
}
function getXml_Update(){ //showForm
  if (xmlHttp.readyState == 4) {
    var response = xmlHttp.responseText; 
  }
}