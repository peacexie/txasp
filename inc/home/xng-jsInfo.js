
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
function Del_YN(YNaddr,msg)
{
    if(confirm(msg))
     {
         location.href = YNaddr;
         return true;
     }
         return false;
}
function Dir_Addr(Direct)
{
    location.href = Direct; //YNaddr;
}

function PicReLoad(xPath,xType,xSess){
  if(xType==undefined) { xType=""; } if(xSess==undefined) { xSess=""; }
  var dImg = getElmID("ChkCImg"); 
  var dTyp = ""; if(xType!="") { dTyp="&cfgPCode="+xType; }
  var dSes = ""; if(xSess!="") { dSes="&cfgPSess="+xSess; }
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
	//if(isIE){ //||isIE8||isIE9 //isFirefox
	if(isIE6||isIE7){ 
	  oItems[i].setAttribute(onAct,function(){eval(doAct+"(this)")}); //IE6
	}else{
	  oItems[i].setAttribute(onAct,doAct+"(this)"); //IE8,isFirefox
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
  var url = "pub_ajax.asp?yAct=GetForm&tMod="+xMod; 
  xmlHttp.open("GET", url, true); //这里的true代表是异步请求
  xmlHttp.onreadystatechange = getXml_Update;
  xmlHttp.send(null);
}
function getXml_Update(){ //showForm
  if (xmlHttp.readyState == 4) {
    var response = xmlHttp.responseText; 
  }
}



// ******************************************************************************
// 表单 相关；
// ******************************************************************************

function ySel(){
   var sAct = document.getElementById("yAct");
   var sVal = document.getElementById("yVal");
   var vFlag = yFlag.innerText;
   if (vFlag=="N"){
     yFlag.innerText = "Y";
     for(var i=0;i<document.flist.yID.length;i++)
       {document.flist.yID.item(i).checked=true;}
   }else{
     yFlag.innerText = "N";
     for(var i=0;i<document.flist.yID.length;i++)
       {document.flist.yID.item(i).checked=false;}
   }
}  

function chkF_Mail(MObj,msg)
{
    var Mstr,Mi,Mj,Mk,Mkk,Mjj,Mlen;
    Mstr = MObj.value;
    Mat  = Mstr.indexOf( "@" );
    Mdot = Mstr.indexOf( ".", Mi );
    Md2  = Mstr.indexOf( "," );
    Mblk = Mstr.indexOf( " " );
    Mext = Mstr.lastIndexOf( "." ) + 1;
    Mlen = Mstr.length;
    if ( (Mat<=0)||(Mdot<=2)||(Md2!=-1)||(Mblk!=-1)||(Mlen-Mext<2)||(Mlen-Mext>3) )
	{   jsFlag = "ER";
	}else{
		jsFlag = "OK";};
    if( (jsFlag=="ER")&&(msg!="") )
	{  
	    alert(msg);
		MObj.focus();
	}
	return jsFlag;
}

function chkF_Date(MMObj,msg)
{  
	var MYY,MMM,MDD,MObj ;
	jsFlag = "OK";
	MObj = MMObj.value;
	MObj = MObj.replace("/", "-"); 
	MObj = MObj.replace(".", "-"); 
	strArr=MObj.split("-");
    YY = strArr[0]; //MObj.substr(0,4); //alert(YY);
	MM = strArr[1]; //MObj.substr(5,2); //alert(MM);
	DD = strArr[2]; //MObj.substr(8,2); //alert(DD);
	//if ( !(MObj.length==10)) { jsFlag = "ER"; }
	if ( isNaN(YY) ) { jsFlag = "ER"; 
	}else{
	   if ( (YY<1000) || (YY>9999) ) { jsFlag = "ER";}   }
	if ( isNaN(MM) ) { jsFlag = "ER"; 
	}else{
	   if ( (MM<1) || (MM>12) ) { jsFlag = "ER";}   }
	if ( isNaN(DD) ) { jsFlag = "ER"; 
	}else{
	   if ( (DD<1) || (DD>31) ) { jsFlag = "ER";}   }
	if (MM==2) {
	  if (  ((0 == YY % 4) && (0 != (YY % 100))) || (0 == YY % 400)  ) {
		  if (DD>29) {  jsFlag = "ER";}
	  }else{
		  if (DD>28) {  jsFlag = "ER";}
	  }    
	}else{
	  if (  (MM==4)||(MM==6)||(MM==9)||(MM==11)  ) {
		  if (DD>30) {  jsFlag = "ER";}
	  }else{
		  if (DD>31) {  jsFlag = "ER";}
	  }    
	} 
    if( (jsFlag=="ER")&&(msg!="") )
	{  
	    alert(msg);
		MMObj.focus();
	}
	return jsFlag;
}

function chkF_DCmp(DO1,DO2,msg)
{  
    var dt1,dt2,a1,a2,y1,y2,m1,m2,d1,d2;
	jsFlag = "OK";
	dt1 = DO1.value; 
	dt2 = DO2.value;
	dt1 = dt1.replace("/", "-"); 
	dt1 = dt1.replace(".", "-"); 	
	dt2 = dt2.replace("/", "-"); 
	dt2 = dt2.replace(".", "-"); 
	a1 = dt1.split("-");
	a2 = dt2.split("-");
    y1 = a1[0]; 
	m1 = a1[1]; if (m1.length==1)  {  m1 = "0" + m1; }
	d1 = a1[2]; if (d1.length==1)  {  d1 = "0" + d1; }
    y2 = a2[0]; 
	m2 = a2[1]; if (m2.length==1)  {  m2 = "0" + m2; }
	d2 = a2[2]; if (d2.length==1)  {  d2 = "0" + d2; }
    if (y1>y2) {  jsFlag = "ER"; }
	else       {  m1+d1>m2+d2;   } 
	if ( (y1==y2) && (m1+d1>m2+d2) ) {  jsFlag = "ER"; }
	//alert(y1+m1+d1+' '+y2+m2+d2);
    if( (jsFlag=="ER")&&(msg!="") )
	{  
	    alert(msg);
		DO1.focus();
	}
	return jsFlag;
}

function chkF_ID(PKObj,msg)
{
   var PKLen,PKStr,PKi,PKChr;
   PKLen = PKObj.value.length;
   PKStr = "";
   jsFlag = "OK";
   if ( PKLen==0 ) { jsFlag = "ER" ;}
   for (PKi=0;PKi<PKLen;PKi++)
   {
      PKChr = PKObj.value.substring(PKi,PKi+1);
      if ( ((PKChr>='A')&&(PKChr<='Z')) || ((PKChr>='a')&&(PKChr<='z')) || ((PKChr>='0')&&(PKChr<='9')) || (PKChr=='.') || (PKChr=='_') || (PKChr=='-') || (PKChr=='@') )
      {         PKStr += PKChr;
        }else{ jsFlag = "ER";break;     
      }
   }
   PKChr = PKObj.value.substring(0,1);
   PKObj.value = PKStr; 
   if(!(  ((PKChr>='A')&&(PKChr<='Z')) || ((PKChr>='a')&&(PKChr<='z')) || ((PKChr>='0')&&(PKChr<='9'))  )) 
	  { jsFlag = "ER"; }
    if( (jsFlag=="ER")&&(msg!="") )
	{  
	    alert(msg);
		PKObj.focus();
	}
   return jsFlag;
} 

function chkF_Tel(TObj,msg)
{
   var TLen,TStr,Ti,TChr,chNum;
   TLen = TObj.value.length;
   TStr = "";
   jsFlag = "OK";
   chNum  = 0;
   if ( TLen==0 ) { jsFlag = "ER" ;}
   for (Ti=0;Ti<TLen;Ti++)
   {
      TChr = TObj.value.substring(Ti,Ti+1);
      if ( ((TChr>='0')&&(TChr<='9')) || (TChr=='.') || (TChr=='_') || (TChr=='-') || (TChr=='(') || (TChr==')') )
      {         TStr += TChr;    if ((TChr>='0')&&(TChr<='9')) {chNum++;}
        }else{ jsFlag = "ER";break;     
      }
   }
   if (chNum<5) {jsFlag = "ER";}
   TObj.value = TStr; 
    if( (jsFlag=="ER")&&(msg!="") )
	{  
	    alert(msg);
		TObj.focus();
	}
   return jsFlag;
}

function chkF_PW(PKObj,msg)
{
   var PKLen,PKStr,PKi,PKChr;
   PKLen = PKObj.value.length;
   PKStr = "";
   jsFlag = "OK";  // _  "  '  /  \  ?  &  <  >
   if ( PKLen==0 ) { jsFlag = "ER" ;}
   for (PKi=0;PKi<PKLen;PKi++)
   {
      PKChr = PKObj.value.substring(PKi,PKi+1);
      if (  (PKChr==' ') || (PKChr=='&')  || (PKChr=='<') || (PKChr=='>') || (PKChr=='=') || (PKChr=='\'') || (PKChr=='\"')   )
	  {         jsFlag = "ER"; break; 
        }else{ jsFlag = "OK"; PKStr += PKChr;    
      }
   }
   PKObj.value = PKStr; 
    if( (jsFlag=="ER")&&(msg!="") )
	{  
	    alert(msg);
		PKObj.focus();
	}
   return jsFlag;
} 

function chkF_Dot(NDObj,msg)
{  jsFlag = "OK";
   if ( isNaN(NIObj.value) )
   {  jsFlag = "ER";
      NIObj.value = 0.0;
   }  
    if( (jsFlag=="ER")&&(msg!="") )
	{  
	    alert(msg);
		NDObj.focus();
	}
   return jsFlag;
} 

function chkF_Int(NIObj,msg)
{  jsFlag = "OK";
   if ( isNaN(NIObj.value) || NIObj.value.indexOf(".")>=0 )
   {  jsFlag = "ER";
      NIObj.value = 0;
   }  
    if( (jsFlag=="ER")&&(msg!="") )
	{  
	    alert(msg);
		NIObj.focus();
	}
   return jsFlag;
} 

function chkF_Blank(GObj,msg)
{  jsFlag = "OK";
   if ( GObj.value.length==0 ){ F = "ER";}
    if( (jsFlag=="ER")&&(msg!="") )
	{  
	    alert(msg);
		GObj.focus();
	}
   return jsFlag;
}

function chkF_Len(LNObj,LNLen,msg)
{  jsFlag = "OK";
   if ( LNObj.value.length>LNLen )
   {  jsFlag = "ER";
      LNObj.value = LNObj.value.substr(0,LNLen-3)+"..."
   }   
    if( (jsFlag=="ER")&&(msg!="") )
	{  
	    alert(msg);
		LNObj.focus();
	}
   return jsFlag;
}

function chkF_Pub(frmObj,frmIteN,frmMsg,frmItem)
{  
    frmflag=1;
    frm_Item = frmItem.split(";");
    frm_Msg  = frmMsg.split(";");    
    frmn = frmIteN ;//n = frm_Msg.dimensions();//ubound();
    for(frmi=0;frmi<frmn;frmi++)
    {
       if (frm_Item[frmi].length==0)
       {
          alert(frm_Msg[frmi]);
          frmflag=0;
          break;
       } 
    }
   if (frmflag==1) frmObj.submit()
}
