
// ******************************************************************************
// 判断浏览器，操作系统相关；
// ******************************************************************************

//属性 描述 IE F O 
//appCodeName 返回浏览器的代码名。 4 1 9 
//appMinorVersion 返回浏览器的次级版本。 4 No No 
//appName 返回浏览器的名称。 4 1 9 
//appVersion 返回浏览器的平台和版本信息。 4 1 9 
//browserLanguage 返回当前浏览器的语言。 4 No 9 
//cookieEnabled 返回指明浏览器中是否启用 cookie 的布尔值。 4 1 9 
//cpuClass 返回浏览器系统的 CPU 等级。 4 No No 
//onLine 返回指明系统是否处于脱机模式的布尔值。 4 No No 
//platform 返回运行浏览器的操作系统平台。 4 1 9 
//systemLanguage 返回 OS 使用的默认语言。 4 No No 
//userAgent 返回由客户机发送服务器的 user-agent 头部的值。 4 1 9 
//userLanguage 返回 OS 的自然语言设置。 4 No 9 

var uaStr = navigator.userAgent.toLowerCase();
brsCheck = function(r){
  return r.test(uaStr);
};
isOpera = brsCheck(/opera/);
isChrome = brsCheck(/chrome/);
isWebKit = brsCheck(/webkit/);
isSafari = !isChrome && brsCheck(/safari/);
isSafari2 = isSafari && brsCheck(/applewebkit\/4/); // unique to Safari 2
isSafari3 = isSafari && brsCheck(/version\/3/);
isSafari4 = isSafari && brsCheck(/version\/4/);
isIE = !isOpera && brsCheck(/msie/);
isIE7 = isIE && brsCheck(/msie 7/);
isIE8 = isIE && brsCheck(/msie 8/);
isIE9 = isIE && brsCheck(/msie 9/);
isIE6 = isIE && !isIE7 && !isIE8 && !isIE9;
isGecko = !isWebKit && brsCheck(/gecko/); //isFirefox
isGecko2 = isGecko && brsCheck(/rv:1\.8/);
isGecko3 = isGecko && brsCheck(/rv:1\.9/);
isFirefox = brsCheck(/firefox/);
isFirefox1 = brsCheck(/firefox\/1/);
isFirefox2 = brsCheck(/firefox\/2/);
isFirefox3 = brsCheck(/firefox\/3/);

//isFirefox = /firefox/.test(navigator.userAgent.toLowerCase());
  
//document.write('<br>'+uaStr);
//if(isGecko) { document.write('<br>isGecko'); }
//if(isFirefox3) { document.write('<br>isFirefox3'); }
//if(isIE7) { document.write('<br>ie7'); }

function browserSniffer() 
{ 
  var uBrows = "";
  var agt = navigator.userAgent.toLowerCase(); 
  var agt = navigator.userAgent.toLowerCase(); 
  var is_major = parseInt(navigator.appVersion); 
  var is_minor = parseFloat(navigator.appVersion); 
  var is_nav = ((agt.indexOf('mozilla')!=-1) && (agt.indexOf('spoofer')==-1) && (agt.indexOf('compatible') == -1) && (agt.indexOf('opera')==-1) && (agt.indexOf('webtv')==-1)); 
  var is_nav2 = (is_nav && (is_major == 2)); 
  var is_nav3 = (is_nav && (is_major == 3)); 
  var is_nav4 = (is_nav && (is_major == 4)); 
  var is_nav4up = (is_nav && (is_major >= 4)); 
  var is_navonly = (is_nav && ((agt.indexOf(";nav") != -1) || (agt.indexOf("; nav") != -1)) ); 
  var is_nav5 = (is_nav && (is_major == 5)); 
  var is_nav5up = (is_nav && (is_major >= 5)); 
  var is_ie = (agt.indexOf("msie") != -1); 
  var is_ie3 = (is_ie && (is_major < 4)); 
  var is_ie4 = (is_ie && (is_major == 4) && (agt.indexOf("msie 5.0")==-1) ); 
  var is_ie4up = (is_ie && (is_major >= 4)); 
  var is_ie5 = (is_ie && (is_major == 4) && (agt.indexOf("msie 5.0")!=-1) ); 
  var is_ie5up = (is_ie && !is_ie3 && !is_ie4); 
  var is_aol = (agt.indexOf("aol") != -1); 
  var is_aol3 = (is_aol && is_ie3); 
  var is_aol4 = (is_aol && is_ie4); 
  var is_opera = (agt.indexOf("opera") != -1); 
  var is_webtv = (agt.indexOf("webtv") != -1); 
  if (is_nav4up) { 
    location.href = netscape4URL; // netscape 4+ but not NS5 
  }else if (is_ie4up) { //IE4 & IE5 but returns IE4 
    location.href = explorer4URL; 
  }else if (is_webtv) { // Web TV 
    location.href = webtvURL; 
  }else if (is_aol || is_aol3 || is_aol4) { //AOL 
    location.href = aolURL; 
  }else if (is_opera) { // Opera 
    location.href = operaURL; 
  }else if (is_ie3||is_nav3) { // 3.0 version browsers 
    location.href = version3URL; 
  }else if (is_nav5up) { // Netscape 5 
    location.href = w3cURL; 
  } 
} 

function getIEVers()
{
  var sVer = "";
  var uaStr = navigator.userAgent;
  if(uaStr.indexOf("MSIE")>0)//是否是IE浏览器
  { 
    if(uaStr.indexOf("MSIE 5.0")>0) { 
       sVer = "IE5.0";
    }else if(uaStr.indexOf("MSIE 6.0")>0) { 
       sVer = "IE6.0";
    }else if(uaStr.indexOf("MSIE 7.0")>0) {
       sVer = "IE7.0";
    }else if(uaStr.indexOf("MSIE 8.0")>0) {
       sVer = "IE8.0";
    }else if(uaStr.indexOf("MSIE 9.0")>0) {
       sVer = "IE9.0";
    } else {
       sVer = "IE4.0";
    }
  } else {
      sVer = "!IE";
  }
}



// ******************************************************************************
// 在 FireFox 中，可以用outerHTML属性；
// ******************************************************************************
if(typeof(HTMLElement)!="undefined" && !window.opera) 
{ 
    HTMLElement.prototype.__defineGetter__("outerHTML",function() 
    { 
        var a=this.attributes, str="<"+this.tagName, i=0;for(;i<a.length;i++) 
        if(a[i].specified) 
            str+=" "+a[i].name+'="'+a[i].value+'"'; 
        if(!this.canHaveChildren) 
            return str+" />"; 
        return str+">"+this.innerHTML+"</"+this.tagName+">"; 
    }); 
    HTMLElement.prototype.__defineSetter__("outerHTML",function(s) 
    { 
        var r = this.ownerDocument.createRange(); 
        r.setStartBefore(this); 
        var df = r.createContextualFragment(s); 
        this.parentNode.replaceChild(df, this); 
        return s; 
    }); 
    HTMLElement.prototype.__defineGetter__("canHaveChildren",function() 
    { 
        return !/^(area|base|basefont|col|frame|hr|img|br|input|isindex|link|meta|param)$/.test(this.tagName.toLowerCase()); 
    }); 
	
  // innerText
  HTMLElement.prototype.__defineGetter__("innerText",
   function(){  
   var anyString = "";  
   var childS = this.childNodes;  
   for(var i=0; i<childS.length; i++) {
    if(childS[i].nodeType==1)  
       anyString += childS[i].tagName=="BR" ? '\n' : childS[i].textContent;  
     else if(childS[i].nodeType==3)  
       anyString += childS[i].nodeValue;  
    }  
    return anyString;  
   }  
  );  
  HTMLElement.prototype.__defineSetter__("innerText",  
   function(sText){  
    this.textContent=sText;  
    }  
   );  
	
}




// ******************************************************************************
// http://hi.baidu.com/xyue13/blog/item/c5e691194bac9cbf4bedbc86.html
// 在 FireFox 中有 event、event.srcElement、event.fromElement、event.toElement 属性；
// ******************************************************************************

if(window.addEventListener) { FixPrototypeForGecko(); }  

function  FixPrototypeForGecko()  
{  
	HTMLElement.prototype.__defineGetter__("runtimeStyle",element_prototype_get_runtimeStyle);  
	window.constructor.prototype.__defineGetter__("event",window_prototype_get_event);  
	Event.prototype.__defineGetter__("srcElement",event_prototype_get_srcElement);  
	Event.prototype.__defineGetter__("fromElement",  element_prototype_get_fromElement);  
	Event.prototype.__defineGetter__("toElement", element_prototype_get_toElement);
	
}  

function  element_prototype_get_runtimeStyle() { return  this.style; }  
function  window_prototype_get_event() { return  SearchEvent(); }  
function  event_prototype_get_srcElement() { return  this.target; }  

function element_prototype_get_fromElement() {  
	var node;  
	if(this.type == "mouseover") node = this.relatedTarget;  
	else if (this.type == "mouseout") node = this.target;  
	if(!node) return;  
	while (node.nodeType != 1) 
		node = node.parentNode;  
	return node;  
}

function  element_prototype_get_toElement() {  
		var node;  
		if(this.type == "mouseout")  node = this.relatedTarget;  
		else if (this.type == "mouseover") node = this.target;  
		if(!node) return;  
		while (node.nodeType != 1)  
		   node = node.parentNode;  
		return node;  
}
 
function  SearchEvent()  
{  
	if(document.all) return  window.event;  
	 
	func = SearchEvent.caller;  

	while(func!=null){  
		var  arg0=func.arguments[0];  
		
		if(arg0 instanceof Event) {  
			return  arg0;  
		}  
	   func=func.caller;  
	}  
	return   null;  
}
