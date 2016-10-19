/********
  * 主程序来源：http://apps.hi.baidu.com/share/detail/16611600 
  demo: (include: convtab.js)
       <a href="javascript:convU13('nul');void 0;" class="as01" id="VerSetCX">简体版</a></div>
       <a href="javascript:convU13('t2s');void 0;" class="as01" id="VerSetC0">繁體版</a></div>
   or 
       <a href="javascript:convU12();void 0;" class="as01" id="VerSetCX">繁體版</a></div>
*******/
 
String.prototype.format = function(){ 
  var s = this; 
  for (var i=0,j=arguments.length; i<j; i++)
    s = s.replace("{" + (i) + "}", arguments[i]);
  return(s);
}

var convCookie= {
  Set : function (){
	var name = arguments[0], value = escape(arguments[1]),
	days = (arguments.length > 2)?arguments[2]:365,
	path = (arguments.length > 3)?arguments[3]:"/";
	with(new Date()){
	  setDate(getDate()+days);
	  days=toUTCString();
	}
	document.cookie = "{0}={1};expires={2};path={3}".format(name, value, days, path);
  },
  Get : function (){
	var returnValue=document.cookie.match(new RegExp("[\b\^;]?" + arguments[0] + "=([^;]*)(?=;|\b|$)","i"));
	return returnValue?unescape(returnValue[1]):returnValue;
  },
  Delete : function (){
	var name=arguments[0];
	document.cookie = name + "=1 ; expires=Fri, 31 Dec 1900 23:59:59 GMT;";
  }
}

var convMain = function(s2t){
  var s=convTabGB();
  var t=convTabB5();
  s2t = !!s2t || false;
  //convCookie.Set("langCookie",s2t?"s2t":"t2s");
  var stt = function(str){
	var r = "",i,j,k,c;
	for (i=0,j=str.length; i<j; i++)
	{
	  c = str.charAt(i);
	  k = (s2t)?s.indexOf(c):t.indexOf(c);
	  r+= (k==-1)?c:(s2t)?t.charAt(k):s.charAt(k);
	}
	return r;
  }
  return (function(o){
	if(!o)return;
	if(o.nodeType == 3){
	  o.nodeValue = stt(o.nodeValue);
	  return;
	}
	if (o.nodeType != 1)
	  return;
	if (o.tagName && ",OBJECT,FRAME,FRAMESET,IFRAME,SCRIPT,EMBD,STYLE,BR,HR,TEXTAREA,".indexOf(","+o.tagName.toUpperCase()+",")>-1)
	  return;
	if(o.title)
	  o.title = stt(o.title);
	if(o.alt)
	  o.alt = stt(o.alt);
	if(o.tagName && o.type && o.tagName.toUpperCase()=="INPUT" && ",button,submit,reset,".indexOf(o.type.toLowerCase())>-1)
	  o.value = stt(o.value);
	for (var i=0,j=o.childNodes.length; i<j; i++)
	{
	  arguments.callee(o.childNodes[i]);
	}
  })(document.body);
}

function convCText(xText,flag)
{
  var oStr='',xss,xtt;
  if("(s2t;t2s)".indexOf(flag)<0) return xText;
  if(flag=='s2t'){ xss=convTabGB();xtt=convTabB5(); }
  if(flag=='t2s'){ xss=convTabB5();xtt=convTabGB(); }  
  for(var i=0;i<xText.length;i++){
    if(xText.charCodeAt(i)>10000&&xss.indexOf(xText.charAt(i))!=-1)oStr+=xtt.charAt(xss.indexOf(xText.charAt(i)));
    else oStr+=xText.charAt(i);
  }
  return oStr;    
}

function convReset(){ 
  var flag = convCookie.Get("langCookie");
  if("(s2t;t2s)".indexOf(flag)<0) return;
  if(flag == "s2t"){ convMain(true); }
  if(flag == "t2s"){ convMain(false); }
  convUExt(flag);
}

// var convUser = function(){
function convU12(){ 
  var f = convCookie.Get("langCookie") == "s2t";
  var flag = f?"t2s":"s2t";
  convCookie.Set("langCookie",flag);
  convMain(!f);
  convUExt(flag);
}

function convU13(flag){ // flag=s2t,t2s,nul
  convCookie.Set("langCookie",flag);
  if("(s2t;t2s)".indexOf(flag)<0) document.location.href = document.location.href;
  if(flag == "s2t") { convMain(true); }
  if(flag == "t2s") { convMain(false); }
  convUExt(flag);
}

function convUExt(flag){ 
  document.title=convCText(document.title,flag);
  window.status=convCText('完成!'); //'+'js[简繁体中文转换]22(由Peace改写开发)'
  //return 'js简繁体中文转换\n\t['+o.substring(xYYYY-975,xYYYY-974)+String.fromCharCode(27704)+o.substring(xMMDD-60,xMMDD-59)+']';
  switch(flag){
    case 's2t':
      { 
	    document.getElementById("VerSetCX").innerHTML = "简体版";
		convImgs("/vchs/","/vcht/");
		break; 
	  }
    case 't2s':
      { 
	    document.getElementById("VerSetCX").innerHTML = "繁體版";
		convImgs("vcht","/vchs/");
	    break; 
	  }
    default:
      { 
	    convImgs("vcht","/vchs/");
		break; 
	  } // 原始的,不转化
  }
}

function convImgs(aDir,bDir){
  sideObj = document.getElementById("sideTable");
  sideItems = sideObj.getElementsByTagName("img");
  for(i=0;i<sideItems.length;i++){
	iUrl = sideItems[i].getAttribute("src");
	iUrl = iUrl.replace(aDir,bDir);
	sideItems[i].setAttribute("src",iUrl+"?r="+Math.random()+"");
  }
}

function convOption(){
  var verOption = document.getElementById('VerSelect');
  verOption.value = "Big5"; //以上两行为select:option下拉菜单设置
  verOption.options[0].text = "中文简体";
  verOption.options[1].text = "中文繁體";
  //以上4行为select:option下拉菜单设置 
}

function convDirect(flag,e){
  convCookie.Set("langCookie",flag);
  document.location.href = e.getAttribute("href");
}

function convDOption(e){
  var flag = e.options[e.selectedIndex].value;
  if(flag=="xEng") { document.location.href = "/peng/"; }
  convCookie.Set("langCookie",flag);
  if(flag == "s2t"){ convMain(true); }
  if(flag == "t2s"){ convMain(false); }
} 

setTimeout("convReset()",130);

function convTest(e){
  var msg = e.getAttribute("href");
  alert(msg);
}