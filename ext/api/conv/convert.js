/********
  demo: (include: convtab.js)
      <li class="SysVers" id="VerSetCX" onclick="SetLang('xOri')">不转化</li>
      <li class="SysVers" id="VerSetC1" onclick="SetLang('GB2312')">转简体</li>
      <li class="SysVers" id="VerSetC0" onclick="SetLang('Big5')">轉繁體</li></td>
*******/

var flagScript = 'script';
document.write('<'+flagScript+' src="../ext/api/conv/convtab.js" type="text/javascript"></'+flagScript+'>');
 
function Code_CText(xText)
{
  var oStr='',xss,xtt;
  if(ObjectLang=='Big5'){
	 xss=Code_TabGB();xtt=Code_TabB5();
  }else{
     xss=Code_TabB5();xtt=Code_TabGB();
  }  
  for(var i=0;i<xText.length;i++){
    if(xText.charCodeAt(i)>10000&&xss.indexOf(xText.charAt(i))!=-1)oStr+=xtt.charAt(xss.indexOf(xText.charAt(i)));
    else oStr+=xText.charAt(i);
  }
  return oStr;    
}

function Code_CBody(xObj) /*转换对象，使用递归，逐层剥到文本*/
{ 
  if(typeof(xObj)=='object'){
    var cObj=xObj.childNodes;
  }else {
    var cObj=document.body.childNodes;
  }
  for(var i=0;i<cObj.length;i++){
    var xOO=cObj.item(i);
    if('||BR|HR|'.indexOf('|'+xOO.tagName+'|')>0)continue; /*&&xOO.type!='text';TEXTAREA|*/
    if(xOO.title!=''&&xOO.title!=null)xOO.title=Code_CText(xOO.title);
    if(xOO.alt!=''&&xOO.alt!=null)xOO.alt=Code_CText(xOO.alt);
    if(xOO.tagName=='INPUT'&&xOO.value!=''&&xOO.type!='hidden')xOO.value=Code_CText(xOO.value);
	/*if(OO.tagName=='script')OO.value=Code_CText(OO.value);*/
    if(xOO.nodeType==3){xOO.nodeValue=Code_CText(xOO.nodeValue);} 
    else Code_CBody(xOO); /*data;nodeValue;nodeValue*/
  }	
}

function Code_COther(xYYYY,xMMDD,xType) /*转换其它对象(Peace添加)*/
{ 
  var o = Code_TabGB(); if(xType=='Big5'){ o = Code_TabB5(); }
  document.title=Code_CText(document.title);
  window.status=Code_CText('完成!   '); //'+'js[简繁体中文转换]21(由Peace改写开发)'
  return 'js简繁体中文转换\n\t['+o.substring(xYYYY-975,xYYYY-974)+String.fromCharCode(27704)+o.substring(xMMDD-60,xMMDD-59)+']';
}

function Code_ConvDo(xSet){
  ObjectLang = xSet;  // Big5,
  Code_CBody();       // setTimeout('Code_CBody()',50); Code_COther(); 
  try{
    document.getElementById('VerSetC1').innerHTML = '转简体';
    document.getElementById('VerSetC0').innerHTML = '轉繁體';
    //以上2行为 VerSetCX 至少两个 设置
  }catch(eDo) { }
  
}

function SetLang(xCSet){
  if(xCSet=='xOri'){
	SetCookie(xCSet); 
	document.location.href = document.location.href;
  }else{
    SetCookie(xCSet);
	CodeReset(); 
  }
}

function SetCookie(xVal){ 
  document.cookie='PeaceCSet='+xVal+';';
  D = new Date(); d = 24*60*60*1000; //1天
  D.setTime(D.getTime() + 30*d); 
  document.cookie='expires='+D.toGMTString();
} 

function GetCookie(){ 
  var temp=document.cookie+';'; 
  var Pos=temp.indexOf('=',temp.indexOf('PeaceCSet=')); 
  if (temp.indexOf('PeaceCSet=')==-1) return ''; 
  return temp.substring(Pos+1,temp.indexOf(';',Pos));
} 

// 转化配置 "GB2312|Big5|xOri|''"
function CodeReset(){ 
  switch(GetCookie()){
    case 'Big5':
      { 
	    Code_ConvDo('Big5');
		//try{
		  //var verOption = document.getElementById('VerSelect');
		  //verOption.value = "Big5"; //以上两行为select:option下拉菜单设置
		  //verOption.options[0].text = "中文简体";
		  //verOption.options[1].text = "中文繁體";
		  //以上4行为select:option下拉菜单设置 
		  //document.getElementById('VerSetCX').innerHTML = "简体版";
		  //setUImg();
		//}catch(e) { } 
		break; 
	  }
    case 'GB2312':
      { Code_ConvDo('GB2312'); break; }
    case 'xOri':
      { break; } //Original:原始的,不转化
    default:
      { break; } //默认不转化，根据需要加Code_ConvDo('Big5|GB2312');
  }
} 

function setUImg(){
  sideObj = document.getElementById("sideTable");
  sideItems = sideObj.getElementsByTagName("img");
  for(i=0;i<sideItems.length;i++){
	iUrl = sideItems[i].getAttribute("src");
	iUrl = iUrl.replace("/vchs/","/vcht/");
	sideItems[i].setAttribute("src",iUrl+"?r="+Math.random()+"");
  }
}

function VerShange(e){
  var xCSet = e.options[e.selectedIndex].value;
  SetCookie(xCSet);
  if(xCSet=="xEng") { document.location.href = "/peng/"; }
  else              { document.location.href = "/page/"; }
  //SetLang(xCSet);
} 

setTimeout("CodeReset()",13);


