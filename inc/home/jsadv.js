

//****************** 侧边代码(公共) */
//var getElmObj = document.documentElement;
function AdvSideY(xDivN,xAdvY){ 
  var diffY;
  if (getElmObj && getElmObj.scrollTop) diffY = getElmObj.scrollTop;
  else if (document.body) diffY = document.body.scrollTop;
  else {  } 
  percent=.1*(diffY-xAdvY); 
  if(percent>0)percent=Math.ceil(percent); 
  else percent=Math.floor(percent); 
  getElmID(xDivN).style.top=parseInt(getElmID(xDivN).style.top)+percent+"px";
  return xAdvY+percent; 
}


//****************** 对联代码 */ 
//function AdvPair1(xDiv){AdvPTopY=AdvSideY(xDiv,AdvPTopY);}
//AdvPTopY=0;
//window.setInterval("AdvPair1('AdvP01')",30);
//window.setInterval("AdvPair1('AdvP02')",30);
//****************** 左边代码 */ 
//function AdvSide01(xDiv){AdvSY01=AdvSideY(xDiv,AdvSY01);}
//AdvSY01=0;window.setInterval("AdvSide01('AdvLeft01')",41);
//****************** 右边代码 */ 
//function AdvSide02(xDiv){AdvSY02=AdvSideY(xDiv,AdvSY02); }
//AdvSY02=0;window.setInterval("AdvSide02('AdvRight01')",47);
//****************** QQ代码 */ 
//function AdvSide03(xDiv){AdvSY03=AdvSideY(xDiv,AdvSY03); }
//AdvSY03=0;window.setInterval("AdvSide03('AdvQQ01')",49);





//****************** 浮动代码 01  --/
// 加一个浮动: f1-=>f3,f5,f7,f9; Float01-=>Float03,05,07,09
function Float01(){ 
  var L=0,T=0; fdBod = document.body;
  var R = fdBod.clientWidth-F01Obj.offsetWidth;
  var B = fdBod.clientHeight-F01Obj.offsetHeight;
  F01Obj.style.left = xF01 + fdBod.scrollLeft+"px";
  F01Obj.style.top = yF01 + fdBod.scrollTop+"px";
  xF01 = xF01 + F01Step*(xFlgGo?1:-1) ;
  if (xF01<L) { xFlgGo=true; xF01=L} 
  if (xF01>R) { xFlgGo=false; xF01=R} 
  yF01 = yF01 + F01Step*(yFlgGo?1:-1) ;
  if (yF01<T) { yFlgGo=true; yF01=T } 
  if (yF01>B) { yFlgGo=false; yF01=B } 
}  
var xFlgGo=true, yFlgGo=true; 
var F01Step=3,   F01Delay=89;  //1/30
//var xF01=240,    yF01=120;
//var F01Obj;      F01Obj = document.getElementById('AdvFloat01'); 
//var IntF01 = setInterval("Float01()", F01Delay); 
//F01Obj.onmouseover=function(){clearInterval(IntF01)} 
//F01Obj.onmouseout=function(){IntF01=setInterval("Float01()",F01Delay)} 


//****************** 浮动代码 02 */
// 加一个浮动: f2-=>f4,f6,f8,f0; Float02-=>Float04,06,08,00
function Float02() { 
 var fwDoc,fwElem; fwDoc = isDBody(); 
 fwElem = getElmID('AdvFloat02'); 
 fwElem.style.top = f2PosY+'px'; 
 fwDoc.visibility = "visible"; 
 if (f2Pause){
   width = fwDoc.clientWidth; 
   height = fwDoc.clientHeight; 
   f2OffH = fwElem.offsetHeight; 
   f2OffW = fwElem.offsetWidth; 
   //fwElem.style.left = f2PosX + fwDoc.scrollLeft+"px"; 
   //fwElem.style.top = f2PosY + fwDoc.scrollTop+"px"; 
   fwElem.style.left = f2PosX + "px"; 
   fwElem.style.top = f2PosY + "px"; 
   if (f2Yon) { f2PosY=f2PosY+f2Step; } 
   else { f2PosY = f2PosY - f2Step; } 
   if (f2PosY<0) { f2Yon=1; f2PosY=0; } 
   if (f2PosY>=(height-f2OffH)) { f2Yon=0; f2PosY=(height-f2OffH); }  
   if (f2Xon) { f2PosX=f2PosX+f2Step; } 
   else { f2PosX=f2PosX-f2Step; } 
   if (f2PosX<0) { f2Xon=1; f2PosX=0; } 
   if (f2PosX>= (width-f2OffW)) { f2Xon=0; f2PosX = (width-f2OffW); }   
 }
}
var f2OffH=0,     f2OffW=0; 
var f2Step=3,     f2Delay=83; 
var f2Yon=0,      f2Xon=0; 
var f2Int,        f2Pause = true;  
//var f2PosX=720,   f2PosY=480;
//f2Int = setInterval('Float02()', f2Delay); 




