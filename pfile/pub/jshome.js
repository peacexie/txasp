//**************************************************************
//*** 前台效果代码：兼容IE6,IE7,FF；各效果不冲突；兼容有无DOCTYPE定义；
//*** 但建议使用标准DOCTYPE定义。       by Peace(XieYS) 2009-03-28
//***************************************************************/


sw01_lnk1 = "info.asp?TypID=N110016";
sw01_lnk2 = "info.asp?TypID=N110020";
function sw01_Chang(xID)
{
  var iMax = 2;
  for(var i=1;i<=iMax;i++){ 
    ItemID = getElmID("sw01_id"+i+"");
	IDivID = getElmID("sw01_box"+i+"");
	if (xID==i){
	  IDivID.style.display='';
	  IDivID.style.visibility='visible';
      ItemID.className="sw01_css"+i+"a";
	}else{
	  IDivID.style.display='none';
	  IDivID.style.visibility='hidden';
	  ItemID.className="sw01_css"+i+"b";
  } }
  getElmID("sw01_link").href=eval("sw01_lnk"+xID+"");
}
//sw01_Chang; 


//****************** 选项切换代码 */
function SwhChang(xDiv)
{
  for(var i=1;i<4;i++){ 
    ItemID = getElmID("SwhItem0"+i+"");
	IDivID = getElmID("SwhIDiv0"+i+"");
	if (xDiv==i){
	  IDivID.style.display='';
	  IDivID.style.visibility='visible';
      ItemID.className="SwhItmB";
	}else{
	  IDivID.style.display='none';
	  IDivID.style.visibility='hidden';
	  ItemID.className="SwhItmA";
  } }
}
SwhChang(1);


//****************** 上滚动代码 */
function mTop0Scroll(){     
  if(getElmID("mTop01").scrollTop+getElmID("mTop01").offsetHeight>=getElmID("mTop02").scrollHeight){ 
  getElmID("mTop01").scrollTop=-40; 
  }else{     
  getElmID("mTop01").scrollTop++;
  }
} 
var mTop0Speed=79;   
var mTop0Mar=setInterval(mTop0Scroll,mTop0Speed);     
getElmID("mTop01").onmouseover=function() {clearInterval(mTop0Mar);}     
getElmID("mTop01").onmouseout=function() {mTop0Mar=setInterval(mTop0Scroll,mTop0Speed);}  

//****************** 下滚动代码 */
function fBot0Scroll(){
  if(getElmID("mBot01").offsetTop-getElmID("mBot00").scrollTop>=0)
    getElmID("mBot00").scrollTop+=getElmID("mBot02").offsetHeight;
  else{
    getElmID("mBot00").scrollTop--;
  }
}
var mBot0Speed=83;
getElmID("mBot02").innerHTML=getElmID("mBot01").innerHTML;
getElmID("mBot00").scrollTop=getElmID("mBot00").scrollHeight;
var Bot0Obj=setInterval(fBot0Scroll,mBot0Speed);
getElmID("mBot00").onmouseover=function() {clearInterval(Bot0Obj);}
getElmID("mBot00").onmouseout=function() {Bot0Obj=setInterval(fBot0Scroll,mBot0Speed);}


//****************** 左滚动代码 */
function mLeft0Mar(){
  if(getElmID("mLeft0Div2").offsetWidth-getElmID("mLeft0Div0").scrollLeft<=0)
    getElmID("mLeft0Div0").scrollLeft-=getElmID("mLeft0Div1").offsetWidth;
  else {getElmID("mLeft0Div0").scrollLeft++;}
}
var mLeft0Speed = 29;//速度数值越大速度越慢
getElmID("mLeft0Div2").innerHTML=getElmID("mLeft0Div1").innerHTML;
var mLeft0Obj=setInterval(mLeft0Mar,mLeft0Speed);
getElmID("mLeft0Div0").onmouseover=function() {clearInterval(mLeft0Obj);}
getElmID("mLeft0Div0").onmouseout=function() {mLeft0Obj=setInterval(mLeft0Mar,mLeft0Speed);}


//****************** 右滚动代码 */
function mRight0Mar(){
  if( getElmID("mRight0Div2").offsetWidth-getElmID("mRight0Div0").scrollLeft>=0)
    getElmID("mRight0Div0").scrollLeft+=getElmID("mRight0Div1").offsetWidth;
  else {getElmID("mRight0Div0").scrollLeft--;}
}
var mRight0Speed = 31;//速度数值越大速度越慢
getElmID("mRight0Div2").innerHTML=getElmID("mRight0Div1").innerHTML;
var mRight0Obj=setInterval(mRight0Mar,mRight0Speed);
getElmID("mRight0Div0").onmouseover=function() {clearInterval(mRight0Obj);}
getElmID("mRight0Div0").onmouseout=function() {mRight0Obj=setInterval(mRight0Mar,mRight0Speed);}

