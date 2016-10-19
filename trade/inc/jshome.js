//**************************************************************
//*** 前台效果代码：兼容IE6,IE7,FF；各效果不冲突；兼容有无DOCTYPE定义；
//*** 但建议使用标准DOCTYPE定义。       by Peace(XieYS) 2009-03-28
//***************************************************************/


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
  getElmID("mTop01").scrollTop=0; 
  }else{     
  getElmID("mTop01").scrollTop++;
  }
} 
//var mTop0Speed=79;   
//var mTop0Mar=setInterval(mTop0Scroll,mTop0Speed);     
//getElmID("mTop01").onmouseover=function() {clearInterval(mTop0Mar);}     
//getElmID("mTop01").onmouseout=function() {mTop0Mar=setInterval(mTop0Scroll,mTop0Speed);}  

//****************** 下滚动代码 */
function fBot0Scroll(){
  if(getElmID("mBot01").offsetTop-getElmID("mBot00").scrollTop>=0)
    getElmID("mBot00").scrollTop+=getElmID("mBot02").offsetHeight;
  else{
    getElmID("mBot00").scrollTop--;
  }
}
//var mBot0Speed=83;
//getElmID("mBot02").innerHTML=getElmID("mBot01").innerHTML;
//getElmID("mBot00").scrollTop=getElmID("mBot00").scrollHeight;
//var Bot0Obj=setInterval(fBot0Scroll,mBot0Speed);
//getElmID("mBot00").onmouseover=function() {clearInterval(Bot0Obj);}
//getElmID("mBot00").onmouseout=function() {Bot0Obj=setInterval(fBot0Scroll,mBot0Speed);}


//****************** 左滚动代码 */
function mLeftMar(){
  if(getElmID("mLeftDiv2").offsetWidth-getElmID("mLeftDiv0").scrollLeft<=0)
    getElmID("mLeftDiv0").scrollLeft-=getElmID("mLeftDiv1").offsetWidth;
  else {getElmID("mLeftDiv0").scrollLeft++;}
}
//var mLeftSpeed = 29;//速度数值越大速度越慢
//getElmID("mLeftDiv2").innerHTML=getElmID("mLeftDiv1").innerHTML;
//var mLeftObj=setInterval(mLeftMar,mLeftSpeed);
//getElmID("mLeftDiv0").onmouseover=function() {clearInterval(mLeftObj);}
//getElmID("mLeftDiv0").onmouseout=function() {mLeftObj=setInterval(mLeftMar,mLeftSpeed);}


//****************** 右滚动代码 */
function mRightMar(){
  if( getElmID("mRightDiv2").offsetWidth-getElmID("mRightDiv0").scrollLeft>=0)
    getElmID("mRightDiv0").scrollLeft+=getElmID("mRightDiv1").offsetWidth;
  else {getElmID("mRightDiv0").scrollLeft--;}
}
var mRightSpeed = 31;//速度数值越大速度越慢
getElmID("mRightDiv2").innerHTML=getElmID("mRightDiv1").innerHTML;
var mRightObj=setInterval(mRightMar,mRightSpeed);
getElmID("mRightDiv0").onmouseover=function() {clearInterval(mRightObj);}
getElmID("mRightDiv0").onmouseout=function() {mRightObj=setInterval(mRightMar,mRightSpeed);}

