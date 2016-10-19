
function getElmID(xID){
  return document.getElementById(xID);
}

function ImgPlayMain(id,xWidth,xHeight)
{
  var aPic = eval("ImgPlay_"+id+"_Pics").replace(" ","").split("|");
  var aAlt = eval("ImgPlay_"+id+"_Subj").replace(" ","").split("|"); 
  var Speed = eval("ImgPlay_"+id+"_Speed");
  var eDiv = getElmID("ImgPlayer_"+id);
  var n = eval("ImgPlay_"+id+"_Now"); n++; 
  if((n>0)&&(aPic[n]=="")){ n=0; }
  if(n>=aPic.length){ n=0; } 
  var sExt = aPic[n].substring(aPic[n].lastIndexOf(".")).toLowerCase();
  var sSrc = "src='"+aPic[n]+"' width='"+xWidth+"' height='"+xHeight+"' alt='"+aAlt[n]+"'"; 
  if(sExt==".swf") { 
    eDiv.innerHTML = "<embed "+sSrc+" quality='high' wmode='transparent' type='application/x-shockwave-flash'></embed>"; 
  } else { 
    eDiv.innerHTML = "<img "+sSrc+" onclick=\"ImgPlayOpen('"+id+"')\" id='ImgPlay_"+id+"_Trans' border='0' style='cursor:pointer;FILTER: revealTrans(duration=1.87);' />";
  }
  ImgPlayTrans(id);
  eval("ImgPlay_"+id+"_Now = n;");
  setTimeout("ImgPlayMain('"+id+"',"+xWidth+","+xHeight+")",Speed); 
}

function ImgPlayOpen(id)
{
  var aUrl = eval("ImgPlay_"+id+"_Urls").replace(" ","").split("|");
  var sUrl = aUrl[eval("ImgPlay_"+id+"_Now")];
  if((sUrl=="#")||(sUrl=="")) { ; }
  else { window.open(sUrl,"_blank"); }
}

function ImgPlayTrans(xID)
{
  if (document.all){
    try {
	  eval('ImgPlay_'+xID+'_Trans').filters.revealTrans.Transition=Math.floor(Math.random()*23);
      eval('ImgPlay_'+xID+'_Trans').filters.revealTrans.apply(); //变换效果
	  eval('ImgPlay_'+xID+'_Trans').filters.revealTrans.play(); //运行效果
	} catch (ePlay) { }
  }
  //<meta http-equiv="Page-Enter" content="blendTrans(Duration=1.0)">
}

// Demo

//ImgPlay_AdvID_Pics = "y_biaoshi.gif|y_logo.gif|../../../img/logo/wj_chacha.gif";
//ImgPlay_AdvID_Urls = "upRow.htm    |fQQ2.htm  |xx";
//ImgPlay_AdvID_Subj = "说明一       |说明22    |xx测试"; 
//ImgPlay_AdvID_Speed = 1800; // 时间间隔
//ImgPlay_AdvID_Now = -1;     // 记录当前第几个图片，一定要-1开始;

// 每增加一个广告位，就增加一组以上的变量，把AdvID换成如myAd001
// 在显示广告的地方，用以下代码，把AdvID换成如myAd001
//<div id="ImgPlayer_AdvID" style="width:200px; height:150px; overflow:hidden; border:1px solid #CCC">
//<-script language="javascript">setTimeout("ImgPlayMain('AdvID',200,150)",120); <-/script->
//</div>
//支持swf文件,支持IE6，IE7，FireFox,支持DOCTYPE


