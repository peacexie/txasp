
//var ModID="InfN124";
//var KeyID="dtinf-2010-78-B6WV.5EWMM";
//var ModTab="InfoNews"

var vmImgPath = "../img/mood/";
var vmStr = ""; //"128|10|2|3|40|5|6|7|8";

function vmSend(id){	
  var vmPara = "ModTab="+ModTab+"&ModID="+ModID+"&KeyID="+KeyID+"&NO="+id+"";
  //getID("vmMsg").innerHTML = vmUrl;
  try{
    getOutScript("../tools/nmood/vote.asp?"+vmPara,"utf-8");
	getID("vmMsg").textContent = "提交处理中...";
	//alert('f11'); 
  }catch(er){ 
    getID("vmMsg").textContent = "提交错误！";
  }
}

function vmShow(xStr){ 

  //alert(xStr);
  var vmArr = xStr.split("|");
  getID("vmMsg").innerHTML = "共有:"+vmArr[0]+"人次参入投票!";
  getID("vmRateN").style.display = '';
  for(i=1;i<=8;i++) { getID("vmItem"+i).disabled = true; }
  for(i=1;i<=8;i++) { 
	var iSrc = "V20A.gif"; if(vmArr[i]>(vmArr[0]/8)) { iSrc = "V20B.gif"; }
	var iHeight = 80*(vmArr[i]/vmArr[0]); 
	getID("vmRate"+i).innerHTML = vmArr[i]+"<BR><IMG width=20 height="+iHeight+" src='"+vmImgPath+iSrc+"'>";
  }
}

function getID(id){ return document.getElementById(id); }
