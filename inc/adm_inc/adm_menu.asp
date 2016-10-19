<!--#include file="../../sadm/func2/func_const.asp"-->
<!--#include file="../../sadm/func2/func_perm.asp"-->
<%
Call Chk_Perm9("","(Adm)")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="robots" content="noindex, nofollow">
<title><%=Config_Name%>后台管理中心</title>
<script src="../../inc/home/jsInfo.js" type="text/javascript"></script>
<link rel="stylesheet" href="../../inc/adm_img/style.css" type="text/css" />
</head>
<body style="overflow-x:hidden;overflow-y:scroll;text-align:left;">
<div id="TabPage1" style="width:160px; overflow:hidden">
  <div class="left_top"></div>
  <ul class="left_conten" style="padding:0px 12px">
    <div id="ModTBar">Loading...</div>
  </ul>
  <div class="left_end"></div>
  <!-- // Menu Start //////////// -->
  <!--#include file="../../upfile/sys/config/adm_menu.config"-->
  <!-- // Menu End /////////// -->
  <div class="left_top"></div>
  <ul class="left_conten">
    <div id="ModBBar">Loading...</div>
    <div align="center"> Copyright by 东莞网 </div>
  </ul>
  <div class="left_end"></div>
</div>

<div id="sModBar">
  <div align="center">
  <a target="mainFrame" href="../../sadm/user/index.asp"><font color="#FF00FF">管理中心</font></a>
   | 
  <a target="_parent"href="../../sadm/<%=Config_RAdm%>.asp?send=out&UsrType=adm"><font color="#FF0000">安全退出</font></a>
  </div>
</div>
<div id="sModGroup">
  <!--#include file="../../upfile/sys/config/sf_Admin.htm"-->
</div>
<script type="text/javascript">

//var mStr += ";ModInfo;ModPics;ModGbook;ModMember;ModHome;ModSystem"; 
var mArr = mStr.split(";");
var uPerm = "<%=Session(UsrPStr)%>";
<!--#include file="../../upfile/sys/config/sf_Admin.js"-->

function ShowGroup(xID){
  if(getElmID("ModGroup"+xID).className=="ModGroupX ModGroupA") return; //当前退出
  // uPerm中(Mod个数, 多于5个就分组
  var gStr = aGroup[xID]; //alert(gs);
  for(var i=1;i<mArr.length;i++){ 
	var e2 = getElmID("sSub"+mArr[i]).parentNode;
    if(gStr.indexOf(mArr[i])>=0){ 
	  e2.style.display="";  
	  e2.style.visibility='visible';
	}else{
	  e2.style.display="none";  
	  e2.style.visibility='hidden'; 
	}
  } 
  for(var i=0;i<nGroup;i++){ 
    getElmID("ModGroup"+i).className="ModGroupX ModGroupB";
  }
  getElmID("ModGroup"+xID).className="ModGroupX ModGroupA";
  mDef = gStr.split(","); ShowMenu(mDef[0]);
}

function ShowMenu(xID){
  for(var i=1;i<mArr.length;i++){  
	iElm = getElmID("sSub"+mArr[i]);
	/* Close All Div */ 
	if(xID!=mArr[i])
	  iElm.style.display="none"; 
	/* Hidden No Perm */ 
	if( uPerm.indexOf(""+mArr[i]+"")<0 && uPerm.lastIndexOf("{Admin}")<0 ) 
	  iElm.parentNode.style.display="none";
  }  
  /*Close/Open MenuID*/
  var jElm = getElmID("sSub"+xID);  
  if(uPerm.lastIndexOf("{Admin}")<0) HiddMSub(jElm);
  //jElm.style.display=""; 
  if(jElm.style.display == "none") jElm.style.display="";  
  else                             jElm.style.display="none";  
}

function HiddMSub(eGrp){
  var itms = eGrp.getElementsByTagName("li");
  for(var i = 0;i<itms.length;i++){
	var sID = itms[i].id.toString();
	if(sID.length>5){
	  if( uPerm.lastIndexOf(sID.substring(5))<0 ) 
	    getElmID(sID).style.display="none";
	}
  }
  var itms = eGrp.getElementsByTagName("a"); //admaHomAdvert
  for(var i = 0;i<itms.length;i++){
	var sID = itms[i].id.toString();
	if(sID.length>5){
	  if( uPerm.lastIndexOf(sID.substring(4))<0 ) {
	    getElmID(sID).className="fntCCC";
	    getElmID(sID).href="javascript:alert('无权限!')";
	  }
	}
  }
}

////////////////////
<%
Response.Write "var defStr = '"&Request("g")&"';"&vbcrlf ' //1,ModMember
%>
if(defStr!=""){
  var a = defStr.split(",");
  if(a.length==2){
	GropDef = a[0];
	GropSub = a[1];
  }
}
if(GropSwitch=="Y"){
  ShowGroup(GropDef);
  ShowMenu(GropSub);
  getElmID("ModTBar").innerHTML = getElmID("sModGroup").innerHTML; 
  getElmID("ModBBar").innerHTML = getElmID("sModBar").innerHTML; 
}else{
  ShowMenu(GropSub);	
  ShowMenu(GropSub);	
  getElmID("ModTBar").innerHTML = getElmID("sModBar").innerHTML;
  getElmID("ModBBar").style.display="none";
  //getElmID("ModBBar").innerHTML = "";
}
ShowDiv("sModGroup"); ShowDiv("sModBar");

</script>
</body>
</html>