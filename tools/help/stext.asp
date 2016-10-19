<!--#include file="stcfg.asp"-->
<%
act = Show_Text(Request("act")) :If act="" Then act="index"
smd = Show_Text(Request("smd")) :If smd="" Then smd="web"
usr = Show_Text(Request("usr")) :If usr="" Then usr="(Guest)"
ver = Show_Text(Request("ver")) :If ver="" Then ver="(unKnow)"
frm = Request.Servervariables("HTTP_REFERER")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>服务跟踪页@东莞网</title>
<script src="../../inc/home/jsInfo.js" type="text/javascript"></script>
<link rel="stylesheet" type="text/css" href="strace.css"/>
<style type="text/css">
.bdFull {
	border:0px solid #639;
}
.bdTop {
	border-top:1px solid #CCC;
}
.bdBot {
	border-bottom:1px solid #CCC;
}
.noBot {
	border-bottom:1px solid #F0F0FF;
}
.idCom {
	background-color:#F0F0FF;
	cursor:pointer;
}
.idAct {
	background-color:#FFF;
	border:1px solid #639;
	cursor:pointer;
}
.idRed {
	color:#F00;
}
.pdCell {
	padding:8px;
}
.ntSide {
	width:240px;
	background-color:#F0F0F0;
	font-size:12px;
}
</style>
</head>
<body>
<table width="99%" border="0" align="center" cellspacing="5" style="margin:5px; border:1px solid #669">
  <tr>
    <td colspan="2" valign="top" class="bdFull noBot"><table width="100%" border="0">
        <tr>
          <th width="16%" class="idAct" id="sw01_id1" onMouseOver="msgReady('1')" onmousedown="msgShow('1')">服务首页</th>
          <th width="16%" class="idCom" id="sw01_id2" onMouseOver="msgReady('2')" onmousedown="msgShow('2')">最新动态</th>
          <th width="16%" class="idCom" id="sw01_id3" onMouseOver="msgReady('3')" onmousedown="msgShow('3')">推荐1</th>
          <th width="16%" class="idCom" id="sw01_id4" onMouseOver="msgReady('4')" onmousedown="msgShow('4')">推荐2</th>
          <th width="16%" class="idCom" id="sw01_id5" onMouseOver="msgReady('5')" onmousedown="msgShow('5')">推荐3</th>
          <th width="16%" class="idCom" id="sw01_id6" onMouseOver="msgReady('6')" onmousedown="msgShow('6')">申明/刷新</th>
        </tr>
      </table></td>
  </tr>
  <tr>
    <th class="bdBot" id="msgT">东莞网 服务跟踪页</th>
    <td rowspan="2" align="left" valign="top" class="pdCell ntSide" id="msgB">&nbsp;</td>
  </tr>
  <tr>
    <td valign="top" id="msgA" class="pdCell">&nbsp;</td>
  </tr>
</table>
<%
Response.Write "<div style='display:none'>"
%>
<div id="box1A">以下为你提交过来的信息：
  <ul>
    <li>项目名称(usr)=<%=usr%></li>
    <li>服务项目(smd)=<%=smd%></li>
    <li>操作动作(act)=<%=act%></li>
    <li>使用版本(ver)=<%=ver%></li>
    <li>来源地址(frm)=<%=frm%></li>
  </ul>
  <p>服务跟踪 资料</p>
  <ul>
    <li>公司名称：东莞市科成信息科技有限公司 </li>
    <li>联系电话：0769-22028848　22028858 <br />
      客服联络：电话:0769-1234-5678, QQ:80893510, E-mail:xxx@domain.com, XX某<br />
      技术探讨：QQ:80893510, E-mail:xxx@domain.com, XX某</li>
    <li>公司传真：0769-22028800 </li>
    <li>邮政编码：523072 </li>
    <li>联系地址：东莞市莞太大道34号东莞软件企业孵化园3C楼 </li>
    <li>E-mail： info@96327.com.cn </li>
    <li>网 址：http://www.dg.gd.cn/</li>
  </ul>
</div>
<div id="box1B">1. ...！<br />
  2. ...！</div>
<div id="box2A">最新动态 信息：
  <ul>
    <li>最新动态</li>
  </ul>
</div>
<div id="box2B">申明：最新动态 </div>
<div id="box3A">推荐1 信息：
  <ul>
    <li>推荐1</li>
  </ul>
</div>
<div id="box3B">申明：推荐1 </div>
<div id="box4A">推荐2 信息：
  <ul>
    <li>推荐2</li>
  </ul>
</div>
<div id="box4B">申明：推荐3 </div>
<div id="box5A">推荐3 信息：
  <ul>
    <li>推荐3</li>
  </ul>
</div>
<div id="box5B">申明：推荐3 </div>
<div id="box6A">
  <blockquote>
    <p>申明/刷新 信息： </p>
  </blockquote>
  <ol>
    <li>设置本栏目的目的是让项目(网站)拥有者能够及时联系上项目开发公司的相关人员(即使人员变动),并为项目拥有者提供相关服务； </li>
    <li>本栏目相关连接及内容由项目开发公司维护更新； </li>
    <li>项目开发公司拥有对本栏目的最终解吸权利！</li>
  </ol>
  <p align="center"><a href="<%=frm%>?act=reset">返回并刷新&gt;&gt;&gt;</a></p>
</div>
<div id="box6B"> 1. 如果仅显示本错误，并不影响你网站本身的正常运行！<br />
  2. 显示本错误提示，表示不能连接到服务商网站，这可能是暂时的；请不必要紧张，过一会儿再试尝！ </div>
</div>
<%
Response.Write "</div>"
%>
<script type="text/javascript">
var nMax = 6;
function msgReady(n){
  for(var i=1;i<=nMax;i++){ 
    swID = getElmID("sw01_id"+i+"");
	if(swID.className!="idAct idRed"){
	  if(n==i){ swID.className="idAct"; }
	  else{ swID.className="idCom"; } 
  } }
}
function msgShow(n){
  for(var i=1;i<=nMax;i++){ 
    swID = getElmID("sw01_id"+i+"");
	if(n==i){
      swID.className="idAct idRed";
	  getElmID("msgT").innerHTML = swID.innerHTML;
	  getElmID("msgA").innerHTML = getElmID("box"+i+"A").innerHTML;
	  getElmID("msgB").innerHTML = getElmID("box"+i+"B").innerHTML;
	}else{
	  swID.className="idCom";
  } }
}
msgShow(1);
</script>
</body>
</html>
