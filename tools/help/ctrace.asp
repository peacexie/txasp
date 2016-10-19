<!--#include file="../himg/tconfig.asp"-->
<!--#include file="../../upfile/sys/pcfg/link.asp"-->
<%
Call Chk_Perm1("","")

'act=index(def) [action]
'mod=web(def) [module]
'usr=客户名称 [user]
'ver=版本号 [version]


act = Request("act") 'send,reset,update,error
If act="" Then act="send" 


If act="reset" Or act = "send" Then
  Dim bUrl :bUrl = LinkService 
  Dim sRnd :sRnd = "?Peace_"&Rnd_ID("",12)&"_RndID="&Timer()&""
  Dim tUrl :tUrl = bUrl&"strace.htm"&sRnd
  Dim sUrl :sUrl = bUrl
  aTrace = getStatus(tUrl) '// 发送 Http对象，返回状态数组
  If(cStr(aTrace(0))="200") Then
    act = "send" 'Response.Redirect sUrl
	tpos = inStr(aTrace(2),"(")
	tfil = Mid(aTrace(2),tpos+1)
	tfil = Replace(tfil,")","")
	sUrl = sUrl&tfil&sRnd 
	'echo sUrl
  Else
    act = "error"
  End If
ElseIf act = "update" Then
  '更新...
  tVal = RequestS("tVal","C",96)
  tID = RequestS("tID","C",24)
  Call rs_DoSql(conn,"UPDATE AdmPara SET ParText='"&tVal&"' WHERE ParCode='"&tID&"'")
End If
'echo act
'echo aTrace(0)&aTrace(1)&aTrace(21)
'Response.End()


%>
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>服务跟踪 - 后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<script src="../../inc/home/jsInfo.js" type="text/javascript"></script>
<style type="text/css">
p,blockquote,form {
	PADDING: 1px;
	MARGIN: 1px auto;
}
.bdBot {
	border-bottom:1px solid #CCC;
}
.bdFull {
	border:1px solid #CCC;
}
.idAct {
	background-color:#FFF;
	cursor:pointer;
}
.idCom {
	background-color:#E0E0E0;
	cursor:pointer;
}
.noBot {
	border-bottom:1px solid #F0F0F0;
}
.ntSide {
	width:240px;
	background-color:#F0F0F0;
	font-size:12px;
}
.pdCell {
	padding:8px;
}
</style>
</head>
<body>
<%If act="send" Then%>
<form id="fmSend" name="fmSend" method="get" action="<%=sUrl%>">
  <input name="smd" type="hidden" value="web" />
  <input name="usr" type="hidden" value="<%=Config_Name%>" />
  <input name="ver" type="hidden" value="<%=Config_Vers%>" />
</form>
<%ElseIf act="update" Then%>
<p align="center" class="SysCont"> 更新成功！<br>
1. <a href="../../sadm/system/upd_para.asp" target="_blank">刷新参数&gt;&gt;&gt;</a><br>
2. <a href="?">连接试尝&gt;&gt;&gt;</a></p>
<%ElseIf act="error" Then%>
<table width="99%" border="0" align="center" cellspacing="5" style="margin:5px; border:1px solid #CCC">
  <tr>
    <th class="bdBot" id="msgT"> 服务跟踪页 HTTP 错误 404</th>
    <td height="360" rowspan="2" align="left" valign="top" class="pdCell SysCont ntSide" id="msgB">提示：<br>
      &nbsp;&nbsp;&nbsp;&nbsp;0. 错误信息：<%=aTrace(0)%>; <%=aTrace(1)%>；<br>
      &nbsp;&nbsp;&nbsp;&nbsp;1. 如果仅显示本错误，并不影响你网站本身的正常运行！<br>
      &nbsp;&nbsp;&nbsp;&nbsp;2. 显示本错误提示，表示不能连接到服务商网站，这可能是暂时的；请不必要紧张，过一会儿再试尝！<br></td>
  </tr>
  <tr>
    <td height="360" align="left" valign="top" class="pdCell SysCont" id="msgA"><blockquote>
        <p>说明：</p>
      </blockquote>
      <ul>
        <li>设置本栏目的目的是让项目(网站)拥有者能够及时联系上项目开发公司的相关人员(<span class="colF0F">即使人员变动</span>),并为项目拥有者提供相关服务；</li>
        <li>本栏目相关连接及内容由项目开发公司维护更新；</li>
        <li>项目开发公司拥有对本栏目的最终解吸权利！</li>
        <li>本错误，您可通过 如下方式解决：</li>
      </ul>
      <blockquote>
        <p>1. 服务跟踪页 可能暂时不能访问，过一会儿再试尝连接：</p>
      </blockquote>
      <p align="center"> <a href="<%=sUrl%><%=sRnd%>&smd=web&usr=<%=Server.URLEncode(Config_Name)%>&ver=<%=Config_Vers%>">连接试尝&gt;&gt;&gt;</a>&nbsp;</p>
      <blockquote>
        <p>2. 服务跟踪页 可能更换地址或你的参数有误，请联系客服修改参数：</p>
      </blockquote>
      <form name="form1" method="post" action="?">
        <p align="center">服务跟踪页:
          <label for="tVal"></label>
          <input name="tVal" type="text" id="tVal" value="<%=LinkService%>" size="32">
          <input type="submit" name="button" id="button" value="修改参数">
          <input name="tID" type="hidden" id="tID" value="LinkService" />
          <input name="act" type="hidden" id="act" value="update" />
        </p>
      </form></td>
  </tr>
</table>
<%End If%>
</body>
</html>
<%
'/// asp函数 //////////////////////////////////////////////
Function getStatus(xUrl)
  Dim url,oHttp,aHttp(2)
  url = xUrl ':echo "1. "&url
  Set oHttp = Server.CreateObject("Msxml2.xmlHttp")
  With oHttp 
	.Open "GET",url,False," ", " " 
	.Send()
	aHttp(0) = oHttp.Status '200
	aHttp(1) = oHttp.StatusText 'OK
	aHttp(2) = oHttp.ResponseText 'Code
  End With 
  getStatus = aHttp
  Set oHttp = Nothing
End Function
%>
<script type="text/javascript">
  try{ document.fmSend.submit(); }
  catch(e) { ; }
/// js函数 //////////////////////////////////////////////
var url = "<%=url%>";
var oHttp = getXmlHttp(); 
function traceSend() {
  oHttp.open("GET",url, true); //拒绝访问
  oHttp.onreadystatechange = traceCheck;
  oHttp.send(null);
}
function traceCheck(){
  if (oHttp.readyState == 4) {
	location.href = traceUrl; 
  }
}
//traceSend(); 
</script>