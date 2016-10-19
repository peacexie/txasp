<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/func2.asp"-->
<!--#include file="../../sadm/func1/md5_func.asp"-->
<!--#include file="form_app.asp"-->
<%

Call Chk_URL() 
get30Time = RequestS(App30Code&"0","",48) 
get30TSN = RequestS(App30Code&"1","",96) 

'// act=setInput,getForm,getFAdm,
goRef = Request.ServerVariables("HTTP_REFERER")
act = Request("act") :if act="" Then act="[def]"
goPage = Request("goPage")
verMemb = Request("verMemb")
urlBase = "&goPage="&goPage&"&verMemb="&verMemb&""
scr = "script" '//die()


if act="setInput" Then
	
  flag = true
  if A99ChkTime(get30Time,0.1) Then flag = false
  if get30TSN<>MD5_32(App30Code&"@"&get30Time) Then flag = false
	
  sType = Request("t") :if sType="" Then sType="text"
  sName = Request("n") :sName = Get_IDUnc(sName,0)
  sValue= Request("v")
  sSize = Request("s")
  sMax  = Request("m")
  sClass= Request("c")
  sStyle= Request("y")
  sEvent= Request("e") '//:sEvent = Replace()
  
  s = "<input type='"&sType&"' name='"&sName&"' id='"&sName&"' value='"&sValue&"' "
  s =s& " size='"&sSize&"' maxlength='"&sMax&"' "
  
  if sClass<>"" Then s=s& " class='"&sClass&"' "
  if sStyle<>"" Then s=s& " style='"&sStyle&"' "
  if sEvent<>"" Then s=s& " "&sEvent&" "
  
  s = s& " />"
  s = Replace(s,"'","\'")
  s = Replace(s,"""","\""")
  if flag Then 
    Response.Write "document.write('"&s&"')"
  else 
    Response.Write "document.write('<span style=\'color:#f00\'>Timeout Error[setInput]!</span>')"
  end if
  
  Response.End()

elseif act="showMessage" Then
'// 还原显示信息(base64_encode(Get_IDEnc("msg")))
	
  flag = true
  if A99ChkTime(get30Time,0.1) Then flag = false
  if get30TSN<>MD5_32(App30Code&"@"&get30Time) Then flag = false
  
  s = Request("s")
  's = base64_decode(s)
  s = Get_IDUnc(s,0)
  s = Replace(s,"'","\'")
  s = Replace(s,"""","\""")
  if flag Then 
    Response.Write "document.write('"&s&"')"
  else 
    Response.Write "document.write('<span style=\'color:#f00\'>Timeout Error[showMessage]!</span>')"
  end if
  
  Response.End()
  
elseif act="showMessag2" Then
'// 还原显示信息(base64_encode(Get_IDEnc("msg")))

  s = Request("s")
  's = base64_decode(s)
  s = Get_IDUnc(s,0)
  echo "document.write('s')"

  Response.End()
  
elseif act="xxxExt1" Then

  flag = true
  if A99ChkTime(get30Time,0.1) Then flag = false
  if get30TSN<>MD5_32(App30Code&"@"&get30Time) Then flag = false
  
  '//...
  
  if flag Then 
    Response.Write "document.write('"&s&"')"
  else 
    Response.Write "document.write('<span style=\'color:#f00\'>Timeout Error[showMessage]!</span>')"
  end if
  
  die("")

'//}else if($act=="getFAdm"){
elseif inStr("[Check1,[def]]",act)>0 Then
	
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=act%></title>
<link href="spub.css" rel="stylesheet" type="text/css">
<script src="../../inc/home/jsInfo.js" type="text/javascript"></script>
</head>
<body>
<%

if act="Check1" Then 
	
  smax = 180 '//0.1
  flag = true
  if A99ChkTime(get30Time,0.1) Then flag = false
  if get30TSN<>MD5_32(App30Code&"@"&get30Time) Then flag = false
	
  if flag Then 
  
  app99Tim = Now() 'date("Y-m-d H:i:s")
  app30Arr = App30Set(app99Tim,"","")
  app30Tab = Split(app30Arr(2),"-") 'explode("-",)
  
%> 
<div class="line12">&nbsp;</div>
<table width="480" border="0" align="center" cellpadding="5" cellspacing="3">
  <tr>
    <td colspan="2" align="center" style="line-height:150%; border:1px solid #063"><div id="msgLoad" class="cRed f14px"><img src="<?php echo $Config_upRoot ?>myfile/collect/Load.gif" width="18" height="18" align="absmiddle" />检测中</div></td>
  </tr>
  <tr>
    <td align="center" valign="top">注意</td>
    <td align="left" valign="top" class="cRed">请不要刷新！...</td>
  </tr>
  <tr>
    <td width="20%" align="center" valign="top">js检测结果</td>
    <td align="left" valign="top"><div id="msgJava" class="cRed">你禁用了js或不支持，你不能继续下一步操作！</div></td>
  </tr>
  <tr>
    <td align="center" valign="top">限时提示</td>
    <td align="left" valign="top">请在<span id="msgTime" class="cGreen" style="display:inline-block; padding:0px 3px;"><%=smax%></span>表内点[确定]，继续下一步操作！<br />
      否则超时。</td>
  </tr>
  <tr>
    <td colspan="2" align="center" style="line-height:150%; color:#333; border:1px solid #063">&nbsp;点 [<span id="btnOK"><a href="#" target="_self" onclick="sendForm()">确定</a></span>] 注册会员！&nbsp;点 [<a href="../../member/login.asp?act=<%=urlBase%>">取消</a>] 退出注册！</td>
  </tr>
</table>
<form id="fmApply" name="fmApply" method="post" action="../../member/mu_<%=AppRandom%>.asp">
  <input name="<%=App30Code%>0" type="hidden" value="<%=app30Arr(0)%>" />
  <input name="<%=App30Code%>1" type="hidden" value="<%=app30Arr(1)%>" />
  <input name="<%=App30Code%>2" type="hidden" value="<%=app30Arr(2)%>" />
  <input name="goPage" type="hidden" value="<%=goPage%>" />
  <input name="verMemb" type="hidden" value="<%=verMemb%>" />
</form>
<table width="480" border="0" align="center" cellpadding="5" cellspacing="3">
  <tr>
    <td align="left" style="font-size:13px; line-height:150%; color:#333;"><p class="fB">反“自动注册机器人”善意提醒：</p>
      <p>&nbsp;&nbsp;&nbsp;&nbsp;1. 如果有一天，你看到留言或会员资料中，每隔几秒钟就增加一条乱七八糟的资料，共有上十万或百万的这样的资料(99.9%都是乱七八糟的资料)……那时候再来作处理成本就太大了！</p>
      <p>&nbsp;&nbsp;&nbsp;&nbsp;2. 所以，所有留言，评论，注册会员等入口，一般都有认证码等相关的额外的操作，这都是防止一些自动程序在短时间内发布或注册大量信息。所以，请理解这些额外的操作！</p>
      <p>&nbsp;&nbsp;&nbsp;&nbsp;3. 如要求去掉这些额外的操作，如有自动注册或发布信息等现象，本程序概不负责！如保留这些额外的操作，还有自动注册或发布信息现象，我们会全力以赴，去阻止类似现象发生，感谢您的理解和支持。</p>
      <p align="right">Peace(XieYS)<br />
        2011-03-18 </p></td>
  </tr>
</table>
<script type="text/javascript">

var idTime = getElmID("msgTime")
var idOK = getElmID("btnOK")
function App_Check(){
  document.fmApply.submit(); //显示这页，则注释掉...
  var idJava = getElmID("msgJava")
  idJava.innerHTML = "你正在使用的系统支持js。"
  idJava.className = "cGreen"
  setTimeout("App_CTime()",1000)
  var idLoad = getElmID("msgLoad")
  idLoad.innerHTML = "检测完成！"
  idLoad.className = "cGreen"
}
App_Check()
var maxSec = "<%=smax%>"
function App_CTime(){
  idTime.innerHTML -= 1;
  if(idTime.innerHTML==0){
	idTime.innerHTML= "000"
	btnOK.innerHTML = "确定" 
	btnOK.className = "cDGray"
	idTime.className = "cRed"
  }else{
    setTimeout("App_CTime()",1000)
  }
}
function sendForm(){
  document.fmApply.submit()	
}

</script>
<%else%>
<table width="480" border="0" align="center" cellpadding="5" cellspacing="3">
  <tr>
    <td colspan="2" align="center" style="line-height:150%; border:1px solid #063"><div class="cRed f14px">错误：可能超时, 请重新提交！</div></td>
  </tr>
  <tr>
    <td width="20%" align="center" valign="top">错误提示</td>
    <td align="left" valign="top" class="cRed">请不要刷新！...</td>
  </tr>
  <tr>
    <td colspan="2" align="center" style="line-height:150%; color:#333; border:1px solid #063">点 [<a href="<%=goRef%>?">返回</a>] 退到前作页！</td>
  </tr>
</table>
<%End If%>
<%End If%>
</body>
</html>
<%End If%>