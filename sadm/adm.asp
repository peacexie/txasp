<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="func1/func1.asp"-->
<!--#include file="func2/func2.asp"-->
<!--#include file="func1/md5_func.asp"-->
<!--#include file="func2/cch_Class.asp"-->
<!--#include file="../inc/form/form_app.asp"-->

<%		

Call ChkSpider(1)
'Response.Write MD5_Adm("123456")
'Response.Write Session("ChkCode")&":"&uCase(Request("ChkCode")) 
'Session("UsrID") = "sitesa" ''// 忘记密码用这三行
'Session(UsrPStr) = "{Admin}" 
'Response.Redirect "../inc/adm_inc/adm_main.asp" 


Function admLogin(xSql)
 set rs = Server.CreateObject("adodb.recordset") ': Response.Write sql
 rs.Open xSql,conn,1,3
 if not rs.EOF then
   If rs("UsrType")="AdmStop" Then
	Msg = "禁止登陆 !"
   Else
	Session("UsrID") = UsrID
	Session(UsrPStr) = "{(MemFEditor);(MemFUpload)}"&rs("UsrPerm")&"{FileInfo,FileEditor,FileView,FileAdmin}" 
	rs("UsrLogIP") = Get_CIP()
	rs("UsrLTime") = Now()
	rs.UPDATE()
	Msg = "Welcom !"
   End If
  else
	Msg = "-=> 帐号不存在; 密码错误; 帐号禁用 或帐号到期 ! "
  end if
  rs.Close()
  set rs = nothing
  '/////////////////// Check for 子站
  If Len(Config_Path)>5 And lCase(Left(Config_Path,3))="/u/" Then
   Session("Pub_Subs") = Config_Path
  End If
  '//////////////////////////////////
  if Msg = "Welcom !" then
		Call Add_Log(conn,Session("UsrID"),"登入系统","[sadm_login]","")
	    g = rs_Val(conn,"SELECT ParText FROM AdmPara WHERE ParCode='Grop_"&Session("UsrID")&"'")
	    if(g<>"") Then g="?g="&g
		Response.Redirect "../inc/adm_inc/adm_main.asp"&g 
  end if
  admLogin = msg
End Function

uOuts = "|" '//外部登陆地址列表,请先检测外部地址的权限!
uOuts =uOuts& "http://b.dg.gd.cn/admin.php?action=index|"
uOuts =uOuts& "http://www.dg.gd.cn/sys/sadm/login.asp|"
uOuts =uOuts& "http://trade.dg.gd.cn/sys/sadm/login.asp|"
uOuts =uOuts& "http://old.dg.gd.cn/sys/sadm/login.asp|"
'//echo "$uOuts<br>$rPage";
rPage = Request.Servervariables("HTTP_REFERER")&"" :If rPage="" Then rPage="#@!"
If inStr(uOuts,rPage)>0 Then
  UsrID = RequestF("UsrID","C",48) '; //echo "$uOuts<br>$rPage"; 
  msg = admLogin("SELECT * FROM [AdmUser"&Adm_aUser&"] WHERE UsrID='"&UsrID&"' AND UsrType LIKE 'Adm%'")
  send = "@#$"
Else
  send = Request("send") 
End If	


If send = "send" Then
Call Chk_URL()

  get30Time = Request(App30Code&"0")
  get30TSN = Request(App30Code&"1") 
  get30TStr = Request(App30Code&"2")
  get30Tab = Split(get30TStr,"-")
  
  UsrID   = Trim(RequestF(get30Tab(3),3,96))
  UsrPW   = Trim(RequestF(get30Tab(5),3,96))
  ChkCode = Request.Form("ChkCode"&get30Tab(7))&""
  UsrType = RequestF("UsrType"&get30Tab(1),3,84)

 App26Str = "" : App26Chk = ""
 App26Day = Get_AppDay()+47 '每天变化一次
 For i=1 To Len(App26Code)
   App26Str = App26Str&Request(Mid(App26Code,i,1))  'Chr(96+i)
   App26Chk = App26Chk&Chr(App26Day+i)              'Chr(96+i)
 Next
 'Response.Write "<br>"&App26Str&"<br>"&App26Chk : Response.End()

 if Len(UsrID&"")<2 or Len(UsrPW&"")<5 then
   Msg = "-=> 错误 ! 请输入帐号密码! "
 elseif ChkCode="" Or Session("ChkAdmin")<>uCase(ChkCode) then
   Msg = "-=> 认证码错误! ["&Session("ChkAdmin")&"]-["&uCase(Request("ChkCode"))&"]"
   
 '/// 以下3个判断，增强安全性
 elseif Request("xxID")<>"(xxID)" Or Request("Name")<>"(Name)" Then
   msg = " 可能是软件自动填写的信息:系统禁止操作1!"
   Call Add_Log(conn,MemID,"ain:"&Request("xxID")&Request("Name"),"(*|)Login",sql)
 elseif Request("User")<>"(User)" Or Request("Pass")<>"(Pass)" Then
   msg = " 可能是软件自动填写的信息:系统禁止操作2!"
   Call Add_Log(conn,MemID,"ain:"&Request("User")&Request("Pass"),"(*|)Login",sql)
 elseif App26Chk="" Or App26Str<>App26Chk Then
   msg = "超时：请重新提交!"
 '/// 以上3个判断，增强安全性
   
 elseif Get_A30Chk(get30Time,get30TSN,60,"s")<>"" Then
    msg = "错误！超时错误 或 安全认证错误"
   
 elseif UsrType = "Login" then   
   eStr = MD5_Adm(UsrPW&UsrID)
   sql  = "SELECT * FROM [AdmUser"&Adm_aUser&"] "
   sql = sql& " WHERE UsrID='" &UsrID&"' AND UsrPW='"&eStr&"' "
   sql = sql& " AND (UsrType LIKE 'Adm%')  " 'AND UsrExp>#"&Date()&"# 
   msg = admLogin(sql)
 
 elseIf UsrType = "ReSet" then 
	'////
 else
    Msg = "-=> 错误 ! "   
 end if
 
 
ElseIf send = "TimOut" Then
    Msg = "-=> 超时，请重新登陆! "
ElseIf send = "out" Then
    Msg = "-=> 已经安全退出!"
	Session("UsrID") = empty
	Session(UsrPStr) = empty
	'Session("Pub_Subs") = ""
ElseIf send = "NoLogin" Then
    Msg = "-=> 请登陆!"
ElseIf send = "UsrID_ERROR" Then
    Msg = "-=> 错误!请登陆！"
ElseIf send = "getPW" Then 
    Msg = "-=> 取回密码 成功！谢谢使用!<br>请登陆系统修改您好记忆的密码！"
ElseIf send = "Illegal" Then
	Msg = "-=> Illegal Request!!"
	Call Add_Log(conn,Session("UsrID"),"登出系统_Illegal","[sadm_login]",Msg)
	Session("UsrID") = ""
	Session(UsrPStr) = ""
Else ' NoLogin
    Msg = "-=> 请输入帐号密码！" 
End If
u_agent = uCase(Request.ServerVariables("HTTP_USER_AGENT"))
u_yrDif = 0
If inStr(u_agent,"MSIE 5")>0 Then
  u_yrDif = 1999 :u_iever = 5
End If
If inStr(u_agent,"MSIE 6")>0  Then
  u_yrDif = 2001 :u_iever = 6
End If
u_yrNow = Year(Now) ':echo u_agent&u_yrDif&u_isie6
u_yrDif = u_yrNow-u_yrDif


app99Tim = Now() 'date("Y-m-d H:i:s");
app30Arr = App30Set(app99Tim,"s","")
app30Tab = Split(app30Arr(2),"-")

urlEvent = Server.HTMLEncode("class='lgnBtn1 lgnInpt' tabindex=1 onClick='ChkAShow()'") 
url30Par = "&"&App30Code&"0="&Server.HTMLEncode(app30Arr(0))&"&"&App30Code&"1="&app30Arr(1)&"" '//&".$App30Code."2=$app30Arr[2]
urlParas = "?act=setInput&n="&Get_IDEnc(app30Tab(3),0)&"&s=30&m=16&e="&urlEvent&url30Par&""
tmpStr = "?act=showMessage&s="&Get_IDEnc("PEACE123abc45&#20320;6",0)&url30Par&""

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta name="robots" content="noindex, nofollow">
<meta name="robots" content="noarchive">
<title><%=Config_Name%>-后台管理中心</title>
<link rel="stylesheet" type="text/css" href="../inc/adm_img/login.css"/>
<script src="../inc/home/jsInfo.js" type="text/javascript"></script>
<style type="text/css">
<!--
#hackTest {
	color:#fff;
}
.hackColor {
	background-color: #CC00FF;   /*所有浏览器都会显示为紫色*/
	background-color: #FF0000\9; /*IE6、IE7、IE8会显示红色*/
	*background-color: #0066FF;  /*IE6、IE7会变为蓝色*/
	_background-color: #009933;  /*IE6会变为绿色*/
}
-->
</style>
</head>
<body>
<script src="x../inc/form/form_js.asp<%=tmpStr%>" type='text/javascript'></script>
<div class="line20">&nbsp;</div>
<div class="line50">&nbsp;</div>
<!-- 仅限子目录Start -->
<table width="480" border="0" align="center" cellpadding="8" cellspacing="3" style="margin:auto; display:none">
  <tr>
    <td bgcolor="#FFFFCC" style="line-height:150%; border:1px solid #369">申明：本程序<span class="colF0F">只管理 XXX 网站 - YYY 部分</span>；与 其它栏目不相关；<br>
    即 如下页面对应的资料：<a href="../../(dir)/" target="_blank">http://www.(domain).com/(dir)/</a>。</td>
  </tr>
</table 仅限子目录>
<table width="480" border="0" align="center" cellpadding="0" cellspacing="0" style="margin:auto">
  <tr>
    <td colspan="2" class="lgnCnr1"><div class="lgnSubj"><%=Config_Name%> 后台管理登陆</div></td>
    <td class="lgnCnr2">&nbsp;</td>
  </tr>
  <tr>
    <td class="lgnSid1">&nbsp;</td>
    <td valign="top" class="lgnCont">
      
      <% If u_yrDif<>u_yrNow Then %>
      <table width="100%" border="0" cellpadding="2" cellspacing="0">
        <tr>
          <td colspan="2" align="left" valign="top" style="font-size:14px; padding:1px 3px 8px 5px">&nbsp;&nbsp;&nbsp;&nbsp;<span class="colF00">提示</span>： 您还在使用 <%=u_yrDif%> 年前的浏览器 --- "<span class="colF00">Internet Explorer <%=u_iever%></span>"！现在是 <%=u_yrNow%> 年，是一个现代Web标准的时代！是<span class="colF00">升级你的浏览器</span>的时候了！专题：<a href="http://www.ie6countdown.com/" target="_blank">全球IE6倒计时</a>(<a href="http://theie6countdown.cn/" target="_blank">中文版</a>)！</td>
        </tr>
      </table>
      <% End If%>
      
      <table width="100%" border=0 cellpadding=1 cellspacing=1 class="lgnTab">
      <form name='fm1<%=app30Tab(9)%>' method=post  action="?" >
        <tr>
          <td width="25%" align="right" valign=bottom nowrap>
                <div style="display:none;">
                  xxID(帐号)<input name="xxID" type="text" id="xxID" value="(xxID)">
                  Name(姓名)<input name="Name" type="text" id="Name" value="(Name)">
                  User(帐号)<input name="User" type="text" id="User" value="(User)">
                  Pass(密码)<input name="Pass" type="password" id="Pass" value="(Pass)">
                </div>
          &nbsp;用户名
          </td>
          <td valign=bottom nowrap bgcolor="#CCCCCC">
            <script src="../inc/form/form_js.asp<%=urlParas%>" type='text/javascript'></script>
            <input name="send" type="hidden" id="send" value="send"></td>
        </tr>
        <tr>
          <td align="right" valign=bottom nowrap>密　码</td>
          <td valign=bottom nowrap bgcolor="#CCCCCC">
            <input name="<%=app30Tab(5)%>" type=password id="<%=app30Tab(5)%>" class="lgnBtn1 lgnInpt" tabindex=2 size=30 maxlength="16" onClick="ChkAShow()">
            <input name="UsrType<%=app30Tab(1)%>" type="hidden" id="UsrType<%=app30Tab(1)%>" value="Login"></td>
        </tr>
        <tr>
          <td align="right" nowrap>认证码</td>
          <td nowrap>
            <div id="ChkText" style="color:#999; padding:5px 0px">点用户名输入框,激活认证码输入框。</div>
            <div id="ChkCode" style="display:none">
            <input name="ChkCode<%=app30Tab(7)%>" type="text" id="ChkCode<%=app30Tab(7)%>" size="6" maxlength="12" class="lgnInpt" tabindex="3" onFocus="ChkCShow()">
            <img src="../img/blank.gif?../sadm/pcode/img_frnd.asp?xConfig_PCode=hij&Config_PSess=ChkAdmin" alt="点击图片换一个" name="ChkCImg" align="absmiddle" id="ChkCImg" style="cursor:pointer;display:none" onClick="PicReLoad('../','','ChkAdmin')" ('../','hij') />
            <input name="<%=App27Random%>" type="hidden" id="<%=App27Random%>" value="<%=app30Tab(9)%>">
            </div>
          </td>
        </tr>
        <tr>
          <td align="center" nowrap><span class="col666"><a href="../" target="_blank">转到首页</a></span></td>
          <td nowrap><input type="submit" name="Submit" value="提 交" class="lgnBtn2" tabindex="4" />
              <span class="lgnBtnSp">&nbsp;</span>
            <input type="reset" name="Submit2" value="重 设" class="lgnBtn2" /></td>
        </tr>
        <input name="<%=App30Code%>0" type="hidden" id="<%=App30Code%>0" value="<%=app30Arr(0)%>" />
        <input name="<%=App30Code%>1" type="hidden" id="<%=App30Code%>1" value="<%=app30Arr(1)%>" />
        <input name="<%=App30Code%>2" type="hidden" id="<%=App30Code%>2" value="<%=app30Arr(2)%>" />
        <%
		
  App26Day = Get_AppDay()+47 '每天变化一次
  App26Str = ""
  For i=1 To Len(App26Code)
    c = Mid(App26Code,i,1)
    If c=Chr(34) Then
	  c="'"&c&"'"
	Else
	  c=""""&c&""""
	End If
	App26Str = App26Str&vbcrlf&"<input name="&c&" type='hidden' value='"&Chr(App26Day+i)&"' />"
  Next
  Response.Write App26Str
		
		%>
      </form>
    </table>
      <div class="line05"> &nbsp; </div>
      <table width="100%" 
            border="0" cellpadding="2" cellspacing="0">
        <tr align="center">
          <td colspan="2" align="left" nowrap="nowrap" class="lgnMsg">提示: <span id="ChkTMsg"><%=msg%></span></td>
        </tr> 
        <tr>
          <td colspan="2" align="left" valign="top" nowrap="nowrap" class="lgnRBox"><p class="lgnRead13 lgnLogoAdmin">
          &middot; <span title="其它浏览器，请自行设置好" style="cursor:pointer">推荐浏览器:IE7,IE8,</span>
          <span id="hackTest" class="hackColor">&nbsp;请启用Javascript&nbsp;</span>,
          <a href="../ext/xfile/ieTester.htm" target="_blank">测试浏览器</a>；<br />
          &middot; <a href="../tools/help/xhelp.asp" target="_blank"><span class="colF0F">第一次使用，请查看 《后台管理帮助文件》&gt;&gt;&gt;</span></a>；<br />
          &middot; 帐号密码区分大小写;帐号&gt;=2位,密码&gt;=5位；<br />
          &middot; 认证码不分大小写,如错误,请<a href="?"><span class="col00F">刷新</span></a>登陆；<br />
          &middot; 技术支持: <a href="http://www.dg.gd.cn" target="_blank">东莞网(www.dg.gd.cn)</a>。
          </p></td>
        </tr>
    </table></td>
    <td class="lgnSid2">&nbsp;</td>
  </tr>
  <tr>
    <td colspan="2" class="lgnCnr3"></td>
    <td class="lgnCnr4"></td>
  </tr>
</table>

<%
If send = "out" Then
	Call cchClear()
End If
%>

<script type="text/javascript">

var bStr = "其它浏览器"; //Maxthon,Tencent,360se
if(isSafari) { bStr = "Safari"; } // Opera,WebKit,Gecko
if(isFirefox) { bStr = "Firefox"; } //Gecko
if(isChrome) { bStr = "Chrome"; } //WebKit

if(isIE) { bStr = "IE6以下"; }
if(isIE6) { bStr = "IE6"; }
if(isIE7) { bStr = "IE7"; }
if(isIE8) { bStr = "IE8"; }
if(isIE9) { bStr = "IE9"; }
getElmID("hackTest").innerHTML = "&nbsp;<span title='"+brsUAgent+"'>当前内核:"+bStr+"</span>&nbsp;";

  if(top.location!==self.location){top.location=self.location;}
  
  var idText = getElmID('ChkText'); 
  var idCode = getElmID('ChkCode');
  var idCImg = getElmID('ChkCImg');
  var idTMsg = getElmID('ChkTMsg'); //提示
  function ChkAShow(){ 
	if(idCode.style.display=='none')
	  { idCode.style.display=''; }
	idText.style.display='none';
	idTMsg.innerHTML = "点认证码输入框,显示认证码, 认证码不分大小写";
  }
  function ChkCShow(){
	if(idCImg.style.display=='none')
	  { idCImg.style.display=''; PicReLoad('../','','ChkAdmin'); }
  }
  
</script>
<!-- 
项目名称: <%="[ "&Config_Name&" ]"%> 后台
程序开发: 东莞网: Peace(谢永顺) 
开发时间: 20yy-mm-dd ~ 20yy-mm-dd 
-->
</body>
</html>