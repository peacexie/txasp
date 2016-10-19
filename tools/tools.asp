<!--#include file="himg/tconfig.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<base target="_blank" />
<title>目录</title>
<style type="text/css">
<!--
.div { 
    width:157px; 
	height:120px;
	font-size: 12px;
	overflow-x:hidden;
	overflow-y:scroll;
	float:left;
	padding:5px 1px 1px 5px;
	margin:0px 10px 10px 0px;
	border:1px solid #CCC;
}
div.div a {
	display:inline-block;
}
/* ul, ol, li,  */
div, form, img, ul, ol, li, dl, dt, dd {
	margin: 0;
	padding: 0;
	border: 0;
}
body {
	margin:5px;
}
td {
	background-color:#FFF;
	text-align:center;
}
form {
	margin:0px;
	padding:2px;
}
.fsel {
	width:60px;
}
.fbtn {
	width:60px;
}
.bWidth {
	width:720px;
	margin:auto;
	text-align:left;
}
.bBorder {
	border:1px #CCCCCC solid;
}
.bItem {
	width:98%;
	margin:1px;
	padding:5px 5px;
	text-align:left;
	border-bottom:1px #E0E0E0 dashed;
}
.bItm1 {
	width:110px;
	height:21px;
	float:left;
	text-align:right;
	padding:0 3px 0 1px;
	color:#999999;
}
.bItm2 {
	display:inline-block;
}
.bItm3 {
	width:160px;
	height:18px;
	display:inline-block;
	float:left;
	line-height:100%;
}
body, td, th {
	font-size: 14px;
}
.Fnt00F {
	color:#00F;
	padding:0 3px 0 3px;
}
.FntF00 {
	color:#00F;
	padding:0 3px 0 3px;
}
a:link {
	color: #330066;
	text-decoration: none;
}
a:visited {
	text-decoration: none;
	color: #330066;
}
a:hover {
	text-decoration: underline;
	color: #FF0000;
}
a:active {
	text-decoration: none;
	color: #FF00FF;
}
.mAdmBox {
	width:320px;
	height:240px;
	overflow:scroll;
	padding:8px;
}
-->
</style>
</head>
<body>
<%
Call Chk_Perm1("SysTools","") 
%>
<center>
  <div class="bWidth">
    <fieldset style="padding:3px;">
    <legend> &nbsp; <b class="Fnt00F"> 帮助 与 说明 </b>&nbsp; </legend>
    <table width="700" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
      <tr>
        <td width="20%">帮助文件</td>
        <td width="20%"><a href="help/xhelp.asp" target="_blank">后台管理帮助文件</a></td>
        <td width="20%"><a href="help/uhelp.asp" target="_blank">会员帮助文件</a></td>
        <td width="20%"><a href="../smod/gbook/info_list.asp?ModID=GboU124">站务笔记</a></td>
        <td width="20%"><a href="../tools/out/admlogs.asp" target="_blank">导出笔记</a></td>
        </tr>
      <tr>
        <td width="20%">工具与扩展</td>
        <td width="20%">&nbsp;</td>
        <td width="20%">&nbsp;</td>
        <td><a href="../ext/api/scan/aspcheck.asp" target="_blank">阿江ASP 探针</a></td>
        <td><a href="../tools/help/xhelp.asp#FlagComm" target="_blank">常用配置</a></td>
        </tr>
    </table>
    </fieldset>
  </div>
</center>
<center>
</center>
<center>
  <div class="bWidth">
    <fieldset style="padding:3px;">
    <legend> &nbsp; <b class="Fnt00F"> JS 与 其它 </b>&nbsp; </legend>
    <table width="700" border="0" align="center" cellpadding="1" cellspacing="1" bgcolor="#CCCCCC">
      <%If Request("Act")="Admin" Or Request("Act")="Hidden" Then%>
      <tr>
        <td width="20%"><strong>[开发常用]</strong></td>
        <td width="20%"><a href="../ext/xfile/ieTester.htm">IE测试01</a></td>
        <td width="20%"><a href="../ext/xfile/ieTest02.htm">测试02</a></td>
        <td width="20%"><a href="help/vver2x.asp">版本更新Ver2.X</a></td>
        <td width="20%"><a href="help/vhelp.asp" target="_self">开发者说明文件</a></td>
      </tr>
      <tr>
        <td><strong>[管理常用]</strong></td>
        <td><a href="../sadm/system/logs_list.asp?send=dReset">清理Logs</a></td>
        <td><a href="../smod/gbook/info_list.asp?ModID=GboU224">私人秘密</a></td>
        <td><a href="../msg/admin/logs.asp?yAct=dReset">Sms</a></td>
        <td><a xhref="#../sadm/admin/sys_dbacc.asp?send=PExt02">压缩PubData</a></td>
      </tr>
      <tr>
        <td>刷新杂项</td>
        <td><a href="../smod/type/type_data.asp?Group=Inf">数据处理/移动</a></td>
        <td><a href="../smod/adupd/upd_data.asp">数据转化/清理</a></td>
        <td><a href="../smod/adupd/upd.asp?send=TimTest">效率测试</a></td>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td>系统配置</td>
        <td><a href="../sadm/admin/sys_config.asp">站点配置</a>-<a href="../sadm/admin/sys_1key.asp">1Key配置</a></td>
        <td><a href="../sadm/admin/typs_list.asp">信息类别</a>-<a href="../sadm/admin/type_list.asp">系统类别</a></td>
        <td><a href="../sadm/admin/sys_config.asp?yAct=upd">刷新配置</a></td>
        <td><a href="../sadm/admin/typs_list.asp?yAct=ResType">初始化菜单</a></td>
      </tr>
      <tr>
        <td>广告计数</td>
        <td><a href="../tools/base/omaster.asp">统计计数</a></td>
        <td><a href="../smod/link/ad_list.asp">浮动广告</a></td>
        <td><a href="../smod/link/ad_pic.asp">图片广告</a></td>
        <td><a href="../smod/link/ad_text.asp">文字广告</a></td>
      </tr>
      <tr>
        <td><strong>[专题应用]</strong></td>
        <td><a href="base/clear.asp?SysMod=DBOut">导出表结构</a></td>
        <td><a href="../member/ecard/admimp_code.asp" target="_blank">代码生成器</a></td>
        <td><a href="out/code.asp" target="_self">代码导出</a></td>
        <td><a href="out/down.asp" target="_self">批量下载</a></td>
      </tr>
      <tr>
        <td><a href="../ext/api/scan/sind.asp">Asp木马扫描器</a></td>
        <td><a href="base/clear.asp" target="_blank">木马扫把</a></td>
        <td>IIS查询分析器</td>
        <td><a href="base/sadm.asp" target="_self">安全中心</a></td>
        <td><a href="../msg/out/smsapi.asp" target="_self">短信API</a></td>
      </tr>
      <tr>
        <td><strong>[JS,CSS]</strong></td>
        <td>&nbsp;</td>
        <td><a href="../ext/xtest/js-keys.htm">js-关键字</a></td>
        <td><a href="../ext/xtest/js-page.asp">js-分页</a></td>
        <td><a xhref="#js/peace.htm">JS 效果收集</a></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td>&nbsp;</td>
        <td><a href="../ext/api/menu/peaceMenu01.htm">menu-peace菜单</a></td>
        <td><a href="../ext/api/menu/SpryMenu21Bar.asp">menu-spyr菜单</a></td>
        <td><a href="../img/rnd_nid/rbox_nid.htm">圆角效果综合</a></td>
      </tr>
      <tr>
        <td><strong>[ASP代码]</strong></td>
        <td><a href="../ext/xtest/uDepart.asp">uDepart</a></td>
        <td><a href="../ext/xtest/uInclude.asp">uInclude</a></td>
        <td><a href="../ext/api/scan/aspcheck.asp?T=E">ServerVars</a></td>
        <td><a href="out/rate-fee.asp">rate-fee</a></td>
      </tr>
      <tr>
        <td>&nbsp;</td>
        <td><a href="../ext/xtest/func-ID.asp">func-ID测试</a></td>
        <td><a href="../ext/xtest/func-pcode.asp">func-pcode测试</a></td>
        <td><a href="../ext/xtest/func-rs.asp">func-rs测试</a></td>
        <td><a href="../ext/xtest/func-sfile.asp">func-sfile测试</a></td>
      </tr>
      <tr>
        <td><a href="../ext/xtest/tperm.asp" target="_blank"></a></td>
        <td><a href="../ext/xtest/tperm.asp" target="_blank">Test Perm</a></td>
        <td>&nbsp;</td>
        <td><a href="../ext/xtest/tperm.asp" target="_blank"></a></td>
        <td><a href="out/outinfo.asp">外部连接抓取</a></td>
      </tr>
      <%End If%>
    </table>
    </fieldset>
  </div>
</center>
<%If Request("Act")="Admin" Then%>

<%

str = File_Read("../inc/adm_inc/bak_menu.asp","utf-8") ' Replace(str,"../../","../") 
str = Replace(str,"onClick=","") 
str = Replace(str,"<li class='left_top_right right'></li>","") 
str = Replace(str,"<div class='left_end'></div>","") 
str = Replace(str,"<div>","<div class='div'>") 
str = Replace(str,"../../","../")

st1 = File_Read("../upfile/sys/config/sf_Groups.htm","utf-8")
st1 = Replace(st1,"?ModID","../sadm/system/system.asp?ModID")
st2 = File_Read("../upfile/sys/config/sf_Module.htm","utf-8")
st2 = Replace(st2,"?ModID","../sadm/system/system.asp?ModID")
st3 = File_Read("../upfile/sys/config/sf_Para.htm","utf-8")
st3 = Replace(st3,"=para_","=../sadm/system/para_")
st3 = Replace(st3,"<br>","<br> 参数：")
%>

<center>
  <div class="bWidth">
    <fieldset style="padding:3px;">
    <legend> &nbsp; <b class="Fnt00F"> 备份管理菜单 </b>&nbsp; </legend>
    <div class='div'>模块：<%=st2%><br>
    系统：<%=st1%></div>
    <div class='div'>参数：<%=st3%></div>
	<%=str%>
    </fieldset>
  </div>
</center>
<%End If%>
<div style="line-height:8px;">&nbsp;</div>
<center>
  <div class="bWidth">
    <fieldset style="padding:3px;">
    <legend> &nbsp; <b class="Fnt00F"> 说明：</b>&nbsp; </legend>
    <div class="bItem" style="line-height:5px;">&nbsp;</div>
    <div class="bItem"><span class="bItm1">说明1&nbsp;:</span><span class="bItm2"> 请用于合法场合；必要的话，使用前请设置好参数.</span></div>
    <div style="line-height:8px;">&nbsp;</div>
    </fieldset>
  </div>
</center>
</body>
</html>
