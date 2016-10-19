<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
<title>支付结果</title>
</head>
<body>


<%
'Option Explicit
'Session.CodePage = 65001
'Session.LCID = 2057
Response.charset = "utf-8"
%>
<!--#include file="AliPay_Class.asp"-->
<%
    Dim alipay
	Set alipay = New Qlwz_AliPay
%>
<div style="line-height:12px">&nbsp;</div>
<table width="420" border="0" align="center" cellpadding="2" cellspacing="1">
  <%If alipay.Return_url Then%>
  <tr>
    <td width="180" align="center"><img src="../../img/logo/logo.jpg" width="160" height="60" /></td>
    <td style="line-height:180%;"><span class="colF00">付款成功</span>! 您可以：<br />
      <a href="/member/mu_info.asp?yAct=ViewOrder"><font color="#0000FF">查看定单</font></a>&nbsp; 或进入 <a href="../inc/mem_inc/mem_main.asp" target="_blank"><font color="#0000FF">会员后台</font></a><br />
      或  <a href="/">返回首页</a></td>
  </tr>
  <%Else%>
  <tr>
    <td align="center"><img src="../../img/logo/logo.jpg" width="160" height="60" /></td>
    <td><span class="colF00">付款失败</span>！请 <a href="/">返回首页</a></td>
  </tr>
  <%End If%>
</table>
<!--
	If alipay.Return_url Then    'Return_url返回 True或False
		response.write "交易成功！"
	Else
		response.write "交易失败！"
	End If
-->
<%
	Set alipay = Nothing
%>



</body>
</html>