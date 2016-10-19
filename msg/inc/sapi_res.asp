<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Response</title>
<link rel="stylesheet" type="text/css" href="../inc/spub.css"/>
<style type="text/css">
.pTitle{
	padding-left:120px;
}
</style>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
</head>
<body>
<div class="line15">&nbsp;</div>
<table width="610" border="1" align="center" cellpadding="3" style="margin:5px;">
  <%If act="Readme" Then%>
  <tr>
    <td colspan="2" align="left" class="fB pTitle">Notes-返回值说明</td>
  </tr>
  <tr>
    <td align="center">0</td>
    <td>OK! 成功！</td>
  </tr>
  <!--#include file="inc_error.asp"-->
  <%Else%>
  <tr>
    <td colspan="2" align="left" class="fB pTitle">RequireInfo-接收参数</td>
  </tr>
  <tr>
    <td align="center">tm</td>
    <td><%=tm%></td>
  </tr>
  <%If tel<>"" Then%>
  <tr>
    <td align="center">tel</td>
    <td><%=tel%></td>
  </tr>
  <%End If%>
  <%If ct1<>"" Then%>
  <tr>
    <td align="center">ct1</td>
    <td><%=ct1%></td>
  </tr>
  <%End If%>
  <%If npw<>"" Then%>
  <tr>
    <td align="center">npw</td>
    <td><%=npw%></td>
  </tr>
  <%End If%>
  <tr>
    <td align="center">cs</td>
    <td><%=cs%></td>
  </tr>
  <tr>
    <td align="center">act</td>
    <td><%=act%></td>
  </tr>
  <tr>
    <td align="center">re</td>
    <td><%=re%></td>
  </tr>
  <tr>
    <td colspan="2" align="left" class="fB pTitle">ResponseInfo-返回参数</td>
  </tr>
  <tr>
    <td align="center">reState</td>
    <td><%=reState%></td>
  </tr>
  <tr>
    <td align="center">reInt</td>
    <td><%=reInt%></td>
  </tr>
  <tr>
    <td align="center">reInfo</td>
    <td><%=reInfo%></td>
  </tr>
  <tr>
    <td align="center">reTime</td>
    <td><%=tm%></td>
  </tr>
  <tr>
    <td colspan="2" align="left" class="fB pTitle">Return-返回处理</td>
  </tr>
  <%End If%>
  <tr>
    <td width="20%" align="center">&nbsp;</td>
    <td align="center">请<a href="<%=rePage%>">返回</a>！</td>
  </tr>
</table>
<div class="line15">&nbsp;</div>
</body>
</html>