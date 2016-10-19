<!--#include file="inc/_config.asp"-->
<!--#include file="inc/info_code.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=MDName%>-<%=Config_Name%></title>
<link href="inc/style.css" rel="stylesheet" type="text/css">
<script src="../inc/home/jsInfo.js" type="text/javascript"></script>
<script src="../ext/api/play/jsPlayer.js" type="text/javascript"></script>
</head>
<body>
<!--#include file="_head.asp"-->
<table cellspacing="0" cellpadding="0" width="950" align="center" border="0">
  <tr>
    <td width="10"><img height="42" src="../pfile/pimg/qqimg1_06_01.jpg" width="10" /></td>
    <td width="120" align="center" background="../pfile/pimg/qqimg1_06_02.jpg" class="SysTitle"><%=MDName%></td>
    <td width="10"><img height="42" src="../pfile/pimg/qqimg1_06_04.jpg" width="43" /></td>
    <td valign="top" background="../pfile/pimg/qqimg1_06_05.jpg"><table width="100%" border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td align="right" class="SysTit03" height="30"><%=vPMsg_WSite%> <a href="../index.asp"><%=vPMsg_WHome%></a> &gt;&gt; <a href="../trade/">供求黄页</a> &gt;&gt; <%=MDName%> &gt;&gt;&nbsp; <%=TPName%> &nbsp; </td>
        </tr>
      </table></td>
  </tr>
</table>
<table cellspacing="0" cellpadding="0" width="950" align="center" border="0">
  <tr>
    <td valign="top" style="BORDER:#E0E0E0 1px solid;"><table cellspacing="0" cellpadding="0" width="99%" border="0">
        <tr>
          <td height="500" colspan="4" align="left" valign="top" style="padding:5px;"><!--Item Start-->
            <!--#include file="inc/info_page.asp"-->
            <!--Item End--></td>
        </tr>
      </table></td>
    <td width="8"></td>
    <td style="BORDER: #b6d9f7 1px solid;" valign="top" width="210" height="240">
	<!--Side Start-->
	<!-- #include file="_side.asp" -->
    <!--Side End-->
    </td>
  </tr>
</table>
<!--#include file="_foot.asp"-->
</body>
</html>
