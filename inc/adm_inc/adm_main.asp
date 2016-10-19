<!--#include file="../../sadm/func2/func_const.asp"-->
<!--#include file="../../sadm/func2/func_perm.asp"-->

<%
Call Chk_Perm9("","(Adm)")
%>
<%Response.CodePage=65001%>
<%Response.Charset="UTF-8"%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta name="robots" content="noindex, nofollow">
<title><%=Config_Name%>后台管理</title>
<STYLE type=text/css>
BODY {
	SCROLLBAR-FACE-COLOR: #d9e5f6; FONT-SIZE: 12px; MARGIN: 0px; SCROLLBAR-HIGHLIGHT-COLOR: #ffffff; OVERFLOW: hidden; SCROLLBAR-3DLIGHT-COLOR: #d9e5f6; SCROLLBAR-TRACK-COLOR: #f3faf4; SCROLLBAR-DARKSHADOW-COLOR: #f3faf4; BACKGROUND-COLOR: #f7fbff
}
td,tr,th,body,p {
	FONT-SIZE: 12px; LINE-HEIGHT: 150%; LETTER-SPACING: 1px
}
A:link {
	COLOR: #06305a; TEXT-DECORATION: none
}
A:visited {
	COLOR: #06305a; TEXT-DECORATION: none
}
A:hover {
	COLOR: #ff0000; TEXT-DECORATION: none
}
A:active {
	COLOR: #0067bf; TEXT-DECORATION: none
}
.btnSystem{ width:67px; height:24px; background-image:url(../../inc/adm_img/login01.gif); border:0; cursor:hand; }
</STYLE>
</head>
<body>
<table width="100%" height="100%"  border="0" cellpadding="1" cellspacing="1" bgcolor="#418DC3">
  <tr bgcolor="#FFFFFF">
    <td width="160" rowspan="2" align="center" valign="top"><IFRAME name=LeftMenu src="adm_menu.asp?g=<%=Request("g")%>" 
      frameBorder=0 width="160" scrolling="no" height="100%"></IFRAME></td>
    <td height="10" valign="top"><table width="100%"  border="0" cellpadding="1" cellspacing="1">
        <tr>
          <td width="1%" rowspan="2" align="left" nowrap style="border-right:1px solid #F0F0F0"><img src="../../pfile/pub/admin.jpg" height="60" hspace="2" vspace="2"></td>
          <td colspan="2" align="left" nowrap><a 
	href="../../" target="_blank"><font color="#FF0000">前台</font></a> | <a 
	href="../../sadm/user/index.asp" target="mainFrame"><font color="#FF0000FF">后台</font></a>
<%If Chk_PermSP() Then%>
<!--显示[配置,工具]快捷键；管理模式Config_Mode=isExpert下可见-->
[
<a target='mainFrame' href='../../sadm/admin/sys_config.asp'>配置</a>
<a target='mainFrame' href='../../tools/tools.asp?Act=Admin'>工具</a>
]
<%End If%>    
          </td>
          <td width="1%" align="center" valign="bottom"><input type="button" name="Button" value="<<后退" onClick="RefMainFrame(-1)" class="btnSystem"></td>
          <td width="1%" align="center" valign="bottom"><input type="button" name="Button" value="前进>>" onClick="RefMainFrame(1)" class="btnSystem"></td>
        </tr>
        <tr>
          <td valign="middle" nowrap>
		  当前管理员:<font color="#0000FF"><%=Session("UsrID")%></font>
		  </td>
          <td align="right" valign="middle" nowrap><IFRAME 
		  name=RefFrame src="../../tools/out/check.asp?rem=OK" frameBorder=0 width="30" scrolling=no height="15"></IFRAME></td>
          <td width="1%" align="center" valign="top"><input type="button" name="Button" value="当前页" onClick="RefMainFrame(0)" class="btnSystem"></td>
          <td width="1%" align="center" valign="top"><input type="button" name="Button" value="[退出]" onClick="javascript:location='../../sadm/<%=Config_RAdm%>.asp?send=out&UsrType=adm'" class="btnSystem"></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF"><IFRAME name=mainFrame src="../../sadm/user/index.asp" 
      frameBorder=0 width="100%" scrolling="yes" height="100%"></IFRAME></td>
  </tr>
</table>
</body>
</html>
<script type="text/javascript">
function RefMainFrame(xPage)
{
self.parent.frames["mainFrame"].history.go(xPage);
}
</script>
