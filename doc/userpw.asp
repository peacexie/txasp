<!--#include file="dinc/_config.asp"-->
<%
p = Request("p") 
If p="MemB624" Then 
  p="../smod/gbook/info_list.asp?ModID=MemB624&PrmFlag=(Inn)"
ElseIf p="MemB524" Then 
  p="../smod/gbook/info_list.asp?ModID=MemB524&PrmFlag=(Inn)"
ElseIf p="MemB424" Then 
  p="../smod/gbook/info_list.asp?ModID=MemB424&PrmFlag=(Inn)"
Else
  p="../sadm/user/user_editpw.asp?PrmFlag=(Inn)"
End If

w = RequestS("w","N",950) 
h = RequestS("h","N",350)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=sysName%></title>
<link rel="stylesheet" type="text/css" href="dinc/style.css">
<script src="../inc/home/jsInfo.js" type="text/javascript"></script>
<script type="text/javascript">
<!--#include file="../upfile/sys/doc/list_depart.js"-->
</script>
<script src="dinc/_funcs.js" type="text/javascript"></script>
</head>
<body>
<!--#include file="dinc/inc_top.asp"-->
<div align="center" class="sysCMid">
  <table width="950" border="0" cellpadding="0" cellspacing="0">
    <tr>
      <td width="8%" align="center">会员中心</td>
      <td align="left"> &nbsp;<a href="?">修改密码</a> || <a href="?p=MemB624">个人笔记</a> | <a href="?p=MemB524">系统通知</a> | <a href="?p=MemB424">写给网管</a></td>
    </tr>
  </table>
  <IFRAME name=LeftMenu src="<%=p%>" frameBorder=0 width="<%=w%>" height="<%=h%>" style="overflow-x:hidden;overflow-y:scroll;"></IFRAME>
</div>
<%jsFlag="mTopUser"%>
<!--#include file="dinc/inc_bot.asp"-->
</body>
</html>
