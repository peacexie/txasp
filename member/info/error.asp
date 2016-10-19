<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/func2.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>用户中心 (Member Center)</title>
<link href="../../inc/mem_inc/mem_style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%
send = Request("send")
if send = "NoPerm" then
  'SysID = RequestS("SysID",3,24)
  'SysID2= rs_val("","SELECT MSTName FROM MemSyst WHERE MSTID='"&SysID&"'")
  Act = LCase(Request("Act"))
  if Act&"" = "" then
    Act = "^NULL^"
  elseif Act&"" = "upd" then
    Act = "修改"
  elseif Act&"" = "ins" then
    Act = "增加"
  elseif Act&"" = "del" then
    Act = "删除"
  elseif Act&"" = "sel" then
    Act = "查询"
  end if
  msg2 = " <font color=red><b> 无权限 </b></font>&nbsp;<font color=blue>["&SysID&"] "&SysID2 &"</font> - <font color=blue>["&Act&"] 操作</font> <br> 无权限！请与管理员联系。"
  'Call Add_Log(conn,Session("UsrID"),"无权限 ("&Act&") "&"["&Trim(SysID)&": "&SysID2&"]","[xSys]",Replace(sql,"'","|"))
  elseIf send = "SysKeyWD" Then
%>
    <div style="padding:8px;">
	<!--#include file="../../upfile/sys/para/kwords2.asp"-->
    </div>
<%
		  
    Response.Write "</body></html>"
	Response.End()
  else
    msg2 = "&nbsp;超时啦! 要么你还没有登陆?"
  end if


%>
<br>
<table  border="0" align="center" cellpadding="3" cellspacing="5" style="border:1px solid #CCC;">
  <tr>
    <td colspan="3" align="center"><%=msg2%></td>
  </tr>
  <tr>
    <td align="center" ><a href="/">转到首页</a></td>
    <td ><a href="../mu_<%=AppRand12%>.asp" target="_blank">注册帐号</a></td>
    <td align="center" ><a href="?">重载本页</a></td>
  </tr>
</table>
</body>
</html>
