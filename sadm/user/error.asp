<!--#include file="../func1/func1.asp"-->
<!--#include file="../func2/func2.asp"-->
<html>
<head>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=Config_Name%>后台管理中心</title>
<link href="../../inc/adm_inc/adm_style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%

		  if Request("send") = "NoPerm" then
		    SysID = RequestS("SysID",3,24)
			SysID2= rs_val("","SELECT SysName FROM AdmSyst WHERE SysID='"&SysID&"'")
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
			msg2 = " <font color=red><b> 无权限操作！ </b></font>&nbsp;<font color=blue>["&SysID&"] "&SysID2 &"</font> <br> 该模块无权限！请与管理员联系。"
			'Call Add_Log(conn,Session("UsrID"),"无权限 ("&Act&") "&"["&Trim(SysID)&": "&SysID2&"]","[xSys]",Replace(sql,"'","|"))
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
    <td >&nbsp;</td>
    <td align="center" ><a href="?">重载本页</a></td>
  </tr>
</table>
</body>
</html>
