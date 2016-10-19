<!--#include file="_config.asp"-->
<%

If smcDebug = "isDebug" Then '调试标记
  Response.Redirect "../page/message.asp?mType=Debug&msg="&msg&"&goPage=goRef"
  Response.End()   
End If 

Dim res,act,actName,msg,tNumb,tCont
Dim sndType,sndUser,sndMaxs
act = Request("act")
If act="" Then act="Send"

sndType = "Member"
sndUser = Session("MemID")
sndMaxs = doBalance(sndType,sndUser,-1) 

If act="doSend" Then 
  actName = "发送短信"
  msg = doSend(sndType,sndUser,"")
ElseIf act="Group" Then 
  actName = "短信群发"
  chkFuncs = "chkGroup()"
ElseIf act="doGroup" Then 
  actName = "短信群发"
  msg = doGroup()
  chkFuncs = "chkGroup()"
Else 'act="Send"
  actName = "发送短信"
End If

%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>发送短信</title>
<link rel="stylesheet" type="text/css" href="../inc/spub.css"/>
<script src="../../sadm/func1/Func_JS.js" type="text/javascript"></script>
<script src="../inc/send_check.js" type="text/javascript"></script>
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="0">
</head>
<body>
<div class="line15">&nbsp;</div>
<!--#include file="../inc/send_form.asp"-->
<div class="line15">&nbsp;</div>
</body>
</html>
