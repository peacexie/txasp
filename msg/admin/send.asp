<!--#include file="_config.asp"-->
<%
Dim res,act,actName,msg,tNumb,tCont
Dim sndType,sndUser,sndMaxs
act = Request("act")
If act="" Then act="Send"

sndType = "(SmsAPI)" '(SmsAPI),Admin(不限制),Inner,Message,Member(SmsMember)
sndUser = "demo" '
sndMaxs = doBalance(sndType,sndUser,-1) 

chkFuncs = "chkSend()"
If act="Test" Then
  actName = "发送测试"
  tCont = ""&Time()&"测试@"&Get_CIP()&"("&Session("UsrID")&")"
  msg = "已经输入["&Len(tCont)&"]个字"
ElseIf act="doTest" Then 
  actName = "发送测试"
  res = doTest()
  msg = res(2)&" "&Now()
ElseIf act="Group" Then 
  actName = "短信群发"
  chkFuncs = "chkGroup()"
ElseIf act="doGroup" Then 
  actName = "短信群发"
  msg = doGroup()
  chkFuncs = "chkGroup()"
ElseIf act="doSend" Then 
  actName = "发送短信"
  msg = doSend(sndType,sndUser,"")
Else 'act="Send"
  actName = "发送短信"
End If

Session("ChkCode") = Rnd_ID("",8)
If act="Test" Then
  tNumb = "198-1234-5678 [Xie永顺], 198-1234-5648 [Xie永顺]" '测试
  tCont = ""&Time()&"测试@"&Get_CIP()&"("&Session("UsrID")&")" '测试
Else
  tNumb = Request("tNumb")
  tCont = Request("tCont") 
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
<%If inStr(act,"Test")>0 Then%>
<!--#include file="../inc/send_test.asp"-->
<%End If%>
<%If inStr(act,"Send")>0 Or inStr(act,"Group")>0 Then%>
<!--#include file="../inc/send_form.asp"-->
<%End If%>
<div class="line15">&nbsp;</div>
</body>
</html>
