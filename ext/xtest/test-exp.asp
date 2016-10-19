<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../../page/_config.asp"-->
<!--#include file="../../sadm/func2/upremote.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Test Code Demo</title>
</head>
<body>

<!--

' CreateObject( / OpenTextFile( / CreateTextFile(
' eval / execute
' VBScript.Encode / JScript.Encode
' #@~^ / ==^#~@

s = "["&"script|["&"iframe"&_
"|request(|execute |eval |execute(|eval("&_
"|createobject(|createtextfile("  '  |["&"%

[%@ Page Language="C#" Debug="true" trace="false" validateRequest="false"	%]
[%@ import Namespace="System.IO" %]

-->
<%

'过滤姓名: 198-1234-5678 [Xie永顺]
s1 =     "<br>135-1234-4561 [xx"&vbcrlf&"x1]"&vbcrlf
s1 = s1& "<br>135-1234-4562 [谢谢]"&vbcr
s1 = s1& "<br>135-1234-4563 []"&vblf
s1 = s1& "<br>135-1234-4533 ["&vbcrlf
s1 = s1& "<br>135-1234-4564 [谢永]<br>"
s2 = Show_RExp(s1,"\[[^>]*?\]","")
Response.Write "<br>Show_RExp: []("&s2&")<br>"

tStr = File_Read("test-exf.asp","utf-8") ':Response.Write tStr 
sStr = File_Read("test-exf.asp","utf-8") ':Response.Write tStr 
' Response.Write tStr 
' <object runat="server" id="ws" scope="page" classid="xx"></object>
' runat="server" classid="xx" 

Response.Write "<br>Show_RTest: inc("&Check_RTest(tStr,ckv2_Inc)&")<br>"
Response.Write "<br>Show_RTest: srv("&Check_RTest(tStr,ckv3_srv)&")<br>"
Response.Write "<br>Show_RTest: out("&Check_RTest(tStr,ckv5_out)&")<br>"
Response.Write "<br>Show_RTest: code("&Check_RTest(tStr,ckv4_code)&")<br>"
Response.Write "<br>Show_RTest: enc("&Check_RTest(sStr,ckv1_enc)&")<br>"
Response.Write "<br>Show_RTest: net("&Check_RTest(tStr,ckv0_net)&")<br>"


tTimer1 = Timer()

%>
</pre>
</body>
</html>
