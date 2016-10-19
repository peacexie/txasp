<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>rs_AutID Demo</title>
</head>
<body>

<!--#include file="../../page/_config.asp"-->
<!--#include file="../../sadm/func1/md5_func.asp"-->

<%

'session.timeout=5'999 '32768 525,600 
Server.ScriptTimeOut = 5

Response.Write "<pre>"

echo Get_FmtID("yymd-hnsx-r2s2","")
echo Get_FmtID("yymmdd-hnsx-r3xxx","")

echo Get_FmtID("yy-mm-dd","")
echo Get_FmtID("yyyy-md","")
echo Get_FmtID("hnsx-r3-r2","")
echo Get_FmtID("hnsx-s3-s2","")

echo("<br>")



echo("")

echo(Now())
echo("")

Response.Write "</pre>"

Response.Write "<br> "&Get_vPath(0)
Response.Write "<br> "&Get_fName()

'6F9619FF-8B86-D011-B42D-00C04FC964FF 
'GUID的128位可以分为16个字节，
'前8个字节是时间和版本号，
'中间2个字节是UUID变体和时钟序数，
'后6个字节标识地点。0086-0769
'99.99.99.99
Response.Write "<pre>"
Response.Write "<br>"&Session.SessionID()
'Response.Write "<br>"&Sess_ID(n)&" --- Sess_ID(n) "
Response.Write "<br>"&rs_AutID(conn,"InfoNews","KeyID","dtdef","1","")&" --- Def"
Response.Write "<br>"&rs_AutID(conn,"InfoNews","KeyID","dtdef","M","")&" --- M"
Response.Write "<br>"&rs_AutID(conn,"InfoNews","KeyID","dtdef","D","")&" --- D"
Response.Write "<br>------6时间---2版本-2TID-6标识地点--"
Response.Write "<br>12345678-1234-1234-1234-123456789012"
Response.Write "<br>123456789012345678901234567890123456"
Response.Write "<br>"&Get_GUID("2.1","PEACE0ASP0")&" --- myGUID ===xx "
Response.Write "<br>"&Get_GUID("2.1","PEACE0PHP0")&" --- my-PHP ===xx "
Response.Write "<br>"&Get_GUID("640.99","IP")&" --- myGUID "
Response.Write "<br>"&Get_GUID("1.1","IP")&" --- myGUID"
Response.Write "<br>"&Get_GUID("99.99","IP")&" --- myGUID"
Response.Write "<br>"&Request.ServerVariables("HTTP_USER_AGENT")
s = Get_yyyymmdd("")&Get_hhmmss()
Response.Write "<br> Get_AutoID(xx) -------- 123456789012345678901234--"
Response.Write "<br> Get_AutoID(12) -------- "&Get_AutoID(12)
Response.Write "<br> Get_AutoID(15) -------- "&Get_AutoID(15)
Response.Write "<br> Get_AutoID(18) -------- "&Get_AutoID(18)
Response.Write "<br> Get_AutoID(24) -------- "&Get_AutoID(24)
Response.Write "<br> Get_AutoID(xx) -------- 123456789012345678901234--"
Response.Write "<br> Get_AutoID(17) -------- "&Get_AutoID(17)
Response.Write "<br> Get_AutoID(20) -------- "&Get_AutoID(20)
Response.Write "<br> Get_AutoID(xx) -------- 123456789012345678901234--"
Response.Write "<br> Get_AutoID(13) -------- "&Get_AutoID(13)
Response.Write "<br> Get_AutoID(14) -------- "&Get_AutoID(14)
Response.Write "<br>99999:"&CStr(Hex(99999))
Response.Write "<br>99999999:"&CStr(Hex(99999999))
Response.Write "</pre>"
'Response.Write Hex(s)
'Response.End()

Response.Write "<pre>"
Response.Write "<br>123456789012345678901234"
Response.Write "<br>Pic-2010-5-UE810_KSRRU23G"
a = Year("2010-04-01 12:30:30") 'Now()
a = Get_FmtID("mdhnsx","") 
Response.Write "<br>Get_FmtID(mdhnsx):"&a

'Response.End()

T1 = Timer() : Response.Write "<br>Timer():"&T1

Response.Write "<br> rs_AutID ... ---------------"

a = rs_AutID(conn,"InfoPics","KeyCode","uf12","mm-dd","") 
Response.Write "<br>"&a
a = rs_AutID(conn,"InfoPics","KeyCode","uf12","md","2010-04-01 9:08:08") 
Response.Write "<br>"&a
a = rs_AutID(conn,"InfoPics","KeyCode","uf12","mm","2008-06-08 9:08:08") 
Response.Write "<br>"&a
a = rs_AutID(conn,"InfoPics","KeyCode","dat12","mm-dd","") 
Response.Write "<br>"&a
a = rs_AutID(conn,"InfoNews","KeyCode","dtInf","md","") 
Response.Write "<br>"&a&" OK "
a = rs_AutID(conn,"InfoPics","KeyCode","dtPic","mm","") 
Response.Write "<br>"&a&" OK "

T1 = Timer() - T1
Response.Write "<br>"&T1

Response.Write "</pre>"
Response.End()
%>

&nbsp; &nbsp;　
</body>
</html>
