<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="sadm/func1/func1.asp"-->
<!--#include file="sadm/func2/func2.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title><%=Config_Name%></title>
<link rel="shortcut icon" href="/favicon.ico" />
<link rel="Bookmark" href="/favicon.ico" />
</head>
<body>
<%
'' Server.Transfer "page/index.asp"
Response.Redirect "page/"
%>
</body>
</html>