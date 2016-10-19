<!--#include file="_config.asp"-->
<%
MD=RequestS("ModID",3,48) :If MD="" Then MD="InfJ124" 
MDName = "应聘职位" ': Response.Write MD
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title><%=MDName%>-<%=vPMsg_WName%></title>
<link href="../pfile/pimg/style.css" rel="stylesheet" type="text/css">
<script src="../inc/home/jsInfo.js" type="text/javascript"></script>
</head>
<body>
<!--#include file="../pfile/pinc/remjoin.asp"-->
</body>
</html>
