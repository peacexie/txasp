<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../himg/tconfig.asp"-->
<!--#include file="outfunc.asp"-->
<%
'// 处理 显示内容
yUrl = "http://shop.dg.gd.cn/"
yCSet = "UTF-8"
sHtml=OutPage(yUrl,yCSet)&""
sHtml=Replace(sHtml,"src=""js/","src=""http://shop.dg.gd.cn/js/")
sHtml=Replace(sHtml,"href=""themes/","href=""http://shop.dg.gd.cn/themes/")
sHtml=Replace(sHtml,"src=""themes/","src=""http://shop.dg.gd.cn/themes/")
'sHtml=Replace(sHtml,"src='data/","src='http://shop.dg.gd.cn/data/")
Response.Write sHtml 
%>
