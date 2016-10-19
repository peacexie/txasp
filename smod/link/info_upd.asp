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
Set rs=Server.Createobject("Adodb.Recordset")
%>
<!--#include file="../../sadm/func1/func1.asp"-->
<!--#include file="../../sadm/func2/func2.asp"-->
<!--#include file="../../sadm/func1/func_file.asp"-->
<!--#include file="config.asp"-->
<!--#include file="../../inc/home/func3.asp"-->

<%

Set rs2=Server.Createobject("Adodb.Recordset")
sql = " SELECT TypID,TypName FROM WebTyps WHERE TypMod='HomLnk1' ORDER BY TypID"
rs2.Open Sql,conn,1,1
  Do While NOT rs2.EOF
    TypID = rs2("TypID")
	TypName = rs2("TypName")
	sLink = ListLink(TypID,"")
	If TypID="Lnk0012" Then
	  Call File_Add2("../../upfile/sys/xadv/linkn_"&TypID&".htm",getNLink(sLink,18),"UTF-8")
	End If
	If TypID="Lnk0020" Then
	  Call File_Add2("../../upfile/sys/xadv/linkn_"&TypID&".htm",getNLink(sLink,6),"UTF-8")
	End If
	Call File_Add2("../../upfile/sys/xadv/links_"&TypID&".htm",sLink,"UTF-8")
	Response.Write vbcrlf&"<hr>"&TypID&" : "&TypName&" : 刷新完成：<br>"&sLink
  rs2.MoveNext()
  Loop
 
rs2.Close() 
Set rs2 = Nothing
Set rs = Nothing

' 得到 (图片)连接 TopN /////////////////////
Function getNLink(xStr,xTop)
  Dim s,a,n,f,i
  s = xStr :f="</a></div>"
  a = Split(s,f)
  s = ""
  for i=0 to xTop-1
    if i<=uBound(a) Then
	  if a(i)<>"" Then
	    s = s&vbcrlf&a(i)&f
	  end if
	End If
  next
getNLink = s
End Function

%>

</body>
</html>
