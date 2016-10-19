<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Response.Charset="gb2312"%>
<!--#include file="conv_func.asp"-->
<%
str0 = "Test测试字符串的互转"
str1 = Server.URLEncode(str0)
str2 = cUrl_Encode(str0,936,936)
str3 = cUrl_Encode(str0,950,936)
str4 = cUrl_Encode(str0,65001,936)
get0 = Request("str0") 
get1 = Request("str1") 
get2 = Request("str2") 
get3 = Request("str3") 
s5 = conv_Unicode(str0)
Response.Write s5
Response.Write Server.URLEncode(s5)
Response.Write "<br>1."&cUrl_Decode(str1)
Response.Write "<br>2."&cUrl_Decode(str2)
Response.Write "<br>3."&cUrl_Decode(str3)
Response.Write "<br>4."&cUrl_Decode(str4)
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=str0%></title>
<base target="_blank" />
<style type="text/css">
body, td, th {
  font-family:"Courier New", Courier, monospace;
}
</style>
</head>
<body>
<table width="480" border="1" align="center">
  <tr>
    <td colspan="3">Now: GB2312(<%=Session.Codepage%>), <%=str0%></td>
  </tr>
  <tr>
    <td nowrap="nowrap">Link</td>
    <td>Str</td>
    <td nowrap="nowrap">Get</td>
  </tr>
  <tr>
    <td nowrap="nowrap"><a href="ut.asp?str0=<%=str0%>">str0</a></td>
    <td><%=str0%></td>
    <td nowrap="nowrap"><%=get0%></td>
  </tr>
  <tr>
    <td nowrap="nowrap"><a href="ut.asp?str1=<%=str1%>">str1</a></td>
    <td><%=str1%></td>
    <td nowrap="nowrap"><%=get1%></td>
  </tr>
  <tr>
    <td nowrap="nowrap"><a href="ut.asp?str2=<%=str2%>">str2</a></td>
    <td><%=str2%></td>
    <td nowrap="nowrap"><%=get2%></td>
  </tr>
  <tr>
    <td nowrap="nowrap"><a href="ut.asp?str3=<%=str3%>">str3</a></td>
    <td><%=str3%></td>
    <td nowrap="nowrap"><%=get3%></td>
  </tr>
  <tr>
    <td nowrap="nowrap"><a href="ut.asp?str4=<%=str4%>">str4</a></td>
    <td><%=str4%></td>
    <td nowrap="nowrap"><%=get4%></td>
  </tr>
  <tr>
    <td colspan="3" nowrap="nowrap"><%=cUrl_ut2gb(str4)%></td>
  </tr>
</table>
</body>
</html>
